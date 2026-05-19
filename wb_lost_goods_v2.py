"""
wb_lost_goods_v2.py — Анализатор потерянных товаров на Wildberries
Версия 2.1 | Cardamón | Написан строго по документации dev.wildberries.ru

═══════════════════════════════════════════════════════════════════
ИСТОЧНИКИ ДАННЫХ (актуально на май 2026)
═══════════════════════════════════════════════════════════════════

1. ПОСТАВКИ
   GET statistics-api.wildberries.ru/api/v1/supplier/incomes
   Токен:    Statistics
   Поля:     nmId, quantity, status ("Принято"), supplierArticle,
             techSize, warehouseName, lastChangeDate
   Лимит:    100 000 строк / запрос, 1 запрос/мин
   Пагинация: lastChangeDate последней строки → следующий dateFrom

2. ИСТОРИЯ ПРОДАЖ ЗА ВСЁ ВРЕМЯ
   GET statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod
   Токен:    Statistics
   Поля:     nm_id, supplier_oper_name ("Продажа" / "Возврат…"),
             quantity, sa_name, brand_name, subject_name, rrd_id
   Лимит:    100 000 строк / запрос, 1 запрос/мин
   Пагинация: скользящие окна 90 дней + rrd_id внутри окна
   Данные с: 29 января 2024
   ⚠️  БУДЕТ ОТКЛЮЧЁН 15 ИЮЛЯ 2026

3. ОСТАТКИ НА СКЛАДАХ WB (текущий снэпшот)
   POST seller-analytics-api.wildberries.ru
        /api/analytics/v1/stocks-report/wb-warehouses
   Токен:    Analytics
   Поля:     nmId, quantity (физически на складе WB),
             inWayToClient, inWayFromClient, lastChangeDate
   Лимит:    60 000 строк / запрос, 1 запрос/мин
   1 строка = 1 размер артикула на 1 складе WB
   Данные:   текущий снэпшот, обновляется каждые 30 минут

═══════════════════════════════════════════════════════════════════
ФОРМУЛА РАСЧЁТА ПОТЕРЬ
═══════════════════════════════════════════════════════════════════
Нетто-продажи    = Продажи − Возвраты от покупателей
Остатки по учёту = Отгружено − Нетто-продажи
Потери           = Остатки по учёту − (Остатки факт + Едет к клиенту + Едет от клиента)

Важно: товары в пути к клиенту и обратно уже сняты со склада WB
(не входят в ост_факт), но ещё не завершили цикл — без учёта
этих полей возникают ложные "потери".

═══════════════════════════════════════════════════════════════════
УСТАНОВКА
═══════════════════════════════════════════════════════════════════
pip install requests openpyxl pandas

ЗАПУСК
python wb_lost_goods_v2.py
"""

import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import time
import sys

# ─────────────────────────────────────────────────────────────
# НАСТРОЙКИ — заполни перед запуском
# ─────────────────────────────────────────────────────────────

API_KEY     = ""               # ← вставь API-ключ (категории: Statistics + Analytics)
DATE_FROM   = "2019-01-01"    # начало истории для поставок
OUTPUT_FILE = "wb_lost_goods.xlsx"

# reportDetailByPeriod хранит данные только с 29 января 2024
SALES_DATE_FROM = "2024-01-29"

STATS_BASE     = "https://statistics-api.wildberries.ru"
ANALYTICS_BASE = "https://seller-analytics-api.wildberries.ru"
DELAY          = 65   # секунд между запросами (лимит WB: 1 запрос/мин)
DELAY_SHORT    = 5    # пауза при пустом ответе (продолжение пагинации)

# ─────────────────────────────────────────────────────────────
# УТИЛИТЫ
# ─────────────────────────────────────────────────────────────

def auth_headers(key: str) -> dict:
    return {"Authorization": key}

def safe_get(url: str, api_key: str, params: dict, label: str) -> list:
    """GET-запрос с обработкой rate limit и ошибок."""
    while True:
        try:
            r = requests.get(url, headers=auth_headers(api_key),
                             params=params, timeout=90)
        except requests.exceptions.RequestException as e:
            print(f"   ⚠️  Сетевая ошибка ({label}): {e}. Повтор через 30 сек…")
            time.sleep(30)
            continue

        if r.status_code == 200:
            return r.json() or []
        if r.status_code == 401:
            sys.exit(
                f"\n❌ HTTP 401 ({label})\n"
                f"Причина: неверный API-ключ или токен не имеет нужной категории.\n"
                f"Нужны категории: Statistics (поставки/продажи) и Analytics (остатки)."
            )
        if r.status_code == 429:
            print(f"   ⏳ Rate limit ({label}), жду 70 сек...")
            time.sleep(70)
            continue
        print(f"   ⚠️  HTTP {r.status_code} ({label}): {r.text[:300]}")
        return []


def safe_post(url: str, api_key: str, body: dict, label: str) -> list:
    """POST-запрос с обработкой rate limit и ошибок."""
    while True:
        try:
            r = requests.post(url, headers={**auth_headers(api_key),
                                            "Content-Type": "application/json"},
                              json=body, timeout=90)
        except requests.exceptions.RequestException as e:
            print(f"   ⚠️  Сетевая ошибка ({label}): {e}. Повтор через 30 сек…")
            time.sleep(30)
            continue

        if r.status_code == 200:
            resp = r.json()
            # Некоторые эндпоинты WB возвращают {"data": [...]} вместо плоского списка
            if isinstance(resp, dict):
                for key in ("data", "result", "items", "list", "response"):
                    val = resp.get(key)
                    if isinstance(val, list):
                        return val or []
                    # Один уровень вложенности: {"data": {"stocks": [...]}}
                    if isinstance(val, dict):
                        for inner in ("stocks", "items", "list", "data", "result"):
                            inner_val = val.get(inner)
                            if isinstance(inner_val, list):
                                return inner_val or []
                # Не нашли список — печатаем структуру для диагностики
                import json as _json
                print(f"   ℹ️  Неизвестная структура ответа ({label}): "
                      f"{_json.dumps(resp, ensure_ascii=False)[:500]}")
                return []
            return resp or []
        if r.status_code == 401:
            sys.exit(
                f"\n❌ HTTP 401 ({label})\n"
                f"Причина: неверный API-ключ или токен не имеет категории Analytics.\n"
                f"Нужны категории: Statistics (поставки/продажи) и Analytics (остатки)."
            )
        if r.status_code == 429:
            print(f"   ⏳ Rate limit ({label}), жду 70 сек...")
            time.sleep(70)
            continue
        print(f"   ⚠️  HTTP {r.status_code} ({label}): {r.text[:300]}")
        return []

# ─────────────────────────────────────────────────────────────
# 1. ПОСТАВКИ
# ─────────────────────────────────────────────────────────────

def fetch_incomes(api_key: str) -> pd.DataFrame:
    """
    Поставки — пробуем несколько известных URL (WB менял домены в 2024-2025).
    Принимаем только status == "Принято" (регистр нечувствительно).
    Пагинация: lastChangeDate последней строки как следующий dateFrom.
    """
    candidate_urls = [
        f"{STATS_BASE}/api/v1/supplier/incomes",
        "https://marketplace-api.wildberries.ru/api/v1/supplier/incomes",
        "https://suppliers-api.wildberries.ru/api/v1/supplier/incomes",
    ]

    # Находим рабочий URL
    url = None
    for candidate in candidate_urls:
        r = safe_get(candidate, api_key, {"dateFrom": DATE_FROM}, "Поставки-проверка")
        if r is not None and r != []:
            url = candidate
            print(f"   ✅ Рабочий URL поставок: {candidate}")
            break
        # safe_get вернул [] — может быть пустой ответ (ОК) или 404
        # Проверяем напрямую чтобы отличить пустой список от ошибки
        import requests as _req
        try:
            probe = _req.get(candidate, headers=auth_headers(api_key),
                             params={"dateFrom": DATE_FROM}, timeout=30)
            if probe.status_code == 200:
                url = candidate
                print(f"   ✅ Рабочий URL поставок: {candidate}")
                break
            print(f"   ✗ {candidate} → HTTP {probe.status_code}")
        except Exception:
            print(f"   ✗ {candidate} → недоступен")

    if not url:
        print("   ❌ Ни один URL поставок не работает — колонка 'Отгружено' будет пустой.")
        return pd.DataFrame(columns=["nmId", "отгружено", "арт_продавца", "тех_размер"])

    all_rows = []
    date_cursor = DATE_FROM
    page = 0

    print("\n📦 Поставки — загружаем данные")

    while True:
        page += 1
        data = safe_get(url, api_key, {"dateFrom": date_cursor}, "Поставки")

        if not data:
            break

        all_rows.extend(data)
        last = data[-1]
        last_date = last.get("lastChangeDate", "")
        print(f"   стр.{page}: +{len(data)} строк | lastChangeDate={last_date[:19]}")

        if len(data) < 100_000:
            break

        date_cursor = last_date
        time.sleep(DELAY)

    print(f"   Итого строк поставок: {len(all_rows)}")

    if not all_rows:
        return pd.DataFrame(columns=["nmId", "отгружено", "арт_продавца", "тех_размер"])

    df = pd.DataFrame(all_rows)
    df["nmId"] = df["nmId"].astype(str)

    if "status" in df.columns:
        df = df[df["status"].str.strip().str.lower() == "принято"].copy()

    grp = df.groupby("nmId", as_index=False).agg(
        отгружено=("quantity", "sum"),
        арт_продавца=("supplierArticle", "first"),
        тех_размер=("techSize", "first"),
    )
    print(f"   ✅ Уникальных артикулов в поставках: {len(grp)}")
    return grp

# ─────────────────────────────────────────────────────────────
# 2. ИСТОРИЯ ПРОДАЖ И ВОЗВРАТОВ
# ─────────────────────────────────────────────────────────────

def fetch_sales_history(api_key: str) -> pd.DataFrame:
    """
    /api/v5/supplier/reportDetailByPeriod

    Тип операции определяется полем supplier_oper_name:
      - содержит "продажа"  → продажа покупателю
      - содержит "возврат"  → возврат от покупателя
      - логистика, штрафы и пр. — не учитываем

    quantity может быть положительным или отрицательным в зависимости
    от типа операции — берём abs() и считаем по типу.

    Пагинация:
      - Разбиваем всё время на окна по 90 дней начиная с SALES_DATE_FROM
      - Внутри каждого окна итерируем rrdid (начинаем с 0)
      - Обязательная дедупликация по rrd_id (окна могут перекрываться)

    ⚠️  Работает до 15 июля 2026!
    """
    url = f"{STATS_BASE}/api/v5/supplier/reportDetailByPeriod"
    all_rows = []
    seen_rrd = set()

    # Строим окна по 90 дней от SALES_DATE_FROM до сегодня.
    # WB хранит данные этого отчёта только с 29 января 2024 —
    # более ранняя дата даст только пустые ответы и лишние паузы.
    start_dt  = datetime.strptime(SALES_DATE_FROM, "%Y-%m-%d")
    today_dt  = datetime.now()
    windows   = []
    cur = start_dt
    while cur < today_dt:
        end = min(cur + timedelta(days=89), today_dt)
        windows.append((cur.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")))
        cur = end + timedelta(days=1)

    total_windows = len(windows)
    print(f"\n📊 История продаж — /api/v5/supplier/reportDetailByPeriod")
    print(f"   {total_windows} окон по ~90 дней от {SALES_DATE_FROM} до сегодня")

    for wi, (w_from, w_to) in enumerate(windows, 1):
        rrdid = 0
        window_new = 0
        inner_page = 0

        while True:
            inner_page += 1
            params = {
                "dateFrom": w_from,
                "dateTo":   w_to,
                "limit":    100_000,
                "rrdid":    rrdid,
            }
            data = safe_get(url, api_key, params, f"Продажи окно {wi}")

            if not data:
                break

            new_rows = [row for row in data if row.get("rrd_id") not in seen_rrd]
            for row in new_rows:
                seen_rrd.add(row["rrd_id"])
            all_rows.extend(new_rows)
            window_new += len(new_rows)

            rrdid = max(row["rrd_id"] for row in data)

            if len(data) < 100_000:
                break
            time.sleep(DELAY)

        pct = wi / total_windows * 100
        print(f"   [{pct:4.0f}%] окно {wi}/{total_windows} ({w_from}→{w_to}): "
              f"+{window_new} строк | всего: {len(all_rows)}")

        if wi < total_windows:
            time.sleep(DELAY)

    print(f"   Итого уникальных строк в отчёте реализации: {len(all_rows)}")

    if not all_rows:
        return pd.DataFrame(columns=["nmId", "продано", "возвраты", "бренд", "предмет"])

    df = pd.DataFrame(all_rows)
    df["nmId"] = df["nm_id"].astype(str)

    op = df["supplier_oper_name"].str.strip().str.lower().fillna("")
    is_sale   = op.str.contains("продажа")
    is_return = op.str.contains("возврат")

    sales_df   = df[is_sale].copy()
    returns_df = df[is_return].copy()

    s = sales_df.groupby("nmId", as_index=False).agg(
        продано=("quantity", lambda x: int(x.abs().sum())),
        бренд=("brand_name", "first"),
        предмет=("subject_name", "first"),
    )
    rt = returns_df.groupby("nmId", as_index=False).agg(
        возвраты=("quantity", lambda x: int(x.abs().sum())),
    )

    result = s.merge(rt, on="nmId", how="outer").fillna(0)
    result["продано"]  = result["продано"].astype(int)
    result["возвраты"] = result["возвраты"].astype(int)

    print(f"   ✅ Артикулов с продажами: {(result['продано'] > 0).sum()}, "
          f"с возвратами: {(result['возвраты'] > 0).sum()}")
    return result

# ─────────────────────────────────────────────────────────────
# 3. ОСТАТКИ НА СКЛАДАХ WB
# ─────────────────────────────────────────────────────────────

def fetch_stocks(api_key: str) -> pd.DataFrame:
    """
    POST /api/analytics/v1/stocks-report/wb-warehouses
    Токен: Analytics

    1 строка ответа = 1 размер артикула на 1 складе WB.
    Для анализа потерь суммируем quantity по nmId по всем складам.

    Пагинация: lastChangeDate последней строки → следующий dateFrom.
    Лимит: 60 000 строк / запрос.

    Поля inWayToClient и inWayFromClient используются в формуле потерь:
    товары в пути не входят в quantity на складе, но ещё не прошли
    финальную операцию продажи/возврата — без них расчёт даёт
    ложные потери на артикулы с активными доставками.
    """
    url = f"{ANALYTICS_BASE}/api/analytics/v1/stocks-report/wb-warehouses"
    all_rows = []
    date_cursor = DATE_FROM
    page = 0

    print("\n🏭 Остатки на складах — POST /api/analytics/v1/stocks-report/wb-warehouses")

    while True:
        page += 1
        data = safe_post(url, api_key, {"dateFrom": date_cursor}, "Остатки")

        if not data:
            break

        all_rows.extend(data)
        last_date = data[-1].get("lastChangeDate", "")
        print(f"   стр.{page}: +{len(data)} строк | lastChangeDate={last_date[:19]}")

        if len(data) < 60_000:
            break

        date_cursor = last_date
        time.sleep(DELAY)

    print(f"   Итого строк остатков: {len(all_rows)}")

    if not all_rows:
        return pd.DataFrame(columns=["nmId", "ост_факт", "ост_к_клиенту", "ост_от_клиента"])

    df = pd.DataFrame(all_rows)
    df["nmId"] = df["nmId"].astype(str)

    # Суммируем по nmId через все склады (1 строка = 1 артикул на 1 складе)
    agg = {"quantity": "sum"}
    if "inWayToClient" in df.columns:
        agg["inWayToClient"] = "sum"
    if "inWayFromClient" in df.columns:
        agg["inWayFromClient"] = "sum"

    result = df.groupby("nmId", as_index=False).agg(agg).rename(columns={
        "quantity":        "ост_факт",
        "inWayToClient":   "ост_к_клиенту",
        "inWayFromClient": "ост_от_клиента",
    })

    # Если API не вернул поля в пути — подставляем нули
    for col in ("ост_к_клиенту", "ост_от_клиента"):
        if col not in result.columns:
            result[col] = 0

    print(f"   ✅ Артикулов с остатками: {len(result)}")
    return result

# ─────────────────────────────────────────────────────────────
# СБОРКА ИТОГОВОГО ДАТАФРЕЙМА
# ─────────────────────────────────────────────────────────────

def build_report(api_key: str) -> pd.DataFrame:
    incomes = fetch_incomes(api_key)
    print(f"\n⏳ Пауза {DELAY} сек перед следующим запросом…")
    time.sleep(DELAY)

    sales = fetch_sales_history(api_key)
    print(f"\n⏳ Пауза {DELAY} сек перед следующим запросом...")
    time.sleep(DELAY)

    stocks = fetch_stocks(api_key)

    df = (
        incomes
        .merge(sales,   on="nmId", how="outer")
        .merge(stocks,  on="nmId", how="outer")
        .fillna(0)
    )

    for col in ["отгружено", "продано", "возвраты", "ост_факт",
                "ост_к_клиенту", "ост_от_клиента"]:
        if col in df.columns:
            df[col] = df[col].astype(int)

    df["продано_нетто"]    = df["продано"] - df["возвраты"]
    df["остатки_по_учёту"] = df["отгружено"] - df["продано_нетто"]

    # Товары в пути к клиентам и обратно физически не на складе WB,
    # но в учёте ещё не завершили цикл — вычитаем их, чтобы не было
    # ложных потерь на артикулы с активными доставками.
    df["потери"] = (
        df["остатки_по_учёту"]
        - df["ост_факт"]
        - df["ост_к_клиенту"]
        - df["ост_от_клиента"]
    )

    df["потери_%"] = df.apply(
        lambda r: round(r["потери"] / r["отгружено"] * 100, 1)
        if r["отгружено"] > 0 else 0.0,
        axis=1
    )

    def make_status(row):
        if row["потери"] > 0:  return "⚠️ Потери"
        if row["потери"] < 0:  return "❓ Пересорт"
        return "✅ Ок"

    df["статус"] = df.apply(make_status, axis=1)

    return df.sort_values("потери", ascending=False).reset_index(drop=True)

# ─────────────────────────────────────────────────────────────
# EXCEL — СОХРАНЕНИЕ
# ─────────────────────────────────────────────────────────────

COLUMNS = [
    ("nmId",              "Артикул WB"),
    ("арт_продавца",      "Арт. продавца"),
    ("тех_размер",        "Размер"),
    ("бренд",             "Бренд"),
    ("предмет",           "Предмет"),
    ("отгружено",         "Отгружено"),
    ("продано",           "Продано (брутто)"),
    ("возвраты",          "Возвраты"),
    ("продано_нетто",     "Продано нетто"),
    ("остатки_по_учёту",  "Остаток по учёту"),
    ("ост_факт",          "Остаток факт (WB)"),
    ("ост_к_клиенту",     "Едет к клиенту"),
    ("ост_от_клиента",    "Едет от клиента"),
    ("потери",            "Потери (шт)"),
    ("потери_%",          "Потери (%)"),
    ("статус",            "Статус"),
]

COL_WIDTHS = {
    "nmId": 16, "арт_продавца": 20, "тех_размер": 10, "бренд": 18,
    "предмет": 20, "отгружено": 14, "продано": 16, "возвраты": 14,
    "продано_нетто": 16, "остатки_по_учёту": 18, "ост_факт": 18,
    "ост_к_клиенту": 16, "ост_от_клиента": 16, "потери": 14,
    "потери_%": 13, "статус": 18,
}

def _thin_border():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)

def _write_sheet(ws, df: pd.DataFrame, title_text: str,
                 title_bg: str, header_bg: str, row_bg_even: str,
                 loss_row_bg: str, start_row: int = 1):
    """Универсальная запись листа."""
    border = _thin_border()
    existing_cols = [(k, v) for k, v in COLUMNS if k in df.columns]
    n_cols = len(existing_cols)

    # Заголовок
    last_col_letter = get_column_letter(n_cols)
    ws.merge_cells(f"A{start_row}:{last_col_letter}{start_row}")
    tc = ws.cell(row=start_row, column=1, value=title_text)
    tc.font      = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    tc.fill      = PatternFill("solid", start_color=title_bg)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[start_row].height = 30

    # Подзаголовок-формула
    formula_row = start_row + 1
    ws.merge_cells(f"A{formula_row}:{last_col_letter}{formula_row}")
    fc = ws.cell(row=formula_row, column=1,
                 value="Потери = Отгружено − (Продано − Возвраты) − Остаток факт − Едет к клиенту − Едет от клиента")
    fc.font      = Font(name="Arial", size=9, italic=True, color="777777")
    fc.fill      = PatternFill("solid", start_color="F8F8F8")
    fc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[formula_row].height = 16

    # Заголовки колонок
    hdr_row = start_row + 2
    for ci, (_, label) in enumerate(existing_cols, 1):
        c = ws.cell(row=hdr_row, column=ci, value=label)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill      = PatternFill("solid", start_color=header_bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = border
    ws.row_dimensions[hdr_row].height = 38

    # Данные
    for ri, (_, row) in enumerate(df.iterrows(), hdr_row + 1):
        is_loss = row.get("потери", 0) > 0
        is_err  = row.get("потери", 0) < 0
        fill_clr = (loss_row_bg  if is_loss
                    else ("FFF8DC" if is_err
                          else (row_bg_even if ri % 2 == 0 else "FFFFFF")))

        for ci, (col_key, _) in enumerate(existing_cols, 1):
            val = row.get(col_key, "")
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = border
            c.fill   = PatternFill("solid", start_color=fill_clr)

            if col_key == "потери" and is_loss:
                c.font = Font(name="Arial", size=10, bold=True, color="C00000")
            elif col_key == "статус":
                c.font = Font(name="Arial", size=10, bold=True,
                              color=("C00000" if is_loss else
                                     ("806000" if is_err else "1E6B1E")))
            else:
                c.font = Font(name="Arial", size=10)

            text_cols = {"nmId", "арт_продавца", "тех_размер", "бренд",
                         "предмет", "статус"}
            if col_key in text_cols:
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif col_key == "потери_%":
                c.alignment    = Alignment(horizontal="right", vertical="center")
                c.number_format = '0.0"%"'
            else:
                c.alignment    = Alignment(horizontal="right", vertical="center")
                c.number_format = '#,##0'

    # Ширина колонок
    for ci, (col_key, _) in enumerate(existing_cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(col_key, 14)

    # Закрепление и фильтр
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
    last_data_row = hdr_row + len(df)
    ws.auto_filter.ref = f"A{hdr_row}:{last_col_letter}{last_data_row}"


def save_excel(df: pd.DataFrame, path: str):
    wb = openpyxl.Workbook()

    # ── Лист 1: Все артикулы ──────────────────────────────────
    ws1 = wb.active
    ws1.title = "Все артикулы"
    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    _write_sheet(
        ws1, df,
        title_text  = f"Анализ потерь — Cardamón @ Wildberries | {ts}",
        title_bg    = "1A237E",
        header_bg   = "283593",
        row_bg_even = "F5F5F5",
        loss_row_bg = "FFEBEE",
    )

    # ── Лист 2: Только потери ────────────────────────────────
    losses = df[df["потери"] > 0].copy()
    ws2 = wb.create_sheet("Только потери")
    _write_sheet(
        ws2, losses,
        title_text  = f"Артикулы с потерями: {len(losses)} шт | "
                      f"Итого потеряно: {int(losses['потери'].sum()):,} ед.",
        title_bg    = "B71C1C",
        header_bg   = "C62828",
        row_bg_even = "FFF5F5",
        loss_row_bg = "FFEBEE",
    )

    # ── Лист 3: Сводка ───────────────────────────────────────
    ws3 = wb.create_sheet("Сводка")
    ws3.column_dimensions["A"].width = 38
    ws3.column_dimensions["B"].width = 20

    total_losses_qty = int(df[df["потери"] > 0]["потери"].sum())
    total_losses_cnt = int((df["потери"] > 0).sum())

    summary = [
        ("ИТОГИ ПО ВСЕМ АРТИКУЛАМ",           None,   True,  "1A237E", "FFFFFF"),
        ("Уникальных артикулов",               len(df),            False, None, None),
        ("",                                   None,   False, None, None),
        ("ДВИЖЕНИЕ ТОВАРА (шт)",               None,   True,  "283593", "FFFFFF"),
        ("Всего отгружено на WB",              int(df["отгружено"].sum()),      False, None, None),
        ("Всего продано (брутто)",             int(df["продано"].sum()),        False, None, None),
        ("Всего возвратов от покупателей",     int(df["возвраты"].sum()),       False, None, None),
        ("Продано нетто",                      int(df["продано_нетто"].sum()),  False, None, None),
        ("",                                   None,   False, None, None),
        ("ОСТАТКИ",                            None,   True,  "283593", "FFFFFF"),
        ("Остатки по учёту (должно быть)",     int(df["остатки_по_учёту"].sum()), False, None, None),
        ("Остатки факт на WB",                 int(df["ост_факт"].sum()),       False, None, None),
        ("Едет к клиентам (в пути)",           int(df["ост_к_клиенту"].sum()), False, None, None),
        ("Едет обратно от клиентов",           int(df["ост_от_клиента"].sum()), False, None, None),
        ("",                                   None,   False, None, None),
        ("⚠️  ПОТЕРИ",                         None,   True,  "B71C1C", "FFFFFF"),
        ("Артикулов с потерями",               total_losses_cnt,   False, None, None),
        ("Потеряно единиц товара",             total_losses_qty,   False, "FFEBEE", "C00000"),
        ("Артикулов без отклонений (ОК)",      int((df["потери"] == 0).sum()), False, None, None),
        ("Артикулов с пересортом/ошибкой",     int((df["потери"] < 0).sum()), False, None, None),
        ("",                                   None,   False, None, None),
        ("Дата формирования",                  datetime.now().strftime("%d.%m.%Y %H:%M"), False, None, None),
        ("⚠️  reportDetailByPeriod отключ.",  "15 июля 2026",  False, "FFF8DC", "806000"),
    ]

    for ri, (label, val, is_header, bg, fg) in enumerate(summary, 1):
        lc = ws3.cell(row=ri, column=1, value=label)
        vc = ws3.cell(row=ri, column=2, value=val)

        for c in (lc, vc):
            c.font = Font(
                name="Arial", size=11,
                bold=is_header or (label.startswith("⚠️") and "потери" in label.lower()),
                color=(fg or ("C00000" if "потери" in label.lower() and val and isinstance(val, int) and val > 0 else "000000"))
            )
            if bg:
                c.fill = PatternFill("solid", start_color=bg)
            c.alignment = Alignment(
                horizontal=("left" if c.column == 1 else "right"),
                vertical="center"
            )
            if isinstance(val, int) and c.column == 2:
                c.number_format = '#,##0'
        ws3.row_dimensions[ri].height = 22

    wb.save(path)
    print(f"\n✅ Excel сохранён: {path}")

# ─────────────────────────────────────────────────────────────
# ТОЧКА ВХОДА
# ─────────────────────────────────────────────────────────────

def main():
    global API_KEY

    print("═" * 62)
    print("  WB Анализатор потерь v2.1 — Cardamón")
    print("  Написан строго по документации dev.wildberries.ru")
    print("═" * 62)

    if not API_KEY:
        API_KEY = input("\n🔑 API-ключ (категории Statistics + Analytics): ").strip()
    if not API_KEY:
        sys.exit("❌ Ключ не введён.")

    print(f"\n📅 Поставки с:     {DATE_FROM}")
    print(f"📅 Продажи с:      {SALES_DATE_FROM} (WB хранит данные с этой даты)")
    print(f"⏱️  Rate limit WB: 1 запрос/мин → паузы {DELAY} сек между запросами")
    print(f"⏱️  Ориентировочное время: 10–30 мин")

    df = build_report(API_KEY)

    losses   = df[df["потери"] > 0]
    пересорт = df[df["потери"] < 0]
    ok_cnt   = (df["потери"] == 0).sum()

    print("\n" + "═" * 62)
    print("  ИТОГИ")
    print("═" * 62)
    print(f"  Артикулов всего:              {len(df):>8,}")
    print(f"  Отгружено (шт):               {df['отгружено'].sum():>8,}")
    print(f"  Продано нетто (шт):           {df['продано_нетто'].sum():>8,}")
    print(f"  Остатки по учёту (шт):        {df['остатки_по_учёту'].sum():>8,}")
    print(f"  Остатки факт на WB (шт):      {df['ост_факт'].sum():>8,}")
    print(f"  Едет к клиентам (шт):         {df['ост_к_клиенту'].sum():>8,}")
    print(f"  Едет от клиентов (шт):        {df['ост_от_клиента'].sum():>8,}")
    print(f"")
    print(f"  ⚠️  Потери:     {losses['потери'].sum():>6,} шт ({len(losses)} артикулов)")
    print(f"  ❓  Пересорт:  {abs(пересорт['потери'].sum()):>6,} шт ({len(пересорт)} артикулов)")
    print(f"  ✅  Ок:                     {ok_cnt:>6} артикулов")
    print("═" * 62)

    save_excel(df, OUTPUT_FILE)

    print("\n📊 Листы в файле:")
    print("   1. Все артикулы   — полная таблица, сортировка по потерям")
    print("   2. Только потери  — только проблемные артикулы")
    print("   3. Сводка         — суммарные показатели")


if __name__ == "__main__":
    main()
