"""
WB Reviews Bot — главный файл приложения
FastAPI сервер + планировщик фоновых задач
"""

import asyncio
import json
import logging
import os
from datetime import datetime
from pathlib import Path

import httpx
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
log = logging.getLogger(__name__)

WB_API_KEY     = os.environ["WB_API_KEY"]
ANTHROPIC_KEY  = os.environ["ANTHROPIC_KEY"]
CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL_MINUTES", "30"))

DATA_FILE = Path("/data/reviews.json")
DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
if not DATA_FILE.exists():
    DATA_FILE.write_text(json.dumps({
        "processed": [], "pending": [], "published": [],
        "templates": [], "template_index": 0,
    }))

CARDAMON_HITS = [
    {"art": "943473684", "name": "костюмчики"},
    {"art": "966951684", "name": "топы"},
    {"art": "340791045", "name": "футболки-поло оверсайз"},
    {"art": "428900845", "name": "футболки по фигуре"},
]

def build_hits_block(current_article: str) -> str:
    hits = [h for h in CARDAMON_HITS if h["art"] != current_article][:3]
    if not hits:
        return ""
    parts = ", ".join(f"{h['name']} арт. {h['art']}" for h in hits)
    return f"Также в нашей новой коллекции: {parts}."

def build_claude_prompt(review_text: str, rating: int, article: str, product_name: str, has_photo: bool = False) -> str:
    hits_block = build_hits_block(article)

    tone_instruction = {
        5: "Ответ тёплый, женственный и с достоинством — как письмо от любимого бренда. Обращайся на «вы». Отреагируй на конкретную деталь из отзыва (посадка, цвет, ткань, ощущение). Пиши каждый раз по-разному — живо, но без разговорного стиля. Никаких восклицаний через слово.",
        4: "Тёплый и достойный тон, обращение на «вы». Поблагодари с изяществом, кратко отреагируй на суть. Если есть замечание — признай спокойно и с уважением, без оправданий.",
        3: "Спокойный и уважительный тон, обращение на «вы». Мягко отреагируй на то, что не устроило. Покажи, что слышите — без заискивания.",
        2: "Сдержанный и участливый тон, обращение на «вы». Признай проблему конкретно, без лишних слов. Предложи написать в личные сообщения для решения.",
        1: "Серьёзный и уважительный тон, обращение на «вы». Признай проблему, не оправдывайся. Предложи написать в личные сообщения.",
    }.get(rating, "Ответь вежливо, с достоинством, на «вы».")

    photo_instruction = ""
    if has_photo and rating >= 4:
        photo_instruction = "- Покупательница приложила фото. Впиши изящный комплимент — что вещь на ней смотрится безупречно, образ очень элегантный, именно то настроение коллекции. Утончённо, не по-разговорному, каждый раз по-разному.\n"

    if rating >= 4:
        ending = f"{hits_block} Подпишитесь на бренд Cardamón — так вы первой узнаете о новинках и специальных ценах 🤍".strip()
        ending_instruction = f"2. В конце ответа обязательно добавь дословно: «{ending}»"
    else:
        ending_instruction = "2. Сосредоточься на решении проблемы. Никакой рекламы — сейчас неуместно."

    review_display = review_text if review_text.strip() else "Покупатель не оставил текст, только оценку."

    return f"""Ты — менеджер бренда женской одежды Cardamón на Wildberries. Пишешь живые ответы на отзывы.

Товар: {product_name or 'не указан'} (артикул {article or 'неизвестен'})
Оценка: {rating}/5
Отзыв: «{review_display}»

ЗАДАЧА:
1. {tone_instruction}
{ending_instruction}

ПРАВИЛА:
{photo_instruction}- Запрещено: «Ваш отзыв важен», «будем рады», «спасибо за доверие», «рады слышать», «всегда рады»
- Запрещено: разговорные обороты, уменьшительные слова в тексте ответа, избыточные эмодзи
- Не копируй слова из отзыва дословно
- Обращение только на «вы» / «вам» / «ваш»
- Тон — тепло, но с достоинством, уровень премиум
- Весь ответ включая концовку — не более 500 символов
- Только текст ответа, без кавычек"""

def load_data() -> dict:
    try:
        d = json.loads(DATA_FILE.read_text())
        d.setdefault("templates", [])
        d.setdefault("template_index", 0)
        return d
    except Exception:
        return {"processed": [], "pending": [], "published": [], "templates": [], "template_index": 0}

def save_data(data: dict):
    DATA_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2))

def wb_headers():
    return {"Authorization": WB_API_KEY, "Content-Type": "application/json"}

def get_next_template(data: dict, product_name: str) -> str:
    """Возвращает следующий шаблон по ротации с подстановкой {product}."""
    templates = data.get("templates", [])
    if not templates:
        return ""
    idx = data.get("template_index", 0) % len(templates)
    text = templates[idx]
    data["template_index"] = (idx + 1) % len(templates)
    return text.replace("{product}", product_name or "")

async def wb_get_unanswered() -> list[dict]:
    url = "https://feedbacks-api.wildberries.ru/api/v1/feedbacks"
    all_feedbacks = []
    skip = 0
    take = 100
    async with httpx.AsyncClient(timeout=20) as client:
        while True:
            params = {"isAnswered": "false", "take": take, "skip": skip}
            r = await client.get(url, headers=wb_headers(), params=params)
            r.raise_for_status()
            feedbacks = r.json().get("data", {}).get("feedbacks") or []
            all_feedbacks.extend(feedbacks)
            log.info(f"Загружено {len(all_feedbacks)} отзывов...")
            if len(feedbacks) < take:
                break
            skip += take
    return all_feedbacks

async def wb_post_answer(feedback_id: str, text: str) -> bool:
    url = "https://feedbacks-api.wildberries.ru/api/v1/feedbacks"
    async with httpx.AsyncClient(timeout=15) as client:
        r = await client.patch(url, headers=wb_headers(), json={"id": feedback_id, "text": text})
        if r.status_code != 200:
            log.warning(f"WB ответил {r.status_code}: {r.text}")
        return r.status_code == 200

async def generate_reply(review_text: str, rating: int, article: str = "", product_name: str = "", has_photo: bool = False) -> str:
    prompt = build_claude_prompt(review_text, rating, article, product_name, has_photo)
    async with httpx.AsyncClient(timeout=40) as client:
        r = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"},
            json={"model": "claude-sonnet-4-20250514", "max_tokens": 400, "messages": [{"role": "user", "content": prompt}]},
        )
        r.raise_for_status()
        return r.json()["content"][0]["text"].strip()

async def process_new_reviews():
    log.info("🔍 Проверяю новые отзывы WB...")
    data = load_data()
    processed_ids = set(data["processed"])

    try:
        feedbacks = await wb_get_unanswered()
    except Exception as e:
        log.error(f"Ошибка WB API: {e}")
        return

    new_reviews = [fb for fb in feedbacks if fb.get("id") and fb["id"] not in processed_ids]
    log.info(f"Новых отзывов: {len(new_reviews)}")

    for fb in new_reviews:
        fid     = fb["id"]
        rating  = fb.get("productValuation", 0)
        pros    = (fb.get("pros") or "").strip()
        cons    = (fb.get("cons") or "").strip()
        comment = (fb.get("text") or "").strip()

        parts = []
        if pros:    parts.append(f"Достоинства: {pros}")
        if cons:    parts.append(f"Недостатки: {cons}")
        if comment: parts.append(f"Комментарий: {comment}")
        text = " | ".join(parts) if parts else ""

        product = (fb.get("productDetails") or {}).get("productName", "")
        article = str((fb.get("productDetails") or {}).get("nmId", ""))
        created = fb.get("createdDate", "")

        review_obj = {
            "id": fid, "rating": rating, "text": text,
            "product": product, "article": article, "created": created,
            "processed_at": datetime.now().isoformat(),
        }

        if rating == 5:
            reply = get_next_template(data, product)
            if not reply:
                # Шаблоны не настроены — генерируем через Claude
                try:
                    reply = await generate_reply(text, rating, article, product)
                except Exception as e:
                    log.error(f"Ошибка Claude для {fid}: {e}")
                    reply = ""

            if reply:
                ok = await wb_post_answer(fid, reply)
                if ok:
                    log.info(f"✅ Авто-ответ на 5⭐ {fid}")
                    data["published"].append({**review_obj, "reply": reply, "mode": "auto"})
                else:
                    log.warning(f"⚠️ WB не принял ответ {fid} — в pending")
                    data["pending"].append({**review_obj, "draft": reply})
            else:
                log.warning(f"⚠️ Нет шаблонов и Claude недоступен для {fid} — в pending")
                data["pending"].append({**review_obj, "draft": ""})
        else:
            try:
                draft = await generate_reply(text, rating, article, product)
            except Exception as e:
                log.error(f"Ошибка Claude для {fid}: {e}")
                draft = ""
            data["pending"].append({**review_obj, "draft": draft})
            log.info(f"📝 Черновик {rating}⭐ {fid}")

        data["processed"].append(fid)
        save_data(data)

    log.info(f"✅ Готово. Обработано новых: {len(new_reviews)}")

app = FastAPI(title="WB Reviews Bot")
app.mount("/static", StaticFiles(directory="static"), name="static")
scheduler = AsyncIOScheduler()

@app.on_event("startup")
async def startup():
    scheduler.add_job(process_new_reviews, "interval", minutes=CHECK_INTERVAL, id="check_reviews")
    scheduler.start()
    log.info(f"🚀 Запущен. Проверка каждые {CHECK_INTERVAL} мин.")
    asyncio.create_task(process_new_reviews())

@app.on_event("shutdown")
async def shutdown():
    scheduler.shutdown()

@app.get("/", response_class=HTMLResponse)
async def index():
    return Path("static/index.html").read_text()

@app.get("/api/pending")
async def get_pending():
    return JSONResponse(load_data()["pending"])

@app.get("/api/published")
async def get_published():
    return JSONResponse(load_data()["published"][-50:])

@app.get("/api/stats")
async def get_stats():
    data = load_data()
    pub = data["published"]
    job = scheduler.get_job("check_reviews")
    return {
        "pending": len(data["pending"]),
        "published": len(pub),
        "auto": sum(1 for p in pub if p.get("mode") == "auto"),
        "manual": sum(1 for p in pub if p.get("mode") == "manual"),
        "next_check": job.next_run_time.isoformat() if job and job.next_run_time else None,
    }

@app.get("/api/templates")
async def get_templates():
    data = load_data()
    return {"templates": data.get("templates", [])}

class TemplatesPayload(BaseModel):
    templates: list[str]

@app.post("/api/templates")
async def save_templates_endpoint(payload: TemplatesPayload):
    cleaned = [t.strip() for t in payload.templates if t.strip()]
    if not cleaned:
        raise HTTPException(400, "Список шаблонов не может быть пустым")
    data = load_data()
    data["templates"] = cleaned
    save_data(data)
    return {"ok": True, "count": len(cleaned)}

class PublishPayload(BaseModel):
    review_id: str
    text: str

@app.post("/api/publish")
async def publish_answer(payload: PublishPayload):
    if len(payload.text) > 500:
        raise HTTPException(400, f"Текст превышает 500 символов ({len(payload.text)})")
    data = load_data()
    review = next((r for r in data["pending"] if r["id"] == payload.review_id), None)
    if not review:
        raise HTTPException(404, "Отзыв не найден в очереди")
    ok = await wb_post_answer(payload.review_id, payload.text)
    if not ok:
        raise HTTPException(502, "WB API вернул ошибку при публикации")
    data["pending"] = [r for r in data["pending"] if r["id"] != payload.review_id]
    data["published"].append({**review, "reply": payload.text, "mode": "manual", "published_at": datetime.now().isoformat()})
    save_data(data)
    return {"ok": True}

class RegeneratePayload(BaseModel):
    review_id: str
    has_photo: bool = False

@app.post("/api/regenerate")
async def regenerate(payload: RegeneratePayload):
    data = load_data()
    review = next((r for r in data["pending"] if r["id"] == payload.review_id), None)
    if not review:
        raise HTTPException(404, "Отзыв не найден")
    try:
        draft = await generate_reply(review["text"], review["rating"], review.get("article", ""), review.get("product", ""), payload.has_photo)
    except Exception as e:
        raise HTTPException(502, f"Ошибка Claude API: {e}")
    for r in data["pending"]:
        if r["id"] == payload.review_id:
            r["draft"] = draft
            r["has_photo"] = payload.has_photo
    save_data(data)
    return {"draft": draft}

@app.post("/api/check-now")
async def check_now():
    asyncio.create_task(process_new_reviews())
    return {"ok": True}
