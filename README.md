# WB Reviews Bot 🤖

Автоматические и полуавтоматические ответы на отзывы Wildberries.

## Логика работы

| Оценка | Режим | Что происходит |
|--------|-------|----------------|
| ⭐⭐⭐⭐⭐ (5) | **Авто** | Выбирается шаблон из списка (ротация), публикуется сразу |
| ⭐⭐⭐⭐ (4) | **Полуавто** | Claude генерирует черновик → ты проверяешь и публикуешь |
| ⭐⭐⭐ (3) | **Полуавто** | То же самое |
| ⭐⭐ (2) | **Полуавто** | То же самое |
| ⭐ (1) | **Полуавто** | То же самое |

---

## Установка на сервер

### 1. Требования
- Docker + Docker Compose
- Открытый порт 8000 (или любой другой)

### 2. Скопируй файлы на сервер
```bash
scp -r wb-reviews/ user@your-server:/opt/wb-reviews
ssh user@your-server
cd /opt/wb-reviews
```

### 3. Создай .env файл
```bash
cp .env.example .env
nano .env
```

Заполни:
```
WB_API_KEY=твой_ключ_от_wb
ANTHROPIC_KEY=sk-ant-твой_ключ_claude
CHECK_INTERVAL_MINUTES=30
```

**Где взять WB API ключ:**
Личный кабинет WB → Настройки → Доступ к API → раздел **«Отзывы»** → создать новый токен

**Где взять Anthropic ключ:**
https://console.anthropic.com → API Keys → Create Key

### 4. Запуск
```bash
docker compose up -d --build
```

### 5. Открыть веб-интерфейс
```
http://your-server-ip:8000
```

---

## Управление

```bash
# Посмотреть логи
docker compose logs -f

# Перезапустить
docker compose restart

# Остановить
docker compose down

# Обновить код (после изменений)
docker compose up -d --build
```

---

## Настройка шаблонов для 5 звёзд

В веб-интерфейсе — вкладка **«Шаблоны ⭐⭐⭐⭐⭐»**.

Шаблоны чередуются по ротации. Рекомендую 5–7 разных вариантов.

---

## Nginx (опционально, для домена + HTTPS)

```nginx
server {
    listen 80;
    server_name reviews.yourstore.ru;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

Затем:
```bash
certbot --nginx -d reviews.yourstore.ru
```

---

## Структура данных

Все данные хранятся в Docker volume в файле `/data/reviews.json`:
- `processed` — ID уже обработанных отзывов (чтобы не дублировать)
- `pending` — черновики, ждущие одобрения
- `published` — история опубликованных ответов
