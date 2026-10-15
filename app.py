import os
import asyncio
import logging
from flask import Flask, request, Response
from aiogram import types

# Импортируем бота, диспетчер и планировщик из ras.py
from ras import bot, dp, init_scheduler, log

app = Flask(__name__)

# На Render автоматически есть переменная RENDER_EXTERNAL_URL
# Она вида https://твой-сервис.onrender.com
WEBHOOK_PATH = "/webhook"
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "secret")


@app.route("/")
def index():
    return "Bot is running"


@app.route("/health")
def health():
    return "OK"


@app.route(WEBHOOK_PATH, methods=["POST"])
async def webhook():
    if request.headers.get("content-type") == "application/json":
        update = types.Update.model_validate(request.json)
        await dp.feed_webhook_update(bot, update, secret=WEBHOOK_SECRET)
        return Response("ok", status=200)
    return Response("bad request", status=400)


def setup_webhook():
    """Устанавливает вебхук в Telegram при старте."""
    external_url = os.environ.get("RENDER_EXTERNAL_URL")
    if not external_url:
        log.warning("RENDER_EXTERNAL_URL не задан — вебхук не установлен.")
        return
    webhook_url = f"{external_url}{WEBHOOK_PATH}"
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(
            bot.set_webhook(
                url=webhook_url,
                secret_token=WEBHOOK_SECRET,
                drop_pending_updates=True,
            )
        )
        log.info(f"Вебхук установлен: {webhook_url}")
    except Exception as e:
        log.exception(f"Ошибка установки вебхука: {e}")


if __name__ == "__main__":
    # Запускаем планировщик
    init_scheduler()
    # Ставим вебхук
    setup_webhook()
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
