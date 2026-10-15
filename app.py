import os
import asyncio
import threading
from flask import Flask, request, Response
from aiogram import types

from ras import bot, dp, init_scheduler, log, WEBHOOK_SECRET

app = Flask(__name__)

loop = asyncio.new_event_loop()
asyncio.set_event_loop(loop)


@app.route("/")
def index():
    return "Bot is running"


@app.route("/health")
def health():
    return "OK"


@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        update = types.Update.model_validate(request.json)
        asyncio.run_coroutine_threadsafe(
            dp.feed_webhook_update(bot, update),
            loop,
        )
        return Response("ok", status=200)
    except Exception as e:
        log.exception(f"Ошибка обработки вебхука: {e}")
        return Response("error", status=500)


def run_bot_loop():
    asyncio.set_event_loop(loop)
    init_scheduler()

    external_url = os.environ.get("RENDER_EXTERNAL_URL")
    if external_url:
        webhook_url = f"{external_url}/webhook"
        try:
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

    loop.run_forever()


if __name__ == "__main__":
    t = threading.Thread(target=run_bot_loop, daemon=True)
    t.start()
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
