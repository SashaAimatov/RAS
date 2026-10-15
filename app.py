import os
import asyncio
import threading
from flask import Flask

# Импортируем bot и dp из твоего ras.py
from ras import bot, dp

app = Flask(__name__)


@app.route("/")
def home():
    return "Bot is running"


@app.route("/health")
def health():
    return "OK"


def run_bot():
    """Запускает polling бота в отдельном потоке."""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        loop.run_until_complete(dp.start_polling(bot))
    except Exception as e:
        print(f"Ошибка polling: {e}")


# На Render переменная RENDER выставляется автоматически
if os.environ.get("RENDER"):
    t = threading.Thread(target=run_bot, daemon=True)
    t.start()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)