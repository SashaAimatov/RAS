import json
import asyncio
import io
import calendar
import logging
import time
from datetime import datetime, date, timedelta
from pathlib import Path
from urllib.parse import urlencode

import requests
from docx import Document
from aiogram import Bot, Dispatcher, F, BaseMiddleware
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    Message, InlineKeyboardMarkup, InlineKeyboardButton,
    CallbackQuery, InlineQuery, InlineQueryResultArticle,
    InputTextMessageContent
)
from apscheduler.schedulers.asyncio import AsyncIOScheduler

# ================= НАСТРОЙКИ =================
TELEGRAM_TOKEN = "8299017583:AAE2QEwHva2OLUpL_h_39oE0R7jEKyEeNLE"

OWNER_USERNAME = "pampampa1"
INITIAL_ADMINS = ["ginytimary"]
CONTACT_USERNAME = "raspisach"

SUBSCRIBERS_FILE = Path("subscribers.json")
ADMINS_FILE = Path("admins.json")
BANNED_FILE = Path("banned.json")
BANNED_USERNAMES_FILE = Path("banned_usernames.json")
CANCELLED_FILE = Path("cancelled.json")
MUTE_FILE = Path("mute.json")
SETTINGS_FILE = Path("settings.json")
OWNER_FILE = Path("owner.json")
LOG_FILE = Path("bot.log")

DEFAULT_SETTINGS = {
    "ref_monday": "2026-10-12",
    "ref_is_numerator": True,
}

WEEKDAYS_RU = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
WEEKDAYS_ALIASES = {
    "пн": "Понедельник", "понедельник": "Понедельник",
    "вт": "Вторник", "вторник": "Вторник",
    "ср": "Среда", "среда": "Среда",
    "чт": "Четверг", "четверг": "Четверг",
    "пт": "Пятница", "пятница": "Пятница",
}

THIRD_BUILDING_DAYS = {"Понедельник", "Пятница"}
YADISK_PUBLIC_KEY = "https://disk.yandex.ru/d/vCkixjXynkqHwQ"

MONTHS_RU_SHORT = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн",
                   "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]

START_TIME = time.time()

# ================= ЛОГИ =================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("bot")

# ================= РАСПИСАНИЕ =================
SCHEDULE = {
    "group": "ТМ-25-202",
    "times": {
        "1": "8:20–9:50", "2": "10:00–11:30", "3": "12:10–13:40",
        "4": "14:00–15:30", "5": "15:40–17:10", "6": "17:20–18:50",
        "7": "19:00–20:30"
    },
    "days": {
        "Понедельник": [
            {"pair": 2, "subject": "История России", "room": "107/3к/у"},
            {"pair": 3, "subject": "Процессы формообразования и инструменты", "room": "201м"},
            {"pair": 4, "subject": "Охрана труда", "room": "207м"},
            {"pair": 5, "subject": "Инженерная графика", "room": "207м"}
        ],
        "Вторник": [
            {"pair": 1, "subject": "Иностранный язык", "room": "120о / 03"},
            {
                "pair": 2,
                "numerator":   {"subject": "Математика в профессиональной деятельности", "room": "214"},
                "denominator": {"subject": "Технология машиностроения", "room": "122о"}
            },
            {"pair": 3, "subject": "Математика в профессиональной деятельности", "room": "214"},
            {"pair": 4, "subject": "Физическая культура", "room": ""}
        ],
        "Среда": [
            {"pair": 3, "subject": "Техническая механика", "room": "106о"},
            {
                "pair": 4,
                "numerator":   {"subject": "Техническая механика", "room": "201м"},
                "denominator": {"subject": "Инженерная графика", "room": "105о"}
            }
        ],
        "Четверг": [
            {
                "pair": 1,
                "numerator":   {"subject": "История России", "room": "310"},
                "denominator": {"subject": "—", "room": ""}
            },
            {"pair": 2, "subject": "Материаловедение", "room": "01"},
            {"pair": 3, "subject": "Техническая механика", "room": "106о"},
            {"pair": 4, "subject": "Материаловедение", "room": "01"},
            {
                "pair": 5,
                "numerator":   {"subject": "—", "room": ""},
                "denominator": {"subject": "Материаловедение", "room": "01"}
            }
        ],
        "Пятница": [
            {"pair": 3, "subject": "Технологическое оборудование", "room": "201м"},
            {"pair": 4, "subject": "Технология машиностроения", "room": "201м"},
            {"pair": 5, "subject": "Процессы формообразования и инструменты", "room": "201м"},
            {"pair": 6, "subject": "Основы бережливого производства", "room": "201м"}
        ]
    }
}

TIMES = SCHEDULE["times"]
DAYS = SCHEDULE["days"]


# ================= ФАЙЛЫ =================
def _load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path: Path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_subscribers(): return set(_load_json(SUBSCRIBERS_FILE, []))
def save_subscribers(s): _save_json(SUBSCRIBERS_FILE, sorted(s))
def load_admins(): return set(_load_json(ADMINS_FILE, INITIAL_ADMINS))
def save_admins(s): _save_json(ADMINS_FILE, sorted(s))
def load_banned(): return set(_load_json(BANNED_FILE, []))
def save_banned(s): _save_json(BANNED_FILE, sorted(s))
def load_banned_usernames(): return set(_load_json(BANNED_USERNAMES_FILE, []))
def save_banned_usernames(s): _save_json(BANNED_USERNAMES_FILE, sorted(s))
def load_cancelled(): return _load_json(CANCELLED_FILE, {})
def save_cancelled(d): _save_json(CANCELLED_FILE, d)
def load_mute(): return _load_json(MUTE_FILE, {})
def save_mute(d): _save_json(MUTE_FILE, d)


def load_settings():
    s = _load_json(SETTINGS_FILE, {})
    return {**DEFAULT_SETTINGS, **s}


def save_settings(s): _save_json(SETTINGS_FILE, s)


def add_subscriber(chat_id: int):
    subs = load_subscribers()
    if chat_id not in subs:
        subs.add(chat_id)
        save_subscribers(subs)


def remove_subscriber(chat_id: int):
    subs = load_subscribers()
    if chat_id in subs:
        subs.remove(chat_id)
        save_subscribers(subs)


def get_owner_chat_id():
    return _load_json(OWNER_FILE, {}).get("chat_id")


def set_owner_chat_id(chat_id: int):
    _save_json(OWNER_FILE, {"chat_id": chat_id})


def _username(message: Message) -> str:
    return (message.from_user.username or "").lower()


def is_owner(message: Message) -> bool:
    return _username(message) == OWNER_USERNAME.lower()


def is_admin(message: Message) -> bool:
    if is_owner(message):
        return True
    return _username(message) in load_admins()


def is_banned(message: Message) -> bool:
    if message.chat.id in load_banned():
        return True
    u = _username(message)
    if u and u in load_banned_usernames():
        return True
    return False


def is_muted(chat_id: int) -> bool:
    m = load_mute()
    until = m.get(str(chat_id))
    if not until:
        return False
    try:
        return datetime.strptime(until, "%Y-%m-%d").date() >= date.today()
    except Exception:
        return False


# ================= ЛОГИКА НЕДЕЛИ =================
def get_ref_monday() -> date:
    try:
        return datetime.strptime(load_settings()["ref_monday"], "%Y-%m-%d").date()
    except Exception:
        return date(2026, 10, 12)


def ref_is_numerator() -> bool:
    return load_settings().get("ref_is_numerator", True)


def is_numerator(d: date) -> bool:
    weeks_passed = (d - get_ref_monday()).days // 7
    return (weeks_passed % 2 == 0) == ref_is_numerator()


def days_until_week_change(d: date) -> int:
    days_to_monday = (7 - d.weekday()) % 7
    return days_to_monday if days_to_monday > 0 else 7


# ================= ЯНДЕКС.ДИСК =================
def yadisk_list_items(public_key: str, path: str = "/"):
    params = {"public_key": public_key, "limit": 200}
    if path and path != "/":
        params["path"] = path
    url = "https://cloud-api.yandex.net/v1/disk/public/resources?" + urlencode(params)
    try:
        r = requests.get(url, timeout=20)
        if r.status_code != 200:
            log.warning(f"Яндекс.Диск список — код {r.status_code}")
            return []
        return r.json().get("_embedded", {}).get("items", [])
    except Exception as e:
        log.exception(f"Ошибка Яндекс.Диск список: {e}")
        return []


def yadisk_download(public_key: str, path: str):
    params = {"public_key": public_key, "path": path}
    url = "https://cloud-api.yandex.net/v1/disk/public/resources/download?" + urlencode(params)
    try:
        r = requests.get(url, timeout=20)
        if r.status_code != 200:
            log.warning(f"Яндекс.Диск download — код {r.status_code}")
            return None
        href = r.json().get("href")
        if not href:
            return None
        file_r = requests.get(href, timeout=40)
        return file_r.content if file_r.status_code == 200 else None
    except Exception as e:
        log.exception(f"Ошибка Яндекс.Диск download: {e}")
        return None


MONTHS_RU_GEN = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля", 5: "мая", 6: "июня",
    7: "июля", 8: "августа", 9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
}


def find_replacement_file(target_date: date):
    day = target_date.day
    month_name = MONTHS_RU_GEN[target_date.month]
    month_folder = month_name[:-1] + "ь"
    weekday = WEEKDAYS_RU[target_date.weekday()]
    is_third = weekday in THIRD_BUILDING_DAYS

    candidates = [
        f"{day} {month_name}",
        f"{day}.{target_date.month:02d}.{target_date.year}",
        f"{day}.{target_date.month}.{target_date.year}",
    ]
    has_date = lambda n: any(c in n for c in candidates)
    is_docx = lambda n: n.lower().endswith(".docx")
    is_xlsx = lambda n: n.lower().endswith(".xlsx")
    is_3k = lambda n: "3 корпус" in n.lower()

    all_files = []
    root = yadisk_list_items(YADISK_PUBLIC_KEY, "/")
    for item in root:
        if item.get("type") == "file":
            all_files.append(item)
    month_path = None
    for item in root:
        if item.get("type") == "dir" and item.get("name", "").lower() in (
            month_folder.lower(), month_name.lower()
        ):
            month_path = item.get("path")
            break
    if month_path:
        for item in yadisk_list_items(YADISK_PUBLIC_KEY, month_path):
            if item.get("type") == "file":
                all_files.append(item)

    for item in all_files:
        n = item.get("name", "")
        if is_docx(n) and has_date(n) and not is_3k(n):
            return item.get("path"), n
    if is_third:
        for item in all_files:
            n = item.get("name", "")
            if is_xlsx(n) and has_date(n) and is_3k(n):
                return item.get("path"), n
    for item in all_files:
        n = item.get("name", "")
        if is_docx(n) and has_date(n):
            return item.get("path"), n
    if is_third:
        for item in all_files:
            n = item.get("name", "")
            if is_xlsx(n) and has_date(n):
                return item.get("path"), n
    return None, None


def parse_replacement_docx(content: bytes) -> dict:
    out = {}
    try:
        doc = Document(io.BytesIO(content))
    except Exception as e:
        log.exception(f"Ошибка чтения docx: {e}")
        return out
    roman = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7}
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        lower = text.lower()
        if not ("тм 202" in lower or "тм-202" in lower or "тм202" in lower):
            continue
        pair_num = None
        for w in text.split():
            ww = w.strip(".,;:()")
            if ww in roman:
                pair_num = roman[ww]
                break
        if pair_num is None:
            continue
        if "→" in text:
            repl = text.split("→", 1)[1].strip()
        elif "->" in text:
            repl = text.split("->", 1)[1].strip()
        else:
            repl = text
        if repl.lower().startswith("на "):
            repl = repl[3:].strip()
        out[pair_num] = repl
    return out


def parse_replacement_xlsx(content: bytes) -> dict:
    out = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    except Exception as e:
        log.exception(f"Ошибка чтения xlsx: {e}")
        return out
    roman = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7}
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            cells = [str(c) if c is not None else "" for c in row]
            line = " | ".join([c for c in cells if c.strip()])
            if not line:
                continue
            lower = line.lower()
            if not ("тм 202" in lower or "тм-202" in lower or "тм202" in lower):
                continue
            pair_num = None
            for w in line.split():
                ww = w.strip(".,;:()|")
                if ww in roman:
                    pair_num = roman[ww]
                    break
            if pair_num is None:
                continue
            if "→" in line:
                repl = line.split("→", 1)[1].strip()
            elif "->" in line:
                repl = line.split("->", 1)[1].strip()
            else:
                ne = [c.strip() for c in cells if c.strip()]
                repl = " | ".join(ne[-2:]) if len(ne) >= 2 else line
            repl = repl.strip(" |")
            if repl.lower().startswith("на "):
                repl = repl[3:].strip()
            out[pair_num] = repl
    return out


def get_replacements_for_date(target_date: date) -> dict:
    path, name = find_replacement_file(target_date)
    if not path:
        return {}
    content = yadisk_download(YADISK_PUBLIC_KEY, path)
    if not content:
        return {}
    if name.lower().endswith(".docx"):
        return parse_replacement_docx(content)
    if name.lower().endswith(".xlsx"):
        return parse_replacement_xlsx(content)
    return {}


# ================= РАСПИСАНИЕ НА ДАТУ =================
def get_schedule_for_date(d: date, use_replacements: bool = True, use_cancelled: bool = True):
    weekday = WEEKDAYS_RU[d.weekday()]
    numerator = is_numerator(d)
    week_type = "числитель" if numerator else "знаменатель"

    if weekday not in DAYS or d.weekday() >= 5:
        return None, week_type, []

    result = []
    for item in DAYS[weekday]:
        pair_num = item["pair"]
        if "subject" in item:
            subject = item["subject"]
            room = item.get("room", "")
        else:
            variant = item["numerator"] if numerator else item["denominator"]
            subject = variant["subject"]
            room = variant.get("room", "")
        if subject == "—":
            continue
        result.append({
            "pair": pair_num,
            "subject": subject,
            "room": room,
            "time": TIMES.get(str(pair_num), ""),
        })

    changes = []

    if use_cancelled:
        cancelled = load_cancelled()
        date_key = d.strftime("%Y-%m-%d")
        c_today = cancelled.get(date_key, [])
        if c_today:
            removed = [p for p in result if p["pair"] in c_today]
            result = [p for p in result if p["pair"] not in c_today]
            for r in removed:
                changes.append({
                    "pair": r["pair"],
                    "old": r["subject"] + (f" ({r['room']})" if r["room"] else ""),
                    "new": "ОТМЕНЕНО",
                })

    if use_replacements and d in (date.today(), date.today() + timedelta(days=1)):
        try:
            replacements = get_replacements_for_date(d)
            if replacements:
                for pair in result:
                    pn = pair["pair"]
                    if pn in replacements:
                        old = pair["subject"]
                        if pair["room"]:
                            old += f" ({pair['room']})"
                        pair["subject"] = replacements[pn]
                        pair["room"] = ""
                        changes.append({"pair": pn, "old": old, "new": replacements[pn]})
        except Exception as e:
            log.exception(f"Ошибка при обработке замен: {e}")

    return result, week_type, changes


# ================= ФОРМАТИРОВАНИЕ =================
def format_message(d: date, title_prefix: str = "📅 Расписание на") -> str:
    weekday = WEEKDAYS_RU[d.weekday()]
    pairs, week_type, changes = get_schedule_for_date(d)
    date_str = d.strftime("%d.%m.%Y")
    header = f"{title_prefix} {date_str}, {weekday.lower()} ({week_type})\n\n"
    if not pairs:
        return header + "Пар нет 🎉"
    lines = []
    for p in pairs:
        room_part = f" — {p['room']}" if p["room"] else ""
        time_part = f" ({p['time']})" if p["time"] else ""
        lines.append(f"{p['pair']}. {p['subject']}{room_part}{time_part}")
    text = header + "\n".join(lines)
    if changes:
        text += "\n\n🔄 Изменения:\n"
        for c in changes:
            text += f"• {c['pair']} пара: {c['old']} → {c['new']}\n"
    text += f"\n💬 Вопросы — @{CONTACT_USERNAME}"
    return text


def format_preview(d: date, title: str) -> str:
    pairs, week_type, changes = get_schedule_for_date(d, use_replacements=False)
    date_str = d.strftime("%d.%m.%Y")
    weekday = WEEKDAYS_RU[d.weekday()]
    text = f"{title}\n{date_str}, {weekday.lower()} ({week_type})\n\n"
    if not pairs:
        return text + "Пар нет 🎉"
    for p in pairs:
        room_part = f" — {p['room']}" if p["room"] else ""
        text += f"{p['pair']}. {p['subject']}{room_part}\n"
    return text


def format_week() -> str:
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    text = f"📚 Расписание на неделю ({SCHEDULE['group']})\n\n"
    for i in range(5):
        d = monday + timedelta(days=i)
        weekday = WEEKDAYS_RU[d.weekday()]
        pairs, week_type, _ = get_schedule_for_date(d, use_replacements=False, use_cancelled=False)
        text += f"▫️ {weekday} ({d.strftime('%d.%m')})\n"
        if not pairs:
            text += "   — пар нет\n"
        else:
            for p in pairs:
                room_part = f" — {p['room']}" if p["room"] else ""
                text += f"   {p['pair']}. {p['subject']}{room_part}\n"
        text += "\n"
    text += f"💬 Вопросы — @{CONTACT_USERNAME}"
    return text.strip()


def format_bells() -> str:
    text = "🔔 Расписание звонков\n\n"
    for pair_num in sorted(TIMES.keys(), key=lambda x: int(x)):
        text += f"{pair_num} пара — {TIMES[pair_num]}\n"
    text += f"\n💬 Вопросы — @{CONTACT_USERNAME}"
    return text.strip()


def format_main_menu() -> str:
    return (
        "👋 Привет! Я бот с расписанием группы ТМ-25-202.\n\n"
        "Выбери, что тебе нужно, кнопками внизу 👇\n\n"
        f"💬 Вопросы и предложения — @{CONTACT_USERNAME}"
    )


# ================= КНОПКИ =================
def main_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="📅 Сегодня", callback_data="btn_today"),
            InlineKeyboardButton(text="🌙 Завтра", callback_data="btn_tomorrow"),
        ],
        [
            InlineKeyboardButton(text="📆 Выбрать дату", callback_data="pick_year"),
        ],
        [
            InlineKeyboardButton(text="📚 Вся неделя", callback_data="btn_week"),
            InlineKeyboardButton(text="🔄 Изменения", callback_data="btn_changes"),
        ],
        [
            InlineKeyboardButton(text="🔔 Звонки", callback_data="btn_bells"),
        ],
        [
            InlineKeyboardButton(
                text="💬 Вопросы и предложения",
                url=f"https://t.me/{CONTACT_USERNAME}"
            ),
        ],
    ])


def bells_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="btn_back_main")],
    ])


def years_keyboard() -> InlineKeyboardMarkup:
    y = date.today().year
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text=str(y), callback_data=f"pick_month_{y}"),
            InlineKeyboardButton(text=str(y + 1), callback_data=f"pick_month_{y+1}"),
        ],
        [InlineKeyboardButton(text="⬅️ Назад", callback_data="btn_back_main")],
    ])


def months_keyboard(year: int) -> InlineKeyboardMarkup:
    rows = []
    row = []
    for i, m in enumerate(MONTHS_RU_SHORT, start=1):
        row.append(InlineKeyboardButton(text=m, callback_data=f"pick_day_{year}_{i}"))
        if len(row) == 3:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    rows.append([InlineKeyboardButton(text="⬅️ Назад", callback_data="pick_year")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def days_keyboard(year: int, month: int) -> InlineKeyboardMarkup:
    cal = calendar.monthcalendar(year, month)
    rows = [
        [
            InlineKeyboardButton(text="Пн", callback_data="noop"),
            InlineKeyboardButton(text="Вт", callback_data="noop"),
            InlineKeyboardButton(text="Ср", callback_data="noop"),
            InlineKeyboardButton(text="Чт", callback_data="noop"),
            InlineKeyboardButton(text="Пт", callback_data="noop"),
            InlineKeyboardButton(text="Сб", callback_data="noop"),
            InlineKeyboardButton(text="Вс", callback_data="noop"),
        ]
    ]
    for week in cal:
        row = []
        for d in week:
            if d == 0:
                row.append(InlineKeyboardButton(text=" ", callback_data="noop"))
            else:
                row.append(InlineKeyboardButton(
                    text=str(d),
                    callback_data=f"show_date_{year}_{month}_{d}"
                ))
        rows.append(row)
    rows.append([InlineKeyboardButton(text="⬅️ Назад", callback_data=f"pick_month_{year}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


# ================= РАССЫЛКА =================
async def broadcast(bot: Bot, text: str):
    subs = load_subscribers()
    if not subs:
        log.info("Нет подписчиков — рассылка пропущена")
        return
    for chat_id in subs:
        if is_muted(chat_id):
            continue
        try:
            await bot.send_message(chat_id=chat_id, text=text, reply_markup=main_keyboard())
        except Exception as e:
            log.warning(f"Не удалось отправить {chat_id}: {e}")
            if "blocked" in str(e).lower() or "chat not found" in str(e).lower():
                remove_subscriber(chat_id)


async def notify_owner(bot: Bot, text: str):
    cid = get_owner_chat_id()
    if not cid:
        log.warning("chat_id владельца неизвестен — уведомление пропущено")
        return
    try:
        await bot.send_message(chat_id=cid, text=text)
    except Exception as e:
        log.warning(f"Не удалось отправить владельцу: {e}")


# ================= ДЖОБЫ =================
async def send_today(bot: Bot):
    text = format_message(date.today(), "☀️ Доброе утро! Расписание на сегодня:")
    await broadcast(bot, text)
    await notify_owner(bot, f"✅ Рассылка «на сегодня» выполнена. Подписчиков: {len(load_subscribers())}")


async def send_tomorrow(bot: Bot):
    tomorrow = date.today() + timedelta(days=1)
    text = format_message(tomorrow, "🌙 Расписание на завтра:")
    await broadcast(bot, text)
    await notify_owner(bot, f"✅ Рассылка «на завтра» выполнена. Подписчиков: {len(load_subscribers())}")


async def preview_today(bot: Bot):
    text = format_preview(date.today(), "🔔 Через 10 минут — рассылка на СЕГОДНЯ:")
    await notify_owner(bot, text)


async def preview_tomorrow(bot: Bot):
    tomorrow = date.today() + timedelta(days=1)
    text = format_preview(tomorrow, "🔔 Через 10 минут — рассылка на ЗАВТРА:")
    await notify_owner(bot, text)


async def warn_today(bot: Bot):
    await notify_owner(bot, "⏰ Через 2 минуты будет запущена рассылка «на сегодня».")


async def warn_tomorrow(bot: Bot):
    await notify_owner(bot, "⏰ Через 2 минуты будет запущена рассылка «на завтра».")


# ================= MIDDLEWARE (БАН) =================
class BanMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, data):
        if isinstance(event, Message):
            if is_banned(event):
                log.info(f"Игнор от забаненного: {event.chat.id} @{_username(event)}")
                return
        return await handler(event, data)


# ================= БОТ =================
bot = Bot(token=TELEGRAM_TOKEN)
dp = Dispatcher()

dp.message.middleware(BanMiddleware())
dp.callback_query.middleware(BanMiddleware())
dp.inline_query.middleware(BanMiddleware())


@dp.message(CommandStart())
async def cmd_start(message: Message):
    add_subscriber(message.chat.id)

    if is_owner(message):
        set_owner_chat_id(message.chat.id)
        await message.answer("👑 Ты — ВЛАДЕЛЕЦ бота.\nКоманды владельца: /ownerhelp")
    elif is_admin(message):
        await message.answer("🛠 Ты — администратор.\nКоманды: /edithelp")

    await message.answer(format_main_menu(), reply_markup=main_keyboard())


@dp.message(Command("menu"))
async def cmd_menu(message: Message):
    await message.answer(format_main_menu(), reply_markup=main_keyboard())


@dp.message(Command("stop"))
async def cmd_stop(message: Message):
    remove_subscriber(message.chat.id)
    await message.answer("Отписался. Вернуться — /start.")


@dp.message(Command("today"))
async def cmd_today(message: Message):
    await message.answer(format_message(date.today()), reply_markup=main_keyboard())


@dp.message(Command("tomorrow"))
async def cmd_tomorrow(message: Message):
    await message.answer(format_message(date.today() + timedelta(days=1)), reply_markup=main_keyboard())


@dp.message(Command("week_all"))
async def cmd_week_all(message: Message):
    await message.answer(format_week(), reply_markup=main_keyboard())


@dp.message(Command("bells"))
async def cmd_bells(message: Message):
    await message.answer(format_bells(), reply_markup=bells_keyboard())


@dp.message(Command("week"))
async def cmd_week(message: Message):
    today = date.today()
    numerator = is_numerator(today)
    week_type = "числитель" if numerator else "знаменатель"
    days_left = days_until_week_change(today)
    next_type = "знаменатель" if numerator else "числитель"
    await message.answer(
        f"📆 Сейчас неделя: {week_type}\n"
        f"До смены на {next_type}: {days_left} дн.\n"
        f"Сегодня: {WEEKDAYS_RU[today.weekday()]}, {today.strftime('%d.%m.%Y')}"
    )


@dp.message(Command("find"))
async def cmd_find(message: Message):
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /find <предмет>\nПример: /find Математика")
        return
    query = parts[1].strip().lower()
    found = []
    for weekday, pairs in DAYS.items():
        for item in pairs:
            if "subject" in item:
                if query in item["subject"].lower():
                    found.append((weekday, item["pair"], item["subject"], item.get("room", "")))
            else:
                for variant in ("numerator", "denominator"):
                    v = item[variant]
                    if query in v["subject"].lower():
                        label = "числ" if variant == "numerator" else "знам"
                        found.append((weekday, item["pair"], f"{v['subject']} [{label}]", v.get("room", "")))
    if not found:
        await message.answer(f"По запросу «{parts[1]}» ничего не найдено.")
        return
    text = f"🔍 Найдено по «{parts[1]}»:\n\n"
    for weekday, pair, subject, room in found:
        room_part = f" — {room}" if room else ""
        text += f"• {weekday}, {pair} пара: {subject}{room_part}\n"
    await message.answer(text)


@dp.message(Command("date"))
async def cmd_date(message: Message):
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /date <ДД.ММ>\nПример: /date 20.10")
        return
    try:
        day, month = parts[1].strip().split(".")
        target = date(date.today().year, int(month), int(day))
    except Exception:
        await message.answer("Не понял дату. Формат: ДД.ММ")
        return
    await message.answer(format_message(target), reply_markup=main_keyboard())


@dp.message(Command("mute"))
async def cmd_mute(message: Message):
    parts = message.text.split()
    if len(parts) < 2:
        await message.answer("Использование: /mute <дни>")
        return
    try:
        days = int(parts[1])
    except ValueError:
        await message.answer("Число дней должно быть числом.")
        return
    until = (date.today() + timedelta(days=days)).strftime("%Y-%m-%d")
    m = load_mute()
    m[str(message.chat.id)] = until
    save_mute(m)
    await message.answer(f"🔕 Тишина до {until} ({days} дн.)")


@dp.message(Command("unmute"))
async def cmd_unmute(message: Message):
    m = load_mute()
    m.pop(str(message.chat.id), None)
    save_mute(m)
    await message.answer("🔔 Рассылка снова включена.")


# ---------- Владелец ----------
@dp.message(Command("ownerhelp"))
async def cmd_ownerhelp(message: Message):
    if not is_owner(message):
        return
    await message.answer(
        "👑 Команды ВЛАДЕЛЬЦА:\n\n"
        "/ping — аптайм, задержка, состояние\n"
        "/logs [N] — последние N строк лога\n"
        "/errors — ошибки и предупреждения\n"
        "/stats — статистика\n"
        "/addadmin @username — назначить админа\n"
        "/deladmin @username — снять админа\n"
        "/listadmins — список админов\n"
        "/ban @username — забанить по username\n"
        "/banid <chat_id> — забанить по chat_id\n"
        "/unban @username — разбанить\n"
        "/listbanned — список забаненных\n"
        "/broadcast <текст> — рассылка всем\n\n"
        "Плюс все админ-команды: /edithelp"
    )


@dp.message(Command("ping"))
async def cmd_ping(message: Message):
    if not is_owner(message):
        return
    uptime_sec = int(time.time() - START_TIME)
    h = uptime_sec // 3600
    m = (uptime_sec % 3600) // 60
    s = uptime_sec % 60
    t0 = time.time()
    try:
        await bot.get_me()
        api_ms = int((time.time() - t0) * 1000)
    except Exception as e:
        api_ms = f"ошибка: {e}"
    await message.answer(
        f"🏓 Понг!\n"
        f"• Аптайм: {h}ч {m}м {s}с\n"
        f"• Задержка Telegram API: {api_ms} мс\n"
        f"• Время сервера: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n"
        f"• Подписчиков: {len(load_subscribers())}\n"
        f"• Админов: {len(load_admins())}\n"
        f"• Забанено: {len(load_banned()) + len(load_banned_usernames())}"
    )


@dp.message(Command("logs"))
async def cmd_logs(message: Message):
    if not is_owner(message):
        return
    parts = message.text.split()
    n = 30
    if len(parts) >= 2:
        try:
            n = int(parts[1])
        except ValueError:
            pass
    if not LOG_FILE.exists():
        await message.answer("Лог пуст.")
        return
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        lines = f.readlines()
    tail = "".join(lines[-n:]) or "(пусто)"
    if len(tail) > 3500:
        tail = tail[-3500:]
    await message.answer(f"📄 Последние {n} строк лога:\n\n{tail}")


@dp.message(Command("errors"))
async def cmd_errors(message: Message):
    if not is_owner(message):
        return
    if not LOG_FILE.exists():
        await message.answer("Лог пуст.")
        return
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        errs = [l for l in f.readlines() if "[ERROR]" in l or "[WARNING]" in l]
    if not errs:
        await message.answer("✅ Ошибок и предупреждений нет.")
        return
    tail = "".join(errs[-20:])
    if len(tail) > 3500:
        tail = tail[-3500:]
    await message.answer(f"⚠️ Последние ошибки:\n\n{tail}")


@dp.message(Command("stats"))
async def cmd_stats(message: Message):
    if not is_owner(message):
        return
    s = load_settings()
    await message.answer(
        f"📊 Статистика:\n"
        f"• Подписчиков: {len(load_subscribers())}\n"
        f"• Админов: {len(load_admins())}\n"
        f"• Забанено (chat_id): {len(load_banned())}\n"
        f"• Забанено (username): {len(load_banned_usernames())}\n"
        f"• Точка отсчёта: {s['ref_monday']} — {'числитель' if s['ref_is_numerator'] else 'знаменатель'}\n"
        f"• Время сервера: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}"
    )


def normalize_username(raw: str) -> str:
    return raw.strip().lstrip("@").lower()


@dp.message(Command("addadmin"))
async def cmd_addadmin(message: Message):
    if not is_owner(message):
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /addadmin @username")
        return
    u = normalize_username(parts[1])
    admins = load_admins()
    if u in admins:
        await message.answer(f"@{u} уже админ.")
        return
    admins.add(u)
    save_admins(admins)
    await message.answer(f"✅ @{u} назначен администратором.")


@dp.message(Command("deladmin"))
async def cmd_deladmin(message: Message):
    if not is_owner(message):
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /deladmin @username")
        return
    u = normalize_username(parts[1])
    admins = load_admins()
    if u not in admins:
        await message.answer(f"@{u} не админ.")
        return
    admins.remove(u)
    save_admins(admins)
    await message.answer(f"🗑 @{u} больше не админ.")


@dp.message(Command("listadmins"))
async def cmd_listadmins(message: Message):
    if not is_owner(message):
        return
    admins = sorted(load_admins())
    if not admins:
        await message.answer("Админов нет.")
        return
    text = "🛠 Админы:\n"
    for a in admins:
        text += f"• @{a}\n"
    await message.answer(text)


@dp.message(Command("ban"))
async def cmd_ban(message: Message):
    if not is_owner(message):
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /ban @username")
        return
    u = normalize_username(parts[1])
    bu = load_banned_usernames()
    if u in bu:
        await message.answer(f"@{u} уже в бан-листе.")
        return
    bu.add(u)
    save_banned_usernames(bu)
    await message.answer(f"🚫 @{u} забанен.")


@dp.message(Command("banid"))
async def cmd_banid(message: Message):
    if not is_owner(message):
        return
    parts = message.text.split()
    if len(parts) < 2:
        await message.answer("Использование: /banid <chat_id>")
        return
    try:
        cid = int(parts[1])
    except ValueError:
        await message.answer("chat_id должен быть числом.")
        return
    banned = load_banned()
    banned.add(cid)
    save_banned(banned)
    remove_subscriber(cid)
    await message.answer(f"🚫 chat_id {cid} забанен.")


@dp.message(Command("unban"))
async def cmd_unban(message: Message):
    if not is_owner(message):
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /unban @username")
        return
    u = normalize_username(parts[1])
    bu = load_banned_usernames()
    if u in bu:
        bu.remove(u)
        save_banned_usernames(bu)
        await message.answer(f"✅ @{u} разбанен.")
    else:
        await message.answer(f"@{u} не в бан-листе.")


@dp.message(Command("listbanned"))
async def cmd_listbanned(message: Message):
    if not is_owner(message):
        return
    bu = sorted(load_banned_usernames())
    ids = sorted(load_banned())
    text = "🚫 Забаненные:\n\n"
    if bu:
        text += "По username:\n"
        for u in bu:
            text += f"• @{u}\n"
    if ids:
        text += "\nПо chat_id:\n"
        for i in ids:
            text += f"• {i}\n"
    if not bu and not ids:
        text = "Никто не забанен."
    await message.answer(text)


@dp.message(Command("broadcast"))
async def cmd_broadcast(message: Message):
    if not is_owner(message):
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /broadcast <текст>")
        return
    text = parts[1]
    subs = load_subscribers()
    sent, failed = 0, 0
    for cid in subs:
        try:
            await bot.send_message(chat_id=cid, text=text)
            sent += 1
        except Exception:
            failed += 1
    await message.answer(f"📣 Рассылка: отправлено {sent}, ошибок {failed}.")


# ---------- Админ ----------
@dp.message(Command("edithelp"))
async def cmd_edithelp(message: Message):
    if not is_admin(message):
        return
    await message.answer(
        "🛠 Команды АДМИНА (расписание):\n\n"
        "/setpair <день> <пара> <предмет> <аудитория>\n"
        "/delpair <день> <пара>\n"
        "/settime <пара> <время>\n"
        "/setweek <день> <пара> <numerator|denominator> <предмет> <аудитория>\n"
        "/show <день>\n"
        "/replacements\n"
        "/cancel <ДД.ММ> <пара>\n"
        "/uncancel <ДД.ММ> <пара>\n"
        "/setref <ДД.ММ.ГГГГ> <числитель|знаменатель>\n\n"
        "Дни: пн, вт, ср, чт, пт."
    )


def normalize_day(raw: str):
    return WEEKDAYS_ALIASES.get(raw.strip().lower())


@dp.message(Command("show"))
async def cmd_show(message: Message):
    if not is_admin(message):
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Использование: /show <день>")
        return
    day = normalize_day(parts[1])
    if not day:
        await message.answer("Не понял день.")
        return
    pairs = DAYS.get(day, [])
    if not pairs:
        await message.answer(f"{day}: пар нет")
        return
    text = f"📋 {day}:\n\n"
    for item in pairs:
        pair = item["pair"]
        t = TIMES.get(str(pair), "")
        if "subject" in item:
            text += f"{pair}. {item['subject']} — {item.get('room','')} ({t})\n"
        else:
            text += f"{pair}. [числ] {item['numerator']['subject']} — {item['numerator'].get('room','')}\n"
            text += f"   [знам] {item['denominator']['subject']} — {item['denominator'].get('room','')}\n"
    await message.answer(text)


@dp.message(Command("setpair"))
async def cmd_setpair(message: Message):
    if not is_admin(message):
        return
    parts = message.text.split(maxsplit=4)
    if len(parts) < 5:
        await message.answer("Использование: /setpair <день> <пара> <предмет> <аудитория>")
        return
    _, day_raw, pair_raw, subject, room = parts
    day = normalize_day(day_raw)
    if not day:
        await message.answer("Не понял день.")
        return
    try:
        pair_num = int(pair_raw)
    except ValueError:
        await message.answer("Номер пары должен быть числом.")
        return
    day_list = DAYS.setdefault(day, [])
    found = False
    for item in day_list:
        if item["pair"] == pair_num:
            item.clear()
            item.update({"pair": pair_num, "subject": subject, "room": room})
            found = True
            break
    if not found:
        day_list.append({"pair": pair_num, "subject": subject, "room": room})
        day_list.sort(key=lambda x: x["pair"])
    await message.answer(f"✅ {day}, пара {pair_num} → «{subject}», ауд. {room}")


@dp.message(Command("delpair"))
async def cmd_delpair(message: Message):
    if not is_admin(message):
        return
    parts = message.text.split()
    if len(parts) < 3:
        await message.answer("Использование: /delpair <день> <пара>")
        return
    day = normalize_day(parts[1])
    if not day:
        await message.answer("Не понял день.")
        return
    try:
        pair_num = int(parts[2])
    except ValueError:
        await message.answer("Номер пары должен быть числом.")
        return
    day_list = DAYS.get(day, [])
    before = len(day_list)
    DAYS[day] = [i for i in day_list if i["pair"] != pair_num]
    if len(DAYS[day]) == before:
        await message.answer(f"Пара {pair_num} в {day} не найдена.")
        return
    await message.answer(f"🗑 Пара {pair_num} в {day} удалена.")


@dp.message(Command("settime"))
async def cmd_settime(message: Message):
    if not is_admin(message):
        return
    parts = message.text.split(maxsplit=2)
    if len(parts) < 3:
        await message.answer("Использование: /settime <пара> <время>")
        return
    _, pair_raw, time_str = parts
    try:
        pair_num = int(pair_raw)
    except ValueError:
        await message.answer("Номер пары должен быть числом.")
        return
    TIMES[str(pair_num)] = time_str.strip()
    await message.answer(f"✅ Пара {pair_num} → время {time_str}")


@dp.message(Command("setweek"))
async def cmd_setweek(message: Message):
    if not is_admin(message):
        return
    parts = message.text.split(maxsplit=5)
    if len(parts) < 6:
        await message.answer("Использование: /setweek <день> <пара> <numerator|denominator> <предмет> <аудитория>")
        return
    _, day_raw, pair_raw, variant_raw, subject, room = parts
    day = normalize_day(day_raw)
    if not day:
        await message.answer("Не понял день.")
        return
    if variant_raw not in ("numerator", "denominator"):
        await message.answer("Вариант должен быть numerator или denominator.")
        return
    try:
        pair_num = int(pair_raw)
    except ValueError:
        await message.answer("Номер пары должен быть числом.")
        return
    day_list = DAYS.setdefault(day, [])
    found = False
    for item in day_list:
        if item["pair"] == pair_num and "numerator" in item:
            item[variant_raw] = {"subject": subject, "room": room}
            found = True
            break
    if not found:
        new_item = {
            "pair": pair_num,
            "numerator": {"subject": "—", "room": ""},
            "denominator": {"subject": "—", "room": ""},
        }
        new_item[variant_raw] = {"subject": subject, "room": room}
        DAYS[day] = [i for i in day_list if i["pair"] != pair_num]
        DAYS[day].append(new_item)
        DAYS[day].sort(key=lambda x: x["pair"])
    label = "числитель" if variant_raw == "numerator" else "знаменатель"
    await message.answer(f"✅ {day}, пара {pair_num} ({label}) → «{subject}», ауд. {room}")


@dp.message(Command("cancel"))
async def cmd_cancel(message: Message):
    if not is_admin(message):
        return
    parts = message.text.split()
    if len(parts) < 3:
        await message.answer("Использование: /cancel <ДД.ММ> <пара>")
        return
    try:
        day, month = parts[1].split(".")
        target = date(date.today().year, int(month), int(day))
    except Exception:
        await message.answer("Не понял дату. Формат: ДД.ММ")
        return
    try:
        pair_num = int(parts[2])
    except ValueError:
        await message.answer("Номер пары должен быть числом.")
        return
    key = target.strftime("%Y-%m-%d")
    cancelled = load_cancelled()
    if key not in cancelled:
        cancelled[key] = []
    if pair_num not in cancelled[key]:
        cancelled[key].append(pair_num)
    save_cancelled(cancelled)
    await message.answer(f"🚫 {target.strftime('%d.%m.%Y')}, пара {pair_num} отменена.")


@dp.message(Command("uncancel"))
async def cmd_uncancel(message: Message):
    if not is_admin(message):
        return
    parts = message.text.split()
    if len(parts) < 3:
        await message.answer("Использование: /uncancel <ДД.ММ> <пара>")
        return
    try:
        day, month = parts[1].split(".")
        target = date(date.today().year, int(month), int(day))
    except Exception:
        await message.answer("Не понял дату. Формат: ДД.ММ")
        return
    try:
        pair_num = int(parts[2])
    except ValueError:
        await message.answer("Номер пары должен быть числом.")
        return
    key = target.strftime("%Y-%m-%d")
    cancelled = load_cancelled()
    if key in cancelled and pair_num in cancelled[key]:
        cancelled[key].remove(pair_num)
        if not cancelled[key]:
            cancelled.pop(key)
        save_cancelled(cancelled)
        await message.answer(f"✅ Отмена пары {pair_num} на {target.strftime('%d.%m.%Y')} снята.")
    else:
        await message.answer("Такой отмены нет.")


@dp.message(Command("setref"))
async def cmd_setref(message: Message):
    if not is_admin(message):
        return
    parts = message.text.split()
    if len(parts) < 3:
        await message.answer("Использование: /setref <ДД.ММ.ГГГГ> <числитель|знаменатель>")
        return
    try:
        ref_date = datetime.strptime(parts[1], "%d.%m.%Y").date()
    except Exception:
        await message.answer("Не понял дату. Формат: ДД.ММ.ГГГГ")
        return
    variant = parts[2].lower()
    if variant not in ("числитель", "знаменатель"):
        await message.answer("Второй аргумент: числитель или знаменатель.")
        return
    s = load_settings()
    s["ref_monday"] = ref_date.strftime("%Y-%m-%d")
    s["ref_is_numerator"] = (variant == "числитель")
    save_settings(s)
    await message.answer(f"✅ Точка отсчёта: {ref_date.strftime('%d.%m.%Y')} — {variant}.")


@dp.message(Command("replacements"))
async def cmd_replacements(message: Message):
    if not is_admin(message):
        return
    await message.answer("⏳ Ищу замены на сегодня...")
    path, name = find_replacement_file(date.today())
    if not path:
        await message.answer("❌ Файл замен на сегодня не найден.")
        return
    content = yadisk_download(YADISK_PUBLIC_KEY, path)
    if not content:
        await message.answer("❌ Не удалось скачать файл.")
        return
    if name.lower().endswith(".docx"):
        repl = parse_replacement_docx(content)
    elif name.lower().endswith(".xlsx"):
        repl = parse_replacement_xlsx(content)
    else:
        repl = {}
    if not repl:
        await message.answer(f"📄 Файл: {name}\n\nЗамен для ТМ 202 не найдено.")
        return
    text = f"📄 Файл: {name}\n\nНайденные замены для ТМ 202:\n"
    for pn, r in sorted(repl.items()):
        text += f"• {pn} пара → {r}\n"
    await message.answer(text)


# ---------- Инлайн ----------
@dp.inline_query()
async def inline_handler(query: InlineQuery):
    q = query.query.strip().lower()
    today = date.today()
    tomorrow = today + timedelta(days=1)

    def make_result(rid, title, d):
        return InlineQueryResultArticle(
            id=rid,
            title=title,
            input_message_content=InputTextMessageContent(
                message_text=format_message(d, "📅 Расписание на")
            )
        )

    results = []
    if q in ("", "сегодня", "today"):
        results.append(make_result("today", "Расписание на сегодня", today))
    if q in ("", "завтра", "tomorrow"):
        results.append(make_result("tomorrow", "Расписание на завтра", tomorrow))
    if q in ("", "неделя", "week"):
        results.append(InlineQueryResultArticle(
            id="week", title="Расписание на неделю",
            input_message_content=InputTextMessageContent(message_text=format_week())
        ))
    await query.answer(results, cache_time=60)


# ---------- Кнопки (редактируем то же сообщение) ----------
@dp.callback_query()
async def cb_handler(callback: CallbackQuery):
    action = callback.data

    text = format_main_menu()
    markup = main_keyboard()

    if action == "btn_today":
        text = format_message(date.today())
        markup = main_keyboard()

    elif action == "btn_tomorrow":
        text = format_message(date.today() + timedelta(days=1))
        markup = main_keyboard()

    elif action == "btn_week":
        text = format_week()
        markup = main_keyboard()

    elif action == "btn_bells":
        text = format_bells()
        markup = bells_keyboard()

    elif action == "btn_changes":
        pairs, week_type, changes = get_schedule_for_date(date.today())
        if not changes:
            text = "🔄 Изменений на сегодня нет."
        else:
            text = "🔄 Изменения на сегодня:\n\n"
            for c in changes:
                text += f"• {c['pair']} пара: {c['old']} → {c['new']}\n"
        text += f"\n💬 Вопросы — @{CONTACT_USERNAME}"
        markup = main_keyboard()

    elif action == "btn_back_main":
        text = format_main_menu()
        markup = main_keyboard()

    elif action == "pick_year":
        text = "Укажи год:"
        markup = years_keyboard()

    elif action.startswith("pick_month_"):
        year = int(action.split("_")[2])
        text = f"Укажи месяц ({year}):"
        markup = months_keyboard(year)

    elif action.startswith("pick_day_"):
        _, _, y, m = action.split("_")
        year, month = int(y), int(m)
        text = f"Укажи день ({MONTHS_RU_SHORT[month-1]} {year}):"
        markup = days_keyboard(year, month)

    elif action.startswith("show_date_"):
        _, _, y, m, d = action.split("_")
        target = date(int(y), int(m), int(d))
        text = format_message(target)
        markup = main_keyboard()

    elif action == "noop":
        await callback.answer()
        return

    try:
        await callback.message.edit_text(text, reply_markup=markup)
    except Exception as e:
        err = str(e).lower()
        if "message is not modified" in err:
            pass
        else:
            try:
                await callback.message.answer(text, reply_markup=markup)
            except Exception:
                pass

    await callback.answer()


# ================= ЗАПУСК =================
# Эту функцию вызывает app.py через отдельный поток
async def run_bot():
    scheduler = AsyncIOScheduler(timezone="Asia/Yekaterinburg")

    scheduler.add_job(preview_today, "cron", hour=5, minute=50, args=[bot])
    scheduler.add_job(preview_tomorrow, "cron", hour=17, minute=50, args=[bot])
    scheduler.add_job(warn_today, "cron", hour=5, minute=58, args=[bot])
    scheduler.add_job(warn_tomorrow, "cron", hour=17, minute=58, args=[bot])
    scheduler.add_job(send_today, "cron", hour=6, minute=0, args=[bot])
    scheduler.add_job(send_tomorrow, "cron", hour=18, minute=0, args=[bot])

    scheduler.start()
    log.info("Бот запущен.")

    await notify_owner(
        bot,
        f"✅ Бот запущен.\n"
        f"Время: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n"
        f"Рассылки: 06:00 и 18:00 (Екатеринбург)\n"
        f"Команды владельца: /ownerhelp"
    )

    print("Бот запущен.")
    await dp.start_polling(bot)


# На Render запуском управляет app.py.
# Локально можно запустить так: python ras.py
if __name__ == "__main__":
    asyncio.run(run_bot())