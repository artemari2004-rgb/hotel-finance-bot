import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

EXCEL_PATH = DATA_DIR / "finance.xlsx"
CHARTS_DIR = DATA_DIR / "charts"
CHARTS_DIR.mkdir(exist_ok=True)

# Токен берётся из переменной окружения TELEGRAM_BOT_TOKEN
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "PASTE_YOUR_TOKEN_HERE")

# Если указать ID через запятую — бот будет отвечать только им.
# Пустой список = доступен всем.
ALLOWED_USER_IDS = [
    # 123456789,
]

# Пароль входа. Пустая строка = без пароля.
# Можно задать переменной BOT_PASSWORD.
BOT_PASSWORD = os.getenv("BOT_PASSWORD", "12345")

CATEGORIES = {
    "hotel_income": {
        "title": "Доход отеля",
        "kind": "income",
        "payment": "mixed",
        "emoji": "🏨",
    },
    "resto_income": {
        "title": "Доход ресторана",
        "kind": "income",
        "payment": "mixed",
        "emoji": "🍽",
    },
    "purchase_cash": {
        "title": "Расход нал",
        "kind": "expense",
        "payment": "cash",
        "emoji": "💵",
    },
    "purchase_card": {
        "title": "Расход безнал",
        "kind": "expense",
        "payment": "card",
        "emoji": "💳",
    },
    "personal_card": {
        "title": "Траты безнал",
        "kind": "personal",
        "payment": "card",
        "emoji": "👤",
    },
    "other_income": {
        "title": "Прочий доход",
        "kind": "income",
        "payment": "mixed",
        "emoji": "➕",
    },
    "cash_to_bank": {
        "title": "Нал в банк",
        "kind": "transfer",
        "payment": "transfer",
        "emoji": "🏦",
    },
    "debt": {
        "title": "Долг",
        "kind": "debt",
        "payment": "cash",
        "emoji": "📝",
    },
    "debt_repay": {
        "title": "Погашение долга",
        "kind": "debt_repay",
        "payment": "cash",
        "emoji": "💵",
    },
}

CURRENCY = "֏"
SHEET_NAME = "operations"
