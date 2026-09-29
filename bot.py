import logging
import re
from datetime import datetime
from pathlib import Path

from telegram import (
    Update,
    ReplyKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardRemove,
    InputFile,
)
from telegram.ext import (
    Application,
    ApplicationHandlerStop,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)

from config import BOT_TOKEN, ALLOWED_USER_IDS, BOT_PASSWORD, CATEGORIES, CURRENCY, EXCEL_PATH, DATA_DIR, CASH_EXPENSE_ITEMS, CARD_EXPENSE_ITEMS
from storage import (
    add_operation,
    add_extra_card_supplier,
    all_card_suppliers,
    prev_manual_remainder,
    month_summary,
    last_operations,
    load_df,
    debt_balances,
    today_revenue,
    start_new_day,
    parse_user_date,
    get_shift_start,
    roll_if_needed,
    remember_chat,
    MONTH_NAME,
)
from charts import pie_expenses, months_compare, year_bars, profit_year_bars, MONTH_FULL

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    level=logging.INFO,
)
log = logging.getLogger("hotel-bot")

(
    WAIT_PAY,
    WAIT_AMOUNT,
    WAIT_COMMENT,
    WAIT_DNAME,
    WAIT_DAMOUNT,
    WAIT_DDATE,
    WAIT_DCOMMENT,
    WAIT_RNAME,
    WAIT_RAMOUNT,
    WAIT_SUPPLIER,
    WAIT_SUPPLIER_NAME,
    WAIT_P_TURN,
    WAIT_P_ACCT,
) = range(13)

TITLE_TO_ID = {v["title"]: k for k, v in CATEGORIES.items()}


AUTH_PATH = DATA_DIR / "auth.json"


def _auth_ids() -> set[int]:
    if not AUTH_PATH.exists():
        return set()
    try:
        import json
        data = json.loads(AUTH_PATH.read_text(encoding="utf-8"))
        return {int(x) for x in data.get("unlocked", [])}
    except Exception:
        return set()


def _save_auth(ids: set[int]):
    import json
    AUTH_PATH.write_text(
        json.dumps({"unlocked": sorted(ids)}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def is_unlocked(uid: int) -> bool:
    if not BOT_PASSWORD:
        return True
    return uid in _auth_ids()


def unlock_user(uid: int):
    ids = _auth_ids()
    ids.add(int(uid))
    _save_auth(ids)


def lock_user(uid: int):
    ids = _auth_ids()
    ids.discard(int(uid))
    _save_auth(ids)


def allowed(update: Update) -> bool:
    uid = update.effective_user.id if update.effective_user else 0
    if ALLOWED_USER_IDS and uid not in ALLOWED_USER_IDS:
        return False
    if BOT_PASSWORD and not is_unlocked(uid):
        return False
    return True


async def auth_gate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not BOT_PASSWORD or not update.effective_user or not update.message:
        return
    uid = update.effective_user.id
    if is_unlocked(uid):
        return
    text = (update.message.text or "").strip()
    if text == BOT_PASSWORD or text.lstrip("/") == BOT_PASSWORD:
        unlock_user(uid)
        await update.message.reply_text(
            "Пароль верный. Доступ открыт.\n/logout — выйти.",
            reply_markup=main_keyboard(),
        )
        raise ApplicationHandlerStop
    await update.message.reply_text(
        "Бот закрыт паролем. Напиши пароль одним сообщением.",
        reply_markup=ReplyKeyboardRemove(),
    )
    raise ApplicationHandlerStop


async def logout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user:
        lock_user(update.effective_user.id)
    await update.message.reply_text(
        "Вышел. Чтобы снова пользоваться — введи пароль.",
        reply_markup=ReplyKeyboardRemove(),
    )


def main_keyboard() -> ReplyKeyboardMarkup:
    rows = [
        [KeyboardButton("🏨 Доход отеля"), KeyboardButton("🍽 Доход ресторана")],
        [KeyboardButton("🍳 Завтраки → ресторан")],
        [KeyboardButton("💵 Расход нал"), KeyboardButton("💳 Расход безнал")],
        [KeyboardButton("➕ Прочий доход"), KeyboardButton("🏦 Нал в банк")],
        [KeyboardButton("📝 Долг"), KeyboardButton("💰 Выручка за сегодня")],
        [KeyboardButton("📌 Чистая прибыль")],
        [KeyboardButton("📊 Итоги месяца"), KeyboardButton("📅 По месяцам")],
        [KeyboardButton("📋 История"), KeyboardButton("📁 Скачать Excel")],
    ]
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


CARD_PAGE_SIZE = 8


def card_exp_kb(page: int = 0) -> ReplyKeyboardMarkup:
    items = CARD_EXPENSE_ITEMS
    start = page * CARD_PAGE_SIZE
    chunk = items[start:start + CARD_PAGE_SIZE]
    rows = []
    for i in range(0, len(chunk), 2):
        rows.append([KeyboardButton(n) for n in chunk[i:i + 2]])
    nav = []
    if page > 0:
        nav.append(KeyboardButton("⬅️ Назад"))
    if start + CARD_PAGE_SIZE < len(items):
        nav.append(KeyboardButton("➡️ Ещё"))
    nav.append(KeyboardButton("✏️ Другой"))
    rows.append(nav)
    rows.append([KeyboardButton("❌ Отмена")])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


def cash_exp_kb() -> ReplyKeyboardMarkup:
    names = CASH_EXPENSE_ITEMS
    rows = []
    for i in range(0, len(names), 2):
        chunk = names[i:i + 2]
        rows.append([KeyboardButton(n) for n in chunk])
    rows.append([KeyboardButton("❌ Отмена")])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


def cancel_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup([[KeyboardButton("❌ Отмена")]], resize_keyboard=True)


def debt_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("📝 Новый долг"), KeyboardButton("💵 Погашение")],
            [KeyboardButton("📋 Список долгов"), KeyboardButton("⬅️ Назад")],
        ],
        resize_keyboard=True,
    )


def revenue_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("🌅 Начать новый день")],
            [KeyboardButton("⬅️ Назад")],
        ],
        resize_keyboard=True,
    )


def pay_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("💵 Нал"), KeyboardButton("💳 Безнал")],
            [KeyboardButton("❌ Отмена")],
        ],
        resize_keyboard=True,
    )


def parse_amount(text: str) -> float | None:
    t = text.strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
    t = re.sub(r"[^\d.\-]", "", t)
    if not t:
        return None
    try:
        val = float(t)
    except ValueError:
        return None
    if val <= 0:
        return None
    return round(val, 2)


def fmt_money(n: float) -> str:
    return f"{n:,.2f}".replace(",", " ").replace(".", ",") + f" {CURRENCY}"


async def ensure_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat:
        remember_chat(update.effective_chat.id)
    info = roll_if_needed(update.effective_chat.id if update.effective_chat else None)
    if not info:
        return
    name = MONTH_NAME[info["closed_month"]]
    await update.message.reply_text(
        f"{name.capitalize()} закрыт.\n"
        f"Остаток на руках {fmt_money(info['opening'])} "
        f"перенесён в {MONTH_NAME[info['new_month']]}.\n"
        "Файл за закрытый месяц ниже. В рабочих листах теперь новый месяц, "
        "старые строки не удалены — они в «Все операции»."
    )
    arch = info["archive"]
    if arch and Path(arch).exists():
        with open(arch, "rb") as f:
            await update.message.reply_document(
                document=InputFile(f, filename=Path(arch).name),
                caption=f"Таблица за {name} {info['closed_year']}",
            )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        await update.message.reply_text("Нет доступа.")
        return
    await ensure_month(update, context)
    await update.message.reply_text(
        "Учёт отеля и ресторана.\n\n"
        "Нажми категорию → введи сумму → при желании комментарий.\n"
        "В конце месяца смотри «Итоги» — круговая диаграмма трат "
        "и столбцы по месяцам, чтобы сравнить оборот.",
        reply_markup=main_keyboard(),
    )


async def pick_category(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return ConversationHandler.END
    await ensure_month(update, context)
    title = update.message.text
    # убрать эмодзи в начале, если пользователь нажал кнопку
    clean = title
    for cid, meta in CATEGORIES.items():
        if title.endswith(meta["title"]) or title == meta["title"]:
            context.user_data["category_id"] = cid
            kind = "доход" if meta["kind"] == "income" else "расход"
            preset = meta.get("payment")
            if cid == "breakfast_move":
                context.user_data["payment"] = "none"
                await update.message.reply_text(
                    "🍳 Завтраки из проживания → ресторан.\n"
                    "Сумма уйдёт из выручки отеля в ресторан, общий оборот не вырастет.\n"
                    "Введи сумму:",
                    reply_markup=cancel_kb(),
                )
                return WAIT_AMOUNT
            if cid in {"purchase_cash", "purchase_card"}:
                context.user_data["payment"] = preset
                if cid == "purchase_cash":
                    await update.message.reply_text(
                        "Расход нал. Выбери статью:",
                        reply_markup=cash_exp_kb(),
                    )
                    return WAIT_SUPPLIER
                context.user_data["card_page"] = 0
                await update.message.reply_text(
                    "Расход безнал. Выбери поставщика (часто используемые сверху).\n"
                    "«Ещё» — следующая страница.",
                    reply_markup=card_exp_kb(0),
                )
                return WAIT_SUPPLIER
            if preset in {"cash", "card", "transfer"}:
                context.user_data["payment"] = preset
                extra = ""
                if preset == "cash":
                    extra = ", нал"
                elif preset == "card":
                    extra = ", безнал"
                elif preset == "transfer":
                    extra = " — сумма уйдёт из кассы на счёт, оборот не увеличится"
                    kind = "инкассация"
                await update.message.reply_text(
                    f"{meta['emoji']} {meta['title']} ({kind}{extra}).\nВведи сумму:",
                    reply_markup=cancel_kb(),
                )
                return WAIT_AMOUNT
            await update.message.reply_text(
                f"{meta['emoji']} {meta['title']} ({kind}).\nЭто нал или безнал?",
                reply_markup=pay_kb(),
            )
            return WAIT_PAY
    await update.message.reply_text("Не понял категорию. Выбери кнопку.", reply_markup=main_keyboard())
    return ConversationHandler.END


async def got_pay(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.message.text or "").lower()
    if "отмена" in text:
        return await cancel(update, context)
    if "безнал" in text or "карт" in text:
        context.user_data["payment"] = "card"
        label = "безнал"
    elif "нал" in text:
        context.user_data["payment"] = "cash"
        label = "нал"
    else:
        await update.message.reply_text("Нажми «Нал» или «Безнал».", reply_markup=pay_kb())
        return WAIT_PAY
    meta = CATEGORIES[context.user_data["category_id"]]
    await update.message.reply_text(
        f"{meta['emoji']} {meta['title']}, {label}.\nВведи сумму:",
        reply_markup=cancel_kb(),
    )
    return WAIT_AMOUNT


async def got_supplier(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.message.text or "").strip()
    if "отмена" in text.lower():
        return await cancel(update, context)
    cid = context.user_data.get("category_id")
    if cid == "purchase_card":
        page = int(context.user_data.get("card_page") or 0)
        if text in {"➡️ Ещё", "Ещё"}:
            context.user_data["card_page"] = page + 1
            await update.message.reply_text(
                f"Страница {page + 2}",
                reply_markup=card_exp_kb(page + 1),
            )
            return WAIT_SUPPLIER
        if text in {"⬅️ Назад", "Назад"}:
            page = max(0, page - 1)
            context.user_data["card_page"] = page
            await update.message.reply_text(
                f"Страница {page + 1}",
                reply_markup=card_exp_kb(page),
            )
            return WAIT_SUPPLIER
        if "другой" in text.lower():
            await update.message.reply_text("Название поставщика:", reply_markup=cancel_kb())
            return WAIT_SUPPLIER_NAME
        if text not in CARD_EXPENSE_ITEMS:
            await update.message.reply_text(
                "Выбери поставщика с кнопки.",
                reply_markup=card_exp_kb(page),
            )
            return WAIT_SUPPLIER
        context.user_data["expense_item"] = text
        await update.message.reply_text(f"{text}. Введи сумму:", reply_markup=cancel_kb())
        return WAIT_AMOUNT
    if text not in CASH_EXPENSE_ITEMS:
        await update.message.reply_text("Выбери кнопку статьи.", reply_markup=cash_exp_kb())
        return WAIT_SUPPLIER
    context.user_data["expense_item"] = text
    await update.message.reply_text(f"{text}. Введи сумму:", reply_markup=cancel_kb())
    return WAIT_AMOUNT


async def got_supplier_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.message.text or "").strip()
    if "отмена" in text.lower():
        return await cancel(update, context)
    if len(text) < 2:
        await update.message.reply_text("Напиши название траты.")
        return WAIT_SUPPLIER_NAME
    context.user_data["expense_item"] = text
    await update.message.reply_text(f"{text}. Введи сумму:", reply_markup=cancel_kb())
    return WAIT_AMOUNT


async def got_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text == "❌ Отмена":
        context.user_data.clear()
        await update.message.reply_text("Отменено.", reply_markup=main_keyboard())
        return ConversationHandler.END

    amount = parse_amount(update.message.text or "")
    if amount is None:
        await update.message.reply_text("Нужно число больше нуля. Пример: 12500 или 12 500,50")
        return WAIT_AMOUNT
    context.user_data["amount"] = amount
    if context.user_data.get("category_id") == "breakfast_move":
        user = update.effective_user
        row = add_operation(
            category_id="breakfast_move",
            amount=amount,
            comment="",
            user_id=user.id,
            username=user.username or user.full_name,
            payment="none",
            category_title="Завтраки → ресторан",
        )
        now = datetime.now()
        s = month_summary(now.year, now.month)
        await update.message.reply_text(
            f"Перенос {fmt_money(amount)}: отель − / ресторан +\n"
            "Общий оборот не изменился.\n\n"
            f"{MONTH_FULL[now.month].capitalize()}: "
            f"доход {fmt_money(s['income'])}, прибыль {fmt_money(s['profit'])}",
            reply_markup=main_keyboard(),
        )
        context.user_data.clear()
        return ConversationHandler.END
    await update.message.reply_text(
        f"Сумма {fmt_money(amount)}.\n"
        "Напиши комментарий (поставщик, номер чека) или отправь «-», чтобы пропустить.",
        reply_markup=cancel_kb(),
    )
    return WAIT_COMMENT


async def got_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text == "❌ Отмена":
        context.user_data.clear()
        await update.message.reply_text("Отменено.", reply_markup=main_keyboard())
        return ConversationHandler.END

    comment = (update.message.text or "").strip()
    if comment in {"-", "—", "нет", "Нет"}:
        comment = ""
    cid = context.user_data.get("category_id")
    amount = context.user_data.get("amount")
    if not cid or amount is None:
        await update.message.reply_text("Сессия сброшена. Выбери категорию снова.", reply_markup=main_keyboard())
        return ConversationHandler.END

    user = update.effective_user
    row = add_operation(
        category_id=cid,
        amount=amount,
        comment=comment,
        user_id=user.id,
        username=user.username or user.full_name,
        payment=context.user_data.get("payment"),
        category_title=context.user_data.get("expense_item"),
    )
    sign = "+" if row["kind"] == "income" else "−"
    pay_ru = {"cash": "нал", "card": "безнал", "transfer": "нал → банк"}.get(row["payment"], "")
    now = datetime.now()
    s = month_summary(now.year, now.month)
    await update.message.reply_text(
        f"Записал: {sign}{fmt_money(amount)}\n"
        f"{row['category']} · {pay_ru}"
        + (f"\n💬 {comment}" if comment else "")
        + f"\n\n{MONTH_FULL[now.month].capitalize()} {now.year}:\n"
        f"доход {fmt_money(s['income'])}\n"
        f"расход {fmt_money(s['expense'])}\n"
        f"прибыль {fmt_money(s['profit'])}",
        reply_markup=main_keyboard(),
    )
    context.user_data.clear()
    return ConversationHandler.END


async def profit_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return ConversationHandler.END
    await update.message.reply_text(
        "Чистая прибыль (не входит в обороты и итоги).\n"
        "Оборот ресторана:",
        reply_markup=cancel_kb(),
    )
    return WAIT_P_TURN


async def profit_got_turn(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "отмена" in (update.message.text or "").lower():
        return await cancel(update, context)
    amount = parse_amount(update.message.text or "")
    if amount is None:
        await update.message.reply_text("Число, например 1500000")
        return WAIT_P_TURN
    context.user_data["p_turn"] = amount
    await update.message.reply_text("Сумма бухгалтера:", reply_markup=cancel_kb())
    return WAIT_P_ACCT


async def profit_got_acct(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "отмена" in (update.message.text or "").lower():
        return await cancel(update, context)
    raw = (update.message.text or "").strip().replace(" ", "").replace(",", ".")
    raw = re.sub(r"[^\d.\-]", "", raw)
    try:
        acct = float(raw)
    except ValueError:
        acct = None
    if acct is None or acct < 0:
        await update.message.reply_text("Число, например 400000")
        return WAIT_P_ACCT
    turn = float(context.user_data.get("p_turn") or 0)
    prev = prev_manual_remainder()
    profit = turn + acct - prev
    user = update.effective_user
    add_operation(
        category_id="manual_profit",
        amount=profit,
        comment=f"{turn}|{acct}|{prev}",
        user_id=user.id,
        username=user.username or user.full_name,
        payment="none",
        category_title="Чистая прибыль",
    )
    await update.message.reply_text(
        f"Оборот ресторана: {fmt_money(turn)}\n"
        f"Бухгалтер: {fmt_money(acct)}\n"
        f"Остаток прошлого месяца: {fmt_money(prev)}\n"
        f"Чистая прибыль: {fmt_money(profit)}\n"
        f"(оборот + бухгалтер − остаток прошлого месяца)\n\n"
        "В общие итоги и кассу не попало.",
        reply_markup=main_keyboard(),
    )
    path = profit_year_bars()
    if path and path.exists():
        with open(path, "rb") as f:
            await update.message.reply_photo(photo=InputFile(f, filename=path.name))
    context.user_data.clear()
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Отменено.", reply_markup=main_keyboard())
    return ConversationHandler.END


async def report_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return ConversationHandler.END
    await ensure_month(update, context)
    try:
        now = datetime.now()
        s = month_summary(now.year, now.month)
        text = (
            f"📊 {MONTH_FULL[now.month].capitalize()} {now.year}\n"
            f"Операций: {s['rows']}\n"
            f"Доход: {fmt_money(s['income'])}\n"
            f"Расход: {fmt_money(s['expense'])}\n"
            f"Прибыль: {fmt_money(s['profit'])}\n"
        )
        if s["by_cat"]:
            text += "\nПо статьям:\n"
            merged = {}
            for item in s["by_cat"]:
                key = (item["kind"], item["category"])
                merged[key] = merged.get(key, 0) + float(item["amount"])
            for (kind, cat), amt in merged.items():
                mark = "＋" if kind == "income" else "−"
                text += f"{mark} {cat}: {fmt_money(amt)}\n"
        await update.message.reply_text(text, reply_markup=main_keyboard())

        pie = pie_expenses(now.year, now.month)
        if pie and pie.exists():
            with open(pie, "rb") as f:
                await update.message.reply_photo(
                    photo=InputFile(f, filename=pie.name),
                    caption="Траты месяца и чистая прибыль в заголовке, драмы",
                )
        else:
            await update.message.reply_text(
                "Пока нет расходов в этом месяце — пирог появится после закупок."
            )
        year_path = year_bars(now.year)
        if year_path and year_path.exists():
            with open(year_path, "rb") as f:
                await update.message.reply_photo(
                    photo=InputFile(f, filename=year_path.name),
                    caption=f"Чистая прибыль по месяцам {now.year}, драмы ({CURRENCY})",
                )
    except Exception:
        log.exception("report_month failed")
        await update.message.reply_text(
            "Не смог построить отчёт. Напиши сюда в чат Grok — разберём.",
            reply_markup=main_keyboard(),
        )
    return ConversationHandler.END


async def report_months(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return ConversationHandler.END
    try:
        now = datetime.now()
        path = year_bars(now.year)
        if not path or not path.exists():
            await update.message.reply_text(
                "Пока мало данных. Внеси операции хотя бы за один месяц.",
                reply_markup=main_keyboard(),
            )
            return ConversationHandler.END
        with open(path, "rb") as f:
            await update.message.reply_photo(
                photo=InputFile(f, filename=path.name),
                caption=(
                    f"Чистая прибыль по месяцам {now.year}, драмы ({CURRENCY}).\n"
                    "На столбце сумма за месяц, в заголовке — итог за год."
                ),
            )
    except Exception:
        log.exception("report_months failed")
        await update.message.reply_text(
            "Не смог построить график по месяцам.",
            reply_markup=main_keyboard(),
        )
    return ConversationHandler.END


async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    df = last_operations(12)
    if df.empty:
        await update.message.reply_text("Записей ещё нет.", reply_markup=main_keyboard())
        return
    lines = ["📋 Последние операции:"]
    for r in df.itertuples():
        sign = "+" if r.kind == "income" else "−"
        comment = f" — {r.comment}" if str(r.comment).strip() not in {"", "nan"} else ""
        lines.append(f"{r.date} {sign}{fmt_money(float(r.amount))} {r.category}{comment}")
    await update.message.reply_text("\n".join(lines), reply_markup=main_keyboard())


async def send_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await ensure_month(update, context)
    if not allowed(update):
        return
    if not EXCEL_PATH.exists():
        await update.message.reply_text("Файл ещё не создан.")
        return
    with open(EXCEL_PATH, "rb") as f:
        await update.message.reply_document(
            document=InputFile(f, filename="finance.xlsx"),
            caption="Excel: Сводка, Отель, Ресторан, Расходы, Нал в банк, Все операции",
        )


def _fmt_shift(dt: datetime) -> str:
    return dt.strftime("%d.%m %H:%M")


async def debt_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return ConversationHandler.END
    rest = sum(x["rest"] for x in debt_balances())
    await update.message.reply_text(
        "Долги только налом. Оборот и прибыль не меняются.\n"
        f"Сейчас должны: {fmt_money(rest)}",
        reply_markup=debt_kb(),
    )
    return ConversationHandler.END


async def debt_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rows = debt_balances()
    if not rows:
        await update.message.reply_text("Долгов нет.", reply_markup=debt_kb())
        return ConversationHandler.END
    lines = ["📋 Долги (нал):"]
    for x in rows:
        if abs(x["rest"]) < 0.01:
            continue
        lines.append(f"{x['name']}: остаток {fmt_money(x['rest'])} (было {fmt_money(x['given'])})")
    if len(lines) == 1:
        lines.append("Все долги погашены.")
    await update.message.reply_text("\n".join(lines), reply_markup=debt_kb())
    return ConversationHandler.END


async def debt_new_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["category_id"] = "debt"
    await update.message.reply_text("Имя должника:", reply_markup=cancel_kb())
    return WAIT_DNAME


async def debt_got_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "отмена" in (update.message.text or "").lower():
        return await cancel(update, context)
    name = (update.message.text or "").strip()
    if len(name) < 2:
        await update.message.reply_text("Напиши имя.")
        return WAIT_DNAME
    context.user_data["debtor"] = name
    await update.message.reply_text(f"{name}. Сумма долга (только нал):")
    return WAIT_DAMOUNT


async def debt_got_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "отмена" in (update.message.text or "").lower():
        return await cancel(update, context)
    amount = parse_amount(update.message.text or "")
    if amount is None:
        await update.message.reply_text("Нужна сумма, например 15000")
        return WAIT_DAMOUNT
    context.user_data["amount"] = amount
    await update.message.reply_text(
        "Дата долга: «сегодня», «вчера» или 24.09.2026"
    )
    return WAIT_DDATE


async def debt_got_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "отмена" in (update.message.text or "").lower():
        return await cancel(update, context)
    when = parse_user_date(update.message.text or "")
    if when is None:
        await update.message.reply_text("Не понял дату. Пример: сегодня или 24.09.2026")
        return WAIT_DDATE
    context.user_data["when"] = when
    await update.message.reply_text("Комментарий (ужин / номер) или «-»:")
    return WAIT_DCOMMENT


async def debt_got_comment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "отмена" in (update.message.text or "").lower():
        return await cancel(update, context)
    comment = (update.message.text or "").strip()
    if comment in {"-", "—"}:
        comment = ""
    user = update.effective_user
    row = add_operation(
        category_id="debt",
        amount=context.user_data["amount"],
        comment=comment,
        user_id=user.id,
        username=user.username or user.full_name,
        when=context.user_data.get("when"),
        payment="cash",
        debtor=context.user_data.get("debtor", ""),
    )
    await update.message.reply_text(
        f"Долг записан: {row['debtor']} {fmt_money(row['amount'])}\n"
        "Оборот и прибыль не изменились.",
        reply_markup=main_keyboard(),
    )
    context.user_data.clear()
    return ConversationHandler.END


async def repay_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    rows = [x for x in debt_balances() if x["rest"] > 0.01]
    if not rows:
        await update.message.reply_text("Открытых долгов нет.", reply_markup=debt_kb())
        return ConversationHandler.END
    names = "\n".join(f"• {x['name']} — {fmt_money(x['rest'])}" for x in rows)
    await update.message.reply_text(
        f"Кто отдаёт долг?\n{names}",
        reply_markup=cancel_kb(),
    )
    return WAIT_RNAME


async def repay_got_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "отмена" in (update.message.text or "").lower():
        return await cancel(update, context)
    name = (update.message.text or "").strip()
    rows = debt_balances()
    match = next((x for x in rows if x["name"].lower() == name.lower()), None)
    if not match:
        match = next((x for x in rows if name.lower() in x["name"].lower()), None)
    if not match:
        await update.message.reply_text("Не нашёл имя. Напиши как в списке.")
        return WAIT_RNAME
    context.user_data["debtor"] = match["name"]
    context.user_data["rest"] = match["rest"]
    await update.message.reply_text(
        f"{match['name']}, остаток {fmt_money(match['rest'])}.\nКакую сумму отдают налом?"
    )
    return WAIT_RAMOUNT


async def repay_got_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "отмена" in (update.message.text or "").lower():
        return await cancel(update, context)
    amount = parse_amount(update.message.text or "")
    rest = float(context.user_data.get("rest") or 0)
    if amount is None:
        await update.message.reply_text("Нужна сумма.")
        return WAIT_RAMOUNT
    if amount - rest > 0.5:
        await update.message.reply_text(f"Больше остатка ({fmt_money(rest)}). Введи меньше.")
        return WAIT_RAMOUNT
    user = update.effective_user
    add_operation(
        category_id="debt_repay",
        amount=amount,
        comment="погашение",
        user_id=user.id,
        username=user.username or user.full_name,
        payment="cash",
        debtor=context.user_data.get("debtor", ""),
    )
    left = rest - amount
    await update.message.reply_text(
        f"Приняли нал {fmt_money(amount)} от {context.user_data.get('debtor')}.\n"
        f"Остаток долга: {fmt_money(left)}\n"
        "Оборот не увеличился, касса выросла.",
        reply_markup=main_keyboard(),
    )
    context.user_data.clear()
    return ConversationHandler.END


async def revenue_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not allowed(update):
        return
    await ensure_month(update, context)
    s = today_revenue()
    await update.message.reply_text(
        "💰 Смена с "
        f"{_fmt_shift(s['start'])}\n"
        f"Остаток с прошлой смены: {fmt_money(s['opening'])}\n"
        f"Оборот смены (без долга): {fmt_money(s['income'])}\n"
        f"Нал расход / в банк: {fmt_money(s['cash_out'])}\n"
        f"Долги за смену (пробиты как нал): {fmt_money(s['debt_new'])}\n"
        f"Вернули долги налом: {fmt_money(s['debt_pay'])}\n\n"
        f"Должно быть в кассе: {fmt_money(s['expected_cash'])}\n"
        f"Минус долг за день: {fmt_money(s['debt_new'])}\n"
        f"На руках у менеджера: {fmt_money(s['on_hand'])}\n\n"
        "«Начать новый день» перенесёт сумму на руках как остаток в следующую смену.",
        reply_markup=revenue_kb(),
    )


async def revenue_new_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    payload = start_new_day()
    await update.message.reply_text(
        f"Смена закрыта.\n"
        f"На руках было {fmt_money(payload['prev_on_hand'])} "
        f"(касса {fmt_money(payload['prev_expected'])} − долг {fmt_money(payload['prev_debt'])}).\n"
        f"Новый день с {_fmt_shift(payload['started'])}.\n"
        f"Стартовый остаток: {fmt_money(payload['opening_cash'])}.",
        reply_markup=main_keyboard(),
    )


async def go_back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Меню", reply_markup=main_keyboard())
    return ConversationHandler.END


async def fallback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.message:
        return
    await update.message.reply_text(
        "Выбери кнопку внизу экрана.",
        reply_markup=main_keyboard(),
    )


def build_app() -> Application:
    if not BOT_TOKEN or BOT_TOKEN == "PASTE_YOUR_TOKEN_HERE":
        raise SystemExit(
            "Укажи токен: export TELEGRAM_BOT_TOKEN='123:ABC' "
            "или впиши его в config.py"
        )
    app = Application.builder().token(BOT_TOKEN).build()

    cat_pattern = "|".join(
        re.escape(v["title"])
        for k, v in CATEGORIES.items()
        if k not in {"debt", "debt_repay", "manual_profit"}
    )
    # кнопки с эмодзи: любое начало + название категории
    cat_re = re.compile(rf"^.*({cat_pattern})$")

    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex(cat_re), pick_category)],
        states={
            WAIT_PAY: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_pay)],
            WAIT_SUPPLIER: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_supplier)],
            WAIT_SUPPLIER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_supplier_name)],
            WAIT_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_amount)],
            WAIT_COMMENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, got_comment)],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            MessageHandler(filters.Regex("^❌ Отмена$"), cancel),
            MessageHandler(filters.Regex("(?i)итоги"), report_month),
            MessageHandler(filters.Regex("(?i)по месяцам"), report_months),
            MessageHandler(filters.Regex("(?i)история"), history),
            MessageHandler(filters.Regex("(?i)скачать excel"), send_excel),
        ],
        allow_reentry=True,
    )
    debt_conv = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Regex("(?i)новый долг"), debt_new_start),
            MessageHandler(filters.Regex("(?i)погашение"), repay_start),
        ],
        states={
            WAIT_DNAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, debt_got_name)],
            WAIT_DAMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, debt_got_amount)],
            WAIT_DDATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, debt_got_date)],
            WAIT_DCOMMENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, debt_got_comment)],
            WAIT_RNAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, repay_got_name)],
            WAIT_RAMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, repay_got_amount)],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            MessageHandler(filters.Regex("(?i)отмена"), cancel),
        ],
        allow_reentry=True,
    )
    app.add_handler(MessageHandler(filters.ALL, auth_gate), group=-1)
    app.add_handler(CommandHandler("logout", logout))
    profit_conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("(?i)чистая прибыль"), profit_start)],
        states={
            WAIT_P_TURN: [MessageHandler(filters.TEXT & ~filters.COMMAND, profit_got_turn)],
            WAIT_P_ACCT: [MessageHandler(filters.TEXT & ~filters.COMMAND, profit_got_acct)],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            MessageHandler(filters.Regex("(?i)отмена"), cancel),
        ],
        allow_reentry=True,
    )
    app.add_handler(profit_conv)
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report_month))
    app.add_handler(MessageHandler(filters.Regex("(?i)итоги"), report_month))
    app.add_handler(MessageHandler(filters.Regex("(?i)по месяцам"), report_months))
    app.add_handler(MessageHandler(filters.Regex("(?i)история"), history))
    app.add_handler(MessageHandler(filters.Regex("(?i)скачать excel"), send_excel))
    app.add_handler(MessageHandler(filters.Regex("(?i)выручка за сегодня"), revenue_menu))
    app.add_handler(MessageHandler(filters.Regex("(?i)начать новый день"), revenue_new_day))
    app.add_handler(debt_conv)
    app.add_handler(MessageHandler(filters.Regex("(?i)список долгов"), debt_list))
    app.add_handler(MessageHandler(filters.Regex(r"(?i)^(📝\s*)?долг$"), debt_menu))
    app.add_handler(MessageHandler(filters.Regex("(?i)назад"), go_back))
    app.add_handler(conv)
    app.add_handler(MessageHandler(filters.COMMAND, start))
    app.add_handler(MessageHandler(filters.TEXT, fallback))
    return app


def main():
    app = build_app()
    log.info("Bot started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
