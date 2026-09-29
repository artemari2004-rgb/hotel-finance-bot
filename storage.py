import json
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter

from config import CATEGORIES, EXCEL_PATH, DATA_DIR

SHIFT_PATH = DATA_DIR / "shift.json"
SUPPLIERS_PATH = DATA_DIR / "extra_suppliers.json"


def extra_card_suppliers() -> list[str]:
    if not SUPPLIERS_PATH.exists():
        return []
    try:
        data = json.loads(SUPPLIERS_PATH.read_text(encoding="utf-8"))
        return [str(x).strip() for x in data.get("card", []) if str(x).strip()]
    except Exception:
        return []


def add_extra_card_supplier(name: str) -> bool:
    name = (name or "").strip()
    if not name:
        return False
    from config import CARD_EXPENSE_ITEMS
    have = extra_card_suppliers()
    low = {x.lower() for x in CARD_EXPENSE_ITEMS + have}
    if name.lower() in low:
        return False
    have.append(name)
    SUPPLIERS_PATH.write_text(
        json.dumps({"card": have}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return True


def all_card_suppliers() -> list[str]:
    from config import CARD_EXPENSE_ITEMS
    return list(CARD_EXPENSE_ITEMS) + extra_card_suppliers()


def prev_manual_remainder(when: datetime | None = None) -> float:
    when = when or datetime.now()
    y, m = when.year, when.month
    if m == 1:
        y, m = y - 1, 12
    else:
        m -= 1
    df = load_df()
    if df.empty:
        return 0.0
    part = df[(df["kind"] == "manual_profit") & (df["year"] == y) & (df["month"] == m)]
    if part.empty:
        return 0.0
    part = part.sort_values(["datetime", "id"])
    return float(part.iloc[-1]["amount"])


SHEET_ALL = "Все операции"
SHEET_HOTEL = "Отель"
SHEET_RESTO = "Ресторан"
SHEET_EXP_CASH = "Расходы нал"
SHEET_EXP_CARD = "Расходы безнал"
SHEET_TRANSFER = "Нал в банк"
SHEET_DEBT = "Долги"
SHEET_PERSONAL = "Траты безнал"
SHEET_PROFIT = "Чистая прибыль"
SHEET_SUM = "Сводка"

COLUMNS = [
    "id",
    "datetime",
    "date",
    "year",
    "month",
    "category_id",
    "category",
    "kind",
    "payment",
    "amount",
    "comment",
    "debtor",
    "user_id",
    "username",
]

PAY_RU = {"cash": "нал", "card": "безнал", "mixed": "смешанно", "transfer": "нал → банк"}

MONTH_NAME = [
    "", "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
]


def _month_word(m) -> str:
    try:
        i = int(m)
    except (TypeError, ValueError):
        return str(m)
    if 1 <= i <= 12:
        return MONTH_NAME[i]
    return str(m)


_VIEW_YM: tuple[int, int] | None = None


def get_active_period() -> tuple[int, int]:
    raw = {}
    path = DATA_DIR / "period.json"
    if path.exists():
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            raw = {}
    now = datetime.now()
    if not raw.get("locked"):
        return now.year, now.month
    y = int(raw.get("year") or now.year)
    m = int(raw.get("month") or now.month)
    return y, m


def save_period(year: int, month: int, **extra):
    path = DATA_DIR / "period.json"
    data = {}
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    data.update(extra)
    data["year"] = int(year)
    data["month"] = int(month)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def remember_chat(chat_id: int | None):
    if not chat_id:
        return
    y, m = get_active_period()
    save_period(y, m, chat_id=int(chat_id))


def _ops_month(df: pd.DataFrame) -> pd.DataFrame:
    """Рабочие листы — только активный месяц. История живёт в «Все операции»."""
    if df.empty:
        return df
    if _VIEW_YM:
        year, month = _VIEW_YM
    else:
        year, month = get_active_period()
    return df[(df["year"] == year) & (df["month"] == month)].copy()

HEADER_FILL = {
    SHEET_ALL: "1F4E79",
    SHEET_HOTEL: "1B4F72",
    SHEET_RESTO: "196F3D",
    SHEET_EXP_CASH: "922B21",
    SHEET_EXP_CARD: "7B241C",
    SHEET_TRANSFER: "1A5276",
    SHEET_DEBT: "7D6608",
    SHEET_PERSONAL: "5D6D7E",
    SHEET_PROFIT: "117A65",
    SHEET_SUM: "6C3483",
}


def _border():
    return Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )


def _style_header(ws, color="1F4E79"):
    fill = PatternFill("solid", fgColor=color)
    font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _border()
    ws.freeze_panes = "A2"
    if ws.max_column >= 1 and ws.max_row >= 1:
        ws.auto_filter.ref = ws.dimensions
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = min(28, max(12, len(str(ws.cell(1, col).value or "")) + 6))


def _empty_df() -> pd.DataFrame:
    return pd.DataFrame(columns=COLUMNS)


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return _empty_df()
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    df = df[COLUMNS].copy()
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    df["year"] = pd.to_numeric(df["year"], errors="coerce").fillna(0).astype(int)
    month_map = {name: i for i, name in enumerate(MONTH_NAME) if name}
    month_map.update({name[:3]: i for i, name in enumerate(MONTH_NAME) if name})
    def _to_month(v):
        if pd.isna(v):
            return 0
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return int(v)
        s = str(v).strip().lower()
        if s in month_map:
            return month_map[s]
        try:
            return int(float(s))
        except ValueError:
            return 0
    df["month"] = df["month"].map(_to_month).fillna(0).astype(int)
    if "date" in df.columns:
        dates = pd.to_datetime(df["date"], errors="coerce")
        miss_m = (df["month"] < 1) & dates.notna()
        df.loc[miss_m, "month"] = dates.loc[miss_m].dt.month.astype(int)
        miss_y = (df["year"] < 1) & dates.notna()
        df.loc[miss_y, "year"] = dates.loc[miss_y].dt.year.astype(int)
    return df


def _read_legacy(path) -> pd.DataFrame:
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return _empty_df()
    name = SHEET_ALL if SHEET_ALL in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=name)
    return _normalize(df)


def init_excel():
    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    if EXCEL_PATH.exists():
        try:
            wb = load_workbook(EXCEL_PATH)
            if SHEET_ALL in wb.sheetnames:
                return
        except Exception:
            pass
        df = _read_legacy(EXCEL_PATH)
        _write_workbook(df)
        return
    _write_workbook(_empty_df())


RU_COL = {
    "№": "id",
    "Дата и время": "datetime",
    "Дата": "date",
    "Год": "year",
    "Месяц": "month",
    "Статья": "category",
    "Тип": "kind",
    "Нал": "cash_amt",
    "Безнал": "card_amt",
    "Сумма": "amount",
    "Комментарий": "comment",
    "Должник": "debtor",
    "Код статьи": "category_id",
    "Оплата": "payment",
    "ID пользователя": "user_id",
    "Кто внёс": "username",
}

KIND_RU = {
    "income": "доход",
    "expense": "расход",
    "transfer": "инкассация",
    "debt": "долг",
    "debt_repay": "погашение долга",
    "personal": "лично",
    "manual_profit": "чистая прибыль",
    "reclass": "перенос",
}
KIND_EN = {
    "доход": "income",
    "расход": "expense",
    "инкассация": "transfer",
    "долг": "debt",
    "погашение долга": "debt_repay",
    "income": "income",
    "debt": "debt",
    "debt_repay": "debt_repay",
    "expense": "expense",
    "transfer": "transfer",
    "лично": "personal",
    "personal": "personal",
    "перенос": "reclass",
    "reclass": "reclass",
}
PAY_EN = {
    "нал": "cash",
    "безнал": "card",
    "нал → банк": "transfer",
    "cash": "cash",
    "card": "card",
    "mixed": "mixed",
    "смешанно": "mixed",
    "transfer": "transfer",
}


def load_df() -> pd.DataFrame:
    init_excel()
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_ALL)
    except ValueError:
        df = _read_legacy(EXCEL_PATH)
    if df is None or df.empty:
        return _empty_df()
    df = df.rename(columns={k: v for k, v in RU_COL.items() if k in df.columns})
    first = str(df.columns[0])
    key = "id" if "id" in df.columns else first
    df = df[~df[key].astype(str).str.contains("оборот|итого", case=False, na=False)]
    if "kind" in df.columns:
        df["kind"] = df["kind"].map(lambda x: KIND_EN.get(str(x).strip().lower(), x))
    if "payment" in df.columns:
        df["payment"] = df["payment"].map(lambda x: PAY_EN.get(str(x).strip().lower(), x))
    return _normalize(df)


def _next_id(df: pd.DataFrame) -> int:
    if df.empty:
        return 1
    return int(pd.to_numeric(df["id"], errors="coerce").fillna(0).max()) + 1


def _append_totals(ws, headers, label, money_names):
    if ws.max_row < 2:
        return
    total = [""] * len(headers)
    total[0] = label
    name_to_i = {h: i for i, h in enumerate(headers)}
    for name in money_names:
        if name not in name_to_i:
            continue
        idx = name_to_i[name]
        s = 0.0
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, idx + 1).value
            if isinstance(v, (int, float)):
                s += float(v)
        total[idx] = s
    ws.append(total)
    last = ws.max_row
    fill = PatternFill("solid", fgColor="F4E6C3")
    font = Font(bold=True)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(last, c)
        cell.fill = fill
        cell.font = font
        if isinstance(cell.value, (int, float)):
            cell.number_format = "#,##0"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last - 1}"


def _emphasize_summary(ws):
    if ws.max_row < 2:
        return
    gold = PatternFill("solid", fgColor="1B4F72")
    green = PatternFill("solid", fgColor="145A32")
    gold_cell = PatternFill("solid", fgColor="D4E6F1")
    green_cell = PatternFill("solid", fgColor="D5F5E3")
    white = Font(bold=True, color="FFFFFF", size=12)
    big = Font(bold=True, size=12, color="1B4F72")
    big_g = Font(bold=True, size=12, color="145A32")
    # Месяц | Оборот | Прибыль
    for col, fill, font in ((2, gold, white), (3, green, white)):
        cell = ws.cell(1, col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col)].width = 20
    for r in range(2, ws.max_row + 1):
        c2 = ws.cell(r, 2)
        c3 = ws.cell(r, 3)
        c2.fill = gold_cell
        c3.fill = green_cell
        c2.font = big
        c3.font = big_g
        c2.number_format = "#,##0"
        c3.number_format = "#,##0"
        c2.alignment = Alignment(horizontal="center")
        c3.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 22


def _write_sheet(wb, title, headers, rows, color, total_label=None, money_names=None):
    if title in wb.sheetnames:
        ws = wb[title]
        wb.remove(ws)
    ws = wb.create_sheet(title)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    _style_header(ws, color)
    # amount column highlight
    for r in range(2, ws.max_row + 1):
        for c, h in enumerate(headers, 1):
            if h in ("Сумма", "Сумма нал", "Сумма безнал", "amount", "Нал", "Безнал", "Всего", "Доход всего", "Расход всего", "Чистая прибыль", "Общий оборот") or (
                isinstance(h, str) and ("нал" in h.lower() or "безнал" in h.lower() or "всего" in h.lower() or "прибыль" in h.lower())
            ):
                cell = ws.cell(r, c)
                cell.number_format = "#,##0"
                if h == "Нал":
                    cell.font = Font(color="1B7A3A")
                elif h == "Безнал":
                    cell.font = Font(color="1F4E79")
    if total_label and money_names:
        _append_totals(ws, headers, total_label, money_names)
    return ws


def _cash_card(payment, amount):
    amount = float(amount or 0)
    if payment == "card":
        return 0.0, amount
    if payment == "cash":
        return amount, 0.0
    return amount, 0.0


def _dept_rows(df: pd.DataFrame, category_ids: list[str]):
    part = _ops_month(df)
    part = part[part["category_id"].isin(category_ids)].sort_values(["date", "id"])
    rows = []
    for r in part.itertuples():
        cash, card = _cash_card(r.payment, r.amount)
        rows.append(
            [
                r.date,
                _month_word(r.month),
                cash or None,
                card or None,
                float(r.amount),
                r.comment or "",
            ]
        )
    return rows


def _breakfast_rows(df: pd.DataFrame, sign: int):
    part = _ops_month(df[df["kind"] == "reclass"]).sort_values(["date", "id"])
    rows = []
    for r in part.itertuples():
        amt = sign * float(r.amount)
        rows.append(
            [
                r.date,
                _month_word(r.month),
                None,
                None,
                amt,
                "завтраки из проживания",
            ]
        )
    return rows


def _expense_rows(df: pd.DataFrame, payment: str):
    if payment == "card":
        mask = (df["kind"] == "expense") & (df["payment"] == "card")
    else:
        mask = (df["kind"] == "expense") & (df["payment"] != "card")
    part = _ops_month(df.loc[mask]).sort_values(["date", "id"])
    rows = []
    for r in part.itertuples():
        rows.append(
            [
                r.date,
                _month_word(r.month),
                r.category,
                float(r.amount),
                r.comment or "",
            ]
        )
    return rows


def _summary_rows(df: pd.DataFrame):
    if df.empty:
        return []
    rows = []
    cur = _ops_month(df)
    if cur.empty:
        return []
    grouped = cur.groupby(["year", "month"], dropna=False)
    for (year, month), g in grouped:
        hotel = g[g["category_id"] == "hotel_income"]
        resto = g[g["category_id"] == "resto_income"]
        other_i = g[(g["kind"] == "income") & (~g["category_id"].isin(["hotel_income", "resto_income"]))]
        exp = g[g["kind"] == "expense"]
        tr = g[g["kind"] == "transfer"]

        def split(part):
            cash = float(part.loc[part["payment"] == "cash", "amount"].sum())
            card = float(part.loc[part["payment"] == "card", "amount"].sum())
            total = float(part["amount"].sum())
            return cash, card, total

        h_c, h_k, h_t = split(hotel)
        r_c, r_k, r_t = split(resto)
        o_c, o_k, o_t = split(other_i)
        e_c, e_k, e_t = split(exp)
        bf = float(g.loc[g["kind"] == "reclass", "amount"].sum())
        h_t -= bf
        r_t += bf
        income = h_t + r_t + o_t
        incash = float(g.loc[(g["kind"] == "income") & (g["payment"] == "cash"), "amount"].sum())
        incard = float(g.loc[(g["kind"] == "income") & (g["payment"] == "card"), "amount"].sum())
        moved = float(tr["amount"].sum())
        repaid = float(g.loc[g["kind"] == "debt_repay", "amount"].sum())
        kassa = incash - e_c - moved + repaid
        bank = incard - e_k + moved
        rows.append(
            [
                _month_word(month),
                income,
                income - e_t,
                moved,
                kassa,
                bank,
                h_c, h_k, h_t,
                r_c, r_k, r_t,
                o_c, o_k, o_t,
                e_c, e_k, e_t,
            ]
        )
    return rows


def _write_workbook(df: pd.DataFrame, path=None):
    df = _normalize(df)
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    all_rows = []
    for r in df.itertuples(index=False):
        rec = dict(zip(COLUMNS, r))
        rec["payment"] = PAY_RU.get(str(rec["payment"]), rec["payment"])
        # store raw payment in operations as ru for readability? keep code in source
        all_rows.append([rec[c] if c != "payment" else rec["payment"] for c in COLUMNS])

    # store payment as cash/card in source for logic — rewrite properly
    src_rows = []
    for r in df.itertuples(index=False):
        rec = dict(zip(COLUMNS, r))
        if rec["kind"] == "manual_profit":
            continue
        cash, card = _cash_card(rec["payment"], rec["amount"])
        if rec["kind"] == "debt":
            cash, card = None, None
        elif rec["kind"] == "debt_repay":
            cash, card = float(rec["amount"]), None
        src_rows.append(
            [
                rec["id"],
                rec["datetime"],
                rec["date"],
                rec["year"],
                _month_word(rec["month"]),
                rec["category"],
                KIND_RU.get(str(rec["kind"]), rec["kind"]),
                cash or None,
                card or None,
                rec["amount"],
                rec["comment"],
                rec.get("debtor", ""),
                rec["category_id"],
                PAY_RU.get(str(rec["payment"]), rec["payment"]),
                rec["user_id"],
                rec["username"],
            ]
        )

    _write_sheet(
        wb,
        SHEET_ALL,
        [
            "№", "Дата и время", "Дата", "Год", "Месяц",
            "Статья", "Тип", "Нал", "Безнал", "Сумма",
            "Комментарий", "Должник", "Код статьи", "Оплата", "ID пользователя", "Кто внёс",
        ],
        src_rows,
        HEADER_FILL[SHEET_ALL],
        total_label="Общий оборот",
        money_names=["Нал", "Безнал", "Сумма"],
    )
    _write_sheet(
        wb,
        SHEET_HOTEL,
        ["Дата", "Месяц", "Нал", "Безнал", "Всего", "Комментарий"],
        _dept_rows(df, ["hotel_income"]) + _breakfast_rows(df, -1),
        HEADER_FILL[SHEET_HOTEL],
        total_label="Общий оборот",
        money_names=["Нал", "Безнал", "Всего"],
    )
    _write_sheet(
        wb,
        SHEET_RESTO,
        ["Дата", "Месяц", "Нал", "Безнал", "Всего", "Комментарий"],
        _dept_rows(df, ["resto_income"]) + _breakfast_rows(df, 1),
        HEADER_FILL[SHEET_RESTO],
        total_label="Общий оборот",
        money_names=["Нал", "Безнал", "Всего"],
    )
    _write_sheet(
        wb,
        SHEET_EXP_CASH,
        ["Дата", "Месяц", "Статья", "Сумма нал", "Комментарий"],
        _expense_rows(df, "cash"),
        HEADER_FILL[SHEET_EXP_CASH],
        total_label="Итого расходов нал",
        money_names=["Сумма нал"],
    )
    _write_sheet(
        wb,
        SHEET_EXP_CARD,
        ["Дата", "Месяц", "Статья", "Сумма безнал", "Комментарий"],
        _expense_rows(df, "card"),
        HEADER_FILL[SHEET_EXP_CARD],
        total_label="Итого расходов безнал",
        money_names=["Сумма безнал"],
    )
    tr_part = _ops_month(df[df["kind"] == "transfer"]).sort_values(["date", "id"])
    tr_rows = [
        [r.date, _month_word(r.month), float(r.amount), r.comment or ""]
        for r in tr_part.itertuples()
    ]
    _write_sheet(
        wb,
        SHEET_TRANSFER,
        ["Дата", "Месяц", "Сумма", "Комментарий"],
        tr_rows,
        HEADER_FILL[SHEET_TRANSFER],
        total_label="Итого сдано в банк",
        money_names=["Сумма"],
    )
    _write_sheet(
        wb,
        SHEET_DEBT,
        ["Дата", "Имя", "Начислено", "Погашено", "Комментарий"],
        _debt_journal_rows(_ops_month(df)),
        HEADER_FILL[SHEET_DEBT],
        total_label="Итого",
        money_names=["Начислено", "Погашено"],
    )
    pers = _ops_month(df[df["kind"] == "personal"]).sort_values(["date", "id"])
    pers_rows = [
        [r.date, _month_word(r.month), float(r.amount), r.comment or ""]
        for r in pers.itertuples()
    ]
    _write_sheet(
        wb,
        SHEET_PERSONAL,
        ["Дата", "Месяц", "Сумма безнал", "Комментарий"],
        pers_rows,
        HEADER_FILL[SHEET_PERSONAL],
        total_label="Итого лично (не в обороте)",
        money_names=["Сумма безнал"],
    )
    prof_rows = []
    for r in _ops_month(df[df["kind"] == "manual_profit"]).sort_values(["date", "id"]).itertuples():
        turn, acct, prev = 0.0, 0.0, 0.0
        parts = str(r.comment or "").split("|")
        try:
            if len(parts) >= 1 and parts[0] != "":
                turn = float(parts[0])
            if len(parts) >= 2 and parts[1] != "":
                acct = float(parts[1])
            if len(parts) >= 3 and parts[2] != "":
                prev = float(parts[2])
        except ValueError:
            turn = float(r.amount)
        prof_rows.append([r.date, _month_word(r.month), turn, acct, prev, float(r.amount)])
    _write_sheet(
        wb,
        SHEET_PROFIT,
        ["Дата", "Месяц", "Оборот ресторана", "Сумма бухгалтера", "Остаток прошлого месяца", "Чистая прибыль"],
        prof_rows,
        HEADER_FILL[SHEET_PROFIT],
        total_label="Итого",
        money_names=["Оборот ресторана", "Сумма бухгалтера", "Остаток прошлого месяца", "Чистая прибыль"],
    )
    sum_headers = [
            "Месяц",
            "Общий оборот", "Чистая прибыль",
            "Нал в банк", "Касса (нал)", "Счёт (безнал)",
            "Отель нал", "Отель безнал", "Отель всего",
            "Ресторан нал", "Ресторан безнал", "Ресторан всего",
            "Прочий доход нал", "Прочий доход безнал", "Прочий доход всего",
            "Расходы нал", "Расходы безнал", "Расходы всего",
        ]
    ws_sum = _write_sheet(
        wb,
        SHEET_SUM,
        sum_headers,
        _summary_rows(df),
        HEADER_FILL[SHEET_SUM],
        total_label="ИТОГО",
        money_names=sum_headers[1:],
    )
    _emphasize_summary(ws_sum)
    # keep Сводка first after all? user wants overview visible — put Сводка first
    order = [SHEET_SUM, SHEET_HOTEL, SHEET_RESTO, SHEET_EXP_CASH, SHEET_EXP_CARD, SHEET_TRANSFER, SHEET_DEBT, SHEET_PERSONAL, SHEET_PROFIT, SHEET_ALL]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))
    wb.save(path or EXCEL_PATH)


def add_operation(
    category_id: str,
    amount: float,
    comment: str,
    user_id: int,
    username: str,
    when: datetime | None = None,
    payment: str | None = None,
    debtor: str = "",
    category_title: str | None = None,
) -> dict:
    init_excel()
    cat = CATEGORIES[category_id]
    when = when or datetime.now()
    df = load_df()
    pay = payment or cat.get("payment") or "mixed"
    if pay not in {"cash", "card", "transfer", "none"}:
        pay = "cash"
    row = {
        "id": _next_id(df),
        "datetime": when.strftime("%Y-%m-%d %H:%M:%S"),
        "date": when.date().isoformat(),
        "year": when.year,
        "month": when.month,
        "category_id": category_id,
        "category": (category_title or cat["title"]).strip() or cat["title"],
        "kind": cat["kind"],
        "payment": pay,
        "amount": round(float(amount), 2),
        "comment": comment or "",
        "debtor": (debtor or "").strip(),
        "user_id": user_id,
        "username": username or "",
    }
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    _write_workbook(df)
    save_period(int(row["year"]), int(row["month"]), locked=True, opening_cash=get_opening_cash())
    return row


def month_summary(year: int, month: int) -> dict:
    df = load_df()
    if df.empty:
        return {"income": 0, "expense": 0, "profit": 0, "by_cat": [], "rows": 0}
    mask = (df["year"] == year) & (df["month"] == month)
    m = df.loc[mask]
    income = float(m.loc[m["kind"] == "income", "amount"].sum())
    expense = float(m.loc[m["kind"] == "expense", "amount"].sum())
    official = m[m["kind"].isin(["income", "expense", "transfer", "debt", "debt_repay"])]
    by_cat = (
        official[official["kind"].isin(["income", "expense"])]
        .groupby(["category", "kind"], as_index=False)["amount"]
        .sum()
        .to_dict("records")
    )
    return {
        "income": income,
        "expense": expense,
        "profit": income - expense,
        "by_cat": by_cat,
        "rows": len(official),
    }


def monthly_series(limit_months: int = 12) -> pd.DataFrame:
    df = load_df()
    if df.empty:
        return pd.DataFrame(columns=["year", "month", "income", "expense", "profit"])
    g = (
        df.groupby(["year", "month", "kind"])["amount"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )
    if "income" not in g.columns:
        g["income"] = 0
    if "expense" not in g.columns:
        g["expense"] = 0
    g["profit"] = g["income"] - g["expense"]
    return g.sort_values(["year", "month"]).tail(limit_months)


def last_operations(n: int = 10) -> pd.DataFrame:
    df = load_df()
    if df.empty:
        return df
    return df.sort_values("id", ascending=False).head(n)


def _debt_journal_rows(df: pd.DataFrame):
    part = df[df["kind"].isin(["debt", "debt_repay"])].copy()
    if part.empty:
        return []
    if "debtor" not in part.columns:
        part["debtor"] = ""
    part["debtor"] = part["debtor"].fillna("").astype(str).str.strip().replace("", "без имени")
    part = part.sort_values(["date", "id"])
    rows = []
    for r in part.itertuples():
        given = float(r.amount) if r.kind == "debt" else None
        paid = float(r.amount) if r.kind == "debt_repay" else None
        rows.append([r.date, r.debtor, given, paid, r.comment or ""])
    return rows


def _debt_balance_rows(df: pd.DataFrame):
    part = df[df["kind"].isin(["debt", "debt_repay"])]
    if part.empty:
        return []
    names = part["debtor"].fillna("").astype(str).str.strip()
    part = part.copy()
    part["debtor"] = names.replace("", "без имени")
    rows = []
    for name, g in part.groupby("debtor"):
        given = float(g.loc[g["kind"] == "debt", "amount"].sum())
        paid = float(g.loc[g["kind"] == "debt_repay", "amount"].sum())
        rows.append([name, given, paid, given - paid])
    rows.sort(key=lambda x: (-x[3], str(x[0])))
    return rows


def debt_balances() -> list[dict]:
    df = load_df()
    rows = _debt_balance_rows(df)
    return [
        {"name": n, "given": g, "paid": p, "rest": r}
        for n, g, p, r in rows
    ]


def _shift_data() -> dict:
    if SHIFT_PATH.exists():
        try:
            return json.loads(SHIFT_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def get_shift_start() -> datetime:
    raw = _shift_data()
    if raw.get("started_at"):
        try:
            return datetime.fromisoformat(raw["started_at"])
        except Exception:
            pass
    now = datetime.now()
    start = now.replace(hour=8, minute=0, second=0, microsecond=0)
    if now < start:
        start = start - timedelta(days=1)
    return start


def get_opening_cash() -> float:
    try:
        return float(_shift_data().get("opening_cash") or 0)
    except (TypeError, ValueError):
        return 0.0


def start_new_day() -> dict:
    snap = today_revenue()
    now = datetime.now()
    DATA_DIR.mkdir(exist_ok=True)
    payload = {
        "started_at": now.isoformat(),
        "opening_cash": round(float(snap["on_hand"]), 2),
        "prev_on_hand": round(float(snap["on_hand"]), 2),
        "prev_expected": round(float(snap["expected_cash"]), 2),
        "prev_debt": round(float(snap["debt_new"]), 2),
    }
    SHIFT_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    payload["started"] = now
    return payload


def parse_user_date(text: str) -> datetime | None:
    t = (text or "").strip().lower()
    now = datetime.now()
    if t in {"сегодня", "today", "-", "сейчас"}:
        return now
    if t in {"вчера"}:
        return now - timedelta(days=1)
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m"):
        try:
            dt = datetime.strptime(t, fmt)
            if fmt == "%d.%m":
                dt = dt.replace(year=now.year)
            return dt.replace(hour=now.hour, minute=now.minute, second=0)
        except ValueError:
            continue
    return None


def today_revenue() -> dict:
    df = load_df()
    start = get_shift_start()
    opening = get_opening_cash()
    empty = {
        "start": start,
        "opening": opening,
        "income": 0,
        "expense": 0,
        "profit": 0,
        "cash_in": 0,
        "cash_out": 0,
        "debt_new": 0,
        "debt_pay": 0,
        "punched_cash": opening,
        "expected_cash": opening,
        "on_hand": opening,
        "rows": 0,
    }
    if df.empty:
        return empty
    dt = pd.to_datetime(df["datetime"], errors="coerce")
    m = df.loc[dt >= start]
    income = float(m.loc[m["kind"] == "income", "amount"].sum())
    expense = float(m.loc[m["kind"] == "expense", "amount"].sum())
    cash_sales = float(m.loc[(m["kind"] == "income") & (m["payment"] == "cash"), "amount"].sum())
    debt_new = float(m.loc[m["kind"] == "debt", "amount"].sum())
    debt_pay = float(m.loc[m["kind"] == "debt_repay", "amount"].sum())
    cash_out = float(m.loc[(m["kind"] == "expense") & (m["payment"] == "cash"), "amount"].sum())
    cash_out += float(m.loc[m["kind"] == "transfer", "amount"].sum())
    # долг пробивается как нал в кассе смены
    punched_cash = opening + cash_sales + debt_new + debt_pay - cash_out
    on_hand = punched_cash - debt_new
    return {
        "start": start,
        "opening": opening,
        "income": income,
        "expense": expense,
        "profit": income - expense,
        "cash_in": cash_sales + debt_pay,
        "cash_out": cash_out,
        "debt_new": debt_new,
        "debt_pay": debt_pay,
        "punched_cash": punched_cash,
        "expected_cash": punched_cash,
        "on_hand": on_hand,
        "rows": int(len(m)),
    }


def month_on_hand(year: int, month: int) -> float:
    raw = {}
    path = DATA_DIR / "period.json"
    if path.exists():
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            raw = {}
    opening = float(raw.get("opening_cash") or get_opening_cash() or 0)
    df = load_df()
    if df.empty:
        return opening
    m = df[(df["year"] == year) & (df["month"] == month)]
    cash_sales = float(m.loc[(m["kind"] == "income") & (m["payment"] == "cash"), "amount"].sum())
    debt_pay = float(m.loc[m["kind"] == "debt_repay", "amount"].sum())
    cash_out = float(m.loc[(m["kind"] == "expense") & (m["payment"] == "cash"), "amount"].sum())
    cash_out += float(m.loc[m["kind"] == "transfer", "amount"].sum())
    return opening + cash_sales + debt_pay - cash_out


def write_month_archive(year: int, month: int):
    global _VIEW_YM
    folder = DATA_DIR / "archive"
    folder.mkdir(exist_ok=True)
    dest = folder / f"finance_{year}_{month:02d}.xlsx"
    _VIEW_YM = (year, month)
    try:
        _write_workbook(load_df(), dest)
    finally:
        _VIEW_YM = None
    return dest


def roll_if_needed(chat_id: int | None = None) -> dict | None:
    """Если календарь ушёл в новый месяц — архивируем прошлый и открываем новый с остатком."""
    remember_chat(chat_id)
    now = datetime.now()
    y, m = get_active_period()
    if (now.year, now.month) <= (y, m):
        return None
    remainder = month_on_hand(y, m)
    archive = write_month_archive(y, m)
    save_period(
        now.year,
        now.month,
        opening_cash=round(remainder, 2),
        prev_month=m,
        prev_year=y,
        chat_id=chat_id or None,
    )
    SHIFT_PATH.write_text(
        json.dumps(
            {
                "started_at": now.isoformat(),
                "opening_cash": round(remainder, 2),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    _write_workbook(load_df())
    return {
        "archive": archive,
        "closed_year": y,
        "closed_month": m,
        "opening": remainder,
        "new_year": now.year,
        "new_month": now.month,
    }
