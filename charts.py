from pathlib import Path
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import numpy as np

from config import CHARTS_DIR, CURRENCY
from storage import month_summary, monthly_series, load_df

_FONT = "/usr/share/fonts/SlidesCarnival/google/Merriweather/static/Merriweather_24pt-Regular.ttf"
_FONT_B = "/usr/share/fonts/SlidesCarnival/google/Merriweather/static/Merriweather_24pt-Bold.ttf"
font_manager.fontManager.addfont(_FONT)
if Path(_FONT_B).exists():
    font_manager.fontManager.addfont(_FONT_B)
_PROP = font_manager.FontProperties(fname=_FONT)
plt.rcParams["font.family"] = _PROP.get_name()
plt.rcParams["axes.titlesize"] = 16
plt.rcParams["axes.labelsize"] = 10


MONTH_RU = [
    "", "Янв", "Фев", "Мар", "Апр", "Май", "Июн",
    "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек",
]
MONTH_FULL = [
    "", "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
]

NAVY = "#1B365D"
LIGHT = "#8FB3E3"


def _fmt(n: float) -> str:
    return f"{n:,.0f}".replace(",", " ")


def pie_expenses(year: int, month: int) -> Path | None:
    s = month_summary(year, month)
    raw = [x for x in s["by_cat"] if x["kind"] == "expense" and x["amount"] > 0]
    if not raw:
        return None
    merged = {}
    for x in raw:
        merged[x["category"]] = merged.get(x["category"], 0) + float(x["amount"])
    labels = list(merged.keys())
    sizes = list(merged.values())

    fig, ax = plt.subplots(figsize=(8, 6), facecolor="white")
    colors = plt.cm.Set2.colors
    wedges, texts, autotexts = ax.pie(
        sizes,
        labels=None,
        autopct=lambda p: f"{p:.1f}%" if p >= 4 else "",
        startangle=90,
        colors=colors[: len(sizes)],
        pctdistance=0.72,
        wedgeprops=dict(width=0.55, edgecolor="white", linewidth=2),
    )
    for t in autotexts:
        t.set_fontsize(9)
        t.set_color("#222")

    ax.legend(
        wedges,
        [f"{l} — {_fmt(v)} {CURRENCY}" for l, v in zip(labels, sizes)],
        title="Статьи трат",
        loc="center left",
        bbox_to_anchor=(1.0, 0.5),
        frameon=False,
    )
    ax.set_title(
        f"Траты за {MONTH_FULL[month]} {year}\n"
        f"расход {_fmt(s['expense'])} {CURRENCY}   ·   чистая прибыль {_fmt(s['profit'])} {CURRENCY}",
        pad=16,
    )
    fig.tight_layout()
    path = CHARTS_DIR / f"pie_{year}_{month:02d}.png"
    fig.savefig(path, dpi=140, bbox_inches="tight")
    plt.close(fig)
    return path


def year_bars(year: int | None = None) -> Path | None:
    """Чистая прибыль по 12 месяцам года в драмах."""
    year = year or datetime.now().year
    now = datetime.now()
    df = load_df()
    profit = [0.0] * 12
    if not df.empty:
        ydf = df[df["year"] == year]
        if not ydf.empty:
            inc = ydf[ydf["kind"] == "income"].groupby("month")["amount"].sum()
            exp = ydf[ydf["kind"] == "expense"].groupby("month")["amount"].sum()
            for m in range(1, 13):
                profit[m - 1] = float(inc.get(m, 0) - exp.get(m, 0))

    if all(v == 0 for v in profit):
        return None

    fig, ax = plt.subplots(figsize=(10.2, 5.6), facecolor="white")
    x = np.arange(12)
    colors = []
    for i, val in enumerate(profit):
        future = year > now.year or (year == now.year and i + 1 > now.month)
        if future:
            colors.append(LIGHT)
        elif val >= 0:
            colors.append(NAVY)
        else:
            colors.append("#8B1E3F")

    ax.bar(x, profit, color=colors, width=0.72, zorder=3)
    ax.axhline(0, color="#999", linewidth=0.8, zorder=2)

    from matplotlib.patches import Patch
    ax.legend(
        handles=[
            Patch(facecolor=NAVY, label="ПРИБЫЛЬ"),
            Patch(facecolor="#8B1E3F", label="УБЫТОК"),
            Patch(facecolor=LIGHT, label="ВПЕРЕДИ"),
        ],
        loc="upper right",
        frameon=False,
        ncol=3,
    )

    year_total = sum(
        v for i, v in enumerate(profit)
        if year < now.year or (year == now.year and i + 1 <= now.month)
    )
    ax.set_title(
        f"Чистая прибыль {year}  ·  {_fmt(year_total)} {CURRENCY}",
        loc="left",
        fontweight="bold",
        color="#222",
        pad=12,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(MONTH_RU[1:])
    ax.set_ylabel("драм")
    ax.yaxis.grid(True, linestyle="--", color="#D0D5DD", zorder=0)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CCCCCC")
    ax.spines["bottom"].set_color("#CCCCCC")

    span = max(abs(min(profit)), abs(max(profit))) or 1
    ax.set_ylim(-span * 1.28, span * 1.28)
    for i, val in enumerate(profit):
        if val == 0:
            continue
        offset = span * 0.04 if val >= 0 else -span * 0.04
        ax.text(
            i,
            val + offset,
            _fmt(val),
            ha="center",
            va="bottom" if val >= 0 else "top",
            fontsize=8,
            color="#333",
        )

    fig.tight_layout()
    path = CHARTS_DIR / f"year_{year}.png"
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def months_compare(limit: int = 12) -> Path | None:
    # оставляем как запасной график доход/расход
    g = monthly_series(limit)
    if g.empty:
        return None

    labels = [f"{MONTH_RU[int(r.month)]}\n{int(r.year)}" for r in g.itertuples()]
    x = range(len(g))
    width = 0.36

    fig, ax = plt.subplots(figsize=(10, 5.4), facecolor="white")
    ax.bar([i - width / 2 for i in x], g["income"], width, label="Доход", color="#2E8B57", zorder=3)
    ax.bar([i + width / 2 for i in x], g["expense"], width, label="Расход", color="#C0392B", zorder=3)
    ax.plot(list(x), g["profit"], color="#1F4E79", marker="o", linewidth=2, label="Прибыль")
    ax.axhline(0, color="#999", linewidth=0.8)
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels)
    ax.set_ylabel(CURRENCY)
    ax.set_title("Доходы и расходы по месяцам")
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    path = CHARTS_DIR / "months_compare.png"
    fig.savefig(path, dpi=140, bbox_inches="tight")
    plt.close(fig)
    return path
