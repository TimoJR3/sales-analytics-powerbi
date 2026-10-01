"""Independent recalculation of the report metrics in pandas.

Repeats the Power Query cleaning (numbers with spaces/NBSP, empty rows, duplicate
reference rows) and computes every measure from the task, so the DAX model can be
checked against a second implementation.

Usage: python scripts/verify_metrics.py [path/to/source.xlsx]
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

DEFAULT_SOURCE = Path(__file__).resolve().parents[1] / "data" / "source.xlsx"
SPACES = re.compile(r"[\s  ]")


def to_number(value) -> float | None:
    """Power Query ЧислоИз: drop spaces/NBSP, parse with en-US decimal point."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = SPACES.sub("", str(value))
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_text(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    return str(value).strip()


def load_tables(source: Path) -> dict[str, pd.DataFrame]:
    sheets = pd.read_excel(source, sheet_name=None, dtype=object)
    return {name.strip(): frame for name, frame in sheets.items()}


def count_sheet_rows(source: Path, sheet: str) -> int:
    """Data rows in the sheet including empty ones (pandas skips trailing blanks)."""
    from openpyxl import load_workbook

    worksheet = load_workbook(source, read_only=True)[sheet]
    return sum(1 for _ in worksheet.iter_rows(min_row=2, values_only=True))


def conflicting_duplicates(raw: pd.DataFrame, key: str, attribute: str) -> pd.DataFrame:
    """Keys that appear several times with different attribute values."""
    frame = raw.assign(**{key: raw[key].map(to_text), attribute: raw[attribute].map(to_text)})
    normalized = frame[key].str.lower()
    repeated = frame[normalized.duplicated(keep=False)]
    return (
        repeated.groupby(normalized[repeated.index].rename(key))[attribute]
        .agg(["count", "nunique"])
        .query("nunique > 1")
        .rename(columns={"count": "строк", "nunique": "разных значений"})
    )


def clean_sales(raw: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    sales = raw.copy()
    for column in ["Order_ID", "Количество", "Выручка", "Прибыль"]:
        sales[column] = sales[column].map(to_number)
    for column in ["Приоритет", "Метод доставки", "Покупатель", "Регион", "Продукт"]:
        sales[column] = sales[column].map(to_text)
    kept = sales[sales["Order_ID"].notna() & sales["Выручка"].notna()].copy()
    kept["Order_ID"] = kept["Order_ID"].astype("int64")
    return kept, len(sales) - len(kept)


def dedupe_reference(raw: pd.DataFrame, key: str) -> tuple[pd.DataFrame, int]:
    """Table.Distinct with Comparer.OrdinalIgnoreCase on the key column."""
    frame = raw.copy()
    for column in frame.columns:
        frame[column] = frame[column].map(to_text)
    frame = frame[frame[key].notna() & (frame[key] != "")]
    lowered = frame[key].str.lower()
    unique = frame[~lowered.duplicated()]
    return unique, len(frame) - len(unique)


def core_metrics(sales: pd.DataFrame, shipping: pd.DataFrame) -> dict[str, float]:
    cost = dict(
        zip(shipping["Метод доставки"], shipping["Стоимость доставки"].map(to_number), strict=True)
    )
    revenue = sales["Выручка"].sum()
    profit = sales["Прибыль"].sum()
    orders = sales["Order_ID"].nunique()
    logistics = sales["Метод доставки"].map(cost).fillna(0).sum()
    return {
        "Выручка": revenue,
        "Прибыль": profit,
        "Рентабельность, %": profit / revenue * 100,
        "АКБ": float(sales["Покупатель"].nunique()),
        "Кол-во заказов": float(orders),
        "Средний чек": revenue / orders,
        "Затраты на логистику": logistics,
        "Доля логистики, %": logistics / revenue * 100,
    }


def abc_by_product(sales: pd.DataFrame, a_share: float = 0.7, b_share: float = 0.9) -> pd.DataFrame:
    """ABC by revenue: cumulative share <= 70% -> A, <= 90% -> B, else C."""
    by_product = (
        sales.groupby("Продукт", as_index=False)["Выручка"]
        .sum()
        .sort_values("Выручка", ascending=False, kind="mergesort")
        .reset_index(drop=True)
    )
    by_product["Накопленная доля"] = by_product["Выручка"].cumsum() / by_product["Выручка"].sum()
    by_product["ABC"] = pd.cut(
        by_product["Накопленная доля"],
        bins=[-float("inf"), a_share, b_share, float("inf")],
        labels=["A", "B", "C"],
    ).astype(str)
    return by_product


def loss_making_orders(sales: pd.DataFrame, managers: pd.DataFrame) -> pd.DataFrame:
    """Orders with negative total profit, per manager (via region)."""
    with_manager = sales.merge(managers[["Регион", "Менеджер"]], on="Регион", how="left")
    orders = with_manager.groupby(["Менеджер", "Order_ID"], as_index=False)["Прибыль"].sum()
    losses = orders[orders["Прибыль"] < 0]
    return (
        losses.groupby("Менеджер", as_index=False)
        .agg(**{"Убыточных заказов": ("Order_ID", "count"), "Убыток": ("Прибыль", "sum")})
        .sort_values("Убыток")
        .reset_index(drop=True)
    )


def main(source: Path = DEFAULT_SOURCE) -> None:
    tables = load_tables(source)
    sales, dropped = clean_sales(tables["Продажи"])
    customers, customer_dupes = dedupe_reference(tables["СПР_покупатели"], "Покупатель")
    _, product_dupes = dedupe_reference(tables["СПР_продукты"], "Продукт")
    managers, _ = dedupe_reference(tables["СПР_Менеджеры"], "Регион")

    sheet_rows = count_sheet_rows(source, "Продажи")
    conflicts = conflicting_duplicates(
        tables["СПР_покупатели"], "Покупатель", "Сегмент покупателя"
    )
    print(
        f"Строк на листе продаж: {sheet_rows}; с данными: {len(sales)}; "
        f"пустых: {sheet_rows - len(sales)}; без Order_ID/выручки: {dropped}"
    )
    print(
        f"Лишних строк в справочнике покупателей: {customer_dupes} "
        f"(покупателей с повторами: {len(conflicts)}, у всех разные сегменты); "
        f"продуктов: {product_dupes}"
    )
    print()
    for name, value in core_metrics(sales, tables["СПР_стоимость_доставки"]).items():
        formatted = f"{value:,.2f}".replace(",", " ")
        print(f"{name:<24} {formatted:>16}")

    abc = abc_by_product(sales)
    summary = abc.groupby("ABC").agg(товаров=("Продукт", "count"), выручка=("Выручка", "sum"))
    summary["доля выручки, %"] = summary["выручка"] / summary["выручка"].sum() * 100
    print("\nABC по номенклатуре:")
    print(summary.round(2).to_string())

    print("\nМенеджеры с убыточными заказами:")
    print(loss_making_orders(sales, managers).round(2).to_string(index=False))
    print(f"\nПокупателей в справочнике: {len(customers)}")


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SOURCE)
