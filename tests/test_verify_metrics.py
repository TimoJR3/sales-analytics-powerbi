from pathlib import Path

import pandas as pd
import pytest

from scripts.verify_metrics import (
    abc_by_product,
    clean_sales,
    conflicting_duplicates,
    core_metrics,
    dedupe_reference,
    loss_making_orders,
    to_number,
)

SOURCE = Path(__file__).resolve().parents[1] / "data" / "source.xlsx"


@pytest.mark.parametrize(
    ("raw", "expected"),
    [("2 781.82", 2781.82), ("1 234.5", 1234.5), ("-695.26", -695.26), ("", None), (None, None), (7, 7.0)],
)
def test_to_number_handles_spaces_and_nbsp(raw, expected) -> None:
    assert to_number(raw) == expected


def _sales() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Order_ID": ["1", "1", "2", None, "3"],
            "Количество": [1, 2, 1, None, 1],
            "Выручка": ["100", "50", "1 000", None, "250"],
            "Прибыль": ["10", "-40", "200", None, "-5"],
            "Метод доставки": ["Express Air", "Regular Air", "Regular Air", None, "Express Air"],
            "Покупатель": [" Ann ", "Ann", "Bob", None, "Cid"],
            "Регион": ["North", "North", "South", None, "South"],
            "Продукт": ["p1", "p2", "p3", None, "p1"],
            "Приоритет": ["High"] * 5,
        }
    )


def test_clean_sales_drops_rows_without_order_or_revenue() -> None:
    sales, dropped = clean_sales(_sales())

    assert dropped == 1
    assert sales["Выручка"].tolist() == [100, 50, 1000, 250]
    assert sales["Покупатель"].nunique() == 3


def test_core_metrics() -> None:
    sales, _ = clean_sales(_sales())
    shipping = pd.DataFrame(
        {"Метод доставки": ["Express Air", "Regular Air"], "Стоимость доставки": ["10", "5"]}
    )
    metrics = core_metrics(sales, shipping)

    assert metrics["Выручка"] == 1400
    assert metrics["Прибыль"] == 165
    assert metrics["Кол-во заказов"] == 3
    assert metrics["Средний чек"] == pytest.approx(1400 / 3)
    assert metrics["Затраты на логистику"] == 30
    assert metrics["АКБ"] == 3


def test_abc_thresholds() -> None:
    sales = pd.DataFrame({"Продукт": list("abcde"), "Выручка": [50, 20, 15, 10, 5]})
    abc = abc_by_product(sales)

    assert dict(zip(abc["Продукт"], abc["ABC"], strict=True)) == {"a": "A", "b": "A", "c": "B", "d": "C", "e": "C"}


def test_loss_making_orders_sum_lines_per_order() -> None:
    sales, _ = clean_sales(_sales())
    managers = pd.DataFrame({"Регион": ["North", "South"], "Менеджер": ["M1", "M2"]})
    losses = loss_making_orders(sales, managers)

    # order 1 = 10 - 40 = -30 (loss), order 2 = +200, order 3 = -5 (loss)
    assert losses.to_dict("records") == [
        {"Менеджер": "M1", "Убыточных заказов": 1, "Убыток": -30.0},
        {"Менеджер": "M2", "Убыточных заказов": 1, "Убыток": -5.0},
    ]


def test_reference_dedup_is_case_insensitive_and_reports_conflicts() -> None:
    customers = pd.DataFrame(
        {"Покупатель": ["Ann", "ann ", "Bob", "Bob"], "Сегмент покупателя": ["A", "B", "C", "C"]}
    )
    unique, removed = dedupe_reference(customers, "Покупатель")
    conflicts = conflicting_duplicates(customers, "Покупатель", "Сегмент покупателя")

    assert removed == 2
    assert unique["Покупатель"].tolist() == ["Ann", "Bob"]
    assert conflicts.index.tolist() == ["ann"]


@pytest.mark.skipif(not SOURCE.exists(), reason="source.xlsx is not available")
def test_report_totals_on_real_data() -> None:
    from scripts.verify_metrics import load_tables

    tables = load_tables(SOURCE)
    sales, _ = clean_sales(tables["Продажи"])
    metrics = core_metrics(sales, tables["СПР_стоимость_доставки"])

    assert len(sales) == 730
    assert metrics["Выручка"] == pytest.approx(1_103_256.31, abs=0.01)
    assert metrics["Прибыль"] == pytest.approx(253_784.61, abs=0.01)
    assert metrics["Кол-во заказов"] == 571
    assert metrics["АКБ"] == 173
    assert metrics["Затраты на логистику"] == 4_405
