"""
Clean "04 Restaurant Sales.xlsx" (course workbook) into data/restaurant_sales.csv.

    python src/prepare_data.py "path/to/04 Restaurant Sales.xlsx"

Keeps the 10 data columns, drops the practice formulas typed next to the data, runs data-quality checks.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

OUT = Path(__file__).resolve().parents[1] / "data" / "restaurant_sales.csv"
COLS = {"Transaction ID": "transaction_id", "Order Date": "order_date", "Restaurant Location": "location", "Order Type": "order_type",
        "Menu Item": "menu_item", "Category": "category", "Quantity": "quantity", "Unit Price": "unit_price", "Order Total": "order_total",
        "Customer Rating": "customer_rating"}


def main(src: str) -> int:
    df = pd.read_excel(src)[list(COLS)].rename(columns=COLS)
    checks = {
        "no missing values": df.isna().sum().sum() == 0,
        "transaction_id is unique": df.transaction_id.is_unique,
        "order_total = quantity x unit_price": ((df.quantity * df.unit_price - df.order_total).abs() < 0.005).all(),
        "rating between 1 and 5": df.customer_rating.between(1, 5).all(),
        "quantity positive": (df.quantity > 0).all(),
        "price positive": (df.unit_price > 0).all(),
        "each menu item has one category": df.groupby("menu_item").category.nunique().max() == 1,
    }
    for k, ok in checks.items():
        print(("OK   " if ok else "FAIL ") + k)
    if not all(checks.values()):
        return 1
    OUT.parent.mkdir(exist_ok=True)
    df.sort_values("order_date").to_csv(OUT, index=False, date_format="%Y-%m-%d")
    print(f"{len(df)} rows -> {OUT.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1]))
