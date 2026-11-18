"""
Draw images/dashboard_preview.png: a picture of the dashboard's numbers (matplotlib), for the README.
The real, interactive dashboard is restaurant_sales_dashboard.xlsx.

    python src/make_preview.py
"""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
df = pd.read_csv(ROOT / "data" / "restaurant_sales.csv", parse_dates=["order_date"])
NAVY, TEAL, AMBER, GREY, INK = "#14213D", "#2A9D8F", "#F4A259", "#6B7785", "#1F2933"
plt.rcParams.update({"font.family": ["Segoe UI", "DejaVu Sans"], "axes.spines.top": False, "axes.spines.right": False, "axes.spines.left": False,
                     "axes.grid": True, "axes.grid.axis": "y", "grid.color": "#E9EDF2", "axes.axisbelow": True, "text.color": INK})

fig = plt.figure(figsize=(15, 8.6), facecolor="white")
fig.patches.append(plt.Rectangle((0, 0.915), 1, 0.085, transform=fig.transFigure, color=NAVY))
fig.text(0.02, 0.96, "Restaurant sales dashboard", color="white", fontsize=20, fontweight="bold", va="center")
fig.text(0.02, 0.928, f"10 locations · 25 menu items · {len(df):,} orders · 2024 to 2025", color="#C9D1DC", fontsize=10, va="center")

kp = [("REVENUE", f"${df.order_total.sum():,.0f}"), ("ORDERS", f"{len(df):,}"), ("AVG ORDER VALUE", f"${df.order_total.mean():.2f}"),
      ("AVG RATING", f"{df.customer_rating.mean():.2f}"), ("ITEMS SOLD", f"{df.quantity.sum():,}"), ("TOP CATEGORY", "Main")]
for i, (n, v) in enumerate(kp):
    x = 0.02 + i * 0.164
    fig.patches.append(plt.Rectangle((x, 0.80), 0.152, 0.095, transform=fig.transFigure, facecolor="white", edgecolor="#E3E8EF"))
    fig.patches.append(plt.Rectangle((x, 0.893), 0.152, 0.003, transform=fig.transFigure, color=TEAL))
    fig.text(x + 0.076, 0.872, n, ha="center", fontsize=8.5, color=GREY); fig.text(x + 0.076, 0.826, v, ha="center", fontsize=20, fontweight="bold", color=NAVY)

m = df.groupby(df.order_date.dt.to_period("M")).order_total.sum()
ax = fig.add_axes([0.04, 0.46, 0.5, 0.29]); ax.plot(m.index.to_timestamp(), m.values, color=TEAL, lw=2.5, marker="o", ms=4)
ax.set_title("Monthly revenue ($)", loc="left", fontweight="bold", color=NAVY)
ax = fig.add_axes([0.6, 0.46, 0.36, 0.29]); c = df.groupby("category").order_total.sum().sort_values(ascending=False)
b = ax.bar(c.index, c.values, color=NAVY, width=.6); ax.bar_label(b, labels=[f"${v:,.0f}" for v in c.values], fontsize=8); ax.set_yticks([]); ax.grid(False)
ax.set_title("Revenue by category ($)", loc="left", fontweight="bold", color=NAVY)
ax = fig.add_axes([0.1, 0.07, 0.3, 0.3]); l = df.groupby("location").order_total.sum().sort_values()
b = ax.barh(l.index, l.values, color=TEAL, height=.65); ax.bar_label(b, labels=[f"${v:,.0f}" for v in l.values], fontsize=8); ax.set_xticks([]); ax.grid(False)
ax.set_title("Revenue by location ($)", loc="left", fontweight="bold", color=NAVY)
ax = fig.add_axes([0.47, 0.07, 0.25, 0.3]); order = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
w = df.groupby(df.order_date.dt.strftime("%a")).order_total.sum().reindex(order)
b = ax.bar(order, w.values, color=AMBER, width=.6); ax.bar_label(b, labels=[f"${v:,.0f}" for v in w.values], fontsize=8); ax.set_yticks([]); ax.grid(False)
ax.set_title("Revenue by weekday ($)", loc="left", fontweight="bold", color=NAVY)
ax = fig.add_axes([0.76, 0.05, 0.22, 0.34]); t = df.groupby("order_type").order_total.sum().sort_values(ascending=False)
_, _, at = ax.pie(t.values, labels=t.index, colors=[NAVY, TEAL, AMBER], autopct="%1.0f%%", pctdistance=.78, wedgeprops=dict(width=.42, edgecolor="white"), textprops=dict(fontsize=9))
for x in at:
    x.set_color("white"); x.set_fontweight("bold")
ax.set_title("Order type mix", loc="left", fontweight="bold", color=NAVY)
fig.text(0.02, 0.012, "Picture of the dashboard content, drawn with matplotlib. The interactive version (dropdown filters, live formulas) is restaurant_sales_dashboard.xlsx.", fontsize=8, color=GREY)
(ROOT / "images").mkdir(exist_ok=True)
fig.savefig(ROOT / "images" / "dashboard_preview.png", dpi=130); print("saved")
