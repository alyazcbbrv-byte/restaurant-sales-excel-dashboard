"""
Build the interactive Excel dashboard  restaurant_sales_dashboard.xlsx  from the cleaned CSV.

    python src/build_workbook.py

Sheets
  Dashboard : KPI cards, 2 dropdown filters (Location, Order type), 5 native Excel charts, key insights
  Analysis  : every number on the dashboard as a live formula (SUMIFS / COUNTIFS / AVERAGEIFS) over the Data sheet
  Data      : 1,000 orders as an Excel table (+ Month and Weekday helper columns as formulas (locale independent))
  Notes     : source, method, how to use

Cached values are written next to every formula (computed with pandas for the default "All" view),
so the file also displays correctly in viewers that do not recalculate.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "restaurant_sales_dashboard.xlsx"
df = pd.read_csv(ROOT / "data" / "restaurant_sales.csv", parse_dates=["order_date"])
N = len(df)
LAST = N + 1  # last data row in Excel (header is row 1)

NAVY, TEAL, AMBER, RED, GREY, INK, PAPER = "#14213D", "#2A9D8F", "#F4A259", "#D1495B", "#6B7785", "#1F2933", "#F4F6F9"

wb = xlsxwriter.Workbook(OUT, {"default_date_format": "yyyy-mm-dd"})
wb.set_properties({"title": "Restaurant sales dashboard", "author": "Alyaz Cabbarova", "comments": "Interactive Excel dashboard, formulas only, no macros"})

f = lambda **k: wb.add_format({"font_name": "Calibri", **k})
F = {
    "title": f(bold=True, font_size=20, font_color="white", bg_color=NAVY, valign="vcenter", indent=1),
    "sub": f(font_size=10, font_color="#C9D1DC", bg_color=NAVY, valign="vcenter", indent=1),
    "label": f(bold=True, font_size=10, font_color=GREY, align="right", valign="vcenter"),
    "select": f(bold=True, font_size=11, font_color=NAVY, bg_color="#FFF4E0", border=1, border_color=AMBER, align="center", valign="vcenter"),
    "kpi_name": f(font_size=9, font_color=GREY, bg_color="white", align="center", valign="bottom", top=2, top_color=TEAL, left=1, right=1, left_color="#E3E8EF", right_color="#E3E8EF"),
    "kpi_money": f(bold=True, font_size=22, font_color=NAVY, bg_color="white", align="center", valign="vcenter", num_format="$#,##0", left=1, right=1, bottom=1, left_color="#E3E8EF", right_color="#E3E8EF", bottom_color="#E3E8EF"),
    "kpi_money2": f(bold=True, font_size=22, font_color=NAVY, bg_color="white", align="center", valign="vcenter", num_format="$#,##0.00", left=1, right=1, bottom=1, left_color="#E3E8EF", right_color="#E3E8EF", bottom_color="#E3E8EF"),
    "kpi_int": f(bold=True, font_size=22, font_color=NAVY, bg_color="white", align="center", valign="vcenter", num_format="#,##0", left=1, right=1, bottom=1, left_color="#E3E8EF", right_color="#E3E8EF", bottom_color="#E3E8EF"),
    "kpi_rate": f(bold=True, font_size=22, font_color=NAVY, bg_color="white", align="center", valign="vcenter", num_format="0.00", left=1, right=1, bottom=1, left_color="#E3E8EF", right_color="#E3E8EF", bottom_color="#E3E8EF"),
    "h2": f(bold=True, font_size=12, font_color=NAVY, bottom=2, bottom_color=TEAL),
    "txt": f(font_size=10, font_color=INK, text_wrap=True, valign="top"),
    "bullet": f(font_size=10, font_color=INK, text_wrap=True, valign="top", indent=1),
    "th": f(bold=True, font_color="white", bg_color=NAVY, border=0),
    "money": f(num_format="$#,##0.00"),
    "money0": f(num_format="$#,##0"),
    "int": f(num_format="#,##0"),
    "dec": f(num_format="0.00"),
    "pct": f(num_format="0.0%"),
    "date": f(num_format="yyyy-mm-dd"),
    "mon": f(num_format="mmm yyyy"),
    "plain": f(),
    "note": f(italic=True, font_color=GREY, font_size=9),
}

dash = wb.add_worksheet("Dashboard")
ana = wb.add_worksheet("Analysis")
dat = wb.add_worksheet("Data")
nts = wb.add_worksheet("Notes")

# ------------------------------------------------------------------ Data sheet
heads = ["Transaction ID", "Order Date", "Restaurant Location", "Order Type", "Menu Item", "Category", "Quantity", "Unit Price", "Order Total", "Customer Rating", "Month", "Weekday"]
rows = []
for i, r in enumerate(df.itertuples(index=False), start=2):
    rows.append([r.transaction_id, r.order_date.to_pydatetime(), r.location, r.order_type, r.menu_item, r.category, int(r.quantity), float(r.unit_price), float(r.order_total), int(r.customer_rating),
                 {"formula": f"=DATE(YEAR(B{i}),MONTH(B{i}),1)", "value": (r.order_date.replace(day=1) - pd.Timestamp("1899-12-30")).days},
                 {"formula": f'=CHOOSE(WEEKDAY(B{i},2),"Mon","Tue","Wed","Thu","Fri","Sat","Sun")', "value": r.order_date.strftime("%a")}])
colfmt = [None, F["date"], None, None, None, None, None, F["money"], F["money"], None, F["mon"], None]
for ri, row in enumerate(rows, start=1):
    for ci, v in enumerate(row):
        if isinstance(v, dict):
            dat.write_formula(ri, ci, v["formula"], colfmt[ci], v["value"])
        elif ci == 1:
            dat.write_datetime(ri, ci, v, colfmt[ci])
        else:
            dat.write(ri, ci, v, colfmt[ci])
dat.add_table(0, 0, N, len(heads) - 1, {"name": "Orders", "style": "Table Style Light 1", "columns": [{"header": h} for h in heads]})
for c, w in enumerate([14, 12, 22, 12, 20, 12, 10, 11, 12, 15, 11, 10]):
    dat.set_column(c, c, w)
dat.freeze_panes(1, 0)

rng = lambda col: f"Data!${col}$2:${col}${LAST}"
DATE, LOC, TYP, ITEM, CAT, QTY, TOT, RAT, MON, WDAY = (rng(c) for c in "BCDEFGIJKL")
# ------------------------------------------------------------------ Analysis sheet
ana.set_column("A:A", 24); ana.set_column("B:F", 14)
ana.write("A1", "Location filter (from Dashboard)", F["label"]); ana.write_formula("B1", '=IF(Dashboard!$C$4="All","*",Dashboard!$C$4)', F["plain"], "*")
ana.write("A2", "Order type filter (from Dashboard)", F["label"]); ana.write_formula("B2", '=IF(Dashboard!$F$4="All","*",Dashboard!$F$4)', F["plain"], "*")
ana.write("D1", "These two cells turn the dropdowns on the Dashboard into SUMIFS criteria; '*' means all.", F["note"])
CR = f"{LOC},$B$1,{TYP},$B$2"  # standard criteria pair

ana.write("A4", "KPI", F["th"]); ana.write("B4", "Value", F["th"])
kpis = [("Revenue", f"=SUMIFS({TOT},{CR})", df.order_total.sum()),
        ("Orders", f"=COUNTIFS({LOC},$B$1,{TYP},$B$2)", N),
        ("Average order value", "=IFERROR(B5/B6,0)", df.order_total.mean()),
        ("Average rating", f"=IFERROR(AVERAGEIFS({RAT},{CR}),0)", df.customer_rating.mean()),
        ("Items sold", f"=SUMIFS({QTY},{CR})", int(df.quantity.sum()))]
for i, (n, fm, v) in enumerate(kpis):
    ana.write(4 + i, 0, n); ana.write_formula(4 + i, 1, fm, F["dec"], float(v))
# KPI rows are 5..9 (1-based)

def table(top, title, headers, labels, formulas, values, fmts, label_fmt=None):
    ana.write(top, 0, title, F["h2"])
    for j, h in enumerate(headers):
        ana.write(top + 1, j, h, F["th"])
    for i, lab in enumerate(labels):
        r = top + 2 + i
        ana.write(r, 0, lab, label_fmt or F["plain"]) if not hasattr(lab, "to_pydatetime") else ana.write_datetime(r, 0, lab.to_pydatetime(), label_fmt)
        for j, (fm, vv, ft) in enumerate(zip(formulas, values, fmts), start=1):
            ana.write_formula(r, j, fm(r + 1), ft, vv[i])
    return top + 2, top + 1 + len(labels)  # first,last zero-based rows of data

# monthly
months = pd.date_range("2024-01-01", "2025-12-01", freq="MS")
m_rev = df.groupby(df.order_date.dt.to_period("M")).order_total.sum().reindex(months.to_period("M")).fillna(0).tolist()
m_ord = df.groupby(df.order_date.dt.to_period("M")).size().reindex(months.to_period("M")).fillna(0).tolist()
MT = 11
m_first, m_last = table(MT, "Monthly trend", ["Month", "Revenue", "Orders"], list(months),
    [lambda r: f'=SUMIFS({TOT},{MON},$A{r},{CR})', lambda r: f'=COUNTIFS({MON},$A{r},{LOC},$B$1,{TYP},$B$2)'], [m_rev, m_ord], [F["money0"], F["int"]], F["mon"])

# category (ignores nothing: respects both filters)
cats = df.groupby("category").order_total.sum().sort_values(ascending=False)
CT = m_last + 3
c_first, c_last = table(CT, "Revenue by category", ["Category", "Revenue", "Share"], list(cats.index),
    [lambda r: f'=SUMIFS({TOT},{CAT},$A{r},{CR})', lambda r: f'=IFERROR(B{r}/$B$5,0)'], [cats.round(2).tolist(), (cats / cats.sum()).tolist()], [F["money0"], F["pct"]])

locs = df.groupby("location").order_total.sum().sort_values(ascending=True)  # ascending so the largest is on top of the bar chart
LT = c_last + 3
l_first, l_last = table(LT, "Revenue by location (respects the order-type filter)", ["Location", "Revenue", "Orders"], list(locs.index),
    [lambda r: f'=SUMIFS({TOT},{LOC},$A{r},{TYP},$B$2)', lambda r: f'=COUNTIFS({LOC},$A{r},{TYP},$B$2)'],
    [locs.round(2).tolist(), df.groupby("location").size().reindex(locs.index).tolist()], [F["money0"], F["int"]])

wd_order = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
wd = df.groupby(df.order_date.dt.strftime("%a")).order_total.sum().reindex(wd_order)
WT = l_last + 3
w_first, w_last = table(WT, "Revenue by weekday", ["Weekday", "Revenue", "Share"], wd_order,
    [lambda r: f'=SUMIFS({TOT},{WDAY},$A{r},{CR})', lambda r: f'=IFERROR(B{r}/$B$5,0)'], [wd.round(2).tolist(), (wd / wd.sum()).tolist()], [F["money0"], F["pct"]])

ty = df.groupby("order_type").order_total.sum().sort_values(ascending=False)
TT = w_last + 3
t_first, t_last = table(TT, "Revenue by order type (respects the location filter)", ["Order type", "Revenue", "Share"], list(ty.index),
    [lambda r: f'=SUMIFS({TOT},{TYP},$A{r},{LOC},$B$1)', lambda r: f'=IFERROR(B{r}/SUM($B${TT + 3}:$B${TT + 5}),0)'], [ty.round(2).tolist(), (ty / ty.sum()).tolist()], [F["money0"], F["pct"]])

it = df.groupby("menu_item").order_total.sum().sort_values(ascending=False).head(10)
IT = t_last + 3
i_first, i_last = table(IT, "Top 10 menu items (ranked on the full data set)", ["Menu item", "Revenue", "Units"], list(it.index),
    [lambda r: f'=SUMIFS({TOT},{ITEM},$A{r},{CR})', lambda r: f'=SUMIFS({QTY},{ITEM},$A{r},{CR})'],
    [it.round(2).tolist(), df.groupby("menu_item").quantity.sum().reindex(it.index).tolist()], [F["money0"], F["int"]])

# ------------------------------------------------------------------ Dashboard sheet
dash.hide_gridlines(2)
dash.set_column("A:A", 2); dash.set_column("B:M", 12.5); dash.set_column("N:N", 2)
dash.set_tab_color(NAVY)
dash.set_row(0, 8)
dash.merge_range("B2:M2", "Restaurant sales dashboard", F["title"]); dash.set_row(1, 36)
dash.merge_range("B3:M3", f"10 locations · 25 menu items · {N:,} orders · {df.order_date.min():%d %b %Y} to {df.order_date.max():%d %b %Y} · all numbers are live formulas", F["sub"]); dash.set_row(2, 20)
dash.write("B4", "Location:", F["label"]); dash.merge_range("C4:D4", "All", F["select"])
dash.write("E4", "Order type:", F["label"]); dash.merge_range("F4:G4", "All", F["select"])
dash.write("H4", "Pick a value in the yellow cells; every KPI and chart updates.", F["note"])
dash.data_validation("C4", {"validate": "list", "source": ["All"] + sorted(df.location.unique()), "input_title": "Location", "input_message": "Choose a location or All"})
dash.data_validation("F4", {"validate": "list", "source": ["All"] + sorted(df.order_type.unique()), "input_title": "Order type", "input_message": "Choose an order type or All"})
dash.set_row(3, 24)

cards = [("B", "C", "REVENUE", "=Analysis!B5", F["kpi_money"], float(df.order_total.sum())),
         ("D", "E", "ORDERS", "=Analysis!B6", F["kpi_int"], N),
         ("F", "G", "AVG ORDER VALUE", "=Analysis!B7", F["kpi_money2"], float(df.order_total.mean())),
         ("H", "I", "AVG RATING (1-5)", "=Analysis!B8", F["kpi_rate"], float(df.customer_rating.mean())),
         ("J", "K", "ITEMS SOLD", "=Analysis!B9", F["kpi_int"], int(df.quantity.sum()))]
dash.set_row(5, 18); dash.set_row(6, 40)
for a, b, name, fm, fmt, val in cards:
    dash.merge_range(f"{a}6:{b}6", name, F["kpi_name"])
    dash.merge_range(f"{a}7:{b}7", "", fmt); dash.write_formula(f"{a}7", fm, fmt, val)
dash.merge_range("L6:M6", "TOP CATEGORY", F["kpi_name"])
dash.merge_range("L7:M7", "", F["kpi_int"])
dash.write_formula("L7", f"=INDEX(Analysis!$A${c_first + 1}:$A${c_last + 1},MATCH(MAX(Analysis!$B${c_first + 1}:$B${c_last + 1}),Analysis!$B${c_first + 1}:$B${c_last + 1},0))",
                   f(bold=True, font_size=18, font_color=NAVY, bg_color="white", align="center", valign="vcenter", left=1, right=1, bottom=1, left_color="#E3E8EF", right_color="#E3E8EF", bottom_color="#E3E8EF"), cats.index[0])

def style(ch, title):
    ch.set_title({"name": title, "name_font": {"name": "Calibri", "size": 11, "bold": True, "color": NAVY}})
    ch.set_legend({"none": True}); ch.set_chartarea({"border": {"none": True}}); ch.set_plotarea({"border": {"none": True}})

A = lambda a, b, c: ["Analysis", a, b, a if c is None else c, b if c is None else c]
# monthly line
ch = wb.add_chart({"type": "line"})
ch.add_series({"name": "Revenue", "categories": ["Analysis", m_first, 0, m_last, 0], "values": ["Analysis", m_first, 1, m_last, 1],
               "line": {"color": TEAL, "width": 2.5}, "marker": {"type": "circle", "size": 5, "fill": {"color": TEAL}, "border": {"color": TEAL}}})
style(ch, "Monthly revenue ($)"); ch.set_x_axis({"num_format": "mmm yy", "num_font": {"size": 8, "rotation": -45}, "major_gridlines": {"visible": False}})
ch.set_y_axis({"num_format": "$#,##0", "major_gridlines": {"visible": True, "line": {"color": "#E9EDF2"}}, "line": {"none": True}}); ch.set_size({"width": 780, "height": 290})
dash.insert_chart("B9", ch, {"x_offset": 0, "y_offset": 4})
# category bar
ch = wb.add_chart({"type": "column"})
ch.add_series({"categories": ["Analysis", c_first, 0, c_last, 0], "values": ["Analysis", c_first, 1, c_last, 1], "fill": {"color": NAVY}, "gap": 60,
               "data_labels": {"value": True, "num_format": "$#,##0", "font": {"size": 8}}})
style(ch, "Revenue by category ($)"); ch.set_y_axis({"visible": False, "major_gridlines": {"visible": False}}); ch.set_size({"width": 400, "height": 290})
dash.insert_chart("J9", ch, {"x_offset": 20, "y_offset": 4})
# location bar
ch = wb.add_chart({"type": "bar"})
ch.add_series({"categories": ["Analysis", l_first, 0, l_last, 0], "values": ["Analysis", l_first, 1, l_last, 1], "fill": {"color": TEAL}, "gap": 40,
               "data_labels": {"value": True, "num_format": "$#,##0", "font": {"size": 8}}})
style(ch, "Revenue by location ($)"); ch.set_x_axis({"visible": False, "major_gridlines": {"visible": False}}); ch.set_y_axis({"num_font": {"size": 9}}); ch.set_size({"width": 500, "height": 330})
dash.insert_chart("B24", ch, {"x_offset": 0, "y_offset": 4})
# weekday column
ch = wb.add_chart({"type": "column"})
ch.add_series({"categories": ["Analysis", w_first, 0, w_last, 0], "values": ["Analysis", w_first, 1, w_last, 1], "fill": {"color": AMBER}, "gap": 50,
               "data_labels": {"value": True, "num_format": "$#,##0", "font": {"size": 8}}})
style(ch, "Revenue by weekday ($)"); ch.set_y_axis({"visible": False, "major_gridlines": {"visible": False}}); ch.set_size({"width": 340, "height": 330})
dash.insert_chart("G24", ch, {"x_offset": 20, "y_offset": 4})
# order type doughnut
ch = wb.add_chart({"type": "doughnut"})
ch.add_series({"categories": ["Analysis", t_first, 0, t_last, 0], "values": ["Analysis", t_first, 1, t_last, 1],
               "points": [{"fill": {"color": NAVY}}, {"fill": {"color": TEAL}}, {"fill": {"color": AMBER}}],
               "data_labels": {"percentage": True, "font": {"size": 9, "color": "white", "bold": True}}})
ch.set_title({"name": "Order type mix", "name_font": {"name": "Calibri", "size": 11, "bold": True, "color": NAVY}})
ch.set_legend({"position": "bottom", "font": {"size": 9}}); ch.set_hole_size(55); ch.set_chartarea({"border": {"none": True}}); ch.set_size({"width": 280, "height": 330})
dash.insert_chart("K24", ch, {"x_offset": 10, "y_offset": 4})

# insights (static text, written from the full-data results)
dash.merge_range("B41:M41", "Key insights (full data set)", F["h2"])
main_share = cats["Main"] / cats.sum()
fri_sun = wd[["Fri", "Sat", "Sun"]].sum() / wd.sum()
ins = [
    f"Main courses are {main_share:.0%} of revenue (${cats['Main']:,.0f} of ${cats.sum():,.0f}) from 47.5% of orders; desserts and beverages together are only 19%.",
    f"Friday to Sunday bring {fri_sun:.0%} of revenue and 58% of orders: demand is concentrated on three days, so staffing should follow the weekly curve.",
    f"Downtown Manhattan leads with ${locs.iloc[-1]:,.0f} ({locs.iloc[-1] / locs.sum():.0%}); the weakest location, Upper West Side, earns ${locs.iloc[0]:,.0f}, but has one of the best ratings (4.05).",
    f"Dine-in is the largest channel ({ty.iloc[0] / ty.sum():.0%}); delivery and takeaway orders are about $1 larger on average ($26.27 vs $25.17).",
    "Customers are generally satisfied: 71.5% of ratings are 4 or 5, and only 2.3% are 1. NY Strip Steak is the top item by revenue ($3,497) with a 4.09 rating.",
]
for i, t in enumerate(ins):
    dash.merge_range(41 + i, 1, 41 + i, 12, "•  " + t, F["bullet"]); dash.set_row(41 + i, 30)
dash.set_landscape(); dash.fit_to_pages(1, 1); dash.set_paper(9)
dash.activate()

# ------------------------------------------------------------------ Notes sheet
nts.set_column("A:A", 120)
notes = [("About", True), ("Interactive Excel dashboard built on 1,000 restaurant orders (Jan 2024 to Dec 2025). No macros, no pivot tables: every figure is a SUMIFS / COUNTIFS / AVERAGEIFS formula, so the logic is visible and auditable.", False), ("", False),
         ("How to use", True), ("Pick a Location and/or an Order type in the yellow cells on the Dashboard. KPI cards and charts recalculate. Select All to reset.", False), ("", False),
         ("Sheets", True), ("Dashboard: presentation layer. Analysis: calculation layer (all formulas, including the filter criteria in B1:B2). Data: the Orders table.", False), ("", False),
         ("Data checks done before building", True), ("1,000 unique transaction IDs, no blanks, Order Total = Quantity x Unit Price for every row, ratings 1 to 5, dates 2024-01-01 to 2025-12-30.", False), ("", False),
         ("Limitations", True), ("Sample data: 1,000 orders over two years is about 1.4 orders per day for 10 restaurants, so it is a sample, not full sales. Monthly differences are small-number noise and should not be over-read. Revenue by location and top-10 items are ranked on the full data set; their order does not change when a filter is applied.", False)]
for i, (t, b) in enumerate(notes):
    nts.write(i, 0, t, F["h2"] if b else F["txt"])
wb.close()
print("written", OUT.name)
