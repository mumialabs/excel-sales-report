"""Turn a messy sales export (CSV) into a formatted Excel sales report.

Usage:
    python build_report.py --input data/sales_raw.csv --output output/sales_report.xlsx

Steps:
  1. Clean the data: drop blank rows, fix region labels, parse "$1,299.00" prices
     and mixed date formats, drop rows with missing values, remove duplicates.
  2. Write a workbook with a Dashboard (KPIs + native Excel charts), Monthly,
     By Region, By Product and By Rep sheets, the cleaned data, and a Data
     Quality log. All summaries are live Excel formulas, so the report updates
     if someone edits the data in Excel.
"""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = "Arial"
NAVY = "1F3A5F"
SERIES = "2A78D6"
GRID = "E3E6EA"
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT)
THIN = Side(style="thin", color="BFC7D1")
MONEY = '$#,##0'
MONEY2 = '$#,##0.00'
PCT = '0.0%'


# ------------------------------------------------------------------ cleaning

def clean(raw: pd.DataFrame) -> tuple[pd.DataFrame, list[tuple[str, int]]]:
    log = [("Rows in raw file", len(raw))]
    df = raw.copy()

    blank = df.isna().all(axis=1)
    df = df[~blank]
    log.append(("Blank rows removed", int(blank.sum())))

    original_region = df["region"]
    df["region"] = df["region"].str.strip().str.title()
    log.append(("Region labels standardised", int((original_region != df["region"]).sum())))

    price_text = df["unit_price"].str.contains(r"[$,]", na=False)
    df["unit_price"] = pd.to_numeric(df["unit_price"].str.replace(r"[$,]", "", regex=True),
                                     errors="coerce")
    log.append(('Prices converted from text (e.g. "$1,299.00")', int(price_text.sum())))

    iso = pd.to_datetime(df["order_date"], format="%Y-%m-%d", errors="coerce")
    us = pd.to_datetime(df["order_date"], format="%m/%d/%Y", errors="coerce")
    log.append(("Dates converted from MM/DD/YYYY", int((iso.isna() & us.notna()).sum())))
    df["order_date"] = iso.fillna(us)

    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    invalid = df[["order_date", "quantity", "unit_price"]].isna().any(axis=1)
    df = df[~invalid]
    log.append(("Rows dropped (missing date, quantity or price)", int(invalid.sum())))

    dupes = df.duplicated(subset="order_id", keep="first")
    df = df[~dupes]
    log.append(("Duplicate orders removed", int(dupes.sum())))

    df = df.sort_values(["order_date", "order_id"]).reset_index(drop=True)
    df["quantity"] = df["quantity"].astype(int)
    df["month"] = df["order_date"].dt.strftime("%Y-%m")
    log.append(("Clean rows in report", len(df)))
    return df, log


# ------------------------------------------------------------------ workbook helpers

def style_table(ws, n_cols: int, n_rows: int, widths: list[int]) -> None:
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for cell in ws[1][:n_cols]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2, max_row=n_rows + 1, max_col=n_cols):
        for cell in row:
            cell.font = BODY_FONT
            cell.border = Border(bottom=THIN)
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"


def set_formats(ws, col: str, fmt: str, first: int, last: int) -> None:
    for r in range(first, last + 1):
        ws[f"{col}{r}"].number_format = fmt


def style_chart(chart, value_fmt: str) -> None:
    chart.legend = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.numFmt = value_fmt
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=GRID))
    chart.x_axis.majorTickMark = "none"
    chart.y_axis.majorTickMark = "none"
    chart.width, chart.height = 17, 8.5


# ------------------------------------------------------------------ report

def build_workbook(df: pd.DataFrame, log: list[tuple[str, int]], source: str, output: Path) -> None:
    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_month = wb.create_sheet("Monthly")
    ws_region = wb.create_sheet("By Region")
    ws_product = wb.create_sheet("By Product")
    ws_rep = wb.create_sheet("By Rep")
    ws_data = wb.create_sheet("Clean Data")
    ws_dq = wb.create_sheet("Data Quality")

    # Clean Data -------------------------------------------------------------
    ws_data.append(["Order ID", "Order Date", "Month", "Region", "Sales Rep", "Product",
                    "Category", "Quantity", "Unit Price", "Revenue"])
    for i, r in enumerate(df.itertuples(index=False), start=2):
        ws_data.append([r.order_id, r.order_date.date(), r.month, r.region, r.sales_rep,
                        r.product, r.category, r.quantity, r.unit_price, f"=H{i}*I{i}"])
    n = len(df) + 1
    style_table(ws_data, 10, len(df), [11, 12, 9, 9, 15, 25, 11, 9, 11, 12])
    set_formats(ws_data, "B", "yyyy-mm-dd", 2, n)
    set_formats(ws_data, "I", MONEY2, 2, n)
    set_formats(ws_data, "J", MONEY2, 2, n)
    ws_data.auto_filter.ref = f"A1:J{n}"

    def rng(col: str) -> str:
        return f"'Clean Data'!${col}$2:${col}${n}"

    # Monthly ----------------------------------------------------------------
    months = sorted(df["month"].unique())
    ws_month.append(["Month", "Revenue", "Orders", "Avg Order Value", "vs Previous Month"])
    for i, m in enumerate(months, start=2):
        ws_month.append([
            m,
            f"=SUMIFS({rng('J')},{rng('C')},A{i})",
            f"=COUNTIFS({rng('C')},A{i})",
            f"=IF(C{i}=0,0,B{i}/C{i})",
            "" if i == 2 else f"=IF(B{i - 1}=0,\"\",B{i}/B{i - 1}-1)",
        ])
    m_last = len(months) + 1
    style_table(ws_month, 5, len(months), [11, 14, 10, 17, 19])
    set_formats(ws_month, "B", MONEY, 2, m_last)
    set_formats(ws_month, "D", MONEY2, 2, m_last)
    set_formats(ws_month, "E", PCT, 2, m_last)

    # By Region --------------------------------------------------------------
    regions = df.groupby("region")["quantity"].count().index.tolist()
    regions.sort(key=lambda reg: -(df.loc[df.region == reg, "quantity"]
                                   * df.loc[df.region == reg, "unit_price"]).sum())
    ws_region.append(["Region", "Revenue", "Orders", "Share of Revenue"])
    r_last = len(regions) + 1
    for i, reg in enumerate(regions, start=2):
        ws_region.append([reg, f"=SUMIFS({rng('J')},{rng('D')},A{i})",
                          f"=COUNTIFS({rng('D')},A{i})", f"=B{i}/SUM($B$2:$B${r_last})"])
    style_table(ws_region, 4, len(regions), [12, 14, 10, 18])
    set_formats(ws_region, "B", MONEY, 2, r_last)
    set_formats(ws_region, "D", PCT, 2, r_last)

    # By Product -------------------------------------------------------------
    revenue = df.assign(rev=df.quantity * df.unit_price)
    products = (revenue.groupby(["product", "category"])["rev"].sum()
                .sort_values(ascending=False).reset_index())
    ws_product.append(["Product", "Category", "Units Sold", "Revenue", "Share of Revenue"])
    p_last = len(products) + 1
    for i, p in enumerate(products.itertuples(index=False), start=2):
        ws_product.append([p.product, p.category, f"=SUMIFS({rng('H')},{rng('F')},A{i})",
                           f"=SUMIFS({rng('J')},{rng('F')},A{i})",
                           f"=D{i}/SUM($D$2:$D${p_last})"])
    style_table(ws_product, 5, len(products), [26, 12, 12, 14, 18])
    set_formats(ws_product, "C", "#,##0", 2, p_last)
    set_formats(ws_product, "D", MONEY, 2, p_last)
    set_formats(ws_product, "E", PCT, 2, p_last)

    # By Rep -----------------------------------------------------------------
    reps = (revenue.groupby(["sales_rep", "region"])["rev"].sum()
            .sort_values(ascending=False).reset_index())
    ws_rep.append(["Sales Rep", "Region", "Orders", "Revenue", "Avg Order Value"])
    s_last = len(reps) + 1
    for i, rp in enumerate(reps.itertuples(index=False), start=2):
        ws_rep.append([rp.sales_rep, rp.region, f"=COUNTIFS({rng('E')},A{i})",
                       f"=SUMIFS({rng('J')},{rng('E')},A{i})", f"=IF(C{i}=0,0,D{i}/C{i})"])
    style_table(ws_rep, 5, len(reps), [16, 10, 10, 14, 17])
    set_formats(ws_rep, "D", MONEY, 2, s_last)
    set_formats(ws_rep, "E", MONEY2, 2, s_last)

    # Data Quality -----------------------------------------------------------
    ws_dq.append(["Cleaning step", "Rows"])
    for step, count in log:
        ws_dq.append([step, count])
    style_table(ws_dq, 2, len(log), [48, 10])
    ws_dq.cell(row=len(log) + 3, column=1,
               value=f"Source: {source}. Report generated by build_report.py on {date.today():%Y-%m-%d}.")
    ws_dq.cell(row=len(log) + 3, column=1).font = Font(name=FONT, italic=True, color="666666")

    # Dashboard --------------------------------------------------------------
    for col in "ABCDEFGHIJ":
        ws_dash.column_dimensions[col].width = 13
    first, last = pd.to_datetime(months[0]), pd.to_datetime(months[-1])
    ws_dash["A1"] = f"Sales Report  |  {first:%b %Y} - {last:%b %Y}"
    ws_dash["A1"].font = Font(name=FONT, bold=True, size=18, color=NAVY)
    ws_dash["A2"] = f"Built automatically from {source} ({len(df):,} clean orders)"
    ws_dash["A2"].font = Font(name=FONT, italic=True, color="666666")

    kpis = [
        ("A", "Total Revenue", f"=SUM(Monthly!B2:B{m_last})", MONEY),
        ("C", "Orders", f"=SUM(Monthly!C2:C{m_last})", "#,##0"),
        ("E", "Avg Order Value", "=IF(C5=0,0,A5/C5)", MONEY2),
        ("G", "Best Month", f"=INDEX(Monthly!A2:A{m_last},MATCH(MAX(Monthly!B2:B{m_last}),"
                            f"Monthly!B2:B{m_last},0))", "@"),
        ("I", "Top Region", f"=INDEX('By Region'!A2:A{r_last},MATCH(MAX('By Region'!B2:B{r_last}),"
                            f"'By Region'!B2:B{r_last},0))", "@"),
    ]
    tile = PatternFill("solid", fgColor="EEF2F7")
    for col, label, formula, fmt in kpis:
        nxt = chr(ord(col) + 1)
        ws_dash.merge_cells(f"{col}4:{nxt}4")
        ws_dash.merge_cells(f"{col}5:{nxt}5")
        ws_dash[f"{col}4"] = label
        ws_dash[f"{col}4"].font = Font(name=FONT, size=10, color="52514E")
        ws_dash[f"{col}5"] = formula
        ws_dash[f"{col}5"].font = Font(name=FONT, size=18, bold=True, color=NAVY)
        ws_dash[f"{col}5"].number_format = fmt
        for r in (4, 5):
            for c in (col, nxt):
                ws_dash[f"{c}{r}"].fill = tile
            ws_dash[f"{col}{r}"].alignment = Alignment(horizontal="left", indent=1)
    ws_dash.row_dimensions[5].height = 32

    line = LineChart()
    line.title = "Monthly Revenue"
    line.add_data(Reference(ws_month, min_col=2, min_row=1, max_row=m_last), titles_from_data=True)
    line.set_categories(Reference(ws_month, min_col=1, min_row=2, max_row=m_last))
    line.series[0].graphicalProperties.line.solidFill = SERIES
    line.series[0].graphicalProperties.line.width = 28575
    line.series[0].smooth = False
    style_chart(line, MONEY)
    ws_dash.add_chart(line, "A7")

    region_chart = BarChart()
    region_chart.type = "bar"
    region_chart.title = "Revenue by Region"
    region_chart.add_data(Reference(ws_region, min_col=2, min_row=1, max_row=r_last),
                          titles_from_data=True)
    region_chart.set_categories(Reference(ws_region, min_col=1, min_row=2, max_row=r_last))
    region_chart.series[0].graphicalProperties.solidFill = SERIES
    region_chart.series[0].graphicalProperties.line.noFill = True
    region_chart.gapWidth = 60
    region_chart.dataLabels = DataLabelList()
    region_chart.dataLabels.showVal = True
    region_chart.dataLabels.numFmt = MONEY
    style_chart(region_chart, MONEY)
    region_chart.x_axis.scaling.orientation = "maxMin"
    region_chart.y_axis.crosses = "max"
    region_chart.width = 13
    ws_dash.add_chart(region_chart, "A25")

    top_n = min(8, len(products))
    product_chart = BarChart()
    product_chart.type = "bar"
    product_chart.title = f"Top {top_n} Products by Revenue"
    product_chart.add_data(Reference(ws_product, min_col=4, min_row=1, max_row=top_n + 1),
                           titles_from_data=True)
    product_chart.set_categories(Reference(ws_product, min_col=1, min_row=2, max_row=top_n + 1))
    product_chart.series[0].graphicalProperties.solidFill = SERIES
    product_chart.series[0].graphicalProperties.line.noFill = True
    product_chart.gapWidth = 50
    style_chart(product_chart, MONEY)
    product_chart.x_axis.scaling.orientation = "maxMin"
    product_chart.y_axis.crosses = "max"
    product_chart.width = 16
    ws_dash.add_chart(product_chart, "F25")

    for ws in wb:
        for row in ws.iter_rows():
            for cell in row:
                if cell.font is None or cell.font.name != FONT:
                    cell.font = Font(name=FONT, bold=cell.font.bold, italic=cell.font.italic,
                                     size=cell.font.size, color=cell.font.color)

    wb.calculation.fullCalcOnLoad = True
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--input", type=Path, default=Path("data/sales_raw.csv"))
    parser.add_argument("--output", type=Path, default=Path("output/sales_report.xlsx"))
    args = parser.parse_args()

    raw = pd.read_csv(args.input, dtype=str, skip_blank_lines=False)
    df, log = clean(raw)
    build_workbook(df, log, args.input.name, args.output)

    for step, count in log:
        print(f"{step:<50} {count:>6,}")
    total = (df.quantity * df.unit_price).sum()
    print(f"\nRevenue {total:,.2f} across {len(df):,} orders -> saved {args.output}")


if __name__ == "__main__":
    main()
