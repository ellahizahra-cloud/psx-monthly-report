"""
build_report.py — Turn prices.json (from fetch_prices.py) into the
monthly Excel report.

Usage:
    python fetch_prices.py
    python build_report.py
    python send_report.py

Output: PSX_Share_Prices_<YYYY-MM>.xlsx
"""

import json
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from tickers import COMPANY_NAMES, SECTORS

ARIAL = "Arial"
NAVY = "1F3864"
GREY = "595959"
AMBER_FILL = "FFF2CC"
AMBER_TEXT = "BF8F00"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def build(prices_path: str = "prices.json") -> str:
    with open(prices_path) as f:
        data = json.load(f)

    ref_month = data["reference_month"]  # e.g. "2026-07" — the fully-closed month being reported
    ref_label = dt.datetime.strptime(ref_month, "%Y-%m").strftime("%B %Y")

    wb = Workbook()
    ws = wb.active
    ws.title = "Share Prices"

    ws["A1"] = f"PSX Share Price Report — {ref_label}"
    ws["A1"].font = Font(name=ARIAL, size=14, bold=True, color=NAVY)
    ws["A2"] = (
        f"Generated {data['generated']}. Source: PSX Data Portal EOD feed. "
        "Month-End Close cross-checked vs Sarmaaya.pk."
    )
    ws["A2"].font = Font(name=ARIAL, size=9, italic=True, color=GREY)

    headers = [
        "Ticker", "Company", "Sector",
        f"Month-Start Close, {ref_label} (PKR)",
        f"Month-End Close, {ref_label} (PKR)",
        "Change (%)", "Cross-check Status",
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    row = 5
    for r in data["results"]:
        tkr = r["ticker"]
        ws.cell(row=row, column=1, value=tkr).font = Font(name=ARIAL, size=10, bold=True)
        ws.cell(row=row, column=2, value=COMPANY_NAMES.get(tkr, tkr)).font = Font(name=ARIAL, size=10)
        ws.cell(row=row, column=3, value=SECTORS.get(tkr, "")).font = Font(name=ARIAL, size=10)

        if r.get("error"):
            for c in range(4, 8):
                cell = ws.cell(row=row, column=c, value="FETCH FAILED — fill manually" if c == 4 else "")
                cell.fill = PatternFill("solid", start_color="FFC7CE")
                cell.border = BORDER
            row += 1
            continue

        def price_cell(col, value):
            cell = ws.cell(row=row, column=col, value=value)
            cell.number_format = "#,##0.00"
            cell.font = Font(name=ARIAL, size=10, color="0000FF")
            if value is None:
                cell.value = "[TO VERIFY]"
                cell.fill = PatternFill("solid", start_color=AMBER_FILL)
                cell.font = Font(name=ARIAL, size=10, color=AMBER_TEXT, italic=True)
            return cell

        price_cell(4, r["month_start_close"])
        price_cell(5, r["month_end_close"])

        start_px, end_px = r["month_start_close"], r["month_end_close"]
        change_pct = (start_px - end_px) / end_px if start_px is not None and end_px else None
        c6 = ws.cell(row=row, column=6, value=change_pct if change_pct is not None else "n/a")
        c6.number_format = "0.0%"
        c6.font = Font(name=ARIAL, size=10)

        status = "Flagged — check manually" if r.get("flagged") else "OK"
        c7 = ws.cell(row=row, column=7, value=status)
        c7.font = Font(name=ARIAL, size=10, color="C00000" if r.get("flagged") else "375623")
        if r.get("flagged"):
            c7.fill = PatternFill("solid", start_color="FFC7CE")

        for c in range(1, 8):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    if data.get("warnings"):
        row += 1
        ws.cell(row=row, column=1, value="Warnings:").font = Font(name=ARIAL, size=9, bold=True)
        for w in data["warnings"]:
            row += 1
            ws.cell(row=row, column=1, value=w).font = Font(name=ARIAL, size=8.5, color="C00000")

    widths = {"A": 10, "B": 38, "C": 16, "D": 26, "E": 26, "F": 12, "G": 22}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    out_path = f"PSX_Share_Prices_{ref_month}.xlsx"
    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    path = build()
    print(f"Saved {path}")
