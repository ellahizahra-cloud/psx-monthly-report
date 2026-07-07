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

ARIAL = "Arial"
NAVY = "1F3864"
GREY = "595959"
AMBER_FILL = "FFF2CC"
AMBER_TEXT = "BF8F00"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

COMPANY_NAMES = {
    "OGDC": "Oil & Gas Development Company Limited",
    "MARI": "Mari Energies Limited",
    "FATIMA": "Fatima Fertilizer Company Limited",
    "AATM": "Ali Asghar Textile Mills Limited",
}


def build(prices_path: str = "prices.json") -> str:
    with open(prices_path) as f:
        data = json.load(f)

    ref_month = data["reference_month"]  # e.g. "2026-07"
    ref_label = dt.datetime.strptime(ref_month, "%Y-%m").strftime("%B %Y")

    wb = Workbook()
    ws = wb.active
    ws.title = "Share Prices"

    ws["A1"] = f"PSX Share Price Report — {ref_label}"
    ws["A1"].font = Font(name=ARIAL, size=14, bold=True, color=NAVY)
    ws["A2"] = f"Generated {data['generated']}. Source: PSX Data Portal EOD feed, cross-checked vs Sarmaaya.pk."
    ws["A2"].font = Font(name=ARIAL, size=9, italic=True, color=GREY)

    headers = [
        "Ticker", "Company",
        f"Close, 1st Trading Day of {ref_label} (PKR)",
        "Latest Close (PKR)", "Latest Date", "Change (%)", "Cross-check Status",
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

        if r.get("error"):
            for c in range(3, 8):
                cell = ws.cell(row=row, column=c, value="FETCH FAILED — fill manually" if c == 3 else "")
                cell.fill = PatternFill("solid", start_color="FFC7CE")
                cell.border = BORDER
            row += 1
            continue

        c3 = ws.cell(row=row, column=3, value=r["month_start_close"])
        c3.number_format = "#,##0.00"
        c3.font = Font(name=ARIAL, size=10, color="0000FF")
        if r["month_start_close"] is None:
            c3.value = "[TO VERIFY]"
            c3.fill = PatternFill("solid", start_color=AMBER_FILL)
            c3.font = Font(name=ARIAL, size=10, color=AMBER_TEXT, italic=True)

        c4 = ws.cell(row=row, column=4, value=r["latest_close"])
        c4.number_format = "#,##0.00"
        c4.font = Font(name=ARIAL, size=10, color="0000FF")

        ws.cell(row=row, column=5, value=r["latest_date"]).font = Font(name=ARIAL, size=10)

        c6 = ws.cell(row=row, column=6,
                      value=f"=IF(ISNUMBER(C{row}),(D{row}-C{row})/C{row},\"n/a\")")
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

    widths = {"A": 10, "B": 38, "C": 30, "D": 18, "E": 14, "F": 12, "G": 22}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    out_path = f"PSX_Share_Prices_{ref_month}.xlsx"
    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    path = build()
    print(f"Saved {path}")
