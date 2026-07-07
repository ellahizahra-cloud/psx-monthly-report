"""
build_report.py — Turn prices.json (from fetch_prices.py) into the
monthly Excel report, grouped by sector, with an optional dividend column.

Usage:
    python fetch_prices.py [--month YYYY-MM] [--div-start YYYY-MM-DD --div-end YYYY-MM-DD]
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
SECTOR_BG = "D9E2F3"
GREY = "595959"
AMBER_FILL = "FFF2CC"
AMBER_TEXT = "BF8F00"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# Extend this as you add tickers. Falls back to the ticker itself if missing.
COMPANY_NAMES = {
    "MCB": "MCB Bank Limited", "HBL": "Habib Bank Limited", "BAFL": "Bank Alfalah Limited",
    "BAHL": "Bank Al Habib Limited", "MEBL": "Meezan Bank Limited", "HMB": "Habib Metropolitan Bank Limited",
    "UBL": "United Bank Limited",
    "LUCK": "Lucky Cement Limited", "KOHC": "Kohat Cement Company Limited", "DGKC": "D.G. Khan Cement Company Limited",
    "CHCC": "Cherat Cement Company Limited", "FCCL": "Fauji Cement Company Limited", "MLCF": "Maple Leaf Cement Factory Limited",
    "ENGROH": "Engro Holdings Limited", "EFERT": "Engro Fertilizers Limited",
    "FATIMA": "Fatima Fertilizer Company Limited", "FFC": "Fauji Fertilizer Company Limited",
    "POL": "Pakistan Oilfields Limited", "PSO": "Pakistan State Oil Company Limited",
    "MARI": "Mari Energies Limited", "PPL": "Pakistan Petroleum Limited", "OGDC": "Oil & Gas Development Company Limited",
    "ALTN": "Altern Energy Limited", "NCPL": "Nishat Chunian Power Limited", "KOHE": "Kohinoor Energy Limited",
    "NPL": "Nishat Power Limited", "HUBC": "Hub Power Company Limited",
    "ISL": "International Steels Limited", "AGHA": "Agha Steel Industries Limited",
    "ASTL": "Amreli Steels Limited", "ASL": "Aisha Steel Mills Limited",
    "CEPB": "Century Paper & Board Mills Limited", "INDU": "Indus Motor Company Limited",
    "NML": "Nishat Mills Limited", "AGIL": "Agritech Limited", "ORIX": "Orix Leasing Pakistan Limited",
    "EPCL": "Engro Polymer & Chemicals Limited",
    "ABOT": "Abbott Laboratories Pakistan Limited", "FEROZ": "Ferozsons Laboratories Limited",
    "GLAXO": "GlaxoSmithKline Pakistan Limited", "HINOON": "Highnoon Laboratories Limited",
    "CPHL": "Citi Pharma Limited", "AGP": "AGP Limited",
    "SYS": "Systems Limited", "TRG": "TRG Pakistan Limited", "AVN": "Avanceon Limited",
    "AATM": "Ali Asghar Textile Mills Limited",
}


def build(prices_path: str = "prices.json") -> str:
    with open(prices_path) as f:
        data = json.load(f)

    ref_month = data["reference_month"]
    ref_label = dt.datetime.strptime(ref_month, "%Y-%m").strftime("%B %Y")
    div_window = data.get("dividend_window")

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
    if div_window:
        headers.append(f"Dividend Payout (PKR/share), {div_window['start']} to {div_window['end']}")
    n_cols = len(headers)

    HEADER_ROW = 4
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=h)
        cell.font = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    # Group results by sector, preserving first-seen sector order.
    sectors_order = []
    by_sector = {}
    for r in data["results"]:
        s = r.get("sector", "Other")
        if s not in by_sector:
            by_sector[s] = []
            sectors_order.append(s)
        by_sector[s].append(r)

    row = HEADER_ROW + 1
    for sector in sectors_order:
        # Sector header row
        cell = ws.cell(row=row, column=1, value=sector)
        cell.font = Font(name=ARIAL, size=10.5, bold=True, color=NAVY)
        cell.fill = PatternFill("solid", start_color=SECTOR_BG)
        for c in range(2, n_cols + 1):
            ws.cell(row=row, column=c).fill = PatternFill("solid", start_color=SECTOR_BG)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        row += 1

        for r in by_sector[sector]:
            tkr = r["ticker"]
            ws.cell(row=row, column=1, value=tkr).font = Font(name=ARIAL, size=10, bold=True)
            ws.cell(row=row, column=2, value=COMPANY_NAMES.get(tkr, tkr)).font = Font(name=ARIAL, size=10)

            if r.get("error"):
                for c in range(3, n_cols + 1):
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

            if div_window:
                dv = r.get("dividend_payout")
                c8 = ws.cell(row=row, column=8, value=dv if dv is not None else "[TO VERIFY]")
                if dv is None:
                    c8.fill = PatternFill("solid", start_color=AMBER_FILL)
                    c8.font = Font(name=ARIAL, size=10, color=AMBER_TEXT, italic=True)
                else:
                    c8.number_format = "#,##0.00"
                    c8.font = Font(name=ARIAL, size=10, color="0000FF")

            for c in range(1, n_cols + 1):
                ws.cell(row=row, column=c).border = BORDER
            row += 1

    if data.get("warnings"):
        row += 1
        ws.cell(row=row, column=1, value="Warnings:").font = Font(name=ARIAL, size=9, bold=True)
        for w in data["warnings"]:
            row += 1
            ws.cell(row=row, column=1, value=w).font = Font(name=ARIAL, size=8.5, color="C00000")

    widths = {"A": 10, "B": 38, "C": 30, "D": 18, "E": 14, "F": 12, "G": 22, "H": 30}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"A{HEADER_ROW + 1}"

    out_path = f"PSX_Share_Prices_{ref_month}.xlsx"
    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    path = build()
    print(f"Saved {path}")
