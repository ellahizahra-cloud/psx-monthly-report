"""
dividend_workbook.py — Build/rewrite the persistent PSX_Dividend_Tracker.xlsx
from the tracking log (the single source of truth for announcement history).

Three sheets:
  - "Summary": one row per ticker — Dividend This Month, Dividend YTD
    (resets every January since it's computed from the current calendar
    year's announcements), Last Announcement.
  - "Audit Log": one row per announcement ever recorded, for traceability
    (Date, Period, raw PSX details, per-share amount, split-adjusted
    amount).
  - "Cash Flow Reconciliation": one row per ticker — Quantity (from
    holdings.py / Cash Dividend 1 (1).xlsx), Dividend Announced this
    period, Gross/Tax/Net Cash Dividend, with a Total row. Tickers with no
    announcement this period show "-" for the dividend-derived columns
    (never 0) but Quantity always populates.

All figures are actual announced per-share cash-dividend amounts
(PSX payout % x face value), adjusted for any stock splits recorded in
stock_splits.json so historical amounts sit on a current-share-count
basis. Never TTM/annualized/"last declared x4".
"""

import datetime as dt
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import holdings
from tickers import UNIVERSE, COMPANY_NAMES, SECTORS

ARIAL = "Arial"
NAVY = "1F3864"
GREY = "595959"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

SPLITS_PATH = Path("stock_splits.json")
WORKBOOK_PATH = "PSX_Dividend_Tracker.xlsx"

# Withholding tax rate applied to gross cash dividends, per the cash-flow
# reconciliation addendum. Named constant so it can be updated if the rate
# or filer status changes without hunting through the formula.
TAX_WITHHOLDING_RATE = 0.15


def load_splits() -> dict:
    if SPLITS_PATH.exists():
        return json.loads(SPLITS_PATH.read_text())
    return {}


def split_adjustment_factor(ticker: str, date_iso: str, splits: dict) -> float:
    """Cumulative multiple for all splits after this announcement's date."""
    factor = 1.0
    for split in splits.get(ticker, []):
        if date_iso < split["date"]:
            factor *= split["ratio"]
    return factor


def adjusted_amount(entry: dict, ticker: str, splits: dict) -> float:
    factor = split_adjustment_factor(ticker, entry["date_iso"], splits)
    return round(entry["amount_per_share"] / factor, 4)


def _header_row(ws, row, headers):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def build(history: dict) -> str:
    splits = load_splits()
    today = dt.date.today()

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"

    summary["A1"] = "PSX Dividend Tracker"
    summary["A1"].font = Font(name=ARIAL, size=14, bold=True, color=NAVY)
    summary["A2"] = (
        f"Last updated {dt.datetime.now().isoformat(timespec='seconds')}. "
        "Source: PSX company payout announcements (board-meeting/announcement date)."
    )
    summary["A2"].font = Font(name=ARIAL, size=9, italic=True, color=GREY)

    headers = [
        "Ticker", "Company", "Sector",
        "Dividend This Month (PKR/share)",
        f"Dividend YTD {today.year} (PKR/share)",
        "Last Announcement Date",
    ]
    _header_row(summary, 4, headers)

    quantities = holdings.load_quantities()
    per_ticker = []  # collected for the Cash Flow Reconciliation sheet

    row = 5
    for entry in UNIVERSE:
        tkr = entry["ticker"]
        announcements = sorted(history.get(tkr, []), key=lambda a: a["date_iso"])

        this_month_total = 0.0
        ytd_total = 0.0
        last_date = None
        for a in announcements:
            a_date = dt.date.fromisoformat(a["date_iso"])
            amt = adjusted_amount(a, tkr, splits)
            if a_date.year == today.year:
                ytd_total += amt
                if a_date.month == today.month:
                    this_month_total += amt
            if last_date is None or a_date > last_date:
                last_date = a_date

        per_ticker.append({
            "ticker": tkr,
            "quantity": quantities.get(tkr, 0),
            "dividend_this_month": this_month_total,
        })

        summary.cell(row=row, column=1, value=tkr).font = Font(name=ARIAL, size=10, bold=True)
        summary.cell(row=row, column=2, value=COMPANY_NAMES.get(tkr, tkr)).font = Font(name=ARIAL, size=10)
        summary.cell(row=row, column=3, value=SECTORS.get(tkr, "")).font = Font(name=ARIAL, size=10)

        c4 = summary.cell(row=row, column=4, value=round(this_month_total, 2))
        c4.number_format = "#,##0.00"
        c4.font = Font(name=ARIAL, size=10)

        c5 = summary.cell(row=row, column=5, value=round(ytd_total, 2))
        c5.number_format = "#,##0.00"
        c5.font = Font(name=ARIAL, size=10)

        summary.cell(row=row, column=6, value=last_date.isoformat() if last_date else "").font = Font(name=ARIAL, size=10)

        for c in range(1, 7):
            summary.cell(row=row, column=c).border = BORDER
        row += 1

    widths = {"A": 10, "B": 38, "C": 16, "D": 24, "E": 24, "F": 20}
    for col, w in widths.items():
        summary.column_dimensions[col].width = w
    summary.freeze_panes = "A5"

    audit = wb.create_sheet("Audit Log")
    _header_row(audit, 1, [
        "Ticker", "Announcement Date", "Period", "PSX Details (raw)",
        "Cash Dividend Amount (PKR/share)", "Split-Adjusted Amount (PKR/share)",
        "Non-Cash Notes",
    ])
    row = 2
    all_entries = []
    for tkr, announcements in history.items():
        for a in announcements:
            all_entries.append((tkr, a))
    all_entries.sort(key=lambda pair: pair[1]["date_iso"], reverse=True)

    for tkr, a in all_entries:
        audit.cell(row=row, column=1, value=tkr).font = Font(name=ARIAL, size=10)
        audit.cell(row=row, column=2, value=a["date"]).font = Font(name=ARIAL, size=10)
        audit.cell(row=row, column=3, value=a["period"]).font = Font(name=ARIAL, size=10)
        audit.cell(row=row, column=4, value=a["raw_details"]).font = Font(name=ARIAL, size=10)
        audit.cell(row=row, column=5, value=a["amount_per_share"]).font = Font(name=ARIAL, size=10)
        audit.cell(row=row, column=6, value=adjusted_amount(a, tkr, splits)).font = Font(name=ARIAL, size=10)
        audit.cell(row=row, column=7, value=", ".join(a.get("non_cash_notes", []))).font = Font(name=ARIAL, size=9, color=GREY)
        for c in range(1, 8):
            audit.cell(row=row, column=c).border = BORDER
        row += 1

    widths2 = {"A": 10, "B": 22, "C": 16, "D": 22, "E": 20, "F": 22, "G": 18}
    for col, w in widths2.items():
        audit.column_dimensions[col].width = w
    audit.freeze_panes = "A2"

    reconciliation = wb.create_sheet("Cash Flow Reconciliation")
    reconciliation["A1"] = f"Cash Flow Reconciliation — {today.strftime('%B %Y')}"
    reconciliation["A1"].font = Font(name=ARIAL, size=14, bold=True, color=NAVY)
    reconciliation["A2"] = (
        f"Quantity sourced from {holdings.HOLDINGS_FILE}. Dividend Announced is this "
        "period's (current month's) actual per-share cash dividend. Tax withheld at "
        f"{TAX_WITHHOLDING_RATE:.0%}."
    )
    reconciliation["A2"].font = Font(name=ARIAL, size=9, italic=True, color=GREY)

    _header_row(reconciliation, 4, [
        "Ticker", "Company", "Quantity", "Dividend Announced (PKR/share)",
        "Gross Cash Dividend (PKR)", f"Tax Amount ({TAX_WITHHOLDING_RATE:.0%}) (PKR)",
        "Net Cash Dividend (PKR)",
    ])

    row = 5
    totals = {"quantity": 0, "gross": 0.0, "tax": 0.0, "net": 0.0}
    for entry in per_ticker:
        tkr = entry["ticker"]
        quantity = entry["quantity"]
        dividend = entry["dividend_this_month"]

        reconciliation.cell(row=row, column=1, value=tkr).font = Font(name=ARIAL, size=10, bold=True)
        reconciliation.cell(row=row, column=2, value=COMPANY_NAMES.get(tkr, tkr)).font = Font(name=ARIAL, size=10)

        cq = reconciliation.cell(row=row, column=3, value=quantity)
        cq.number_format = "#,##0"
        cq.font = Font(name=ARIAL, size=10)
        totals["quantity"] += quantity

        if dividend:
            gross = quantity * dividend
            tax = gross * TAX_WITHHOLDING_RATE
            net = gross - tax
            totals["gross"] += gross
            totals["tax"] += tax
            totals["net"] += net
            values = [dividend, gross, tax, net]
        else:
            values = ["-", "-", "-", "-"]

        for offset, value in enumerate(values, start=4):
            cell = reconciliation.cell(row=row, column=offset, value=value)
            cell.font = Font(name=ARIAL, size=10)
            if isinstance(value, (int, float)):
                cell.number_format = "#,##0.00"

        for c in range(1, 8):
            reconciliation.cell(row=row, column=c).border = BORDER
        row += 1

    reconciliation.cell(row=row, column=1, value="Total").font = Font(name=ARIAL, size=10, bold=True)
    total_values = [None, totals["quantity"], None, totals["gross"], totals["tax"], totals["net"]]
    for offset, value in enumerate(total_values, start=2):
        if value is None:
            continue
        cell = reconciliation.cell(row=row, column=offset, value=value)
        cell.font = Font(name=ARIAL, size=10, bold=True)
        cell.number_format = "#,##0" if offset == 3 else "#,##0.00"
    for c in range(1, 8):
        reconciliation.cell(row=row, column=c).border = BORDER
        reconciliation.cell(row=row, column=c).fill = PatternFill("solid", start_color="F2F2F2")

    widths3 = {"A": 10, "B": 32, "C": 14, "D": 22, "E": 20, "F": 18, "G": 20}
    for col, w in widths3.items():
        reconciliation.column_dimensions[col].width = w
    reconciliation.freeze_panes = "A5"

    wb.save(WORKBOOK_PATH)
    return WORKBOOK_PATH
