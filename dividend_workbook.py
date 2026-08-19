"""
dividend_workbook.py — Build/rewrite the persistent PSX_Dividend_Tracker.xlsx
from the tracking log (the single source of truth for announcement history).

Single sheet, one row per ticker (DIVIDEND_TICKERS — the holdings-only
universe, narrower than the price report's; see tickers.py):
  Ticker, Company, Sector, Quantity (from holdings.py / Cash Dividend 1
  (1).xlsx), Q1/Q2/Q3/Q4 (PKR/share — each quarter's own declared amount,
  populated once that quarter's dividend has actually been announced;
  blank otherwise), Dividend YTD (= SUM of the four quarters, resets every
  January since it's computed from the current calendar year's
  announcements), Gross/Tax/Net Cash Dividend, Last Announcement Date, and
  an Audit Log column (every Date: Amount [Period] on record for that
  ticker, for auditability). Tickers with no current-year dividend at all
  show "-" for the dividend-derived columns (never 0) but Quantity always
  populates. A Total row sums Quantity, Gross, Tax and Net across all
  tickers (summing the per-share quarter/YTD columns across different
  companies wouldn't mean anything, so those are left out of the total).

  An entry is bucketed into a quarter by its own fiscal period's end date
  (e.g. "quarter ended June 30" -> Q2), not by the date PSX announced it —
  consistent with the fiscal-period filtering below. Entries from before
  that feature existed (no period_end_date on record) fall back to the
  announcement date's own quarter as the best available proxy.

  Q1-Q4/YTD/Gross/Tax/Net are all written as live Excel formulas (YTD =
  SUM of the four quarters, Gross = Quantity x YTD, Tax = Gross x the
  tax-rate cell B3, Net = Gross - Tax; Total row is =SUM(...) over the
  data rows), not baked-in values — an intentional exception to the
  "write final values, no formula/recalc dependency" rule elsewhere in
  this project, since this file is only ever consumed by opening it
  directly in Excel/LibreOffice (both recalculate on open), and being
  able to see/audit how each number was derived matters more here than
  avoiding a recalc dependency.

Only announcements whose underlying fiscal period (per
dividend_period.classify) falls in the current year count toward the
quarter columns / YTD / Gross/Tax/Net — a prior-year final dividend
announced this year is logged in the Audit Log column (so it's still
auditable) but excluded from the sums, and anything needs_review is
likewise logged but not summed. Entries pre-dating the period-
classification feature (no "period_classification" recorded) fall back to
being counted, to avoid retroactively changing historical totals.

All figures are actual announced per-share cash-dividend amounts
(PSX payout % x face value — see FACE_VALUE in tickers.py, since PSX
reports a percentage and this project has caught cases where the
blanket-Rs10 assumption was wrong, e.g. BAFL's real face value is Rs5),
adjusted for any stock splits recorded in stock_splits.json so historical
amounts sit on a current-share-count basis. Never TTM/annualized/"last
declared x4".
"""

import datetime as dt
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import holdings
from tickers import DIVIDEND_UNIVERSE, COMPANY_NAMES, SECTORS

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


def adjusted_amount(entry: dict, ticker: str, splits: dict):
    """None when the amount itself is unknown (a needs_review entry from
    the supplementary announcement-tab channel, e.g. a scanned PDF with no
    confirmed figure yet) — never coerced to 0, which would be indistinguishable
    from a genuine NIL dividend."""
    if entry["amount_per_share"] is None:
        return None
    factor = split_adjustment_factor(ticker, entry["date_iso"], splits)
    return round(entry["amount_per_share"] / factor, 4)


def _counts_toward_totals(entry: dict) -> bool:
    """Only announcements for the current fiscal year count toward
    Dividend This Month/YTD/Dividend Announced. Entries pre-dating the
    period-classification feature (no key at all) are counted as before,
    so existing totals aren't retroactively changed."""
    if "period_classification" not in entry:
        return True
    pc = entry["period_classification"]
    return pc is None or pc["status"] == "included"


def _audit_label(entry: dict) -> str:
    pc = entry.get("period_classification")
    return f" [{pc['period_label']}]" if pc else ""


def _entry_quarter(entry: dict) -> int:
    """Which calendar quarter (1-4) this entry's own fiscal period belongs
    in, by its period end-date (e.g. 'quarter ended June 30' -> Q2) —
    never by the date PSX happened to announce it. Falls back to the
    announcement date's own quarter for entries that pre-date the
    period-classification feature (no period_end_date on record)."""
    pc = entry.get("period_classification")
    if pc and pc.get("period_end_date"):
        end_date = dt.date.fromisoformat(pc["period_end_date"])
    else:
        end_date = dt.date.fromisoformat(entry["date_iso"])
    return (end_date.month - 1) // 3 + 1


def _header_row(ws, row, headers):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def build(history: dict) -> str:
    """Gross/Tax/Net are always computed from Dividend YTD so the file is
    meaningful whether it was just rebuilt from a live trigger or opened
    cold days later (rather than tied to whichever announcement happened
    to trigger the current run — see run_dividend_tracker.py for how new
    announcements are still listed in the notification email)."""
    splits = load_splits()
    today = dt.date.today()

    wb = Workbook()
    ws = wb.active
    ws.title = "Dividend Tracker"

    ws["A1"] = "PSX Dividend Tracker"
    ws["A1"].font = Font(name=ARIAL, size=14, bold=True, color=NAVY)
    ws["A2"] = (
        f"Last updated {dt.datetime.now().isoformat(timespec='seconds')}. Source: PSX "
        "company payout announcements (board-meeting/announcement date). Quantity "
        f"sourced from {holdings.HOLDINGS_FILE}. Q1-Q4 are each quarter's own declared "
        f"dividend (by fiscal period end-date, not announcement date). Gross/Tax/Net "
        f"Cash Dividend are Quantity x Dividend YTD {today.year} (see formulas in I:L; "
        "tax rate in B3)."
    )
    ws["A2"].font = Font(name=ARIAL, size=9, italic=True, color=GREY)

    ws["A3"] = "Tax Rate:"
    ws["A3"].font = Font(name=ARIAL, size=9, bold=True, color=GREY)
    ws["B3"] = TAX_WITHHOLDING_RATE
    ws["B3"].number_format = "0%"
    ws["B3"].font = Font(name=ARIAL, size=9, bold=True, color=NAVY)

    headers = [
        "Ticker", "Company", "Sector", "Quantity",
        f"Q1 {today.year} (PKR/share)",
        f"Q2 {today.year} (PKR/share)",
        f"Q3 {today.year} (PKR/share)",
        f"Q4 {today.year} (PKR/share)",
        f"Dividend YTD {today.year} (PKR/share)",
        "Gross Cash Dividend (PKR)",
        f"Tax Amount ({TAX_WITHHOLDING_RATE:.0%}) (PKR)",
        "Net Cash Dividend (PKR)",
        "Last Announcement Date",
        "Audit Log (Date: Amount/share)",
    ]
    _header_row(ws, 4, headers)

    quantities = holdings.load_quantities()

    row = 5
    for entry in DIVIDEND_UNIVERSE:
        tkr = entry["ticker"]
        announcements = sorted(history.get(tkr, []), key=lambda a: a["date_iso"], reverse=True)

        quarter_totals = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
        last_date = None
        audit_parts = []
        for a in announcements:
            a_date = dt.date.fromisoformat(a["date_iso"])
            amt = adjusted_amount(a, tkr, splits)
            if amt is None:
                audit_parts.append(f"{a['date_iso']}: unconfirmed{_audit_label(a)}")
            else:
                audit_parts.append(f"{a['date_iso']}: Rs {amt}{_audit_label(a)}")
                if a_date.year == today.year and _counts_toward_totals(a):
                    quarter_totals[_entry_quarter(a)] += amt
            if last_date is None or a_date > last_date:
                last_date = a_date

        quantity = quantities.get(tkr, 0)

        ws.cell(row=row, column=1, value=tkr).font = Font(name=ARIAL, size=10, bold=True)
        ws.cell(row=row, column=2, value=COMPANY_NAMES.get(tkr, tkr)).font = Font(name=ARIAL, size=10)
        ws.cell(row=row, column=3, value=SECTORS.get(tkr, "")).font = Font(name=ARIAL, size=10)

        cq = ws.cell(row=row, column=4, value=quantity)
        cq.number_format = "#,##0"
        cq.font = Font(name=ARIAL, size=10)

        for offset, q in enumerate((1, 2, 3, 4), start=5):
            cell = ws.cell(row=row, column=offset, value=round(quarter_totals[q], 2) if quarter_totals[q] else "-")
            cell.number_format = "#,##0.00"
            cell.font = Font(name=ARIAL, size=10)

        # Real formulas (not baked-in values) so YTD/Gross/Tax/Net are
        # auditable in Excel: YTD = SUM of the four quarter columns
        # (treating "-" as 0), Gross = Quantity x YTD, Tax = Gross x
        # tax-rate cell, Net = Gross - Tax. IF(...) keeps the "-"
        # convention for tickers with no counted current-year dividend
        # instead of showing 0.
        c9 = ws.cell(row=row, column=9, value=f"=SUM(E{row}:H{row})")
        c9.font = Font(name=ARIAL, size=10)
        c9.number_format = "#,##0.00"

        formulas = [
            f'=IF(I{row}=0,"-",D{row}*I{row})',
            f'=IF(I{row}=0,"-",J{row}*$B$3)',
            f'=IF(I{row}=0,"-",J{row}-K{row})',
        ]
        for offset, formula in enumerate(formulas, start=10):
            cell = ws.cell(row=row, column=offset, value=formula)
            cell.font = Font(name=ARIAL, size=10)
            cell.number_format = "#,##0.00"

        ws.cell(row=row, column=13, value=last_date.isoformat() if last_date else "").font = Font(name=ARIAL, size=10)
        ws.cell(row=row, column=14, value="; ".join(audit_parts)).font = Font(name=ARIAL, size=8.5, color=GREY)

        for c in range(1, 15):
            ws.cell(row=row, column=c).border = BORDER
        row += 1

    last_data_row = row - 1
    ws.cell(row=row, column=1, value="Total").font = Font(name=ARIAL, size=10, bold=True)
    total_formulas = {
        4: f"=SUM(D5:D{last_data_row})",
        10: f"=SUM(J5:J{last_data_row})",
        11: f"=SUM(K5:K{last_data_row})",
        12: f"=SUM(L5:L{last_data_row})",
    }
    for offset, formula in total_formulas.items():
        cell = ws.cell(row=row, column=offset, value=formula)
        cell.font = Font(name=ARIAL, size=10, bold=True)
        cell.number_format = "#,##0" if offset == 4 else "#,##0.00"
    for c in range(1, 15):
        ws.cell(row=row, column=c).border = BORDER
        ws.cell(row=row, column=c).fill = PatternFill("solid", start_color="F2F2F2")

    widths = {
        "A": 10, "B": 34, "C": 16, "D": 12, "E": 14, "F": 14,
        "G": 14, "H": 14, "I": 18, "J": 20, "K": 16, "L": 18,
        "M": 18, "N": 60,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    wb.save(WORKBOOK_PATH)
    return WORKBOOK_PATH
