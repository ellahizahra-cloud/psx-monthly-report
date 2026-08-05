"""
run_dividend_tracker.py — Entry point for the daily dividend tracker cron
job.

Each run does a lightweight check per ticker (PSX only returns the most
recent handful of payout announcements per symbol — never a full-history
scrape) and diffs against tracking_log.json (the record of announcements
already processed) to find genuinely new announcements. Ticker universe
is DIVIDEND_TICKERS (holdings only) — narrower than the price report's
full universe; see tickers.py.

Every new cash-dividend announcement is classified by the fiscal period
it was actually declared for (dividend_period.classify, which opens the
linked PSX announcement PDF) — only announcements for the current year
count toward "Dividend This Month"/"Dividend YTD"; prior-year finals
announced this year are recorded but excluded, and anything that can't be
reliably classified (scanned PDF, unparseable text, no matching PDF) is
flagged "needs_review" rather than guessed at.

In addition to PSX's Payouts table (the primary source), each ticker is
also checked against the Financial Results/Board Meetings announcement
tabs directly (dividend_period.check_supplementary_announcements) — the
Payouts table has been observed to lag a day or more behind the
underlying announcement (e.g. FFC's Jul 29, 2026 interim dividend). Finds
from this channel are recorded with source="financial_results_tab" and,
when the PDF is scanned or otherwise unparseable, "provisional": True and
status "needs_review" (never guessed) with direct PDF/image links for
manual confirmation. If the Payouts table later confirms the same date,
the provisional entry is replaced by the authoritative one rather than
double-counted.

- No new announcements anywhere -> do nothing. No email, no file change.
- One or more new announcements -> update the persistent
  PSX_Dividend_Tracker.xlsx, send one email listing exactly which
  ticker(s) triggered the update and what was announced, and record the
  new announcements in tracking_log.json so they aren't re-flagged.

Whole run wrapped in try/except: on failure, send an error notice instead
of leaving a half-updated tracker.
"""

import json
import sys
from pathlib import Path

import dividend_period
from dividend_fetch import fetch_recent
from dividend_workbook import build
from mailer import send_email, send_error_notice
from tickers import DIVIDEND_TICKERS, COMPANY_NAMES

LOG_PATH = Path("tracking_log.json")


def load_history() -> dict:
    if LOG_PATH.exists():
        return json.loads(LOG_PATH.read_text())
    return {}


def save_history(history: dict):
    LOG_PATH.write_text(json.dumps(history, indent=2))


def run():
    history = load_history()
    new_by_ticker = {}
    fetch_errors = []

    for ticker in DIVIDEND_TICKERS:
        known_dates = {a["date"] for a in history.get(ticker, [])}
        try:
            announcements = fetch_recent(ticker)
        except Exception as e:
            fetch_errors.append(str(e))
            continue

        new_entries = [a for a in announcements if a["date"] not in known_dates]
        for a in new_entries:
            if a["cash_dividend_pct"] > 0:
                a["period_classification"] = dividend_period.classify(
                    ticker, a["date_iso"], payout_period=a["period"]
                )
            else:
                a["period_classification"] = None  # non-cash payout, not subject to the filter

        if new_entries:
            # If the Payouts table has now confirmed a date we'd earlier
            # only had via the announcement-tab fallback, drop the
            # provisional entry so it isn't double-counted alongside the
            # authoritative one.
            confirmed_isos = {a["date_iso"] for a in new_entries}
            history[ticker] = [
                e for e in history.get(ticker, [])
                if not (e.get("provisional") and e["date_iso"] in confirmed_isos)
            ]
            history.setdefault(ticker, [])
            history[ticker].extend(new_entries)
            new_by_ticker[ticker] = new_entries

        # Supplementary channel: scan the Financial Results/Board Meetings
        # tabs directly, independent of the Payouts table, to catch
        # announcements it hasn't indexed yet. known_isos reflects history
        # as of just above, so anything the Payouts check just added this
        # run is already excluded.
        known_isos = {a["date_iso"] for a in history.get(ticker, [])}
        try:
            supplementary = dividend_period.check_supplementary_announcements(ticker, known_isos)
        except Exception as e:
            supplementary = []
            fetch_errors.append(f"{ticker} (announcement tabs check): {e}")

        for s in supplementary:
            ocr_suffix = " (OCR)" if s.get("ocr") else ""
            entry = {
                "ticker": ticker,
                "date": s["date"],
                "date_iso": s["date_iso"],
                "period": s["title"],
                "raw_details": s["title"],
                "cash_dividend_pct": None,
                "amount_per_share": s["amount_per_share"],
                "non_cash_notes": [],
                "source": "financial_results_tab",
                "provisional": True,
                "period_classification": {
                    "status": s["status"],
                    "period_label": (s["period_label"] or s["reason"]) + ocr_suffix,
                    "period_end_date": s["period_end_date"],
                    "pdf_url": s["pdf_url"],
                    "image_url": s["image_url"],
                    "reason": s["reason"],
                    "source": "financial_results_tab_ocr" if s.get("ocr") else "financial_results_tab",
                },
            }
            history.setdefault(ticker, [])
            history[ticker].append(entry)
            new_by_ticker.setdefault(ticker, []).append(entry)

    if fetch_errors and not new_by_ticker:
        # Nothing new to report, but something was wrong with the scan itself.
        raise RuntimeError("; ".join(fetch_errors))

    if not new_by_ticker:
        print("No new dividend announcements — nothing to do.")
        return

    xlsx_path = build(history)

    lines = []
    for ticker, entries in new_by_ticker.items():
        company = COMPANY_NAMES.get(ticker, ticker)
        for a in entries:
            if a.get("source") == "financial_results_tab":
                pc = a["period_classification"]
                if a["amount_per_share"] is not None:
                    lines.append(
                        f"- {ticker} ({company}): possible cash dividend PKR {a['amount_per_share']}/share "
                        f"found via Financial Results/Board Meetings tab (PSX Payouts table hasn't "
                        f"listed it yet), announced {a['date']} [\"{a['raw_details']}\"] — {pc['period_label']}. "
                        f"Source: {pc['pdf_url']}"
                    )
                else:
                    links = pc["pdf_url"]
                    if pc.get("image_url"):
                        links += f" | scanned page: {pc['image_url']}"
                    lines.append(
                        f"- {ticker} ({company}): NEEDS REVIEW — {pc['reason']}, "
                        f"announced {a['date']} [\"{a['raw_details']}\"]. {links}"
                    )
                continue
            if a["cash_dividend_pct"] <= 0:
                lines.append(
                    f"- {ticker} ({company}): non-cash payout {a['raw_details']}, "
                    f"announced {a['date']} [{a['period']}] — not counted in dividend totals"
                )
                continue
            pc = a["period_classification"]
            lines.append(
                f"- {ticker} ({company}): {a['cash_dividend_pct']}% cash dividend "
                f"= PKR {a['amount_per_share']}/share, announced {a['date']} "
                f"[PSX label {a['period']}] — {pc['period_label']}"
            )

    body = (
        "Hi,\n\n"
        "New PSX dividend announcement(s) detected. The tracker has been updated "
        "and the current file is attached.\n\n"
        "Triggered by:\n" + "\n".join(lines) + "\n\n"
        "Only announcements for the current fiscal year count toward Dividend This "
        "Month/YTD; excluded and needs-review entries are still logged for "
        "auditability (see the Audit Log column) but not summed. Items found via the "
        "Financial Results/Board Meetings tabs (rather than PSX's Payouts table) are "
        "marked provisional and need manual confirmation if flagged NEEDS REVIEW.\n\n"
        "Best,\nAutomated PSX Dividend Tracker"
    )
    if fetch_errors:
        body += "\n\nNote: some tickers could not be checked this run:\n" + "\n".join(fetch_errors)

    tickers_label = ", ".join(new_by_ticker.keys())
    send_email(
        subject=f"PSX Dividend Update — {tickers_label}",
        body=body,
        attachment_path=xlsx_path,
    )

    save_history(history)
    print(f"Sent dividend update for: {tickers_label}")


if __name__ == "__main__":
    try:
        run()
    except Exception as e:
        send_error_notice("PSX Dividend Tracker", e)
        print(f"FAILED: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(1)
