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
                a["period_classification"] = dividend_period.classify(ticker, a["date_iso"])
            else:
                a["period_classification"] = None  # non-cash payout, not subject to the filter

        if new_entries:
            history.setdefault(ticker, [])
            history[ticker].extend(new_entries)
            new_by_ticker[ticker] = new_entries

    if fetch_errors and not new_by_ticker:
        # Nothing new to report, but something was wrong with the scan itself.
        raise RuntimeError("; ".join(fetch_errors))

    if not new_by_ticker:
        print("No new dividend announcements — nothing to do.")
        return

    xlsx_path = build(history, new_entries=new_by_ticker)

    lines = []
    for ticker, entries in new_by_ticker.items():
        company = COMPANY_NAMES.get(ticker, ticker)
        for a in entries:
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
        "auditability (see the Audit Log column) but not summed.\n\n"
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
