"""
run_monthly_report.py — Entry point for the monthly price report cron job.

The report always covers the month that just ended, not the one still in
progress — e.g. a run in early August reports on July's own Month-Start
(Jul 1) and Month-End (Jul 31) close, never August's first trading day.
That means the reported month is always fully closed by the time this
runs, so there's no "no trading data yet" gap to fill in.

Scheduled 8am on the 2nd of each month, but the workflow actually invokes
this daily on days 2-5 as a catch-up window: if a scheduled run was missed
(e.g. the runner was down), the next day's run notices last month's report
hasn't been sent yet and sends it. Once sent, later runs in the same
window are no-ops.

Whole run is wrapped in try/except: on failure, send an error notice email
instead of a broken/partial file.

Set FORCE_SEND=1 to bypass the day 2-5 window and the already-sent check
(e.g. for a manual test send) — the scheduled cron never sets this.
"""

import datetime as dt
import json
import os
import sys
from pathlib import Path

import build_report
import fetch_prices
from fetch_prices import prev_month
from mailer import send_error_notice
from send_report import send

STATE_PATH = Path("monthly_report_state.json")
CATCHUP_WINDOW = range(2, 6)  # days 2-5 inclusive


def _forced() -> bool:
    return os.environ.get("FORCE_SEND", "").lower() in ("1", "true", "yes")


def load_state() -> dict:
    if STATE_PATH.exists():
        return json.loads(STATE_PATH.read_text())
    return {}


def save_state(state: dict):
    STATE_PATH.write_text(json.dumps(state, indent=2))


def run():
    today = dt.date.today()
    ref_year, ref_month = prev_month(today.year, today.month)
    reference_month = f"{ref_year:04d}-{ref_month:02d}"
    forced = _forced()

    if not forced and today.day not in CATCHUP_WINDOW:
        print(f"Day {today.day} is outside the catch-up window (2-5) — nothing to do.")
        return

    state = load_state()
    if not forced and state.get("last_sent_month") == reference_month:
        print(f"{reference_month} report already sent — nothing to do.")
        return

    try:
        results, warnings = fetch_prices.fetch(ref_year, ref_month)
        prices_data = {
            "generated": dt.datetime.now().isoformat(timespec="seconds"),
            "reference_month": reference_month,
            "results": results,
            "warnings": warnings,
        }
        Path("prices.json").write_text(json.dumps(prices_data, indent=2))

        xlsx_path = build_report.build("prices.json")
        send(xlsx_path)

        state["last_sent_month"] = reference_month
        save_state(state)
        print(f"Sent {reference_month} report ({len(warnings)} warning(s)).")
    except Exception as e:
        send_error_notice("Monthly PSX Price Report", e)
        print(f"FAILED: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    run()
