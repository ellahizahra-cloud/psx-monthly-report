"""
send_report.py — Email the monthly PSX Excel report via Gmail SMTP.

Usage:
    python fetch_prices.py
    python build_report.py
    python send_report.py                   # auto-finds the latest xlsx
    python send_report.py --file path.xlsx   # or point at a specific one

For the scheduled/catch-up run (which also skips already-sent months and
sends an error notice on failure instead of a broken file), use
run_monthly_report.py instead of calling this directly.
"""

import argparse
import glob
import sys
from pathlib import Path

from mailer import send_email

SUBJECT_TEMPLATE = "Monthly PSX Share Price Report — {label}"
BODY_TEMPLATE = (
    "Hi,\n\n"
    "Attached is the monthly share price report for the full PSX portfolio "
    "(46 tickers across 9 sectors), covering the Month-Start close and the "
    "prior month's Month-End close.\n\n"
    "This was generated and cross-checked automatically; please flag anything "
    "that looks off.\n\n"
    "Best,\nAutomated PSX Report"
)


def latest_report_file() -> str:
    candidates = sorted(glob.glob("PSX_Share_Prices_*.xlsx"))
    if not candidates:
        sys.exit("No PSX_Share_Prices_*.xlsx file found. Run build_report.py first.")
    return candidates[-1]


def send(xlsx_path: str):
    path = Path(xlsx_path)
    if not path.exists():
        sys.exit(f"File not found: {xlsx_path}")

    label = path.stem.replace("PSX_Share_Prices_", "")
    send_email(
        subject=SUBJECT_TEMPLATE.format(label=label),
        body=BODY_TEMPLATE.format(label=label),
        attachment_path=str(path),
    )
    print(f"Sent {path.name}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="Path to the xlsx to send (default: latest PSX_Share_Prices_*.xlsx)")
    args = ap.parse_args()
    send(args.file or latest_report_file())
