"""
send_report.py — Email the monthly PSX Excel report to a fixed recipient
list, with the file attached, via Gmail SMTP.

ONE-TIME SETUP
1. Google Account -> Security -> 2-Step Verification -> App passwords
   -> generate one for "Mail". This is NOT your normal Gmail password.
2. Set two environment variables (don't hardcode secrets in this file):
     export GMAIL_SENDER="you@gmail.com"
     export GMAIL_APP_PASSWORD="xxxx xxxx xxxx xxxx"
3. Edit RECIPIENTS below — this whitelist is what keeps the automation
   safe. The script will only ever send to these three addresses.

Usage:
    python fetch_prices.py
    python build_report.py
    python send_report.py                 # auto-finds the latest xlsx
    python send_report.py --file path.xlsx  # or point at a specific one
"""

import argparse
import glob
import os
import smtplib
import ssl
import sys
from email.message import EmailMessage
from pathlib import Path

# --- Fixed recipient whitelist. Edit these three, nothing else needed. ---
RECIPIENTS = [
    "nadeemellahi@hotmail.com",
    "zaraellahi9@gmail.com",
    
]

SUBJECT_TEMPLATE = "Monthly PSX Share Price Report — {label}"
BODY_TEMPLATE = (
    "Hi,\n\n"
    "Attached is the monthly share price report for OGDC, MARI, FATIMA and AATM, "
    "covering the 1st-of-month close and the latest available close.\n\n"
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
    sender = os.environ.get("GMAIL_SENDER")
    app_password = os.environ.get("GMAIL_APP_PASSWORD")
    if not sender or not app_password:
        sys.exit("Set GMAIL_SENDER and GMAIL_APP_PASSWORD environment variables first.")

    path = Path(xlsx_path)
    if not path.exists():
        sys.exit(f"File not found: {xlsx_path}")

    label = path.stem.replace("PSX_Share_Prices_", "")

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(RECIPIENTS)
    msg["Subject"] = SUBJECT_TEMPLATE.format(label=label)
    msg.set_content(BODY_TEMPLATE.format(label=label))

    msg.add_attachment(
        path.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name,
    )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as s:
        s.login(sender, app_password)
        s.send_message(msg)

    print(f"Sent {path.name} to {', '.join(RECIPIENTS)}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="Path to the xlsx to send (default: latest PSX_Share_Prices_*.xlsx)")
    args = ap.parse_args()
    send(args.file or latest_report_file())
