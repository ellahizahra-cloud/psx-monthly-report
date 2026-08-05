"""
mailer.py — Shared Gmail SMTP sending helper for both automations.

ONE-TIME SETUP
1. Google Account -> Security -> 2-Step Verification -> App passwords
   -> generate one for "Mail". This is NOT your normal Gmail password.
2. Set two environment variables (don't hardcode secrets):
     export GMAIL_SENDER="you@gmail.com"
     export GMAIL_APP_PASSWORD="xxxx xxxx xxxx xxxx"
"""

import os
import smtplib
import ssl
import sys
from email.message import EmailMessage
from pathlib import Path

RECIPIENTS = [
    "nadeemellahi@hotmail.com",
    "zaraellahi9@gmail.com",
    "naveedellahi@hotmail.com",
]


def send_email(subject: str, body: str, attachment_path: str | None = None):
    sender = os.environ.get("GMAIL_SENDER")
    app_password = os.environ.get("GMAIL_APP_PASSWORD")
    if not sender or not app_password:
        sys.exit("Set GMAIL_SENDER and GMAIL_APP_PASSWORD environment variables first.")

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(RECIPIENTS)
    msg["Subject"] = subject
    msg.set_content(body)

    if attachment_path:
        path = Path(attachment_path)
        msg.add_attachment(
            path.read_bytes(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=path.name,
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as s:
        s.login(sender, app_password)
        s.send_message(msg)


def send_error_notice(automation: str, error: Exception):
    """Wrap-and-notify: send an error email instead of a broken/partial file."""
    send_email(
        subject=f"[PSX Automation] {automation} FAILED",
        body=(
            f"The {automation} run failed and did not complete.\n\n"
            f"Error: {type(error).__name__}: {error}\n\n"
            "No file was sent. Please check the automation logs."
        ),
    )
