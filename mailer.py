"""
mailer.py — Shared Gmail sending/drafting helper for both automations.

ONE-TIME SETUP
1. Google Account -> Security -> 2-Step Verification -> App passwords
   -> generate one for "Mail". This is NOT your normal Gmail password.
2. Set two environment variables (don't hardcode secrets):
     export GMAIL_SENDER="you@gmail.com"
     export GMAIL_APP_PASSWORD="xxxx xxxx xxxx xxxx"

send_email() sends immediately via SMTP. save_draft() instead writes the
same message straight into the Gmail account's Drafts folder via IMAP
APPEND, reusing the same app-password credentials — no separate OAuth
setup needed. Nothing goes out until a human opens the draft and hits
send.
"""

import imaplib
import os
import smtplib
import ssl
import sys
import time
from email.message import EmailMessage
from pathlib import Path

RECIPIENTS = [
    "nadeemellahi@hotmail.com",
    "zaraellahi9@gmail.com",
    "naveedellahi@hotmail.com",
]


def _build_message(subject: str, body: str, attachment_path: str | None = None):
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

    return msg, sender, app_password


def send_email(subject: str, body: str, attachment_path: str | None = None):
    msg, sender, app_password = _build_message(subject, body, attachment_path)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as s:
        s.login(sender, app_password)
        s.send_message(msg)


def _drafts_folder(imap: imaplib.IMAP4_SSL) -> str:
    """Find the account's Drafts folder by its \\Drafts special-use
    attribute rather than assuming the English "[Gmail]/Drafts" name,
    since Gmail localizes folder names by the account's language."""
    status, folders = imap.list()
    if status == "OK":
        for line in folders:
            decoded = line.decode() if isinstance(line, bytes) else line
            if "\\Drafts" in decoded:
                # Response form: (\Drafts \HasNoChildren) "/" "[Gmail]/Drafts"
                return decoded.rsplit(' "/" ', 1)[-1].strip('"')
    return '"[Gmail]/Drafts"'


def save_draft(subject: str, body: str, attachment_path: str | None = None):
    """Write the message directly into the Gmail Drafts folder instead of
    sending it — a human reviews and sends manually."""
    msg, sender, app_password = _build_message(subject, body, attachment_path)

    with imaplib.IMAP4_SSL("imap.gmail.com") as imap:
        imap.login(sender, app_password)
        folder = _drafts_folder(imap)
        imap.append(
            folder,
            r"(\Draft)",
            imaplib.Time2Internaldate(time.time()),
            msg.as_bytes(),
        )


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
