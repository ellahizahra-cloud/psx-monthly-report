"""
holdings.py — Loads share Quantity per ticker from the existing holdings
file, "Cash Dividend 1 (1).xlsx" (root of repo), for the cash-flow
reconciliation columns in the dividend tracker.

Per spec, only Quantity is pulled from this file — DPS Received, DPS
Receivable, Book Closed Date, Credit Expected Date and X-Date are
explicitly out of scope and never read.

The source file's "Company Name" column is actually a mix of tickers and
loosely-formatted short names (e.g. "Aissha Steel", "HINOON LAB", "Luck").
ALIASES maps each raw label to the real PSX ticker it refers to so the
match is explicit rather than fuzzy-guessed.
"""

from pathlib import Path

from openpyxl import load_workbook

from tickers import TICKERS

HOLDINGS_FILE = "Cash Dividend 1 (1).xlsx"
SHEET_NAME = "Sheet1 (2)"
COMPANY_NAME_HEADER = "Company Name"

# Raw label in the source file -> real PSX ticker (only where they differ
# from the ticker itself once uppercased/stripped).
ALIASES = {
    "PANTHER": "PTL",
    "AISSHA STEEL": "ASL",
    "BYCO": "BYCO",
    "PACKAGE": "PKGS",
    "HINOON LAB": "HINOON",
    "MAPLE": "MLCF",
    "LUCK": "LUCK",
}


def _resolve_ticker(raw_label: str) -> str | None:
    key = raw_label.strip().upper()
    if key in ALIASES:
        return ALIASES[key]
    if key in TICKERS:
        return key
    return None


def load_quantities() -> dict:
    """Returns {ticker: quantity}. Tickers not held (or not found in the
    source file) are simply absent — callers should default to 0."""
    path = Path(HOLDINGS_FILE)
    if not path.exists():
        return {}

    wb = load_workbook(path, data_only=True)
    ws = wb[SHEET_NAME]

    quantities = {}
    header_seen = False
    for row in ws.iter_rows(values_only=True):
        label, quantity = row[0], row[1]
        if not header_seen:
            if label == COMPANY_NAME_HEADER:
                header_seen = True
            continue
        if label is None or str(label).strip().lower() == "total":
            break
        ticker = _resolve_ticker(str(label))
        if ticker and isinstance(quantity, (int, float)):
            quantities[ticker] = quantities.get(ticker, 0) + quantity

    return quantities
