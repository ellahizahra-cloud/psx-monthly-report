"""
dividend_fetch.py — Pull recent payout (dividend) announcements per ticker
from the PSX Data Portal's company payouts endpoint — the same AJAX call
dps.psx.com.pk/payouts makes when you search a symbol
(POST https://dps.psx.com.pk/company/payouts, form field "symbol").

Consistently sourced from PSX for every ticker (never mixed with an
aggregator) so board-meeting/announcement dates and amounts stay
comparable across the whole universe.

Each returned row's "Date" column is the announcement/board-meeting
datetime PSX stamps on the outcome — this is what counts as the dividend
date per the spec (not book closure/ex-date/credit date).

Only cash dividends ("(D)" flag) are counted. Bonus shares ("(B)"),
right issues ("(R)") and other non-cash payouts are parsed but excluded
from the per-share rupee amount, and noted for audit.

NOTE FOR CLAUDE CODE: if PSX changes this endpoint's shape, open
https://dps.psx.com.pk/payouts in the browser tool, search a symbol, and
watch the network request POST https://dps.psx.com.pk/company/payouts.
"""

import datetime as dt
import re
import urllib.error
import urllib.request

from tickers import FACE_VALUE

PAYOUTS_URL = "https://dps.psx.com.pk/company/payouts"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (dividend-tracker-bot)",
    "X-Requested-With": "XMLHttpRequest",
}

ROW_RE = re.compile(
    r"<tr><td>(?P<date>[^<]+)</td><td>(?P<period>[^<]*)</td>"
    r"<td>(?P<details>[^<]*)</td><td>(?P<book_closure>[^<]*)</td></tr>"
)
ENTRY_RE = re.compile(r"([\d.]+)%\(([^)]*)\)\s*\(([A-Za-z]+)\)")
DATE_FORMAT = "%B %d, %Y %I:%M %p"


def _parse_date(date_str: str):
    return dt.datetime.strptime(date_str.strip(), DATE_FORMAT)


def fetch_payouts_html(symbol: str) -> str:
    data = f"symbol={symbol}".encode()
    req = urllib.request.Request(PAYOUTS_URL, data=data, headers=HEADERS, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8", "ignore")


def parse_payouts(html: str, ticker: str):
    """Return a list of announcement dicts, most recent first as given by PSX."""
    face_value = FACE_VALUE.get(ticker, 10.0)
    rows = []
    for m in ROW_RE.finditer(html):
        date_str = m.group("date").strip()
        try:
            announced_at = _parse_date(date_str)
        except ValueError:
            continue

        entries = ENTRY_RE.findall(m.group("details"))
        cash_pct_total = 0.0
        non_cash_notes = []
        for pct, _code, flag in entries:
            if flag.upper() == "D":
                cash_pct_total += float(pct)
            else:
                non_cash_notes.append(f"{pct}%({flag})")

        rows.append({
            "ticker": ticker,
            "date": date_str,
            "date_iso": announced_at.date().isoformat(),
            "period": m.group("period").strip(),
            "raw_details": m.group("details").strip(),
            "cash_dividend_pct": cash_pct_total,
            "amount_per_share": round(cash_pct_total / 100 * face_value, 4),
            "non_cash_notes": non_cash_notes,
        })
    return rows


def fetch_recent(ticker: str):
    try:
        html = fetch_payouts_html(ticker)
    except (urllib.error.URLError, TimeoutError) as e:
        raise RuntimeError(f"{ticker}: payouts fetch failed ({e})") from e
    return parse_payouts(html, ticker)
