"""
fetch_prices.py — Pull PSX closing prices for the monthly report.

The report always covers one fully-closed month — the month that just
ended relative to when the report runs (run_monthly_report.py works this
out; e.g. a report running Aug 2-5 covers July). For each ticker in the
shared universe (tickers.py) it fetches, from the PSX Data Portal's
end-of-day timeseries feed (the same JSON feed the dps.psx.com.pk website
uses), both prices from *that same month*:

  (a) Month-Start Close: close on the 1st trading day of that month
  (b) Month-End Close: close on the last trading day of that month

then cross-checks the Month-End Close (the most recent of the two, so the
closest to "now") against sarmaaya.pk and flags divergence > 1%.

Output: prices.json (consumed by build_report.py / send_report.py)

NOTE FOR CLAUDE CODE: if the PSX endpoint shape has changed, open
https://dps.psx.com.pk/company/OGDC in the browser tool, watch the network
requests, and adapt PSX_EOD_URL / parse_eod() accordingly.
"""

import argparse
import datetime as dt
import json
import re
import sys
import urllib.request

from tickers import TICKERS

PSX_EOD_URL = "https://dps.psx.com.pk/timeseries/eod/{symbol}"
SARMAAYA_URL = "https://sarmaaya.pk/stocks/{symbol}"
DIVERGENCE_THRESHOLD = 0.01  # 1%
HEADERS = {"User-Agent": "Mozilla/5.0 (monthly-report-bot)"}


def http_get(url: str) -> bytes:
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def parse_eod(raw: bytes):
    """PSX EOD feed -> {date: close}. Feed rows are [unix_ts, close, volume]."""
    payload = json.loads(raw)
    rows = payload["data"] if isinstance(payload, dict) else payload
    out = {}
    for row in rows:
        ts, close = row[0], float(row[1])
        d = dt.datetime.fromtimestamp(ts, dt.timezone.utc).date()
        out[d] = close
    return dict(sorted(out.items()))


def first_trading_close(series: dict, year: int, month: int):
    """Close on the first trading day of the given month."""
    for d, px in series.items():
        if d.year == year and d.month == month:
            return d, px
    return None, None


def last_trading_close(series: dict, year: int, month: int):
    """Close on the last trading day of the given month."""
    result = None
    for d, px in series.items():
        if d.year == year and d.month == month:
            result = (d, px)
    return result if result else (None, None)


def prev_month(year: int, month: int):
    return (year - 1, 12) if month == 1 else (year, month - 1)


def sarmaaya_price(symbol: str):
    """Sarmaaya server-renders 'SYMBOL - Rs 348.72 | ...' in the page title."""
    try:
        html = http_get(SARMAAYA_URL.format(symbol=symbol)).decode("utf-8", "ignore")
        m = re.search(r"Rs\.?\s*([\d,]+\.?\d*)", html)
        return float(m.group(1).replace(",", "")) if m else None
    except Exception:
        return None


def fetch(year: int, month: int):
    """Fetch Month-Start / Month-End closes, both from the same (year, month)
    — the month being reported on, which must already be fully closed by
    the time this runs (run_monthly_report.py works that out)."""
    results, warnings = [], []
    for sym in TICKERS:
        try:
            series = parse_eod(http_get(PSX_EOD_URL.format(symbol=sym)))
        except Exception as e:
            warnings.append(f"{sym}: PSX fetch FAILED ({e}) — fill manually")
            results.append({"ticker": sym, "error": str(e)})
            continue

        start_date, start_px = first_trading_close(series, year, month)
        end_date, end_px = last_trading_close(series, year, month)

        if start_px is None:
            warnings.append(f"{sym}: no trading data for {year:04d}-{month:02d} (start) — fill manually")
        if end_px is None:
            warnings.append(f"{sym}: no trading data for {year:04d}-{month:02d} (end) — fill manually")

        # Cross-check the Month-End close (the more recent of the two
        # values, typically only a few days old by send time) against
        # Sarmaaya's live price — the Month-Start close is weeks old by
        # then and isn't meaningful to compare against a live quote.
        check = sarmaaya_price(sym) if end_px is not None else None
        flagged = False
        if check and end_px and abs(check - end_px) / end_px > DIVERGENCE_THRESHOLD:
            flagged = True
            warnings.append(
                f"{sym}: PSX {end_px} vs Sarmaaya {check} diverge >1% — verify before sending"
            )

        results.append({
            "ticker": sym,
            "month_start_date": start_date.isoformat() if start_date else None,
            "month_start_close": start_px,
            "month_end_date": end_date.isoformat() if end_date else None,
            "month_end_close": end_px,
            "sarmaaya_check": check,
            "flagged": flagged,
        })

    return results, warnings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", help="Month to report on, YYYY-MM (default: last fully-closed month)")
    args = ap.parse_args()

    today = dt.date.today()
    if args.month:
        year, month = map(int, args.month.split("-"))
    else:
        year, month = prev_month(today.year, today.month)

    results, warnings = fetch(year, month)

    out = {
        "generated": dt.datetime.now().isoformat(timespec="seconds"),
        "reference_month": f"{year:04d}-{month:02d}",
        "results": results,
        "warnings": warnings,
    }
    with open("prices.json", "w") as f:
        json.dump(out, f, indent=2)

    print(json.dumps(out, indent=2))
    if warnings:
        print("\n*** WARNINGS — review before the email goes out ***", file=sys.stderr)
        sys.exit(2)  # non-zero so the routine pauses instead of auto-sending


if __name__ == "__main__":
    main()
