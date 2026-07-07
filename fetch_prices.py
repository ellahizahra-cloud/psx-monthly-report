"""
fetch_prices.py — Pull PSX closing prices (and optional dividend payouts)
for the monthly report.

For each ticker it fetches:
  (a) the closing price on the 1st trading day of the reference month
      (defaults to the current month; pass --month YYYY-MM to override)
  (b) the latest available close
  (c) if --div-start/--div-end are given, total cash dividend payout
      (PKR/share) with an ex-dividend date in that range
from the PSX Data Portal, then cross-checks the latest price against
sarmaaya.pk and flags divergence > 1%.

Output: prices.json (consumed by build_report.py / send_report.py)

NOTE FOR CLAUDE CODE: if the PSX endpoint shapes have changed, open
https://dps.psx.com.pk/company/OGDC in the browser tool, watch the network
requests, and adapt PSX_EOD_URL / parse_eod() and PSX_DIVIDEND_URL /
parse_dividends() accordingly. Dividend data may only be available as PDF
board announcements rather than a clean JSON feed — if so, fall back to
scstrade.com or sarmaaya.pk, which tend to list dividend history in table
form, and adapt parse_dividends() to scrape those instead.
"""

import argparse
import datetime as dt
import json
import re
import sys
import urllib.request

# --- Tickers grouped by sector. Order here controls report order. ---
SECTORS = {
    "Banks": ["MCB", "HBL", "BAFL", "BAHL", "MEBL", "HMB", "UBL"],
    "Cement": ["LUCK", "KOHC", "DGKC", "CHCC", "FCCL", "MLCF"],
    "Fertilizers": ["ENGROH", "EFERT", "FATIMA", "FFC"],
    "Oil/Gas": ["POL", "PSO", "MARI", "PPL", "OGDC"],
    "Power": ["ALTN", "NCPL", "KOHE", "NPL", "HUBC"],
    "Steel": ["ISL", "AGHA", "ASTL", "ASL"],
    "Others": ["CEPB", "INDU", "NML", "AGIL", "ORIX", "EPCL"],
    "Pharma": ["ABOT", "FEROZ", "GLAXO", "HINOON", "CPHL", "AGP"],
    "Tech": ["SYS", "TRG", "AVN"],
}

# Flattened (ticker, sector) list, built from SECTORS above.
TICKERS = [(t, sector) for sector, tickers in SECTORS.items() for t in tickers]

PSX_EOD_URL = "https://dps.psx.com.pk/timeseries/eod/{symbol}"
PSX_DIVIDEND_URL = "https://dps.psx.com.pk/company/{symbol}"  # placeholder — see note above
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


def sarmaaya_price(symbol: str):
    """Sarmaaya server-renders 'SYMBOL - Rs 348.72 | ...' in the page title."""
    try:
        html = http_get(SARMAAYA_URL.format(symbol=symbol)).decode("utf-8", "ignore")
        m = re.search(r"Rs\.?\s*([\d,]+\.?\d*)", html)
        return float(m.group(1).replace(",", "")) if m else None
    except Exception:
        return None


def parse_dividends(symbol: str, start: dt.date, end: dt.date):
    """
    Fetch cash dividend announcements for `symbol` and sum PKR/share payouts
    with an ex-dividend/book-closure date in [start, end].

    PLACEHOLDER: PSX's per-company corporate-actions data isn't confirmed to
    be available as clean JSON. This function currently returns None
    (meaning "not fetched") so the rest of the pipeline degrades gracefully.
    Claude Code: inspect https://dps.psx.com.pk/company/{symbol} network
    requests (or the PSX announcements/notices page) for a corporate-actions
    endpoint, and replace this implementation. If PSX doesn't expose it
    cleanly, fall back to scraping scstrade.com's dividend history table for
    the symbol instead.
    """
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", help="Reference month YYYY-MM (default: current)")
    ap.add_argument("--div-start", help="Dividend window start date YYYY-MM-DD")
    ap.add_argument("--div-end", help="Dividend window end date YYYY-MM-DD")
    args = ap.parse_args()

    today = dt.date.today()
    if args.month:
        year, month = map(int, args.month.split("-"))
    else:
        year, month = today.year, today.month

    div_start = dt.date.fromisoformat(args.div_start) if args.div_start else None
    div_end = dt.date.fromisoformat(args.div_end) if args.div_end else None
    want_dividends = div_start and div_end

    results, warnings = [], []
    for sym, sector in TICKERS:
        try:
            series = parse_eod(http_get(PSX_EOD_URL.format(symbol=sym)))
        except Exception as e:
            warnings.append(f"{sym}: PSX fetch FAILED ({e}) — fill manually")
            results.append({"ticker": sym, "sector": sector, "error": str(e)})
            continue

        d1, px1 = first_trading_close(series, year, month)
        last_date, last_px = max(series.items())

        check = sarmaaya_price(sym)
        flagged = False
        if check and last_px and abs(check - last_px) / last_px > DIVERGENCE_THRESHOLD:
            flagged = True
            warnings.append(
                f"{sym}: PSX {last_px} vs Sarmaaya {check} diverge >1% — verify before sending"
            )

        dividend = None
        if want_dividends:
            dividend = parse_dividends(sym, div_start, div_end)
            if dividend is None:
                warnings.append(f"{sym}: dividend data not available — needs manual entry or endpoint fix")

        results.append({
            "ticker": sym,
            "sector": sector,
            "month_start_date": d1.isoformat() if d1 else None,
            "month_start_close": px1,
            "latest_date": last_date.isoformat(),
            "latest_close": last_px,
            "sarmaaya_check": check,
            "flagged": flagged,
            "dividend_payout": dividend,
        })

    out = {
        "generated": dt.datetime.now().isoformat(timespec="seconds"),
        "reference_month": f"{year:04d}-{month:02d}",
        "dividend_window": (
            {"start": div_start.isoformat(), "end": div_end.isoformat()}
            if want_dividends else None
        ),
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
