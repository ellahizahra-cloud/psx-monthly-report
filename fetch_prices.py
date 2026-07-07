"""
fetch_prices.py — Pull PSX closing prices for the monthly report.

For each ticker it fetches:
  (a) the closing price on the 1st trading day of the reference month
      (defaults to the current month; pass --month YYYY-MM to override)
  (b) the latest available close
from the PSX Data Portal's end-of-day timeseries feed (the same JSON feed
the dps.psx.com.pk website uses), then cross-checks the latest price
against sarmaaya.pk and flags divergence > 1%.

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

TICKERS = [  # (ticker, sector)
    ("MCB", "Banks"), ("HBL", "Banks"), ("BAFL", "Banks"), ("BAHL", "Banks"),
    ("MEBL", "Banks"), ("HMB", "Banks"), ("UBL", "Banks"),
    ("LUCK", "Cement"), ("KOHC", "Cement"), ("DGKC", "Cement"), ("CHCC", "Cement"),
    ("FCCL", "Cement"), ("MLCF", "Cement"),
    ("ENGROH", "Fertilizers"), ("EFERT", "Fertilizers"), ("FATIMA", "Fertilizers"), ("FFC", "Fertilizers"),
    ("POL", "Oil/Gas"), ("PSO", "Oil/Gas"), ("MARI", "Oil/Gas"), ("PPL", "Oil/Gas"), ("OGDC", "Oil/Gas"),
    ("ALTN", "Power"), ("NCPL", "Power"), ("KOHE", "Power"), ("NPL", "Power"), ("HUBC", "Power"),
    ("ISL", "Steel"), ("AGHA", "Steel"), ("ASTL", "Steel"), ("ASL", "Steel"),
    ("CEPB", "Others"), ("INDU", "Others"), ("NML", "Others"), ("AGIL", "Others"), ("ORIX", "Others"), ("EPCL", "Others"),
    ("ABOT", "Pharma"), ("FEROZ", "Pharma"), ("GLAXO", "Pharma"), ("HINOON", "Pharma"), ("CPHL", "Pharma"), ("AGP", "Pharma"),
    ("SYS", "Tech"), ("TRG", "Tech"), ("AVN", "Tech"),
]
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


def sarmaaya_price(symbol: str):
    """Sarmaaya server-renders 'SYMBOL - Rs 348.72 | ...' in the page title."""
    try:
        html = http_get(SARMAAYA_URL.format(symbol=symbol)).decode("utf-8", "ignore")
        m = re.search(r"Rs\.?\s*([\d,]+\.?\d*)", html)
        return float(m.group(1).replace(",", "")) if m else None
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", help="Reference month YYYY-MM (default: current)")
    args = ap.parse_args()

    today = dt.date.today()
    if args.month:
        year, month = map(int, args.month.split("-"))
    else:
        year, month = today.year, today.month

    results, warnings = [], []
    for sym, sector in TICKERS:
        try:
            series = parse_eod(http_get(PSX_EOD_URL.format(symbol=sym)))
            if not series:
                raise ValueError("no EOD data returned")
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

        results.append({
            "ticker": sym,
            "sector": sector,
            "month_start_date": d1.isoformat() if d1 else None,
            "month_start_close": px1,
            "latest_date": last_date.isoformat(),
            "latest_close": last_px,
            "sarmaaya_check": check,
            "flagged": flagged,
        })

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
