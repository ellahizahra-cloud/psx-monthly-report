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
requests, and adapt PSX_EOD_URL / parse_eod() and PSX_PAYOUTS_URL /
parse_dividends() accordingly.
"""

import argparse
import datetime as dt
import json
import re
import sys
import urllib.parse
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
# Same endpoint the "Payouts" tab's AJAX call on a company page uses.
PSX_PAYOUTS_URL = "https://dps.psx.com.pk/payouts"
SARMAAYA_URL = "https://sarmaaya.pk/stocks/{symbol}"
DIVERGENCE_THRESHOLD = 0.01  # 1%
# PSX ordinary equity face value; dividend %s are of this. Verify manually
# for any ticker whose face value isn't the PKR 10 standard.
DIVIDEND_FACE_VALUE = 10.0
HEADERS = {"User-Agent": "Mozilla/5.0 (monthly-report-bot)"}


def http_get(url: str) -> bytes:
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def http_post(url: str, data: dict) -> str:
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=body, headers=HEADERS, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8", "ignore")


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


def _book_closure_start(text: str):
    """First date in a 'DD/MM/YYYY - DD/MM/YYYY' (or single-date) cell, else None."""
    first = text.strip().split("-")[0].strip()
    try:
        return dt.datetime.strptime(first, "%d/%m/%Y").date()
    except ValueError:
        return None


def _payout_rows(symbol: str):
    """Yield (dividend_announcement_text, book_closure_text) for every payout
    row PSX has for `symbol`, paging through the /payouts endpoint."""
    offset, count, total = 0, 50, None
    while total is None or offset < total:
        html = http_post(PSX_PAYOUTS_URL, {"symbol": symbol, "count": count, "offset": offset})
        m = re.search(r"of\s+(\d+)\s+entries", html)
        total = int(m.group(1)) if m else 0
        for row in re.findall(r"<tr>(.*?)</tr>", html, re.DOTALL):
            cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
            if len(cells) < 6:
                continue  # header row (<th>) or malformed
            div_text = re.sub("<[^>]+>", "", cells[3]).strip()
            bc_text = re.sub("<[^>]+>", "", cells[5]).strip()
            yield div_text, bc_text
        offset += count


def parse_dividends(symbol: str, start: dt.date, end: dt.date):
    """
    Sum cash dividend payouts (PKR/share) for `symbol` whose book-closure
    start date falls in [start, end], via PSX's /payouts endpoint (the same
    one the "Payouts" tab on a company page calls — POST symbol/count/offset,
    returns an HTML table fragment).

    Dividends are announced as a percentage of face value, e.g. "32.50%(iii)
    (D)" = 3rd interim cash dividend of 32.5% of face value. Bonus ("(B)")
    and right ("(R)") issue rows are not cash payouts and are skipped.

    Returns 0.0 if the symbol has no matching cash dividend in the window
    (a real answer), or None if the fetch itself failed (network/parse
    error), so the pipeline can tell "no dividend" from "couldn't check".
    """
    total_payout = 0.0
    try:
        for div_text, bc_text in _payout_rows(symbol):
            if "(D)" not in div_text:
                continue
            m = re.search(r"([\d.]+)\s*%", div_text)
            if not m:
                continue
            bc_date = _book_closure_start(bc_text)
            if bc_date is None or not (start <= bc_date <= end):
                continue
            total_payout += float(m.group(1)) / 100 * DIVIDEND_FACE_VALUE
    except Exception:
        return None
    return round(total_payout, 2)


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
