"""
dividend_period.py — Classifies a dividend announcement by the fiscal
period it was actually declared for (per the PSX announcement PDF), not
by the date PSX published it.

A dividend counts toward the current year's tracker only if its period
end-date falls in the current calendar year — e.g. a "final dividend for
the year ended December 31, 2025" is excluded even if announced in
Feb/March of the following year; a "1st interim dividend for the quarter
ended March 31, 2026" is included regardless of announcement date.

Pipeline per announcement:
  1. find_pdf_url(ticker, date_iso) — match the announcement's date
     against the ticker's company page "Financial Results" / "Board
     Meetings" / "Others" tabs (server-rendered, most-recent-first — this
     is why the daily check only needs to look at recent items) and grab
     the linked PDF.
  2. download + extract text (PyPDF2). Some PSX filings are scanned
     images with no extractable text layer — those are flagged for
     manual review rather than guessed at.
  3. Regex-search the text for "<quarter|half year|year|nine months|
     twelve months> ended <date>" (case-insensitive, several wordings)
     and take the first match — PSX filings state the current period in
     the subject line before any comparative-period figures appear.
  4. Classify included/excluded by comparing the parsed period end-date's
     year against the current year.

Never OCR, never guess: if the PDF has no text layer, no matching
announcement is found, or no period phrase can be parsed, the result is
"needs_review" and the caller must not silently include or exclude it.
"""

import datetime as dt
import io
import re
import urllib.request

import PyPDF2

COMPANY_PAGE_URL = "https://dps.psx.com.pk/company/{symbol}"
HEADERS_BASE = {"User-Agent": "Mozilla/5.0 (dividend-period-bot)"}

TAB_RE = re.compile(
    r'data-name="(Financial Results|Board Meetings|Others)">\1</div>.*?'
    r'<div class="tabs__panel"[^>]*data-name="\1"[^>]*>(.*?)</table>',
    re.S,
)
ROW_RE = re.compile(
    r'<tr><td>([^<]+)</td><td>([^<]*)</td><td>.*?'
    r'href="(/download/(?:document|attachment)/[^"]+\.pdf)"',
    re.S,
)
ROW_DATE_FORMAT = "%b %d, %Y"

PERIOD_RE = re.compile(
    r'(quarter|half\s*year|year|twelve\s+months?|nine\s+months?)\s+ended\s+'
    r'([A-Za-z]+\.?\s+\d{1,2},?\s*\d{4})',
    re.I,
)
DATE_FORMATS = ["%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y"]


def _http_get(url: str, referer: str | None = None) -> bytes:
    headers = dict(HEADERS_BASE)
    if referer:
        headers["Referer"] = referer
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def _parse_row_date(date_str: str):
    try:
        return dt.datetime.strptime(date_str.strip(), ROW_DATE_FORMAT).date()
    except ValueError:
        return None


def _parse_period_date(date_str: str):
    cleaned = re.sub(r"\s+", " ", date_str.replace(",", ", ")).strip()
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    return None


def find_pdf_url(ticker: str, date_iso: str):
    """Search the company page's announcement tabs (in priority order) for
    a row on the same calendar date as the payout announcement."""
    target = dt.date.fromisoformat(date_iso)
    try:
        html = _http_get(COMPANY_PAGE_URL.format(symbol=ticker)).decode("utf-8", "ignore")
    except Exception:
        return None

    for tab_match in TAB_RE.finditer(html):
        table_html = tab_match.group(2)
        for row_match in ROW_RE.finditer(table_html):
            row_date = _parse_row_date(row_match.group(1))
            if row_date == target:
                return "https://dps.psx.com.pk" + row_match.group(3)
    return None


def download_pdf(url: str, referer: str):
    try:
        return _http_get(url, referer=referer)
    except Exception:
        return None


def extract_text(pdf_bytes: bytes) -> str:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception:
        return ""


def _period_label(period_type: str, end_date: dt.date, included: bool) -> str:
    period_type = re.sub(r"\s+", " ", period_type.lower())
    if "year" in period_type or "twelve" in period_type:
        label = f"FY{end_date.year} Final"
    elif "half" in period_type:
        label = f"H{'1' if end_date.month <= 6 else '2'} {end_date.year} Interim"
    elif "nine" in period_type:
        label = f"9M {end_date.year} Interim"
    else:
        quarter = {3: 1, 6: 2, 9: 3, 12: 4}.get(end_date.month)
        label = f"Q{quarter or '?'} {end_date.year} Interim"
    return label if included else f"{label} (excluded — not {dt.date.today().year})"


# PSX's own "Financial Results" column on the payouts page already states
# the period end-date in a structured "DD/MM/YYYY(CODE)" form (e.g.
# "31/03/2026(IIIQ)", "31/12/2025(YR)") — this is the fallback when the
# PDF can't be classified (mostly: it's a scanned image). CODE ~ YR/FYR =
# annual, HYR = half year, everything else = quarter (the quarter number
# is derived from the end-date's calendar month, same as the PDF path, not
# from the company's own internal fiscal-quarter numbering).
PAYOUT_PERIOD_RE = re.compile(r"(\d{2})/(\d{2})/(\d{4})\(([A-Za-z]+)\)")


def classify_from_payout_period(period_str: str, current_year: int | None = None):
    """Fallback classification using PSX's own payout-table period code.
    Returns a classify()-shaped dict, or None if period_str isn't in the
    expected 'DD/MM/YYYY(CODE)' form (e.g. '-')."""
    current_year = current_year or dt.date.today().year
    match = PAYOUT_PERIOD_RE.match((period_str or "").strip())
    if not match:
        return None

    day, month, year, code = match.groups()
    try:
        period_end_date = dt.date(int(year), int(month), int(day))
    except ValueError:
        return None

    code_upper = code.upper()
    if "HYR" in code_upper:
        period_type = "half year"
    elif "YR" in code_upper or "F" == code_upper:
        period_type = "year"
    else:
        period_type = "quarter"

    included = period_end_date.year == current_year
    return {
        "status": "included" if included else "excluded",
        "period_label": _period_label(period_type, period_end_date, included),
        "period_end_date": period_end_date.isoformat(),
        "pdf_url": None,
        "reason": None,
        "source": "payout_table",
    }


def _classify_from_pdf(ticker: str, date_iso: str, current_year: int) -> dict:
    """Returns {status, period_label, period_end_date, pdf_url, reason}.
    status is one of: included, excluded, needs_review."""
    company_url = COMPANY_PAGE_URL.format(symbol=ticker)

    pdf_url = find_pdf_url(ticker, date_iso)
    if not pdf_url:
        return {
            "status": "needs_review",
            "period_label": "NEEDS REVIEW: no matching announcement PDF found",
            "period_end_date": None,
            "pdf_url": None,
            "reason": "no announcement row matched this date in the recent announcements list",
        }

    pdf_bytes = download_pdf(pdf_url, referer=company_url)
    if not pdf_bytes:
        return {
            "status": "needs_review",
            "period_label": "NEEDS REVIEW: PDF download failed",
            "period_end_date": None,
            "pdf_url": pdf_url,
            "reason": "could not download the announcement PDF",
        }

    text = extract_text(pdf_bytes)
    if not text.strip():
        return {
            "status": "needs_review",
            "period_label": "NEEDS REVIEW: scanned PDF, no extractable text",
            "period_end_date": None,
            "pdf_url": pdf_url,
            "reason": "PDF has no text layer (scanned image) — not OCR'd, needs manual review",
        }

    match = PERIOD_RE.search(text)
    if not match:
        return {
            "status": "needs_review",
            "period_label": "NEEDS REVIEW: could not parse period from PDF text",
            "period_end_date": None,
            "pdf_url": pdf_url,
            "reason": "no '<period> ended <date>' phrase found in the extracted text",
        }

    period_end_date = _parse_period_date(match.group(2))
    if not period_end_date:
        return {
            "status": "needs_review",
            "period_label": f"NEEDS REVIEW: unparseable period date ({match.group(2)!r})",
            "period_end_date": None,
            "pdf_url": pdf_url,
            "reason": f"matched phrase but couldn't parse the date: {match.group(0)!r}",
        }

    included = period_end_date.year == current_year
    return {
        "status": "included" if included else "excluded",
        "period_label": _period_label(match.group(1), period_end_date, included),
        "period_end_date": period_end_date.isoformat(),
        "pdf_url": pdf_url,
        "reason": None,
    }


def classify(ticker: str, date_iso: str, payout_period: str | None = None,
             current_year: int | None = None) -> dict:
    """Classify a dividend by fiscal period. Tries the linked PSX
    announcement PDF first (most authoritative); if that can't be
    classified (typically a scanned PDF with no text layer), falls back to
    PSX's own payout-table period code so entries get populated rather than
    left needs_review. Only truly returns needs_review if both fail (e.g.
    the payout table itself has no period code, "-")."""
    current_year = current_year or dt.date.today().year
    result = _classify_from_pdf(ticker, date_iso, current_year)
    result.setdefault("source", "pdf")

    if result["status"] == "needs_review":
        fallback = classify_from_payout_period(payout_period, current_year)
        if fallback:
            fallback["reason"] = f"PDF unclassifiable ({result['reason']}); used PSX payout-table period code instead"
            return fallback

    return result
