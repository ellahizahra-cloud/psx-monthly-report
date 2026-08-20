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
  2. download + extract text (PyPDF2). Many PSX filings are scanned
     images with no extractable text layer — for those, render the PDF's
     own pages ourselves (PyMuPDF, 300 DPI, up to OCR_MAX_PAGES pages)
     and OCR them (pytesseract) as a fallback rather than giving up
     immediately. Rendering our own pages — rather than fetching PSX's
     lower-resolution page-1 thumbnail — works regardless of the PDF's
     URL scheme and gives OCR more resolution to work with. Confirmed to
     work well on real filings: these are typed letterheads, not
     handwriting, and Tesseract reads them cleanly (spot-checked against
     HBL's Aug 4, 2026 announcement).
  3. Regex-search the text for "<quarter|half year|year|nine months|
     twelve months> ended <date>" (case-insensitive, several wordings,
     including ISO YYYY-MM-DD dates) and take the first match — PSX
     filings state the current period in the subject line before any
     comparative-period figures appear.
  4. Classify included/excluded by comparing the parsed period end-date's
     year against the current year.

Still never guess: if neither the PDF's text layer nor OCR of our own
rendered pages yields a matching period phrase, no matching announcement
is found, or the payout-table fallback also comes up empty, the result is
"needs_review" and the caller must not silently include or exclude it.
OCR-sourced classifications are tagged (source="pdf_ocr") so they stay
distinguishable in the audit trail from clean text-layer extraction.

Supplementary detection channel (check_supplementary_announcements):
PSX's Payouts table (dividend_fetch.py, the primary source) sometimes
lags a day or more behind the underlying Financial Results/Board Meetings
announcement it's built from (seen live for FFC's Jul 29, 2026 interim
dividend). check_supplementary_announcements() reads the same two
announcement tabs directly, independent of the Payouts table, and for any
row not already known: tries to extract a newly-declared cash-dividend
amount from the linked PDF's text (falling back to OCR of our own
rendered pages, same as the primary pipeline) — a "CASH DIVIDEND" heading
followed by the first "Rs X per share" figure, since PSX filings state
the new declaration before any "already paid" comparative figure. If both
the PDF text and OCR fail to produce a usable amount/period, it's
returned as needs_review with direct links to the PDF and the scanned
page image so a human can check it — never guessed. Routine filings (AGM
notices, transmissions, corporate briefings, Shariah disclosures) are
skipped by title before ever downloading a PDF, to keep this a
lightweight daily check rather than a full-text scan of every filing.
"""

import datetime as dt
import io
import re
import urllib.request

import PyPDF2
import pytesseract
from PIL import Image

try:
    import pymupdf as fitz
except ImportError:
    import fitz

OCR_DPI = 300
OCR_MAX_PAGES = 3

COMPANY_PAGE_URL = "https://dps.psx.com.pk/company/{symbol}"
IMAGE_URL = "https://dps.psx.com.pk/download/image/{image_id}"
HEADERS_BASE = {"User-Agent": "Mozilla/5.0 (dividend-period-bot)"}

# Matches each announcement tab's panel independently (Financial Results,
# Board Meetings, Others). Deliberately does NOT anchor off the tab-list
# header the way an earlier version did — data-name="X">X</div> for every
# tab sits together at the top of the page, ahead of all three panels, so
# a single non-overlapping re.finditer() pass could only ever consume the
# first tab's header+panel span and would never reach the next tab's
# header (already inside the first match's span). Matching each
# <div class="tabs__panel" ... data-name="X"> block directly sidesteps
# that: every panel carries its own data-name attribute.
PANEL_RE = re.compile(
    r'<div class="tabs__panel"[^>]*data-name="(Financial Results|Board Meetings|Others)"[^>]*>'
    r'.*?<table[^>]*>(.*?)</table>',
    re.S,
)
ROW_RE = re.compile(
    r'<tr><td>([^<]+)</td><td>([^<]*)</td><td>.*?'
    r'data-images="([^"]*)".*?'
    r'href="(/download/(?:document|attachment)/[^"]+\.pdf)"',
    re.S,
)
ROW_DATE_FORMAT = "%b %d, %Y"

PERIOD_RE = re.compile(
    r'(quarter|half\s*year|year|twelve\s+months?|nine\s+months?)\s+ended\s+'
    r'([A-Za-z]+\.?\s+\d{1,2},?\s*\d{4}|\d{1,2}\s+[A-Za-z]+\.?,?\s*\d{4}|\d{4}-\d{2}-\d{2})',
    re.I,
)
# Both "Month Day, Year" (common in PDFs) and "Day Month Year" (common in
# bank filings, e.g. UBL's "30 June 2026") show up in the wild.
DATE_FORMATS = [
    "%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y",
    "%d %B, %Y", "%d %B %Y", "%d %b, %Y", "%d %b %Y",
    "%Y-%m-%d",
]

# Only these announcement-tab rows are worth downloading a PDF for; every
# other title (AGM notices, transmissions, Shariah disclosures, corporate
# briefings, etc.) is routine and never carries a fresh dividend
# declaration, so it's skipped before any network call.
RELEVANT_TITLE_RE = re.compile(r'financial results|board (of directors )?meeting', re.I)
IRRELEVANT_TITLE_RE = re.compile(
    r'notice of|transmission of|shariah|annual general meeting|corporate briefing|'
    r'other than financial results|reschedul|postpone',
    re.I,
)
# A title that names a full date *later* than the row's own announcement
# date is a notice about an upcoming meeting, not that meeting's outcome
# (e.g. "172 Board Meeting of FCCL - 6 August 2026" announced Jul 17,
# 2026) — outcome titles only ever reference past period-end dates.
TITLE_DATE_RE = re.compile(r'(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})')


def _title_references_future_date(title: str, row_date: dt.date) -> bool:
    for day, month, year in TITLE_DATE_RE.findall(title):
        for fmt in ("%d %B %Y", "%d %b %Y"):
            try:
                title_date = dt.datetime.strptime(f"{day} {month} {year}", fmt).date()
            except ValueError:
                continue
            if title_date > row_date:
                return True
    return False

CASH_DIVIDEND_HEADING_RE = re.compile(r'CASH\s+DIVIDEND\b', re.I)
AMOUNT_RE = re.compile(r'Rs\.?\s*([\d,]+\.?\d*)\s*/?-?\s*per\s*share', re.I)
NIL_RE = re.compile(r'\bNIL\b', re.I)
SUPPLEMENTARY_LOOKBACK_DAYS = 45

_page_cache: dict[str, str] = {}


def _http_get(url: str, referer: str | None = None) -> bytes:
    headers = dict(HEADERS_BASE)
    if referer:
        headers["Referer"] = referer
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def _get_company_html(ticker: str) -> str:
    """Company page HTML, cached per ticker for the life of the process —
    both find_pdf_url() and fetch_announcement_rows() need it, and a run
    only ever looks at one ticker's page once."""
    if ticker not in _page_cache:
        try:
            _page_cache[ticker] = _http_get(COMPANY_PAGE_URL.format(symbol=ticker)).decode(
                "utf-8", "ignore"
            )
        except Exception:
            _page_cache[ticker] = ""
    return _page_cache[ticker]


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
    """Search the company page's announcement tabs (in document order —
    Financial Results, Board Meetings, Others) for a row on the same
    calendar date as the payout announcement."""
    target = dt.date.fromisoformat(date_iso)
    html = _get_company_html(ticker)
    if not html:
        return None

    for tab_match in PANEL_RE.finditer(html):
        table_html = tab_match.group(2)
        for row_match in ROW_RE.finditer(table_html):
            row_date = _parse_row_date(row_match.group(1))
            if row_date == target:
                return "https://dps.psx.com.pk" + row_match.group(4)
    return None


def fetch_announcement_rows(ticker: str):
    """All rows from the Financial Results + Board Meetings tabs (Others
    excluded — it's briefings/press items, never a dividend outcome),
    most-recent-first as PSX renders them."""
    html = _get_company_html(ticker)
    if not html:
        return []

    rows = []
    for tab_match in PANEL_RE.finditer(html):
        tab_name = tab_match.group(1)
        if tab_name == "Others":
            continue
        for row_match in ROW_RE.finditer(tab_match.group(2)):
            row_date = _parse_row_date(row_match.group(1))
            if not row_date:
                continue
            image_id = row_match.group(3).strip()
            rows.append({
                "tab": tab_name,
                "date": row_match.group(1).strip(),
                "date_iso": row_date.isoformat(),
                "title": row_match.group(2).strip(),
                "pdf_url": "https://dps.psx.com.pk" + row_match.group(4),
                "image_url": IMAGE_URL.format(image_id=image_id) if image_id else None,
            })
    return rows


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


def ocr_pdf_bytes(pdf_bytes: bytes, max_pages: int = OCR_MAX_PAGES, dpi: int = OCR_DPI) -> str:
    """OCR fallback when a PDF has no text layer: render the PDF's own
    pages ourselves (PyMuPDF, at `dpi`) and run each through Tesseract,
    rather than depending on PSX exposing a separate pre-rendered page
    image for the same document. Works on any downloaded PDF regardless
    of its URL scheme, and gives OCR more resolution to work with than
    PSX's own thumbnail. These are typed letterheads, not handwriting, so
    OCR reads them cleanly (spot-checked against real filings). Returns
    "" on any failure — PDF can't be opened, rendering fails, OCR fails —
    so callers can treat this exactly like empty PDF text."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return ""

    try:
        matrix = fitz.Matrix(dpi / 72, dpi / 72)
        parts = []
        for page in doc[:max_pages]:
            try:
                pix = page.get_pixmap(matrix=matrix)
                image = Image.open(io.BytesIO(pix.tobytes("png")))
                parts.append(pytesseract.image_to_string(image))
            except Exception:
                continue
        return "\n".join(parts).strip()
    finally:
        doc.close()


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
    source = "pdf"
    if not text.strip():
        text = ocr_pdf_bytes(pdf_bytes)
        source = "pdf_ocr"
        if not text.strip():
            return {
                "status": "needs_review",
                "period_label": "NEEDS REVIEW: scanned PDF, OCR found no text",
                "period_end_date": None,
                "pdf_url": pdf_url,
                "reason": "PDF has no text layer (scanned image) and OCR of the rendered "
                          "pages also came up empty — needs manual review",
            }

    match = PERIOD_RE.search(text)
    if not match:
        return {
            "status": "needs_review",
            "period_label": f"NEEDS REVIEW: could not parse period from {'OCR' if source == 'pdf_ocr' else 'PDF'} text",
            "period_end_date": None,
            "pdf_url": pdf_url,
            "reason": f"no '<period> ended <date>' phrase found in the {'OCR-extracted' if source == 'pdf_ocr' else 'extracted'} text",
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
    label = _period_label(match.group(1), period_end_date, included)
    if source == "pdf_ocr":
        label += " (OCR)"
    return {
        "status": "included" if included else "excluded",
        "period_label": label,
        "period_end_date": period_end_date.isoformat(),
        "pdf_url": pdf_url,
        "reason": None,
        "source": source,
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


def _extract_cash_dividend_amount(text: str):
    """Looks for a 'CASH DIVIDEND' heading and returns the first 'Rs X per
    share' figure after it — PSX board-meeting-outcome letters state the
    newly-declared amount first, before any 'in addition to interim
    dividend already paid @ Rs Y per share' comparative figure (confirmed
    against real filings; the comparative figure is typically stated
    without the leading 'Rs', which is why taking the first match works).
    Returns: a float amount, 0.0 for an explicit NIL, or None if a CASH
    DIVIDEND heading exists but no amount could be parsed (needs_review) —
    distinct from no heading found at all, which means this filing simply
    doesn't declare a dividend and should be ignored."""
    heading = CASH_DIVIDEND_HEADING_RE.search(text)
    if not heading:
        return None
    section = text[heading.end():heading.end() + 500]
    if NIL_RE.search(section[:120]):
        return 0.0
    amount_match = AMOUNT_RE.search(section)
    if not amount_match:
        return None
    return float(amount_match.group(1).replace(",", ""))


def check_supplementary_announcements(ticker: str, known_isos, current_year: int | None = None):
    """Independent second detection channel: scans the Financial
    Results/Board Meetings tabs directly rather than relying on PSX's
    Payouts table, to catch cases where the Payouts table lags behind the
    underlying announcement. Only looks at rows newer than
    SUPPLEMENTARY_LOOKBACK_DAYS and not already in known_isos (a set of
    date_iso strings already recorded for this ticker, from either
    channel) — this is meant to run every day, so it must stay cheap and
    must not re-flag the same still-unresolved item forever.

    Returns a list of dicts: {date, date_iso, title, status, period_label,
    period_end_date, reason, pdf_url, image_url, amount_per_share}.
    status is 'included' / 'excluded' / 'needs_review', same vocabulary as
    classify(). amount_per_share is None only when status is
    needs_review — never guessed."""
    current_year = current_year or dt.date.today().year
    cutoff = dt.date.today() - dt.timedelta(days=SUPPLEMENTARY_LOOKBACK_DAYS)
    company_url = COMPANY_PAGE_URL.format(symbol=ticker)
    results = []

    for row in fetch_announcement_rows(ticker):
        if row["date_iso"] in known_isos:
            continue
        row_date = dt.date.fromisoformat(row["date_iso"])
        if row_date < cutoff:
            continue
        if IRRELEVANT_TITLE_RE.search(row["title"]) or not RELEVANT_TITLE_RE.search(row["title"]):
            continue
        if _title_references_future_date(row["title"], row_date):
            continue

        base = {
            "date": row["date"],
            "date_iso": row["date_iso"],
            "title": row["title"],
            "pdf_url": row["pdf_url"],
            "image_url": row["image_url"],
        }

        pdf_bytes = download_pdf(row["pdf_url"], referer=company_url)
        text = extract_text(pdf_bytes) if pdf_bytes else ""
        used_ocr = False
        if not text.strip() and pdf_bytes:
            text = ocr_pdf_bytes(pdf_bytes)
            used_ocr = bool(text.strip())
        if not text.strip():
            results.append({
                **base,
                "status": "needs_review",
                "period_label": None,
                "period_end_date": None,
                "amount_per_share": None,
                "reason": "scanned PDF (no extractable text) found via announcement tabs, "
                          "and OCR of the rendered pages also came up empty — "
                          "could not confirm a dividend amount automatically",
            })
            continue

        amount = _extract_cash_dividend_amount(text)
        if amount is None and not CASH_DIVIDEND_HEADING_RE.search(text):
            continue  # not a dividend-declaring filing at all, ignore silently

        if amount is None:
            results.append({
                **base,
                "status": "needs_review",
                "period_label": None,
                "period_end_date": None,
                "amount_per_share": None,
                "reason": "'CASH DIVIDEND' section found but the per-share amount "
                          "could not be parsed from the text",
            })
            continue

        if amount == 0.0:
            continue  # explicit NIL — nothing declared, nothing to report

        period_match = PERIOD_RE.search(text)
        period_end_date = _parse_period_date(period_match.group(2)) if period_match else None
        if period_end_date:
            included = period_end_date.year == current_year
            results.append({
                **base,
                "status": "included" if included else "excluded",
                "period_label": _period_label(period_match.group(1), period_end_date, included),
                "period_end_date": period_end_date.isoformat(),
                "amount_per_share": amount,
                "reason": None,
                "ocr": used_ocr,
            })
        else:
            results.append({
                **base,
                "status": "needs_review",
                "period_label": None,
                "period_end_date": None,
                "amount_per_share": amount,
                "reason": "cash dividend amount found but the fiscal period "
                          "('<period> ended <date>') could not be parsed",
            })

    return results
