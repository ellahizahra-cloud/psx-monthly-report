# PSX Portfolio Automation

Two independent automations. They do **not** share the same ticker
universe: Automation 1 (monthly price report) covers all 61 tickers in
`tickers.py` (`TICKERS`); Automation 2 (dividend tracker) is scoped to
just the 37 tickers actually held (`tickers.py` `DIVIDEND_TICKERS`) — see
"Ticker universe" below for how each list was built. See
`.github/workflows/` for their schedules — GitHub Actions is the durable
cron: it runs even when nobody has a session open, and commits state back
to the repo so history survives between runs.

## Setup

1. Google Account -> Security -> 2-Step Verification -> App passwords ->
   generate one for "Mail".
2. In the repo's Settings -> Secrets and variables -> Actions, add:
   - `GMAIL_SENDER` — the sending Gmail address
   - `GMAIL_APP_PASSWORD` — the app password from step 1
3. Recipients are hardcoded in `mailer.py` (`RECIPIENTS`) — edit that list,
   nothing else needed.

Both workflows need `contents: write` (already set) so they can commit
updated state files back to the branch they run on.

## Automation 1 — Monthly Price Report

`.github/workflows/monthly-report.yml` runs `run_monthly_report.py` daily
at 08:00 UTC on days 2-5 of the month. Day 2 is the real scheduled run;
days 3-5 exist purely as a catch-up window — the script checks
`monthly_report_state.json` and no-ops if that month's report was already
sent, so a missed day-2 run gets picked up automatically.

The report always covers the month that just ended, not the one still in
progress — e.g. a run in early August reports on July's own Month-Start
Close (1st trading day of July) and Month-End Close (last trading day of
July), never August's first trading day. This means the reported month is
always fully closed by send time, so there's no "no trading data yet" gap
to fill in even if the 1st (or 2nd) of the new month falls on a weekend.
Month-End Close (the more recent of the two, typically only a few days
old by send time) is cross-checked against sarmaaya.pk (>1% divergence
flagged) — Month-Start is weeks old by then and isn't meaningful to
compare against a live quote. Wrapped in try/except — a failure sends an
error-notice email instead of a broken or partial file.

Manual run: `python fetch_prices.py && python build_report.py && python
send_report.py`, or `python run_monthly_report.py` for the full
catch-up-aware flow.

## Automation 2 — Event-Driven Dividend Tracker

`.github/workflows/dividend-tracker.yml` runs `run_dividend_tracker.py`
daily. It's a lightweight per-ticker check against PSX's own payout feed
(`POST https://dps.psx.com.pk/company/payouts`, the same call
dps.psx.com.pk/payouts makes) — never a full-history re-scrape.

- The **board-meeting/announcement date** PSX stamps on each payout row is
  used to detect *new* announcements (not book closure, ex-date, or credit
  date) — but which *fiscal year* a dividend counts toward is decided
  separately (see "Fiscal-period filtering" below), not by this date.
- Only actual announced cash dividends (`(D)` flag) are counted; bonus
  shares/right issues are parsed but excluded from the rupee totals and
  noted for audit. Never TTM/annualized/"last declared x4" figures.
- `tracking_log.json` is the source of truth for every announcement ever
  seen (by ticker); `PSX_Dividend_Tracker.xlsx` (single "Dividend Tracker"
  sheet, one row per ticker) is fully regenerated from it on every update,
  so "Dividend This Month" / "Dividend YTD" are always internally
  consistent (YTD naturally resets each January since it's computed from
  the current calendar year's log entries).
- No new announcements -> no email, no file change, nothing committed.
- New announcement(s) -> tracker updated, one email sent listing exactly
  which ticker(s) triggered it and what was announced, `tracking_log.json`
  updated so it isn't re-flagged next run.

### Stock splits

`stock_splits.json` (`{"TICKER": [{"date": "YYYY-MM-DD", "ratio": 2}]}`)
holds any subdivisions discovered for a ticker — PSX's payout feed doesn't
expose split ratios directly, so entries need to be added by hand when a
split is announced. Every dividend amount is divided by the cumulative
ratio of all splits *after* its announcement date, so historical figures
stay on a current-share-count basis. Defaults to `{}` (no adjustments).

### Fiscal-period filtering

A dividend counts toward the current year's tracker based on the **fiscal
period it was declared for** (e.g. "for the quarter ended March 31,
2026"), never the date PSX happened to announce it. A "final dividend for
the year ended December 31, 2025" is excluded even if announced in
Feb/March 2026 — it belongs to FY2025.

`dividend_period.py` determines this by: matching the announcement's date
against the ticker's PSX company page ("Financial Results" / "Board
Meetings" / "Others" announcement tabs) to find the linked PDF, downloading
it, extracting text with PyPDF2, and regex-searching for `<quarter|half
year|year|nine months|twelve months> ended <date>` (handles common wording
variants including ISO `YYYY-MM-DD` dates, case-insensitive). The parsed
period end-date's year is compared to the current year to classify the
entry as `included` or `excluded`.

Many PSX filings are scanned images with no text layer at all (confirmed
common for MCB, HBL, BAFL and others — typed letterheads that were
signed and scanned rather than filed as native PDFs). When that happens,
`dividend_period.py` falls back to OCR (Tesseract, via `pytesseract`) on
the page-1 scanned image PSX publishes alongside the same document —
these are typed letters, not handwriting, so OCR reads them reliably
(spot-checked against several real filings). OCR-sourced classifications
are tagged (`source: "pdf_ocr"`, and `"(OCR)"` appended to the period
label) so they stay distinguishable in the audit trail from clean
text-layer extraction. Only if OCR *also* yields nothing usable does
`dividend_period.classify()` fall back to PSX's own payout-table period
code (e.g. `31/03/2026(IIIQ)` — a structured date+code PSX already
publishes on the payouts page, no PDF needed) rather than leaving it
unclassified. Only if *that* is also missing/unparseable (e.g. `-`) does
the entry get marked `needs_review` — **never guessed**.
`included`/`excluded`/`needs_review` entries are all recorded in
`tracking_log.json` and shown in the workbook's Audit Log column (so
exclusions are auditable), but only `included` entries count toward
Dividend This Month / Dividend YTD / Dividend Announced (and therefore
Gross/Tax/Net). Entries from before this feature shipped (no
`period_classification` recorded) fall back to counting, so existing
totals aren't retroactively changed.

### Supplementary announcement-tab detection

PSX's Payouts table (the primary source above) has been observed to lag
a day or more behind the underlying board-meeting/Financial Results
announcement it's built from (confirmed live for FFC's Jul 29, 2026
interim dividend, which sat unindexed on the Payouts table for days while
already public on the company page). To catch this, every run also scans
each ticker's "Financial Results" / "Board Meetings" tabs directly
(`dividend_period.check_supplementary_announcements`), independent of the
Payouts table:

- Routine filings (AGM notices, transmissions, Shariah disclosures,
  corporate briefings, reschedule/postponement notices, and any title
  naming a date later than its own announcement — i.e. a notice about an
  *upcoming* meeting, not that meeting's outcome) are skipped by title
  before any PDF is downloaded, to keep this a lightweight daily check.
- For everything else, the linked PDF is downloaded and searched for a
  "CASH DIVIDEND" heading followed by a "Rs X per share" figure (PSX
  filings state the newly-declared amount first, before any "already
  paid" comparative figure). If the PDF is scanned, OCR of the page-1
  image is tried before giving up (same fallback as the primary
  classifier above). If found and the fiscal period parses, the entry is
  recorded exactly like a Payouts-sourced one.
- If neither the PDF text nor OCR yields a usable amount/period, the
  entry is recorded as `needs_review` with direct links to the PDF and
  the scanned page image, so it can be checked by eye — never guessed.

These entries are marked `"source": "financial_results_tab"` and
`"provisional": true` in `tracking_log.json`. If the Payouts table later
confirms the same date, the provisional entry is replaced by the
authoritative one (never double-counted). Once an item has been recorded
via either channel it won't be re-flagged on subsequent runs, so this
doesn't repeat in every day's email — only genuinely new items do.

### First run

`tracking_log.json` and `PSX_Dividend_Tracker.xlsx` are pre-seeded with a
real pull of each ticker's current PSX payout history (as of 2026-07-10)
so the first scheduled run doesn't blast out a giant "everything is new"
email — only genuinely new announcements after that date will trigger one.
If you ever want to re-baseline from scratch, empty `tracking_log.json` to
`{}` and delete `PSX_Dividend_Tracker.xlsx`; the next run will rebuild both
from whatever PSX currently returns and email that as the new baseline.

### Cash flow reconciliation

`PSX_Dividend_Tracker.xlsx` is a single sheet ("Dividend Tracker"), one
row per ticker: Ticker, Company, Sector, Quantity, Dividend This Month,
Dividend YTD, Dividend Announced (the per-share amount from whichever
announcement(s) triggered *this* update — not a calendar-month total),
Gross/Tax/Net Cash Dividend, Last Announcement Date, and an Audit Log
column (every `Date: Amount/share` on record for that ticker). A Total
row sums Quantity, Gross, Tax and Net across all tickers. Quantity is
sourced from `holdings.py`, which reads `Cash Dividend 1 (1).xlsx` (root
of the repo) — only its Quantity column; DPS Received/Receivable, Book
Closed Date, Credit Expected Date and X-Date are explicitly out of scope
and never read. Tickers with no announcement in this update show `-` (not
`0`) for the dividend-derived columns; Quantity always populates. The
withholding rate is `TAX_WITHHOLDING_RATE` in `dividend_workbook.py`
(currently 15%) — update that one constant if the rate or filer status
changes.

## Ticker universe

`tickers.py`'s `UNIVERSE`/`TICKERS` (61 tickers) is the price report's
full universe: the original 9-sector list plus 15 more added because
they're real holdings in `Cash Dividend 1 (1).xlsx` (AICL, PTL, BYCO,
CPPL, FDIBL, SRVI, PKGS, AKBL, MUREB, GGL, TPL, BFAGRO, ATLH, PIOC, SRR).
Symbols were verified live against `dps.psx.com.pk/company/<symbol>`; a
few source names needed resolving to their real PSX ticker (see the
module docstring for the mapping, including `ENGRO` -> `ENGROH` following
Engro's 2025 corporate restructuring, `PANTHER` -> `PTL`, `PACKAGE` ->
`PKGS`). Note `BYCO` and `FDIBL` have thin/stale PSX trading data (BYCO's
feed stops in Dec 2021) — expect the monthly report to flag them under
"fill manually" most months; that's PSX's data, not a bug.

`DIVIDEND_TICKERS` (37 tickers) is the **dividend tracker's own,
narrower** universe — confirmed by the user as exactly what's held by
Ellahi Capital/AATML. It intentionally excludes tickers that are in the
price-tracking universe but not held (e.g. KOHC, POL, ISL, ABOT) — the
dividend tracker must never report on those, per spec. `TPL` and `BFAGRO`
are included but expected to show blank/0 (TPL doesn't currently pay
dividends; BFAGRO hasn't paid any since its March 2025 listing).
