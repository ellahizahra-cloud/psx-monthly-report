# PSX Portfolio Automation

Two independent automations over the same shared PSX ticker universe
(`tickers.py`): the original 46-ticker/9-sector universe plus 15 more
tickers pulled in from actual holdings (see "Ticker universe" below), 61
in total. See `.github/workflows/` for their schedules — GitHub Actions is
the durable cron: it runs even when nobody has a session open, and commits
state back to the repo so history survives between runs.

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
`monthly_report_state.json` and no-ops if this month's report was already
sent, so a missed day-2 run gets picked up automatically.

For each ticker: Month-Start Close (1st trading day of the current month)
and Month-End Close (last trading day of the month that just ended), cross
checked against sarmaaya.pk (>1% divergence flagged). Wrapped in
try/except — a failure sends an error-notice email instead of a broken or
partial file.

Manual run: `python fetch_prices.py && python build_report.py && python
send_report.py`, or `python run_monthly_report.py` for the full
catch-up-aware flow.

## Automation 2 — Event-Driven Dividend Tracker

`.github/workflows/dividend-tracker.yml` runs `run_dividend_tracker.py`
daily. It's a lightweight per-ticker check against PSX's own payout feed
(`POST https://dps.psx.com.pk/company/payouts`, the same call
dps.psx.com.pk/payouts makes) — never a full-history re-scrape.

- The **board-meeting/announcement date** PSX stamps on each payout row is
  what counts as the dividend date (not book closure, ex-date, or credit
  date).
- Only actual announced cash dividends (`(D)` flag) are counted; bonus
  shares/right issues are parsed but excluded from the rupee totals and
  noted for audit. Never TTM/annualized/"last declared x4" figures.
- `tracking_log.json` is the source of truth for every announcement ever
  seen (by ticker); `PSX_Dividend_Tracker.xlsx` (Summary + Audit Log
  sheets) is fully regenerated from it on every update, so "Dividend This
  Month" / "Dividend YTD" are always internally consistent (YTD naturally
  resets each January since it's computed from the current calendar
  year's log entries).
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

### First run

`tracking_log.json` and `PSX_Dividend_Tracker.xlsx` are pre-seeded with a
real pull of each ticker's current PSX payout history (as of 2026-07-10)
so the first scheduled run doesn't blast out a giant "everything is new"
email — only genuinely new announcements after that date will trigger one.
If you ever want to re-baseline from scratch, empty `tracking_log.json` to
`{}` and delete `PSX_Dividend_Tracker.xlsx`; the next run will rebuild both
from whatever PSX currently returns and email that as the new baseline.

### Cash flow reconciliation

The `PSX_Dividend_Tracker.xlsx` "Cash Flow Reconciliation" sheet adds, per
ticker: Quantity (shares held), Dividend Announced (the per-share amount
from whichever announcement(s) triggered *this* update — not a calendar-
month total), Gross Cash Dividend, Tax Amount, and Net Cash Dividend, plus
a Total row. Quantity is sourced from `holdings.py`, which reads
`Cash Dividend 1 (1).xlsx` (root of the repo) — only its Quantity column;
DPS Received/Receivable, Book Closed Date, Credit Expected Date and X-Date
are explicitly out of scope and never read. Tickers with no announcement
in this update show `-` (not `0`) for the dividend-derived columns;
Quantity always populates. The withholding rate is `TAX_WITHHOLDING_RATE`
in `dividend_workbook.py` (currently 15%) — update that one constant if
the rate or filer status changes.

## Ticker universe

`tickers.py` is the single shared list both automations import from.
Symbols were verified live against `dps.psx.com.pk/company/<symbol>`; a
few source names needed resolving to their real PSX ticker (see the
module docstring for the mapping, including `ENGRO` -> `ENGROH` following
Engro's 2025 corporate restructuring).

15 tickers (AICL, PTL, BYCO, CPPL, FDIBL, SRVI, PKGS, AKBL, MUREB, GGL,
TPL, BFAGRO, ATLH, PIOC, SRR) were added beyond the original 9 sectors
because they're real holdings in `Cash Dividend 1 (1).xlsx` — without them
the dividend tracker would silently never catch dividends on those
positions. Two of that file's labels needed resolving to their real
ticker (`PANTHER` -> `PTL`, `PACKAGE` -> `PKGS`); see `holdings.py`'s
`ALIASES` for the full crosswalk. Note `BYCO` and `FDIBL` have thin/stale
PSX trading data (BYCO's feed stops in Dec 2021) — expect the monthly
report to flag them under "fill manually" most months; that's PSX's data,
not a bug.
