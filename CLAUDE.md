# PSX Monthly Report

Workflow: `python fetch_prices.py` -> `python build_report.py` -> `python send_report.py`.

`send_report.py` sends via raw SMTP (Gmail, port 465), which does not work from
Claude Code's remote/sandboxed environments — outbound is HTTP(S)-proxy-only
there, and port 465 isn't tunneled. In that case (or whenever asked to "run
the report"/"send the report"), skip `send_report.py` and instead create a
Gmail draft via the Gmail MCP tool (`create_draft`) addressed to the
`RECIPIENTS` in `send_report.py`, with the generated
`PSX_Share_Prices_<YYYY-MM>.xlsx` base64-encoded as an attachment, subject
`Monthly PSX Share Price Report — <label>`, and a body summarizing any
warnings (fetch failures, >1% divergence flags) from `prices.json` so they're
easy to verify before sending.
