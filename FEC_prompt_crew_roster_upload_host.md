# Call-sheet crew/talent roster fix — needs a bundle rebuild

## What's new

Steven hit this live on MAIN (SONI 005 test, right after the TX merge
went live): the call sheet's crew/talent roster never got extracted at
all. The run log shows why — `/extract-tx-crew-roster` ran for
125409ms (the backend processes every call-sheet page sequentially,
one Claude call at a time, so a multi-day call sheet routinely takes
well over 100s) and then got a 524 from Cloudflare. Every "Call Sheet"
tag and match — Crew Payroll, Crew IC, and the call-sheet Talent
feature — was silently empty for that run as a result.

## What I already built (source only, pushed to `main` as `7913d48`)

`source/Production Binder Wizard.dc.html`, `extractTxCrewRoster()` —
one change: the request now goes through `this.UPLOAD_BACKEND_URL`
instead of `this.BACKEND_URL`, the unproxied upload subdomain that
bypasses Cloudflare's gateway timeout entirely — the same fix already
used for `extractFringe()`, `extract-prodco-subvendors`, and
`extract-crew-freelance`. Client-side `REQ_TIMEOUT_MS` is 360s, plenty
of headroom once Cloudflare's own ~100s limit is out of the way.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` and send it back.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

No new dc.html edits needed — the fix is already written and pushed.
This is purely: rebuild the compiled artifact from current source and
confirm the two markers.
