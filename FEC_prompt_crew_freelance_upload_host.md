# Crew Freelance upload fix — needs a bundle rebuild

## What's new

A coworker hit a real production failure on the live site (Wrigley
Respawn x Rivals, debug logs in WLM 017): `extractCrewFreelance()`
batched 10 invoice PDFs per POST and sent them through the
Cloudflare-proxied backend host. That batch tripped Cloudflare's
payload-size limit — a clean 413 on the first attempt, then "Failed to
fetch" on a retry with an even larger/duplicated batch.

## What I already built (source only, pushed to `main` as `f293707`)

`source/Production Binder Wizard.dc.html`, `extractCrewFreelance()` —
two changes, matching the pattern already used elsewhere in this same
function family (`extractFringe()`, `extract-prodco-subvendors`):

1. `CHUNK` reduced from 10 to 5, matching this endpoint's own
   5-concurrent backend `Semaphore` (a bigger chunk means multiple
   sequential AI-extraction rounds piling up inside one held-open
   request).
2. The request now goes through `this.UPLOAD_BACKEND_URL` instead of
   `this.BACKEND_URL` — the unproxied upload subdomain that bypasses
   Cloudflare's payload-size and gateway-timeout limits entirely. This
   is the change that actually fixes the reported 413; `CHUNK` alone
   would not have.

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
