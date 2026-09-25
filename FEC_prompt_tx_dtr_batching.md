# TX DTR: batching fix for the request-timeout risk you flagged — needs a rebuild

## Thanks for flagging the backend-limit risk

You were right to flag it, and it was real: `extractTxDtr()` sent every
uploaded file in one POST, copied from `extractTxCrewRoster`'s design
(fine there — a call sheet upload is a handful of files). At up to 1,000
DTR files that's one request awaiting hundreds of sequential Claude calls
behind the backend's own `Semaphore(5)` — easily 15+ minutes, nowhere
close to fitting behind Cloudflare's ~100s proxy timeout. Exactly the
524-timeout failure mode `extractFringe` already hit and solved once
before in this codebase (its own `CHUNK = 5` comment: "matching the
backend's own 5-concurrent cap so one request is a single concurrent
round").

## What I already fixed (source only, pushed to `dev` as `570e5c6`)

`source/Production Binder Wizard.dc.html`, `extractTxDtr()`: now batches at
5 files per POST (same `CHUNK = 5` pattern as `extractFringe`), looping
sequentially and merging results. Also re-dedupes names client-side across
the merged batches — the backend's own per-batch dedupe (for a "with
ID"/"no ID" duplicate pair sharing one filename) only sees one batch at a
time, so a duplicate pair split across a batch boundary would otherwise
survive as two entries.

Also synced your `max: 200 → 1000` change (and the hint text) on the
`txDtrForms` zone into this same file, so it carries forward from your
index009 rebuild rather than reverting on the next sync.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` (has both your 1,000-file cap
   and the batching fix) and send it back.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

No new dc.html edits needed — the fix is already written and pushed. This
is purely: rebuild the compiled artifact from current source and confirm
the two markers.
