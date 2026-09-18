# Wizard023's compiled bundle wasn't rebuilt — needs a fresh rebuild before this ships

## What's wrong

I merged `Production Binder Wizard.dc.html` from the Wizard023 export into the repo — that part is good, clean additive diff (retry logic, `txFallbackRow`/`txFallbackReason`, the TX Talent Payroll tab, the `on:[]` selector list, the `downloadIssuesTxt` object-unwrap fix). No content lost, nothing to flag there.

But `backend/frontend/index.html` in that same Wizard023 package **was not rebuilt from that dc.html**. I checked directly — it has neither the retry logic nor any of the Talent Payroll code baked into the compiled bundle, at either the package root or the `handoff/` copy. It's the same compiled artifact as the Sep 1 Wizard01 handoff (`FRONTEND_SOURCE_HANDOFF.md`), just carried forward unrebuilt.

That doc is explicit: `backend/frontend/index.html` is **"Compiled output — never hand-edit,"** and every prior handoff that actually shipped a working deploy included a matching rebuilt bundle as its own deliverable line item. This one didn't. Right now our repo's dc.html *source* is current, but nothing in this export actually deploys any of it — not the retry/fallback reliability work, not TX Talent Payroll.

## What I need from you

1. **Rebuild `backend/frontend/index.html` from the current `source/Production Binder Wizard.dc.html`** (the one that already has retry logic + all five TX AP-family tabs + Talent Payroll wiring) and send it back as its own deliverable, same as the Sep 1 package did.
2. **Re-verify the two post-rebuild injection markers**, per your own doc's checklist:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present
3. **Use the corrected `texas.xlsx`, not your sandbox's copy.** The one in your sandbox still has the old Talent Payroll layout (no On PTIP / On PDF columns) — I already fixed this on our end (typo fix + column insert, re-synced from the master template) but since you can't `git pull`, your next rebuild will silently re-embed the stale version unless I hand you the corrected file directly. I'll attach/send it separately — swap it into `assets/workbooks/texas.xlsx` in your sandbox before rebuilding, or the Talent Payroll tab will keep soft-skipping with the "template layout mismatch" message even though the fix already landed.

## Not asking you to change anything else

No new dc.html edits needed for this — the source is already correct. This is purely: rebuild the compiled artifact from what's already there, with the corrected `texas.xlsx` swapped in first, and confirm the two markers.
