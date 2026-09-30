# Post-TX-merge full rebuild — needs a bundle rebuild

## What's new

MAIN just received a large merge from DEV (commit `4473d91`): the
entire Texas workbook feature set built and tested over many rounds —
Cast & Crew List (aggregated Crew+Talent roster with call-sheet/DTR
matching), call-sheet Talent extraction and matching, the DTR
cross-tab orphan fix, AP extraction fixes, suffix/entity name
handling, several new payroll parsers, plus various GA/IL fixes and
this same-session's crew-freelance Cloudflare-bypass fix. This is a
much bigger rebuild than the last several small-fix rounds — the
source file itself changed by ~2,900 lines.

## What I already built (source only, pushed to `main` as `4473d91`)

Everything is already merged and pushed. There is nothing left to
write — this request is purely: rebuild the compiled frontend from the
current, merged source.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html`.
2. **Use MAIN's own current copies of the workbook templates and
   engine** — all four now live under `source/assets/workbooks/`
   (`georgia.xlsx`, `illinois-local.xlsx`, `illinois-oos.xlsx`,
   `texas.xlsx`) and `source/workbook-engine.js`. **`texas.xlsx` and
   `workbook-engine.js` changed significantly in this merge** — they
   now match DEV's newer versions (this fixes the earlier
   index013/index014 mismatch you flagged; there should be no
   discrepancy to reconcile this time since MAIN's own files are now
   the current ones).
3. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present
4. Also regenerate `backend/frontend/FRONTEND_SOURCE_HANDOFF.md` fresh
   for this round if that's part of your normal output.

## Not asking you to change anything else

No new dc.html edits needed — everything is already merged and
pushed. This is purely: rebuild the compiled artifact from current
source and confirm the two markers.
