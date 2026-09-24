# Call-sheet crew cache leaked across separate runs — needs a bundle rebuild

## What was wrong

Second real-data test: ran SONIC 005, then ABBVIE 022 in the same browser
session without reloading. ABBVIE's workbook came back with SONIC's crew
roster written into it as "on call sheet" — e.g. "Jody Hill / Director" and
"Tova Dann / Executive Producer" (SONIC's own producers) appeared as new
rows in ABBVIE's Crew Payroll tab. Confirmed by diffing the two output
workbooks directly: a 100% overlap between the two runs' "on call sheet"
name lists. ABBVIE's own run log never even shows a
`/extract-tx-crew-roster` request — the cached result from the SONIC run
was silently reused instead of re-extracting.

Root cause: `getCallSheetCrew()` caches its result on
`this._callSheetCrew` so both Crew Payroll and Crew - Indepen. Contractors
share one extraction within a single run (intentional — avoids extracting
twice). Nothing ever cleared that cache between separate runs, though.

## What I already fixed (source only, pushed to `dev` as `c87628d`)

One line, in the existing TX per-run reset block (`source/Production
Binder Wizard.dc.html`, right where every other per-run counter/array
already gets reset with the comment "Per-run counters live on the
instance and survive 'Start a new workbook'"): `this._callSheetCrew =
null;`.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` (has this fix) and send it
   back, same pattern as prior rebuild handoffs.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

No new dc.html edits needed — the fix is already written and pushed. This
is purely: rebuild the compiled artifact from current source and confirm
the two markers.
