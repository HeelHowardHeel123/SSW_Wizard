# TX Crew Payroll: call-sheet crew matching — needs a bundle rebuild

## What's new

Backend and source are done and pushed to `dev` (`eb55428`). This closes a
gap the TX template already had slots for: `Crew Payroll` tab's `AQ8 = "On
Call Sheet"` / `AR8 = "Name on Call sheet"`, and `Crew - Indepen.
Contractors` tab's `AD8 = "On Call Sheet"` / `AE8 = "Name on Call sheet"` —
both previously always blank.

Backend: new `POST /extract-tx-crew-roster` reads the same call sheet
PDF(s) already uploaded to the `txCallSheet` zone (the one you moved into
the Crew section last rebuild) and returns a deduped crew roster
`{"crew": [{"name", "title"}], "issues": [...], "files": [...]}`.

Frontend (`source/Production Binder Wizard.dc.html`): three new methods —
`extractTxCrewRoster()`, `getCallSheetCrew()` (cached per build, shared by
both tabs), `matchCallSheetToRoster()` (calls the existing `/match-names`
endpoint to fuzzy-match call-sheet names against each tab's own roster).
Wired into `parseTxPayrollRows()` and the Crew IC loop — no new upload
zone, no new UI needed, this is pure backend logic sitting behind the
existing call sheet upload.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` (already has this feature) and
   send it back, same pattern as prior rebuild handoffs.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present
3. If you have other pending dc.html changes queued in your sandbox, fold
   this rebuild in with those and confirm in your reply what's included.

## Not asking you to change anything else

No new dc.html edits needed for this feature — it's already written and
pushed. This is purely: rebuild the compiled artifact from current source
and confirm the two markers.
