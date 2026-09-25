# TX DTR matching: new feature, needs a bundle rebuild

## What's new

A new "Valid DTR Form" YES/NO feature, closely mirroring the call-sheet
crew matching you've already rebuilt for a few times this session. Backend
and source are done and pushed to `dev` (`90a4b40`).

Backend: new `POST /extract-tx-dtr` — one Claude vision call per uploaded
DTR ("Declaration of Texas Residency" — a Texas Film Commission form)
file. Unlike a call sheet, a DTR is a one-person document, so this is
simpler: one call per file, no page chunking.

Frontend (`source/Production Binder Wizard.dc.html`):
- New upload zone `txDtrForms` ("DTR Forms"), in the Crew section next to
  Call Sheets, accepting PDF/JPG/PNG.
- `extractTxDtr()`, `getDtrNames()` (cached per build, same pattern as
  `getCallSheetCrew()`), `matchDtrToRoster()` (reuses `POST /match-names`,
  same as call-sheet matching).
- Wired into **four** tabs' existing row-build loops: Crew Payroll (col K),
  Crew - Indepen. Contractors (col O), Talent Payroll (col J), Talent -
  Indep. Contract (col O, shares a write loop with Crew IC). Only Crew
  Payroll adds new rows for unmatched DTR people; the other three tabs only
  flag existing rows.

No new upload-zone UI needed beyond the standard zone declaration — same
generic upload mechanism as every other zone.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` (already has this feature)
   and send it back, same pattern as prior rebuild handoffs.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present
3. If you have other pending dc.html changes queued in your sandbox, fold
   this rebuild in with those and confirm in your reply what's included.

## Not asking you to change anything else

No new dc.html edits needed for this feature — it's already written and
pushed. This is purely: rebuild the compiled artifact from current source
and confirm the two markers.
