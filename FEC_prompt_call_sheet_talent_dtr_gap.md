# Call-sheet Talent + DTR double-count fix — needs a bundle rebuild

## What's new

Thanks for catching this in your last rebuild — you were right. A
person who only appears in the call sheet's Talent section (no Talent
Payroll or Talent IC row of their own) and also submitted a DTR form
was showing up twice on Cast & Crew List: once as a spurious Crew
Payroll orphan (their DTR name matched nobody on Talent Payroll/IC —
the only rosters checked before Crew Payroll runs), and once as the
correct new Talent row from the call-sheet Talent feature.

## What I already built (source only, pushed to `dev` as `a5fef49`, on top of `b90e792`)

`source/Production Binder Wizard.dc.html`: new
`markDtrMatchedCallSheetTalent(dtrMatchedElsewhere, issues)` — fuzzy-
matches (`/match-names`) DTR names against the call sheet's own Talent
list, and marks any match into the shared `dtrMatchedElsewhere` Set
**before** Crew Payroll's DTR-orphan step runs (same placement logic
as why Crew Payroll already runs last among the 4 TX tabs). This only
suppresses the false Crew Payroll row — it doesn't touch
`castCrewEntries` itself; the real Talent row is still built once,
later, by the existing `matchCallSheetTalentIntoCastCrew`.

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
