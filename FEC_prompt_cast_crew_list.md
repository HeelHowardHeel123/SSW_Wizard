# New TX tab populated: Cast & Crew List — needs a bundle rebuild

## What's new

The TX template's "Cast & Crew List" tab has never been populated by the
wizard before — it was just recently updated (filler text removed, new
"From Documents" column added, the old two-block Crew/Cast layout
collapsed into one continuous 85-row band). This is the first code that
writes to it.

## What I already built (source only, pushed to `dev` as `e99aa9e`)

`source/Production Binder Wizard.dc.html`, inside `buildTexasBlob()`:

- A new `castCrewEntries` accumulator that collects one entry per person
  from all 4 existing TX rosters (Crew Payroll, Crew - Indepen.
  Contractors, Talent Payroll, Talent - Indep. Contract) as each of those
  tabs finishes being computed — reusing the `hasDtr` / `onCallSheet`
  flags those tabs already compute, no new extraction or matching calls.
- A new explicit `_orphan: true` marker set on call-sheet/DTR orphan rows
  at the 3 places they're created (previously there was no way to tell a
  "real" payroll/invoice match from an orphan added only via matching).
- After all 4 tabs finish, the Cast & Crew List tab is written: Crew block
  first (alphabetized by last name), Talent block after (alphabetized
  separately), with columns B (row #), C (Crew/Talent), D (Name), E
  (Role), F (DTR Y/N), and the new I ("From Documents" — e.g. "Crew
  Payroll, Call Sheet, DTR"). Same insert-rows-to-grow pattern as every
  other tab if the combined headcount exceeds the 85 pre-built rows.
- New constants: `TX_CAST_CREW_SHEET`, `TX_CAST_CREW_FIRST_ROW`,
  `TX_CAST_CREW_LAST_ROW`, `TX_CAST_CREW_EXPECT`, plus
  `txCastCrewFormulaRow()` (returns `[]` — this sheet has no per-row
  formulas at all) and `txCastCrewTagList()`.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` and send it back, same
   pattern as prior rebuild handoffs.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

No new dc.html edits needed — the fix is already written and pushed. This
is purely: rebuild the compiled artifact from current source and confirm
the two markers.
