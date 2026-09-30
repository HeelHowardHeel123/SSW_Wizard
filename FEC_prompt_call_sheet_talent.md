# Call-sheet Talent on Cast & Crew List — needs a bundle rebuild

## What's new

The Cast & Crew List tab was never getting Talent from the call sheet.
Every call sheet has its own TALENT section (and sometimes a separate
BACKGROUND section) listing on-camera performers, but only the crew
departments were ever extracted from it — that fed Crew Payroll and
Crew IC's "Call Sheet" matching, but Talent Payroll and Talent IC never
got the same treatment. A Talent person who only ever showed up on the
call sheet, or whose call-sheet spelling didn't exactly match their
Talent Payroll/IC name, was invisible on Cast & Crew List.

## What I already built (source + backend, pushed to `dev` as `b90e792`)

1. **`backend/tx_crew_roster_extraction_prompt.txt`** — the call-sheet
   extraction prompt now pulls a second "talent" list alongside "crew",
   covering both a TALENT heading and a separate BACKGROUND heading
   when the call sheet has one. Skips rows that don't name an
   individual (group headcounts, vendor/casting-agency references).
2. **`backend/main.py`** — the crew-roster extraction pipeline threads
   that talent list through the same per-page Claude calls as crew,
   deduped the same way. `/extract-tx-crew-roster` now returns
   `{"crew": [...], "talent": [...], ...}`.
3. **`source/Production Binder Wizard.dc.html`** — new
   `matchCallSheetTalentIntoCastCrew()` fuzzy-matches the call sheet's
   talent list (via `/match-names`, same mechanism already used for
   Crew) against the Talent people already collected from Talent
   Payroll + Talent IC. A match adds a "Call Sheet" tag to that
   person's row; a call-sheet Talent person who matches nobody becomes
   a brand-new Cast & Crew List row. Runs after Crew Payroll, right
   before the Cast & Crew List tab is written. **Talent Payroll and
   Talent IC themselves are untouched** — this only affects the Cast &
   Crew List tab.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` and send it back.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

The backend half of this (prompt + main.py) deploys on its own and
needs no rebuild from you — this request is purely the frontend bundle
rebuild for the dc.html change above.
