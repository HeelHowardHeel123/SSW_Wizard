# Orphan row name formatting fix — needs a bundle rebuild

## What was wrong

Call-sheet and DTR orphan rows wrote the Crew Name cell verbatim from the
source document — e.g. "ROBBIE MECKNA" instead of "Meckna, Robbie". Call
sheets are often printed ALL CAPS, and extraction preserves that literally.
Every matched payroll row's own Crew Name already reads "Last, First" in
normal case, so an orphan row stood out.

## What I already fixed (source only, pushed to `dev` as `9395e02`)

`source/Production Binder Wizard.dc.html`: new `_properLastFirst()`
helper (reorders "First Last" → "Last, First" and proper-cases it,
handling hyphens/apostrophes reasonably), applied only where an orphan
row's name is constructed in `matchCallSheetToRoster` and
`matchDtrToRoster` — not to matched rows, whose name already comes from
payroll/report data in the right shape. The "Name on Call sheet" column
(`callSheetName`) keeps the raw, unreformatted spelling — that's
specifically what it's for.

Same scope as the existing `_lastFirst` helper already in this file (last
word = surname, so a compound surname like "Van Amber" splits wrong) — a
known, accepted limitation, not something this chases further.

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
