# Call-sheet Talent DTR tag fix — needs a bundle rebuild

## What's new

Thanks again — the follow-on gap you flagged was right. The last fix
stopped the spurious Crew Payroll row for a call-sheet-only Talent
person with a DTR form, but their real Talent row (added later by
`matchCallSheetTalentIntoCastCrew`) never learned about that DTR
match — DTR column blank, no "DTR" tag, despite them actually having
one.

## What I already built (source only, pushed to `dev` as `d6eca3e`, on top of `a5fef49`)

`source/Production Binder Wizard.dc.html`:

- `markDtrMatchedCallSheetTalent` now fills a second Set,
  `dtrMatchedCallSheetTalentNames`, keyed by the call sheet's OWN name
  spelling (the existing `dtrMatchedElsewhere` Set is keyed by the DTR
  form's spelling instead, which doesn't help the later step identify
  which of ITS rows to flag).
- `matchCallSheetTalentIntoCastCrew` now takes that Set as a third
  param and consults it when building each row: sets `hasDtr` and adds
  the "DTR" tag whenever the call-sheet Talent name is in that set,
  and leaves the DTR column untouched (not a false "NO") whenever no
  DTR forms were uploaded at all this run — same convention every
  other Cast & Crew List row source already follows.

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
