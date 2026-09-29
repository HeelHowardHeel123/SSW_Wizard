# AP tab row order fix (revised) — needs a bundle rebuild

## What's new

Thanks for the catch on the previous version of this fix — you're
right that "preserve upload order" isn't a safe bet, since a
multi-select file picker can hand files back in OS listing order
rather than true pick order. Steven confirmed: switch to a
numeric-aware sort instead of no-sort at all. Scoped to just the AP tab
for now — not touching the other ~25 steps you flagged with the same
plain-alphabetical-sort pattern; that's a separate ask for later.

## What I already built (source only, pushed to `dev` as `55c2d4c`, supersedes `d341d8a`)

`source/Production Binder Wizard.dc.html`, `extractTxAp(fileList)`:
replaced the removed sort with
`a.name.localeCompare(b.name, undefined, { numeric: true })` — the
exact fix you suggested. This orders "PO 25-012 - 2" before
"PO 25-012 - 10" deterministically, regardless of what order the
browser/OS actually hands the files back in.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` and send it back.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

The ~25 other upload steps with the same plain-alphabetical-sort
pattern (TX Locations, Crew/Talent IC, Petty Cash, ProdCC, DTR, etc.)
are noted but intentionally out of scope for this rebuild — Steven
wants to revisit those separately, not bundle them into this fix.
