# Run-log visibility fix — needs a bundle rebuild

## What's new

Found while reviewing a real test (SONI 005): `/extract-tx-crew-roster`
has returned a "talent" array since your `b90e792` rebuild, but the
run log's response summary only ever printed "N crew" — the talent
count and per-person detail lines never showed up, making the
call-sheet Talent feature much harder to verify from the log alone (I
had to open the workbook directly to confirm it actually worked).

## What I already built (source only, pushed to `dev` as `899d49b`, on top of `d6eca3e`)

`source/Production Binder Wizard.dc.html`, `postExtract()`'s response
logging — added a `" · N talent"` count to the summary line (mirroring
the existing crew count) and a `"  talent: <name> · <role>"` detail
line per person (mirroring the existing crew detail lines). No warning
for an empty talent list, since that's the normal case for any call
sheet with no TALENT section. Log-line change only — no behavior
change to the feature itself.

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
