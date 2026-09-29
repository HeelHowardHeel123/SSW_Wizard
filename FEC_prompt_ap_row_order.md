# AP tab row order fix — needs a bundle rebuild

## What's new

The TX AP tab was writing rows out of upload order because
`extractTxAp()` alphabetically sorted the uploaded PDFs by filename
before processing them. That breaks numeric PO ordering: "PO 25-012 -
10" sorts before "PO 25-012 - 2" as strings, since "1" < "2"
character-by-character. Confirmed on real data — "PO 25-012 - 2 -
Aspen Travel" was the first PDF added to the batch but ended up the
37th row in the AP tab.

## What I already built (source only, pushed to `dev` as `d341d8a`)

`source/Production Binder Wizard.dc.html`, `extractTxAp(fileList)`:
removed the `.sort()` call entirely. Rows are now written in whatever
order the browser's file picker/drop provided them — i.e. the order
the files were actually added — instead of any alphabetical or
numeric re-sort.

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
