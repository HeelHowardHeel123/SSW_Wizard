# TX Locations: rows past row 8 have no formatting — fix is in source, needs a rebuild

## What's wrong

The Locations tab's template only has real, styled `<c>` cells on rows 5-8 (the old hardcoded "DAY 1 of 4"..."DAY 4 of 4" rows, now blanked but still carrying their style). A call sheet with more than 4 locations writes rows 9, 10, ... via `ensureRowInSheet`, which creates a bare `<row r="N"></row>` with no style at all — so those rows render with no borders/formatting while rows 5-8 look normal. Confirmed live: a real 6-location call sheet run through production wrote rows 5-10, and only 5-8 had any `s="..."` on their cells.

## What I already fixed (source only)

Both in this repo, pushed to `dev`:

- **`source/Production Binder Wizard.dc.html`** (`extractTxLocations` write loop, ~line 4062): any row past `TX_LOC_LEGACY_LAST` (8) now passes `styleFrom: col + TX_LOC_LEGACY_LAST` on every cell patch for that row — including the "clear" patches for empty fields (e.g. an unused Notes cell), not just the ones with a value.
- **`source/workbook-engine.js`**: `buildCell` gained a `"blank"` type (a styled cell with no value). The `cellPatches` loop's `clear` branch now calls `ensureRowInSheet` first (previously that only ran for value patches) and, when the target cell doesn't exist yet and `styleFrom` is set, creates a styled blank cell instead of silently no-opping. This was needed because a cleared field on a brand-new row previously got no cell at all, not even a formatted blank one.

Verified by hand-tracing the regex/patch logic against the actual failing case (I9/I10, the Notes column, both cleared on rows the template doesn't have). No other `clear: true` call site in the file passes `styleFrom`, so every other tab's behavior is unchanged — this only activates where a caller opts in.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current `source/Production Binder Wizard.dc.html` + `source/workbook-engine.js` (both already have this fix) and send it back as its own deliverable, same pattern as prior rebuild handoffs.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present
3. If you have other pending dc.html/engine changes queued in your sandbox already, fold this rebuild in with those rather than shipping two round-trips — just confirm in your reply what's included.

## Not asking you to change anything else

No new dc.html/engine edits needed for this specific bug — the fix is already written and pushed. This is purely: rebuild the compiled artifact from current source and confirm the two markers.
