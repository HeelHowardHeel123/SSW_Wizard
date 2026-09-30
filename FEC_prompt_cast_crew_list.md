# Cast & Crew List — fixed per your review, ready to rebuild

Thanks for catching both issues — you were right on both counts, and the
column map is confirmed fine (matches exactly, no changes needed there).

## What I fixed (source only, pushed to `dev` as `84a5fae`)

1. **Dropped row insertion entirely.** You were right that the K/L totals
   block (1st Shoot date, TX-only counts) sharing rows 4-9 with the top of
   the data band meant inserting rows would push those formula cells down.
   Switched to the same pattern `Locations` already uses for the identical
   reason (its column L Rural Uplift list shares rows with Locations' own
   data): every person is written with direct `cellPatches` using
   `ensureRow: true`, which creates the row if it doesn't exist yet without
   shifting anything else — no `inserts` job at all anymore. Rows past the
   85 pre-built (row 89) borrow row 89's cell style, same as Locations does
   past its own legacy band.
2. **L7/L8 range fix — went with your second option** (code rewrites the
   formulas) rather than widening the template: after writing, if the
   actual combined headcount pushes past row 89, the code now rewrites
   `L7 = COUNTIF(F5:F<actual last row>,"Y")` and
   `L8 = COUNTA(C5:C<actual last row + 1>)` to match. Preferred this over a
   template edit so the tab stays correct regardless of which template
   revision ends up loaded for a given run (Replace → Published →
   bundled) — same defensive posture the rest of this tab's structural
   checks already take. Left the ranges untouched for the common case
   (≤85 people), since the template's own default is already correct there.

No other changes — the column map, provenance-tag logic, and everything
else from the first pass stands as-is.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` and send it back.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

This is purely: rebuild the compiled artifact from current source and
confirm the two markers. No template edits needed on your end either —
the fix is entirely code-side.
