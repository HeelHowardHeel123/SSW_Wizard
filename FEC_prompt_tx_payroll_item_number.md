# Crew Payroll ITEM # numbering fix — needs a bundle rebuild

## What was wrong

Confirmed real on a 250-row ABBVIE 022 run: the Crew Payroll tab's ITEM #
column (B) showed `1`, then a ~155-row blank gap, then `2, 3, 4...46`
resuming partway down the sheet with no relationship to the real rows
sitting there.

Root cause: the template's B9-B54 are plain literal numbers 1-46, not a
formula chain. Row-growth-by-insertion shifts those literals down along
with everything else on the sheet, so once a run grows past the original
46 pre-built rows, the leftover literals land wherever the insertion
pushed them — meaningless at that point — while B was never written by the
frontend at all (it was in `TX_PAYROLL_SKIP`, on the now-incorrect
assumption the template's own numbers would keep working after growth).

## What I already fixed (source only, pushed to `dev` as `1938849`)

`source/Production Binder Wizard.dc.html`: removed `B` from
`TX_PAYROLL_SKIP`; `txPayrollToCells(r)` is now `txPayrollToCells(r, idx)`
and writes `B = idx + 1` explicitly for every row, same pattern the Crew -
Indepen. Contractors tab already uses for its own item numbering. The one
call site (`raw.map((r) => this.txPayrollToCells(r))`) now passes the
index too.

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
