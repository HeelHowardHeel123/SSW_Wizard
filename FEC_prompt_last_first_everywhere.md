# Two small TX fixes — needs a bundle rebuild

## What's new

Two follow-up requests after the Cast & Crew List tab went live:

1. Every crew and talent name on Crew Payroll, Crew - Indepen.
   Contractors, Talent Payroll, and Talent - Indep. Contract guaranteed
   "Last, First" proper case — not just the call-sheet/DTR orphan rows
   that already got this treatment.
2. Cast & Crew List should show each person once — the same person can
   legitimately appear on more than one source tab (e.g. a real Crew
   Payroll match AND a separate Crew Freelance Invoice), and was
   showing up as two separate rows.

## What I already built (source only, pushed to `dev` as `294f34e` and `a80ce4d`)

`source/Production Binder Wizard.dc.html`: reused the existing
`_properLastFirst()` helper (already proven on orphan rows earlier this
session — handles both an already-comma-formatted name and a plain
"First Last" one, hyphens/apostrophes included) at 3 new call sites,
one per tab family, right after each roster is finalized and before
that tab's own cell-writing loop runs:

- Crew Payroll (`parseTxPayrollRows()`): every `r.worker` in `raw`.
- Talent Payroll (`buildTexasBlob()`): every `r.talent_name` in `tpRows`.
- Crew IC / Talent IC (shared loop): every `r.worker_name` in `icRows`.

Each mutates the row object's field in place, so the Cast & Crew List
aggregation (which reads these same fields afterward) inherits the fix
automatically — no separate change needed there.

Deliberately NOT touched: `callSheetName` (the "Name on Call sheet"
column) — that's meant to stay the verbatim call-sheet spelling, same
scoping decision as the original orphan-only fix.

For the dedup: a new `txCastCrewDedupe()` collapses the Crew list and the
Talent list (separately) by name — case/whitespace-insensitive, safe now
that names are forced to the same Last, First spelling everywhere by fix
#1 above — unioning each duplicate's "From Documents" tags (in a fixed
display order, not source order) and OR-ing `hasDtr`, keeping the first
non-empty role seen. Runs right before the existing alphabetical sort.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` and send it back.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

No new dc.html edits needed — the fix is already written and pushed. This
is purely: rebuild the compiled artifact from current source and confirm
the two markers.
