# Frontend source handoff — rebuild for the main push (Aug 23, 2026)

## Short answer

Still **no git on my side.** Everything you need is in this repo working tree
right now — commit it yourself. `source/` is the canonical frontend source;
`backend/frontend/index.html` is the build artifact committed alongside it.

Both were rewritten today from the current sandbox state. The previous versions
in this tree were stale (`source/workbook-engine.js` was ~10 KB behind).

## What I just wrote into this tree

| Path | What it is |
|---|---|
| `backend/frontend/index.html` | **Freshly rebuilt production bundle**, 3.37 MB. Self-contained: the four workbook templates are embedded as base64 via `window.__OM_EMBEDDED_WORKBOOKS__` + a `fetch` shim, and `workbook-engine.js` is embedded as module source and handed to the app as a blob URL. **Compiled output — never hand-edit.** |
| `source/Production Binder Wizard.dc.html` | The real editable source (~5,850 lines). All UI, upload zones, column mapping, extraction calls. |
| `source/workbook-engine.js` | xlsx read/patch/insert engine. Now includes `shiftMergeCells` and read-only `inspectStructure`. |
| `source/wrapbook-fringe.js` | Wrapbook fringe parsing helper. |
| `source/support.js` | Runtime the `.dc.html` loads. Generated — don't edit. |
| `source/assets/workbooks/*.xlsx` | The four bundled default templates. |
| `source/CLAUDE.md` | Current template versions + column mappings for every GA tab. **Read this before touching geometry.** |

## Post-rebuild injection — verified in-browser

Both globals present and working (re-check these on any future rebuild):

- `__OM_EMBEDDED_WORKBOOKS__` — 4 keys. `fetch('assets/workbooks/georgia.xlsx')`
  returns **302,169 bytes**, byte-identical to source, no network request.
- `__OM_EMBEDDED_MODULES__` — 1 key; the blob imports and exports
  `generateWorkbook`.

Absent → every template load 404s / Generate Workbook dies on
*"Failed to resolve module specifier"*.

## What's in this bundle that wasn't in the last one

- **Georgia marked Active** (dev pill removed) on the workbook picker.
- **Agency Billings hidden on GA** — no destination tab there, so the whole
  Agency section now drops out on GA runs. Still writes on both IL wizards.
- **Talent (GA)** — Highland Talent added as a third payroll company; multi-file
  PTIP for all three; `loan_out` → AE with YES/NO normalization (a raw `"NO"`
  string was truthy before); column map current to *Template GA … 2026 08 20*
  (Other Fees inserted at S).
- **Loan Out Withholding tab** — accumulates `loan_out_rows` off four endpoints.
- **Payroll Roster tab** — deduped by (name, job title) from
  `payroll_roster_rows`; Y wins over N.
- **GL(BILLING) / GL (PRODCO) tabs** — literal ledgers, never deduped. PRODCO
  crew rows are built client-side from the reconciled Payroll Report set, so
  they can't double-count.
- **Payroll Report (GA)** — reconciliation via `/reconcile-ga-payroll`, timecards
  slot, 3-way sort selector, current 2026-17-10 column layout.
- **Hotel Charges Summary (GA)** — two-axis growth (hotels x guests).
- **Engine**: row insertion now shifts `<mergeCells>`; soft geometry checks
  (`inspectStructure`) for GA Talent / Loan Out / Payroll Roster / GL, hard
  preconditions for GA Payroll Report.

## Backend targeting — unchanged

`BACKEND_URL` resolution is the same as the Aug 6 note: same-origin on
`tealdocwizard.com`, `www.tealdocwizard.com`,
`sswwizard-production.up.railway.app`,
`sswwizard-production-7974.up.railway.app`; `BACKEND_FALLBACK` (dev) anywhere
else, so sandbox runs never touch prod. `X-App-Secret` still `tpcSSW201005`.
Nothing to change for the main sync.

## Known-open, non-blocking (documented, not bugs to fix before merge)

1. **Loan Out Withholding col A (Job Code) ships blank** with a visible run
   issue — there is no Job Code field on the Overview step to stamp it from.
   One-liner once that input exists.
2. **GA Talent geometry is checked softly** because the bundled
   `assets/workbooks/georgia.xlsx` is still the *older* Talent layout (no Misc
   Payment / Signatory Fee / PTIP columns). A first-time user who never
   uploaded a template gets the Talent tab skipped with an issue rather than an
   aborted workbook. **Publishing the current GA template via Wizard 04 fixes
   this for everyone without a redeploy** — worth doing right after the merge.
3. **Talent invoice roll-up caps at 15 distinct invoices** (the block shares
   rows 43–72 with the AICP category block). Accepted stopgap; raises a visible
   issue naming the overflow.
4. **Four tabs have no backend rows yet** — Loan Out Withholding, Payroll
   Roster, GL(BILLING), GL (PRODCO) (and the Highland branch). The frontend
   harvests `loan_out_rows` / `payroll_roster_rows` / `gl_*_rows` off existing
   envelopes, so a response without those keys is the normal case today and
   those tabs simply write nothing. Not a merge blocker — they go live the
   moment the backend starts sending the arrays, no frontend change needed.
5. **GL(BILLING) / GL (PRODCO) have never had a live test run.** Everything
   else in this bundle has.

## Merge

No blockers on my side. Nothing mid-flight. Merge whenever — simultaneous is
fine, and since the frontend ships inside this repo as
`backend/frontend/index.html`, one merge lands both halves in sync anyway.
