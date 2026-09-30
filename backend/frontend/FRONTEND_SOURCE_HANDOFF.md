# Frontend source handoff — bundle rebuild (Sep 18, 2026)

## Why this package exists

The previous Wizard023 export shipped a current `source/Production Binder Wizard.dc.html`
but carried the OLD compiled `backend/frontend/index.html` forward unrebuilt, so none
of the reliability work or the TX Talent Payroll tab actually deployed. This package is
that rebuild. **No dc.html behaviour changed in this pass** beyond what was already in
source — it is purely a recompile, plus the corrected TX template.

## What changed in this tree

| Path | What it is |
|---|---|
| `backend/frontend/index.html` | **Freshly rebuilt production bundle**, 3.33 MB. Self-contained. **Compiled output — never hand-edit.** |
| `assets/workbooks/texas.xlsx` | **Replaced** with *Template TX - State Submission Workbook - MAY 2026 2.1* (118,800 bytes) — the corrected Talent Payroll layout with On PTIP / On PDF at C/D. Also published via Wizard 04. |
| `source/Production Binder Wizard.dc.html` | Canonical source. Unchanged from the Wizard023 export except the TX Talent Payroll wiring noted below. |
| `source/workbook-engine.js`, `support.js`, `wrapbook-fringe.js` | Synced from sandbox, unchanged. |

## Post-rebuild injection — verified in-browser on THIS build

- `__OM_EMBEDDED_WORKBOOKS__` — **4 keys** (georgia 302,169 · illinois-local 799,338 ·
  illinois-oos 807,273 · texas 118,800 bytes). The fetch shim resolves any `*.xlsx` URL
  by basename; a deliberately bogus host still returned the embedded 302,169-byte GA file.
- `__OM_EMBEDDED_MODULE_SOURCE__` — **1 key** (`./workbook-engine.js`). Blob import
  resolves and exports `generateWorkbook` + `inspectStructure`.
- App boots to Step 1 with no console errors.

## New in this bundle vs. the Aug 23 one

- **`postExtract` retries raw network failures** — 3 attempts, 1s/2s backoff, fresh
  AbortController per attempt; only the final exhausted failure counts toward `_runFails`.
- **`txFallbackRow` / `txFallbackReason`** — any TX submission that yields zero rows
  writes a visible placeholder row (name = filename, amount 0, reason in Notes) instead
  of vanishing.
- **Five TX vendor-packet tabs** — AP, Agency Vendor Exps, Post Production,
  Crew - Indepen. Contractors, Talent - Indep. Contract.
- **TX Talent Payroll tab** — `POST /extract-tx-talent-payroll`, one call carrying all
  invoice PDFs + all PTIP/report files (PDFs-only, report-only and both all valid).
  Column map B–AI; U/V/W and AK/AL/AM (+AN/AP/AQ) left as live formulas; J/K/L/X/Z/AD
  blank by design (manual reviewer fields); O/P zeroed. Growth inserts at row 10 so the
  row-40 SUBTOTAL band, the row-43 SUM band and the 45–68 roll-up SUMIFs all extend.
- **Four payroll companies, scoped per workbook family** — `TALENT_COMPANIES` entries now
  carry `on: ["il"|"ga"|"tx"]`. Extreme Reach and Teams everywhere; Highland GA + TX;
  **CMS Productions TX-only**. Adding either to IL later is a one-word array edit.
  `effectiveTalentCo()` falls back to Extreme Reach if the workbook type changes after a
  company was picked.
- **`downloadIssuesTxt`** object-unwrap fix.

## Notes for the merge

1. The TX Talent Payroll tab is checked **softly** (`inspectStructure`) then hard as a
   precondition — same pattern as GA Talent. With the corrected `texas.xlsx` now bundled
   AND published, it writes normally; an out-of-date published template skips just that
   tab with a named-cell message rather than aborting the workbook.
2. TX `check_number` is written to both G (Ref Number) and Y (Pymt #) per the template.
   Every one of the four companies leaves it blank today, so both read empty.
3. `ptip_excel_b64` is always `null` on TX — no Sorted PTIP deliverable on this tab, and
   the success screen's PTIP download panel stays hidden.
4. Everything in the Aug 23 "known-open" list still stands unchanged (GA Talent soft
   geometry, Loan Out col A blank, 15-invoice roll-up cap, the four GA side-tabs awaiting
   backend arrays).

## Backend targeting — unchanged

Same `BACKEND_URL` resolution and `X-App-Secret` as every prior build. Sandbox runs never
touch prod.
