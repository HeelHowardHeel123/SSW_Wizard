# Frontend source handoff — MAIN post-TX-merge rebuild (Sep 30, 2026)

## Why this package exists

MAIN received the full DEV → MAIN Texas merge (`4473d91`). This is the compiled
frontend rebuilt from that merged source. **No dc.html edits in this pass** — pure
recompile using MAIN's own current engine and templates.

## What changed in this tree

| Path | What it is |
|---|---|
| `backend/frontend/index.html` | **Freshly rebuilt production bundle**, 3.56 MB. Self-contained. **Compiled output — never hand-edit.** |
| `source/Production Binder Wizard.dc.html` | Canonical source as pushed in `4473d91` (~9,600 lines). |
| `source/workbook-engine.js` | MAIN's current engine (44,572 chars; exports `generateWorkbook`, `inspectStructure`). Embedded verbatim. |
| `source/assets/workbooks/texas.xlsx` | MAIN's current TX template (118,365 bytes). Embedded verbatim. |

`georgia.xlsx`, `illinois-local.xlsx`, `illinois-oos.xlsx` are byte-identical to the
previous bundle and were carried over unchanged.

## Post-rebuild injection — verified on THIS build

- `__OM_EMBEDDED_WORKBOOKS__` — **4 keys**: georgia 302,169 · illinois-local 799,338 ·
  illinois-oos 807,273 · texas 118,365 bytes.
- `__OM_EMBEDDED_MODULE_SOURCE__` — **1 key**: `./workbook-engine.js`.
- Build check: re-running the same compile over the previous MAIN source reproduces the
  previous bundle's page template byte-for-byte, so the only differences in this bundle
  are the new source, engine and `texas.xlsx`.

## What's in this bundle (from the merge)

- TX Cast & Crew List (aggregated Crew + Talent, one row per person, Y/N DTR column).
- Call-sheet Talent extraction + matching; DTR cross-tab orphan fix.
- "Last, First" everywhere on the TX rosters; personal suffixes stay with the surname,
  entity names never reordered.
- AP extraction fixes, TX Locations, Petty Cash / ProdCC, Talent Payroll.
- Crew freelance uploads in 5-file batches via `UPLOAD_BACKEND_URL` (Cloudflare bypass).
- Start over resets `this.files` and `this._downloaded`.
- Various GA/IL fixes and new payroll parsers already in source.

## Backend targeting — unchanged

Same `BACKEND_URL` / `UPLOAD_BACKEND_URL` resolution and `X-App-Secret` as every prior
build.
