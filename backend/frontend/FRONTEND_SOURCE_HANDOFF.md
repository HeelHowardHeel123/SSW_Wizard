# Frontend source handoff — where the real source lives (Aug 6, 2026)

## Short answer

There is **no git anywhere on my side**. I have never committed or pushed to
SSW_Wizard (or any repo). My working source lives in a design sandbox
filesystem, and production gets updated by handing you a freshly built
single-file `backend/frontend/index.html`. That is why the last git commit
touching these files is yours from July 28 — every frontend change since then
existed only as a rebuilt bundle, not as a commit.

So: **path (b)**. Everything you need is in this repo working tree now; commit
it to dev yourself.

## What I just wrote into this tree

| Path | What it is |
|---|---|
| `backend/frontend/index.html` | **Freshly rebuilt production bundle** — current as of today, includes everything shipped since July 28. Self-contained: the workbook templates are embedded as base64, no asset folder needed at runtime. **Compiled output — never hand-edit.** |
| `source/Production Binder Wizard.dc.html` | The **real editable source** for that bundle (~4300 lines: all UI, all upload zones, all column mapping, all extraction calls). This is the file I actually edit. |
| `source/workbook-engine.js` | xlsx read/patch/insert engine (row + column inserts, formula rewrite, cell patches). |
| `source/wrapbook-fringe.js` | Wrapbook fringe parsing helper. |
| `source/support.js` | Runtime the `.dc.html` loads. Generated — don't edit. |
| `source/assets/workbooks/*.xlsx` | The four bundled templates: `illinois-local`, `illinois-oos`, `georgia`, `texas`. These are the defaults every browser uses out of the box. |
| `source/CLAUDE.md` | Notes on template versions / column mappings, incl. the IL Local 2.2 signatory-fee columns. |

`index.html` is built FROM `source/` — if you or anyone changes behavior, change
the `.dc.html` and rebuild; editing the bundle directly will be overwritten.

## What's in this bundle that wasn't on July 28

- **Georgia workbook** end to end (`georgia.xlsx` + full column mapping).
- **GA AP zone** → `/extract-ga-ap`, incl. BX / PD1 / PD2 code handling and PD1/PD2 promotion logic.
- **GA Call Sheet zone** → crew-position matching via `/match-ap-positions`, positions folded into AP descriptions.
- **GA Petty Cash zone** → `/extract-ga-petty-cash`, one file per request, 200-file cap, per-file failure isolation. Column decisions: `Qualified` is an Excel formula `=Amount-NonQualified`; `non_qualified` written as sent; DNQ notes written verbatim (never parsed); `proof_of_pc_remittance_crew` passes through so "MISSING" lands literally; Payment Entity comes client-side from Project Info prodco_name; leftover pre-built rows are blanked so sample coding can't leak into Legend lookups or reconciliation SUMIFs.
- **IL petty cash / sub-vendor / residency** tab work and the Template In Use indicator on Review.
- Wizard 04 per-browser template override (IndexedDB), banner branding tweaks.

## Backend targeting (read before merging to main)

`BACKEND_URL` in the `.dc.html` resolves like this:

```
BACKEND_HOSTS = tealdocwizard.com, www.tealdocwizard.com,
                sswwizard-production.up.railway.app,
                sswwizard-production-7974.up.railway.app
  -> same-origin, relative API paths (no CORS, nothing hardcoded)

anywhere else (sandbox, local file) -> BACKEND_FALLBACK =
  https://sswwizard-production-7974.up.railway.app   (DEV)
```

Served by the backend on either prod or dev, it calls that same host. The
absolute dev fallback only applies off-host, so sandbox runs never touch prod.
Nothing needs changing for the main sync — but if the prod host ever changes,
that list is the one place to update. `X-App-Secret` is still `tpcSSW201005`.

## Going forward

Cleanest arrangement: treat `source/` as the canonical frontend source in git,
and `backend/frontend/index.html` as a build artifact committed alongside it.
Each time I ship, I'll rewrite both in this tree and tell you, and you commit.
If you'd rather I stop touching `backend/frontend/index.html` directly and only
update `source/`, say so and I'll do that instead.

## Untested

The GA Petty Cash path has never been run against a live endpoint. Once dev is
up, I want one real test run before this is considered production-ready.
