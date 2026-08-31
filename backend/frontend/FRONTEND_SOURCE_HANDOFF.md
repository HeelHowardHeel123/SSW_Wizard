# Frontend source handoff — rebuild for the main push (Aug 31, 2026)

## Short answer

Still **no git on my side.** Everything you need is in this package — commit it
yourself. `source/` is the canonical frontend source; `backend/frontend/index.html`
is the build artifact committed alongside it. Both are current as of today.

Only **Wizard 01 (Name PDFs)** changed since the Aug 23 note. Wizard 02/03/04,
the workbook engine, the bundled templates, and every GA/IL tab mapping are
byte-identical to what you already merged.

## Files in this package

| Path | What it is |
|---|---|
| `backend/frontend/index.html` | **Freshly rebuilt production bundle**, 3.27 MB. Self-contained: the four workbook templates are embedded as base64 via `window.__OM_EMBEDDED_WORKBOOKS__` + a `fetch` shim, and `workbook-engine.js` is embedded as module source handed to the app as a blob URL. **Compiled output — never hand-edit.** |
| `source/Production Binder Wizard.dc.html` | The real editable source. All UI, upload zones, column mapping, extraction calls. |
| `source/workbook-engine.js` | xlsx read/patch/insert engine. **Unchanged** since Aug 23. |
| `source/wrapbook-fringe.js` | Wrapbook fringe parsing helper. **Unchanged.** |
| `source/support.js` | Runtime the `.dc.html` loads. Generated — don't edit. |
| `source/CLAUDE.md` | Template versions + column mappings for every tab, plus the Wizard 01 contract notes. **Read this before touching geometry.** |

Not included because they didn't change: `source/assets/workbooks/*.xlsx` (the
four bundled defaults, already in the repo and already embedded in the bundle
above), `backend/main.py`, `requirements.txt`.

## What's new in this bundle

1. **Vendor invoice naming toggle** on the Wizard 01 intake screen. Posts
   `vendor_naming` = `"invoice_number"` (default) or `"po_number"` on
   `POST /wizard01/jobs`. Label, note, choices and default are read from
   `GET /wizard01/conventions` (`types[].options[]` where
   `key === "vendor_naming"`); only the two values are built in, so it works
   before that endpoint deploys. Not part of the intake gate — always has a
   default.
2. **Intake screen simplified.**
   - **All address fields removed** — they never affected a filename. The POST
     no longer sends `client_address` / `agency_address` / `prodco_address`.
   - **Client removed as a batch source.** "Who is this batch from?" is now
     Production Co. or Agency only (a client never bills a production).
   - **Only the sender's name is required.** The other entity's name is still
     collected and posted, optional — the backend's Bill-To cross-check is what
     catches a mis-declared batch, so more names is still better.
   - On an **Agency** batch an optional **Client Name** card appears and posts as
     `client_name`; on a ProdCo batch that field is hidden and `client_name`
     posts empty.
3. **The Wizard 01 upload POST moves off Cloudflare.** New `UPLOAD_BACKEND_URL`
   getter + `UPLOAD_HOSTS_MAP`: on `tealdocwizard.com` / `www.` the single
   multipart `POST /wizard01/jobs` goes to `https://upload.tealdocwizard.com`
   (DNS-only, unproxied, same Railway service) so a multi-GB batch isn't 413'd
   at the edge by Cloudflare's body-size cap. Everywhere else it resolves to
   `BACKEND_URL` unchanged. Status polling, download and delete all stay on
   `BACKEND_URL`. **No backend code change — the FastAPI app just needs to
   answer that second hostname, and the DNS record has to exist.**
4. **`extractFringe` chunk size 15 → 5**, matching the backend's own
   5-concurrent cap so one request's worst case is a single ~40s round instead
   of three sequential ones. The old value tripped Cloudflare's ~100s proxy
   timeout (**524**) on a real 13-file ADQ 005 run. Don't raise it again without
   moving that call off the Cloudflare proxy.
5. **Wizard 01's landing card pill is back to "Active"** (aquamarine) now that
   the wizard is going live.

## Post-rebuild injection — re-verify on any future rebuild

- `__OM_EMBEDDED_WORKBOOKS__` — 4 keys, present in this build.
- `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key, present in this build.

Absent → every template load 404s and Generate Workbook dies on
*"Failed to resolve module specifier"*.

## Backend targeting — unchanged

`BACKEND_URL` resolution, the `BACKEND_FALLBACK` dev host, and
`X-App-Secret` (`tpcSSW201005`) are all as documented Aug 6. Nothing to change
for this sync.

## Known-open, non-blocking

Everything on the Aug 23 list still stands unchanged (Loan Out col A blank, GA
Talent soft geometry check, Talent invoice roll-up capped at 15, four tabs
awaiting backend arrays, GL tabs never live-tested). Plus, for Wizard 01:

- **The whole `/wizard01/*` surface is still stubbed on the frontend's side of
  the contract** — jobs, status polling, download, delete, conventions. The UI is
  complete and tested against hand-written payloads; nothing has run against a
  real backend. Hence the "In Development" pill.
- **`GET /wizard01/conventions` 404 is expected** — the conventions panel says
  the endpoint is unreachable rather than showing invented patterns, and the
  naming toggle falls back to two built-in labels. Never inline the patterns.

## Merge

No blockers. Nothing mid-flight. One merge lands both halves in sync since the
frontend ships inside this repo as `backend/frontend/index.html`.
