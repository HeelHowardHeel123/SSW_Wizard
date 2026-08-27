# Production Binder Wizard — project notes

## Wizard 01 — Name PDFs
- **Job-based, not request/response.** Classifying 100-250+ PDFs is 30+ minutes
  of LLM work, so `POST /wizard01/jobs` returns `{job_id}` immediately and the
  frontend polls `GET /wizard01/jobs/{id}` every **3s** (backend increments
  `processed` per file, ~5 concurrent). `…/download` streams the zip (409 while
  running, 404 expired); `DELETE` clears a job early. All four take
  `X-App-Secret` like everything else.
- **Job record and zip share ONE 2-day clock** — a bookmarked job that still
  answers `done` is always still downloadable. `job_id` is a random UUID4
  (deliberately unlike `/templates/{id}`, which are human-chosen and few), so it
  is safe in the URL. We write `#job=<id>` and mirror it to localStorage
  (`tpc_namer_job_v1`). **Hash wins over localStorage on mount** and forces
  `view:"namer"`; a job merely remembered in this browser does NOT hijack the
  landing screen.
- **The upload POST alone goes to `UPLOAD_BACKEND_URL`** (`upload.tealdocwizard.com`
  on prod, a DNS-only/unproxied subdomain on the same Railway service; identical
  to `BACKEND_URL` everywhere else). Cloudflare hard-caps request body size at
  the edge, so a multi-GB batch 413s before reaching the backend. Polling,
  download and delete stay on `BACKEND_URL`.
- **Upload is one multipart POST via XHR**, not fetch — purely for
  `upload.onprogress`, since a multi-GB POST with no progress is
  indistinguishable from a hang. Backend streams each file to disk as it
  arrives. A dropped connection costs a re-upload and nothing more (no work has
  started), so chunked upload was deliberately deferred. Soft nudge over
  **120 files** (`NAMER_BATCH_NUDGE`) — a courtesy, not a backend limit.
- **Intake gate (`namerIntakeMissing`)**: `received_from` (prodco or agency,
  never mixed) + **that entity's NAME**. Client was dropped as a source (a client
  never bills a production) — but on an AGENCY batch an optional Client Name
  field appears and is posted as `client_name` (empty on prodco batches) — and **addresses were dropped entirely** — they never
  affected a filename. The non-sender entity's name is still collected and posted
  (optional): the backend compares each invoice's Bill-To against every name it
  was given, which is the guard against the old tool's regression (naming
  invoices after who was *billed* instead of who was *billing*).
- **Sender-mismatch detection is Vendor-only.** Residency and Diversity docs
  have no "who is this addressed to" field, so for those two the intake screen
  carrying the confirmation burden is the whole defense.
- **Vendor invoice naming toggle** (`vendor_naming` on POST): `invoice_number`
  (default) vs `po_number`. Scoped to plain company invoices only — receipts,
  hotel folios and the labor-line-item pattern ignore it, and a PO-less invoice
  falls back to invoice-number naming backend-side. The control's label, note,
  choices and default are read from the conventions payload
  (`types[].options[]` where `key === "vendor_naming"`); only the two VALUES are
  built in, so the toggle still works before that endpoint deploys. It is not
  part of `namerIntakeMissing` — it always has a default.
- **Naming conventions are NOT hardcoded.** `GET /wizard01/conventions` returns
  them as data so they can't drift when a 4th document type ships. A 404 is
  expected until that endpoint deploys — the panel says so rather than showing
  invented patterns. Never inline the patterns as a fallback.
- `reason_code` is a closed enum mapped to real words in `NAMER_REASONS`
  (`missing_name`, `missing_date`, `missing_company`, `sender_mismatch`,
  `low_confidence`, plus the not-readable split `unclassified` = readable but no
  handler yet vs `unreadable` = genuinely corrupt). Unknown codes fall through
  to the raw string, never swallowed. `reason_detail` is free text shown verbatim.
- **Residency splits one upload into many documents** (backend dev 05fa536): a
  batch scan of a stack of physical IDs (1 page front-only, or 2 front+back per
  person) is sliced into one named file per person. Consequences:
  - **`log[]` can be LONGER than `total`.** `total`/`processed` count uploaded
    files only, so the progress counter is unaffected; `renamed` /
    `needs_review` / `not_readable` count OUTPUT documents. Never assume
    `log.length === total` — the log view derives its own totals.
  - Several entries legitimately share one `filename` (the upload's). Rows get a
    **"Split · 2 of 4"** badge, counted across the WHOLE log so the ordinal
    stays right inside a filtered tab, plus a one-line note above the table
    ("5 uploads produced 8 named documents").
  - Log view caps at **500 rows** and says how many more are in `_manifest.csv`
    rather than silently truncating.
- Every log entry carries its own freshly-generated `file_id` (UUID, stable for
  the job's life) — including single-output files, which used to reuse the
  upload's id. One upload splitting into 5 people means 5 distinct ids, which is
  what makes per-document retry/override workable later. Filename is a weak key
  across 250 files with duplicates, and now genuinely ambiguous on splits.
- Zip: one folder per type that had a classified file, with successfully-renamed
  AND `Unable_To_Rename_NNN` files together in their correct type folder; a
  `Not Readable/` folder for anything unclassifiable, original names intact;
  `_manifest.csv` at top level, **server-generated** (guaranteed to match what
  shipped, and outlives the 2-day status payload).
- **Discarding an in-flight run is a two-step confirm** (`namerConfirmStop`);
  finished jobs discard in one click. The footer button sits beside the benign
  "Show file-by-file log" toggle, and the DELETE destroys a 30-minute run plus
  the only two copies of the job id (hash + localStorage) — so the running case
  has to be asked for twice. Label flips to "Discard this run" while live.
- `jDead` (`namerError` with no `namerJob`) means expired/404/deleted: polling
  has already stopped, so the spinner, stats, progress label and discard-confirm
  all suppress on it. Without that flag a dead job spins forever.
- Folder drag-drop recurses `webkitGetAsEntry` — a plain `dataTransfer.files`
  read returns the directory and drops everything inside it. Non-`.pdf` is
  filtered; dedupe is on `name|size`.
- Zero shared state with Wizard 02 — no engine, no template geometry, no
  Electron concerns.

## Templates (Wizard 04)
- The workbook templates the wizards populate live at `assets/workbooks/*.xlsx`
  (`illinois-local.xlsx` is the only one wired to an active wizard today).
- That bundled file is the **default every browser uses out of the box** — a
  first-time user does NOT need to upload anything.
- Resolution chain for every build (`loadTemplateBytes`):
  **1. this browser's Replace** (IndexedDB, key = workbook id) →
  **2. the Published template** (`GET /templates/{id}`, backend Railway volume at
  `/data`, survives redeploys, gated by `X-App-Secret`) →
  **3. the bundled file** in `assets/workbooks/`.
- Wizard 04 has both actions per template. *Replace* = this browser only, good
  for testing a candidate; *Publish* (`PUT /templates/{id}`, multipart `file`) =
  everyone. *Revert* drops the local copy; *Un-publish* (`DELETE`) drops the
  shared one.
- Updating the bundled file in `assets/workbooks/` still requires a bundle
  rebuild + redeploy (templates are base64-embedded into `backend/frontend/index.html`
  with a `fetch` shim), so **Publish is the normal way to ship a template change**.
  The bundled files are the offline floor only. The code's column mapping must
  still match whatever template is in play — uploading swaps bytes, not geometry.\n- Current bundled IL default = **Template IL - Local - State Submission Workbook 2.2**
  (Signatory Fee on **both** tabs: Talent & Extras at **S** (Total = SUM(M:S)),
  and Crew Payroll at **AD** — inserted between Hand and Sub_Total, so every
  Crew column from old-AD onward shifted right by one. Crew rows leave AD blank;
  Sub_Total stays SUM(O:AC) per the template's own archetype formula).
- The default also carries three tabs ported in from a 2.1 base (inserted after
  **Billings**, before **Crew Payroll**): **Residency - Callsheet** (built out —
  header row 7, cols B–P, blank data rows 8–30; a conditional-format rule
  references a `Masterlist` defined name that does not exist in this workbook, so
  it's inert until defined — same as the source file) and **AICP Bid** / **PO
  Log** (blank placeholders). None are wired to an extractor yet. Ported by
  keeping the 2.1 style indices as-is (fonts/fills/borders/numFmts + the used
  cellXfs are identical between 2.1 and 2.2) and inlining their shared strings.
- The Review step shows **"Template In Use"** (Built-in default vs Uploaded ·
  filename) so you can confirm which file a run uses.

## Hotel Charges Summary (GA)
- The published GA template (`GET /templates/ga`, "Template GA - State Submission
  Workbook - MAY 2026 08 10") has ONE hotel section, rows 6–26, last row 26.
  Chrome rows 1–5 appear once (title B1, GSA config H3:K4 referenced as
  `$H$4..$K$4`, Shoot Date(s) B5 = `Overview!C9`).
- Section 6–26: header 6, sub-headers 8, labels 10, one guest block 11–24
  (11 name/date/PO, 12–15 room + taxes, 16 meals, 17 Total, 19 Nights, 21
  disallow, 23 meal sums), TOTALS footer 25–26.
- Growth is two-axis: clone rows **6–27** (footer + a spacer row) per extra
  hotel, then clone rows **11–24** per extra guest inside a section. Footers are
  regenerated over the blocks that actually landed. No leftover-blanking.

## Payroll Report tab (GA) — in progress

> **Template 2026 17 10 remapped the columns.** B/C (On Production Report /
> On Invoice PDF) were inserted after Employee Name and AF/AG (Automation
> Total / Automation Check) after Total Qualified - GA. Old-B onward shifted
> right by two; old-AD onward by four. Current layout: A name, B/C source
> flags, D–L address/corp/job/union, M base rate, N ttl hours, O COVID
> stipend, P wage payment, Q–U taxes, V–AA fringe/advances/handling,
> **AB/AC/AD/AE formulas**, AF automation total (plain number we write),
> **AG formula** (=AB-AF), AH/AI GA corp tax + SIT, AJ/AK first/last W/E,
> AL invoice no, AM loan-out, AN–AQ doc checkboxes, AR residency (mislabeled
> "Labor Code"), AS AICP code (22), **AT formula** (VLOOKUP on AS), AU–AW
> forms/qualify, AX notes, AY/AZ FF1/FF2, BA source code ("PR"), BB invoice
> date, BC payment entity. Skip-on-write = AB, AC, AD, AE, AG, AT, BA.
> Subtotal row is still 575; its SUBTOTAL band moved K:AE → M:AI.

### Reconciliation (`POST /reconcile-ga-payroll`)
- Takes the rows the two extractors already returned (`pdf_rows`,
  `production_report_rows`) plus `sort_option` — no new uploads. Returns
  `{rows, issues}` fully matched, flagged and sorted; the frontend writes them
  in the order they arrive.
- Returns a per-row **`aicpCode`** (int 1-25 or null) from one batched GPT
  classification over the whole reconciled set. We write it into **AS** verbatim
  and never substitute a default — null means AS stays blank (that row drops out
  of the AICP SUMIFs), which is expected degraded mode, not a bug.
- Also returns an optional **`notes`** string → written verbatim to **AX**
  (Notes). Two producers: `"[no Production Report match] <reason>"` on an
  unmatched PDF-only row, and `"Invoices: 330535 ($1,401.16), ..."` in
  person-level reconciliation mode (consolidated Production Report, no invoice
  column, so the row has no invoice number of its own). Usually absent → blank.
- `timecardTotal` (float or null) has **no column** in the template — verified
  against the old layout's header row, which had a plain `Timecard` YES/NO at
  old-AK (→ **AO** after the 2026-17-10 shift) and no total/check pair. It is
  informational, surfaced via `notes`. Known-open backend issue: two Production
  Report rows for one person can both carry the same `timecardTotal` when the
  timecards straddle both date ranges.
- AO is written only when `hasTimecard` is present — blank, not "NO", when no
  timecards were uploaded, so the column never claims someone lacks a timecard
  we never looked for.
- Matching is per invoice number: SSN last-4, then normalized name, then LLM
  fuzzy for stragglers within an already-matched invoice. Production Report
  wins on money; the PDF's own total only ever feeds **AF** (cross-check).
- `sort_option` ∈ `name_invoice` | `invoice_pdf_layout` |
  `production_report_layout`, chosen by a 3-way selector on the Crew upload
  section (GA only). `resolveGaSort()` silently falls back when the option's
  source wasn't uploaded — the selector is never a gate.
- The call is **best-effort**: any failure returns null and the run continues on
  the concatenated rows. It **always** raises an issue in that state (the
  reconciler is the only source of `aicpCode`, so every row's AS goes blank and
  the AICP totals read zero), plus a second issue when both sources were present
  and duplicate people are therefore possible. Flags then come from the `_src`
  tag stamped at extraction.
- Tab name is exactly **"Payroll Report"** (the "5. Payroll Report" form only
  appears in older example jobs, not the current template).
- Flat table, no cloning: header row 3, data rows **4–574** (571 pre-built rows),
  subtotal row **575** = `SUBTOTAL(9,…)` over M:AI across 4:574. Write N rows
  starting at row 4, one source row = one output row, never merged or split.
  Cap at 571 and raise a visible issue rather than spilling past 574; if we ever
  need more, rows must be *inserted before* 575 so the SUBTOTAL ranges expand.
- **Do not write** these — they are formulas identical on every row (verified at
  rows 4, 5, 200, 400, 574): `Z=SUM(M:Y)`, `AA=Y+M`, `AB=Z-AA`, `AC=AB`,
  `AP=VLOOKUP(AO,…)`. Also leave `AO` (AICP Code = 22) and `AW` (Source Code =
  PR) at their pre-filled defaults.
- **Blank every other data column on each row we touch**, and clear the same
  columns from the last written row through `GA_PAYROLL_EXAMPLE_LAST_ROW` (15).
  Rows 4–15 carry leftover example content (AS 4–6 YES/NO-GA/NO-OOS; AU 4–15,
  AV 4–7, AY 4–6 example FF1/FF2/Payment Entity incl. "TPC Productions, Inc.").
  A run with fewer than 12 people would otherwise ship those under the real crew,
  and stale AS values sit in the live SUMIFS criteria range at rows 580–612.
  It is not validation-source data — all five dropdowns
  (AI/AJ, AK/AL/AM/AR, AN, AQ, AS) use inline literal lists, not cell ranges — so
  clearing is safe.
- No hyperlinks or comments anywhere in rows 1–576, so the Hotel Charges
  hyperlink-artifact fix does **not** apply to this tab.
- Column mapping calls (backend-confirmed): K=BASE RATE, L=TTL HOURS. `wages` and
  `corporate` are mutually exclusive per row and BOTH write **N** (WAGE PAYMENT) —
  GA has no separate loan-out wage column; H (Corp Name) is the loan-out signal.
  `other` is **unmapped**: GA's S is OTHER TAXES (a tax-side column), CAPS' `other`
  is fringe-side. Awaiting Steven on whether other-fringe has any home on this tab.
- Out of scope for v1 (columns exist, nothing populates them): AS (Qualify?),
  AU/AV (FF1/FF2), AK/AL/AM/AR (document-presence YES/Missing/N/A checkboxes,
  not dollars), AI/AJ (Loan-Out/Contract Received), AQ (G7 Form), AN.
  **AN is mislabeled** — header says "Labor Code" but its dropdown is GA
  Resident / Non Resident / Non Qualifying, i.e. residency. We already compute
  shoot-state residency for the Qualify rule, so it's the best first automation
  candidate after v1.
- AD (Georgia Corp Tax) / AE (Georgia SIT) ARE live dollar columns. Backend is
  adding `withholdingsGA` / `corpTaxGA` to FRINGE_FIELDS (mirroring the existing
  `withholdingsIL`) so both paths populate them; they read 0 until that ships.

### Inputs
- **Payroll Batches (Timecards)** — `timecards` slot (`gaOnly`), `.pdf`,
  `/extract-ga-timecards` (files[] batches of 15, same shape as
  `/extract-fringe`). Rows are their OWN shape (worker, ssn, weekStart/weekEnd,
  total, batchNo, tcId), **not** FRINGE_FIELDS — they never become output rows,
  they only feed `timecard_rows` on the reconciler so it can set
  `hasTimecard` → **AO** (YES/NO) and `timecardTotal`. Some productions (PJ 004)
  have no invoice PDFs at all, only these. A timecards-ONLY run therefore writes
  no rows and returns early with one accurate issue — it never enters the
  reconciler-outage path, which would otherwise blame a call never made.
- **Crew Payroll PDFs** — the existing `payroll` slot, unchanged. Same files,
  same `/extract-payroll`, just a different destination tab for GA. Deliberately
  NOT duplicated as a GA-specific zone. It also carries the `payrollHints`
  free-text field already.
- **Production Report** — new `prodReport` slot (`gaOnly`), `.xlsx/.xls/.csv`,
  `/extract-ga-production-report`. Backend header-maps arbitrary column names, so
  no fixed geometry on our side.
- Both endpoints return the **same FRINGE_FIELDS shape**, so one mapper handles
  both. Row order is decided backend-side (Production Report order wins when both
  are present); the frontend writes the array in the order it arrives — no
  reordering, no dedupe.

### Structural preconditions
- Because **Publish** changes the file the code writes into with no redeploy,
  geometry is no longer a safe assumption. `generateWorkbook` takes
  `opts.preconditions: [{sheetName, expect}]`, checked by `checkSheetStructure`
  **before any mutation**; a mismatch throws with the specific cell problems and
  writes nothing. Asserted only for tabs a given run actually writes.
- Payroll Report contract (`GA_PAYROLL_EXPECT`, verified against the bundled
  template): header row 3 with A="EMPLOYEE NAME", Z="Total Amount",
  AO="AICP Code", AW="Source Code"; rows 4 and 574 present; row 575 formulas in
  K/Z/AE; Z/AA/AB/AC/AP are formulas on rows 4, 300, 574.
- Row-matching regexes must use `[^>]*?(?:\/>|>[\s\S]*?<\/row>)` — a greedy
  `[^>]*>` eats the slash of a self-closing `<row r="N"/>` and swallows the next
  row whole. The older `insertRowsIntoSheet` / `cloneRowBlock` matchers target a
  specific row number, so their failure mode there is a silent no-op instead.

## Talent tab (GA)

- Written by the existing `/extract-talent` endpoint — the same call and response
  shape the IL Talent & Extras path uses. GA just has more columns for it, so the
  three withholding fields IL summed away get their own homes.
- **Talent payroll companies** are declared in `TALENT_COMPANIES` (id = the
  `payroll_company` value): `er`, `teams`, and `highland` (**GA-only** — no IL
  mapping yet, so it's hidden on IL/OOS and `effectiveTalentCo()` falls back to
  `er` if the workbook type changes after it was picked). Highland's uploads have
  the same shape as Teams (N invoice PDFs + N reconciliation files posted as
  `ptip_files`); it just calls its file a **Payroll Report**, but the upload zone reads
  **"PTIP Report (Excel)"** for all three companies (one consistent label; only
  the per-company hint/help line differs).
  Backend `highland` branch not built yet.
- **PTIP uploads are multi-file for BOTH payroll companies now.** Teams sends one
  .xlsx per commercial ID, Extreme Reach one per invoice (GA). The frontend posts
  every file as repeated `ptip_files` and *additionally* sets `ptip_file` when
  there is exactly one, so an older backend that only reads `ptip_file` still
  works on single-file runs. The `talentPtip` zone has no file cap.
- Geometry (2026-08 **published** template — note the bundled
  `assets/workbooks/georgia.xlsx` is still the OLDER Talent layout, with no Misc
  Payment / Signatory Fee / PTIP columns): header row 3, **31** pre-built data
  rows 4–34, Subtotal 35, blank 36, TOTAL 37. Growth = insert at row 5
  (`FIRST + 1`) so the sheet's own SUBTOTAL/SUM/SUMIFS ranges bump with us — same
  technique as AP and Petty Cash.
- Column map (**Template GA … 2026 08 20**, which inserted **Other Fees at S** —
  everything from old-S onward shifted one right): A name, B–E address/city/zip/
  home state, F cast category, G invoice no, H invoice date, I work state,
  J wages, **K misc payment**, L P&H, M/N/O state/local/disability withheld,
  P employer taxes, Q workers comp, **R signatory fee**, **S other fees (NEW,
  plain data entry)**, T handling, **U/V/W formulas** (U=SUM(J:T), V=T, W=U-V),
  X pay type, Y payment entity, Z/AA FF1/FF2, AB source code ("PR"), AC qualified,
  AD notes, **AE loanout (NEW 2026-08-22, YES/NO from `loan_out`)**, AF AICP code,
  **AG formula** (AICP category VLOOKUP — never written), AH on PTIP?, AI on PDF,
  AJ amount on PTIP, **AK formula** (=AJ-U, PTIP_Check).
  Skip-on-write = U, V, W, AB, AG, AK. Money cols J–T zero-filled.
  (AG/AH were swapped in the template on 2026-08-19 — AG used to be the PTIP
  amount and AH the on-PDF flag.)
- Misc Payment (K) and Signatory Fee (R) were inserted mid-row in the 2026-08
  template, shifting everything from old-K onward right and moving the four PTIP
  reconciliation columns from AD–AG to **AF–AI**. Total Amount now sums all ten
  real components, so no fold-into-Wages workaround is needed.
- **Two roll-up blocks share rows 43–72.** Per-invoice block in cols F–L (data
  44–58, SUM 59) keyed on a hardcoded invoice list in **F** that we populate from
  the run's distinct invoice numbers — without it the block reads zero.
  AICP-category block in cols N–Y (data 44–71, SUM 72) keyed on AE + AB, which
  drives itself as long as AD is populated. Because they occupy the SAME rows,
  the invoice list **cannot** grow past its 15 pre-built rows without pushing
  blank rows into the middle of the category list — a run with more than 15
  distinct invoices writes the first 15 and raises a visible issue naming the
  rest. **Accepted stopgap** (confirmed with Steven/backend, not a bug to fix):
  revisit only if a real production actually exceeds 15 distinct Talent invoices,
  at which point the fix is a template change giving the invoice block its own
  rows.
- Example content runs rows 4–15 (pay type, payment entity, FF1/FF2, qualified,
  AICP) — blanked on every written row and cleared through row 15. AC is a live
  SUMIFS criteria range for both blocks, so a stale "YES" there would corrupt
  the roll-ups.
- `GA_TALENT_EXPECT` asserts header row 3 (A, G, J, K, R, S, T, AD, AF, AG), rows
  4/34, footer 35 formulas in J/T, and T/U/V/AE as formulas on rows 4/5/34. F and
  X headers wrap mid-cell so they are not asserted.
- **The Talent geometry is checked SOFTLY**, via the engine's new read-only
  `inspectStructure(buf, checks)`, *before* the write is assembled. A mismatch
  skips the Talent tab and raises a visible issue naming the bad cells; every
  other tab still writes. This is deliberate: the bundled default is the older
  layout, and a hard precondition would abort the entire workbook for a
  first-time user who never uploaded a template. (The same expect is ALSO pushed
  as a hard precondition on the runs that do pass the soft check, so geometry
  can't shift between check and write.) `GA_PAYROLL_EXPECT` stays hard-only — it
  was verified against the bundled file and passes there.
- Not yet wired: Z/AA (FF1/FF2), S (Other Fees — no extractor field yet; the
  frontend maps `other_fees` there and zero-fills it otherwise), and Talent Freelance invoices
  (`/extract-talent-freelance`) still only feed the IL tab.

## Loan Out Withholding tab (GA) — in progress
- **No upload of its own.** `/extract-fringe`, `/extract-talent`,
  `/extract-ga-ap` and `/extract-ga-production-report` each append a
  `loan_out_rows` array to their EXISTING response whenever that extraction
  meets a loan-out payment. The Production Report is often the ONLY place a crew
  loan-out is identifiable — payroll-company invoice PDFs frequently carry no
  loan-out marker.
  `harvestLoanOuts(j)` is called from all three and accumulates into
  `this._loanOutRows`; the tab is written from the pile. One appearance = one
  row, never consolidated (same granularity as the source tabs). **Backend not
  built yet** — a response with no `loan_out_rows` key is the normal case today.
- **Geometry differs from the backend brief** (which assumed header 4 / data 5).
  Verified against the published template: title A1, Job Code A3/B3, Shoot Date
  A4/B4 + project name C4, **header row 6**, **six** pre-built data rows **7–12**,
  all six carrying example content (SSNs, payroll companies, dates). Growth =
  insert at row 8 (`FIRST + 1`).
- Column map: C name, D position/title, E loan-out corp name, I invoice gross,
  K crew position (mirrors D by design), L payroll company, M address,
  N home state, O work state, P invoice date, Q invoice no, S notes.
- Skip-on-write = **B** (WTH Remitted?, a manual checkbox) and **J** (Amount
  Withheld — the template's own `=I*4.99%`, shared down the column from row 7).
  The inserted-row archetype writes J only.
- F/G/H (Federal Tax Classification / SSN / FEIN) and R (Payment Date) are
  blanked, not written — no pipeline collects them yet (v2), and the example rows
  carry values that would otherwise show through.
- **A (Job Code) is left blank** and a visible issue says so: it's ours to stamp,
  but there is no Job Code field on the Overview step to stamp it from. Adding
  that field and mapping it in `gaLoanOutRowToCells` is a one-liner.
- Checked **softly** (`inspectStructure`) like GA Talent, then hard as a
  precondition on runs that pass — the bundled default may predate this tab, and
  one missing tab shouldn't abort the workbook.

## Payroll Roster tab (GA) — in progress
- **No upload of its own**, like Loan Out Withholding: consolidates people the
  Crew Payroll reconciler, `/extract-talent` and `/extract-ga-ap` already
  identified, each appending `payroll_roster_rows` (one entry per person-role
  APPEARANCE) to its existing response. Crew entries are harvested from
  **`/reconcile-ga-payroll`**, not an extractor — the Y/N flag depends on whether
  the payment matched the Production Report, which is only known post-reconcile.
  `harvestLoanOuts()` collects both side tables off every envelope.
- **This tab is DEDUPED, not raw-accumulated.** `dedupeRoster()` groups by
  (last, first, job title) normalized; same person + different title = separate
  row. The surviving row keeps the **earliest `start_date`** and that
  appearance's other values. `on_payroll_report` = **Y if ANY appearance was Y**
  (one verified payment proves the person is on the report); `comment` is kept
  only while the group is still N.
- Geometry (published template): title A1, **header row 3**, **67** pre-built
  data rows **4–70**, no formulas, no merges. Growth = insert at row 5
  (`FIRST + 1`).
- **`last_name` is B, not A** — A is the sheet's own 1..67 item number, which we
  renumber for the rows we write. **I ships pre-filled "Y"** on all 67 rows, so it
  is always overwritten (never skipped) — otherwise an unwritten row would claim Y.
- Column map: B last, C first, D corp, E job title, F work state, G residence,
  H start date, I on-payroll-report, J vendor name, K comment. Skip list is empty.
- Checked softly then hard, same as GA Talent / Loan Out Withholding.

## GL(BILLING) / GL (PRODCO) tabs (GA) — in progress
- **Sheet names are not symmetrical**: `GL(BILLING)` (no space) and
  `GL (PRODCO)` (with one). Both share ONE column layout, so `GL_MAP` /
  `GL_EXPECT` / `glRowToCells` serve both.
- Geometry (published template): title A1, **header row 3**, **95** pre-built data
  rows **4–98**, no formulas, no merges. Rows **4–9** carry example content in
  A/K/L (FF2 GS/GL/NQ/DL, Source Code "AP") — cleared through row 9. Growth by
  plain insertion at row 5 is safe (nothing below to shift).
- **Literal transaction ledgers — one row per source occurrence, never deduped**
  (opposite of Payroll Roster).
- **A (Sequence Number)** is ours: 1..N across the tab. **B/C/D/M** (Account
  Number, Reference Number, Effective Date, JE Number) and **W** (no header) are
  blanked, never written — manual accounting-system fields.
- Sources — all harvested by `harvestLoanOuts()` off the existing envelopes:
  - GL(BILLING): `gl_billing_rows` from every `/extract-talent` call, plus
    `/extract-billings` **with `vendor_type=prodco`** (the only vendor_type that
    returns them; agency / sub_prodco calls carry none).
  - GL (PRODCO): `gl_prodco_rows` from `/extract-ga-ap`,
    `/extract-ga-petty-cash`, `/extract-ga-prodcc`, **plus crew rows built
    client-side** in `crewRowToGlProdco()` from the reconciled Payroll Report
    rows — the reconciled set only exists in this app, so emitting crew from
    either raw source would double-count nearly the whole crew.
- Crew ledger row: `invoiceNo`→F, `invoiceDate`→G, `worker` as "LAST,FIRST"→I
  (O/P stay blank), "PR"→L, `jobTitle`→N, `resState`→Q, "USD"→T, and
  `wages ?? corporate`→U (mutually exclusive per row; never summed across
  invoices or pay periods). E/H/J/K/R/S/V/X blank; FF2 awaits the GA-resident
  GL/DL rule.
- **ProdCo invoices reuse the existing "Production Company Billings" zone**, which
  was already visible on GA and previously had no GA destination — no new upload
  slot was added. `/extract-billings` also had `work_state` hardcoded to "IL";
  it now sends GA on GA runs.

## Engine notes
- **Row insertion now shifts `<mergeCells>`** (`shiftMergeCells`, called from
  `insertRowsIntoSheet` and `cloneRowBlock`). merge refs are absolute strings, so
  before this a merged range below the insert point (e.g. the GA Talent chrome at
  template row 42, just above the roll-up block header at 43) stayed put and
  landed on a written data row — the merge swallowed S/T so real values read as
  empty in Excel. Ranges straddling the insert point are stretched, not moved.
  `insertColumnSlots`-style column-scoped shifting is NOT merge-aware; if a merge
  ever appears in one of those tabs it will need the same treatment.
- Known cosmetic gap (unchanged): `conditionalFormatting`/`dataValidation`
  `sqref` ranges are also absolute and do not grow with inserted rows, so rows
  past the pre-built block lose dropdowns/formatting.

## TODO
- [ ] Write new-user instructions (how the tool works, that the correct template
      ships as the default so no setup is needed, and how/when to use Wizard 04
      to update a template).
