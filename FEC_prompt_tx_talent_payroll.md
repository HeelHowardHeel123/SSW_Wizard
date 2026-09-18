# TX Talent Payroll — frontend wiring

## What this is

A new TX tab, "Talent Payroll" — the counterpart to the "Crew Payroll" tab, but for on-camera talent instead of crew. The backend is fully built and tested against real TX production files. This prompt covers everything needed to wire it into `Production Binder Wizard.dc.html`: an upload zone, a payroll-company selector, a column map, and the extraction call.

## The core new thing: a payroll-company selector

Talent payroll is processed by one of four different payroll companies, and each one sends genuinely different document formats (different PDF layouts, different PTIP/report spreadsheet formats). The backend already has GA/IL precedent for three of these four (same pattern you may recognize from the existing "Talent & Extras" zone's `payrollCompany` selector) — TX needs its own equivalent selector with a 4th option:

```
payrollCompany: "highland" | "er" | "teams" | "cms"
```

Labels for the UI:
- `highland` → "Highland Talent Payments"
- `er` → "Extreme Reach"
- `teams` → "The Team Companies (Teams)"
- `cms` → "CMS Productions"

## Upload zone

Two file inputs, both accepting multiple files:

1. **Talent Invoice PDFs** — `.pdf`, multiple. Maps to the endpoint's `pdf_files` field.
2. **PTIP / Payroll Report** — `.xlsx`, multiple. Maps to the endpoint's `ptip_files` field. Multiple files matter here — Teams sends one PTIP file per commercial ID, and a production may need to upload several at once.

All three intake scenarios must work, same as Crew Payroll's Production Report handling:
- **PDFs only** (no PTIP/report uploaded)
- **PTIP/Report only** (no PDFs uploaded)
- **Both** (the normal case)

The backend already handles all three scenarios correctly for all four companies — you don't need to build any of that logic, just don't block a submit that only has one of the two file types.

## The endpoint

```
POST /extract-tx-talent-payroll
```

Form fields:
| Field | Type | Notes |
|---|---|---|
| `pdf_files` | file[] | Talent Invoice PDFs |
| `ptip_files` | file[] | PTIP/Payroll Report file(s) — send as a list even if there's only one |
| `prodco_name` | string | Same as every other TX endpoint |
| `work_state` | string | Always `"TX"` for this tab |
| `payroll_company` | string | One of `"highland"`, `"er"`, `"teams"`, `"cms"` from the selector |

At least one of `pdf_files`/`ptip_files` must be present or the endpoint 400s with a clear message — same pattern as the other TX extractors, so `txFallbackRow`/retry handling should treat that the same way.

### Response shape

```json
{
  "rows": [ /* array of row objects, shape below */ ],
  "ptip_excel_b64": null,
  "summary": {
    "total_rows": 0,
    "invoices_with_pdf": ["604117", "604121"],
    "invoices_ptip_only": [],
    "duplicates_found": 0,
    "issues": ["free-text strings for the Files needing attention log"]
  }
}
```

`ptip_excel_b64` is always `null` for TX (no "Sorted PTIP" deliverable is built for this tab) — ignore it.

### Row object fields

Every row uses this exact same shape regardless of which payroll company produced it — some fields are simply blank/zero for companies that don't supply that data point. `talent_name` is already Title-Cased on the backend (e.g. `"Blake J Gibbons"`), matching every other TX tab's convention — no reformatting needed on your end.

```
item_no            int
on_ptip             bool   -- found on the PTIP/report
on_pdf               bool   -- found on a received PDF invoice
work_state          str
talent_name         str    -- already Title-Cased
loan_out            "YES" | "NO"
loan_out_company    str    -- usually blank; not identifiable for CMS/Highland
title               str    -- often blank (see field notes below)
invoice_no          str    -- may be a comma-joined list if this person spans several invoices
invoice_date        str    -- MM/DD/YYYY
wages               number
misc_pymt           number
er_tax              number
wc                  number -- Workers Comp; 0 for companies with no such line item (see notes)
handling            number
sag                 number -- Pension & Health
total               number
check_number        str    -- currently always blank for all four companies
received_invoice    bool   -- same value as on_pdf
payment_entity      str    -- e.g. "Highland Talent Payments, Inc"
type                str    -- e.g. "Session" / "Session Fee"; blank for some companies
home_address        str
city                str
state               str
zip                 str
commercial_id       str
commercial_title    str
notes                str    -- pre-built, human-readable; just drop straight into the Notes column
```

## Column map — "Talent Payroll" tab

Header row 8, data starts row 9 (same convention as every other TX tab). `TX_TALENT_PAYROLL_FIRST = 9`; use the same `bumpFormulaRefs(xml, 10, N)` row-growth pattern as the other tabs (insert at row 10, one past the first data row, so the whole-range footer formulas extend correctly).

| Col | Header | Source field | Notes |
|---|---|---|---|
| B | ITEM # | — | sequential, same as every tab |
| C | On PTIP | `on_ptip` | write `"YES"`/`"NO"` |
| D | On PDF | `on_pdf` | write `"YES"`/`"NO"` |
| E | Invoice # | `invoice_no` | |
| F | Invoice Date | `invoice_date` | |
| G | Ref Number (Check #) | `check_number` | will be blank for now — every company leaves this empty currently |
| H | Talent Name | `talent_name` | |
| I | Title | `title` | often blank — see field notes |
| J | Valid DTR Form | — | **leave blank.** No DTR-matching logic exists for this tab yet (same as Crew Payroll's own Valid DTR Form column) |
| K | Location | — | **leave blank.** No location data available from any of the 4 sources |
| L | TX Resident | — | **leave blank.** Confirmed explicitly: even when `state` happens to be known, don't auto-derive this — it's a manual reviewer field, same treatment as Qualify? |
| M | Gross Wages $ | `wages` | |
| N | Misc Payment | `misc_pymt` | |
| O | Mileage | — | **leave blank/0.** None of the 4 companies itemize this separately |
| P | Kit Rental | — | **leave blank/0.** Same as Mileage |
| Q | ER Payroll Taxes | `er_tax` | |
| R | Workers Comp | `wc` | 0 for Highland/Teams/CMS (no separate WC line item in their formats); real values only from ER |
| S | Handling/Vendor Fee | `handling` | |
| T | Pension & Health | `sag` | |
| U | Total | — | **live formula (`=SUM(M9:T9)`) — never write to this column** |
| V | Ineligible $ | — | **live formula (`=R9`) — never write to this column** |
| W | Submitted Amount | — | **live formula (`=U9-V9`) — never write to this column** |
| X | Pymt Method | — | **leave blank.** No payment-method data from any of the 4 sources |
| Y | Pymt # | `check_number` | same blank caveat as column G |
| Z | Proof of Payment | — | **leave blank.** No proof-of-payment document concept in this pipeline |
| AA | Payment Entity | `payment_entity` | always populated |
| AB | Type | `type` | |
| AC | Loan Out | `loan_out` | already `"YES"`/`"NO"`, write directly |
| AD | Qualify? | — | **leave blank** (human judgment, same as every other TX tab) |
| AE | Home Address | `home_address` | |
| AF | City | `city` | |
| AG | State | `state` | |
| AH | Zip | `zip` | |
| AI | Notes: | `notes` | write directly, already human-readable |
| AK | GROSS | — | **live formula — never write** |
| AL | FRINGES | — | **live formula — never write** |
| AM | TOTAL | — | **live formula — never write** |

Heads up on `title` (col I): Highland can supply it from the PTIP when present; Extreme Reach and Teams pull it from cast category codes on their own PTIP/PDF; CMS never supplies one at all (always blank) since its invoice format has no title/role field. This isn't a bug — just expected variance between the four formats.

## Reliability wiring

Apply the exact same pattern already built for the other 5 TX tabs this session:
- `postExtract`'s retry logic (network-failure-only, 2-3 attempts, backoff `[1000, 2000]`ms)
- `txFallbackRow` / `txFallbackReason` — push a synthetic placeholder row (filename or a generic label as `talent_name`, `total`=0, a note explaining the failure) whenever a submission ends up with zero rows for any reason (network failure, backend error, or an empty 200 response)

## What NOT to build

- No "Sorted PTIP" export for this tab (unlike GA's Teams-specific deliverable) — `ptip_excel_b64` is always `null`.
- No DTR matching, no TX-residency auto-detection, no proof-of-payment detection — all confirmed out of scope for this build and intentionally left blank for manual review.
- No Loan Out / Payroll Roster / GL Billing side-tabs — TX's workbook doesn't have these (unlike GA), so nothing analogous to `_loan_out_rows_from_talent` needs wiring here.
