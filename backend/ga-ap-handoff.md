# GA AP Extraction — Backend/Frontend Handoff

`POST /extract-ga-ap` returns `{ rows, issues, files }`.

Each object in `rows` has the following keys.  These are the canonical field
names the frontend should read to populate the AP tab.

## Row Schema

| Key | Type | Notes |
|-----|------|-------|
| `check_number` | string | Payment or check number; empty string if not found |
| `invoice_number` | string | Invoice or reference number as shown |
| `invoice_date` | string | ISO `YYYY-MM-DD`; empty string if not found |
| `po_number` | string | PO number visible on any document in the PDF |
| `vendor_name` | string | Business or individual who issued the invoice |
| `payee_type` | `"person"` \| `"company"` | Explicit — never infer from last_name |
| `ff1` | string | **v001: always empty string.** User fills in manually. |
| `ff2` | string | **v001: always empty string.** User fills in manually. |
| `je_number` | string | Journal entry number; empty string if not found |
| `distribution_description` | string | Short description of what was purchased |
| `last_name` | string | Last name; populated only when `payee_type === "person"` |
| `first_name` | string | First name; same rule as `last_name` |
| `home_state` | string | 2-letter state code for individual's home state |
| `episode` | string | Episode number; blank for commercials |
| `nights` | number \| null | Integer hotel nights; `null` for non-lodging rows |
| `address` | string | Street address only (no city/state/zip) |
| `city` | string | |
| `state` | string | 2-letter state code |
| `zip` | string | 5-digit zip |
| `amount` | number | Plain float, no `$` or commas |
| `non_qualified` | number | **v001: always 0.** User fills in manually. |
| `payment_method` | string | `Check` `ACH` `Wire` `Credit Card` `P-Card` — or empty string |
| `proof_of_payment` | `"Yes"` \| `"No"` | |
| `payment_entity` | string | Which payer entity paid — matched against `payer_entities` input |
| `received_invoice` | `"Yes"` \| `"No"` | |
| `loan_out` | `"Yes"` \| `"No"` | |
| `loan_out_individual_name` | string | Natural person behind the loan-out corp; empty if `loan_out === "No"` |
| `sales_tax_on_invoice` | `"Yes"` \| `"No"` | |
| `active_sales_tax_account` | `"Yes"` \| `"No"` | |
| `withholding` | string | GA loan-out withholding amount as plain number string; empty if N/A |
| `w9` | `"Yes"` \| `"No"` | |
| `business_license` | `"Yes"` \| `"No"` | |
| `aicp_code` | null | **v001: always null.** User fills in manually. |
| `qualified` | string | **v001: always empty string.** User fills in manually. |
| `notes` | string | Reviewer notes; empty string in normal cases |
| `website_address` | string | Vendor website if printed on invoice |
| `labor` | boolean | `true` = labor/personal services; `false` = goods/equipment/facilities. **Not written to AP tab** — used downstream to route rows to Payroll Roster and other tabs |

## Column Map (AP tab)

For reference only — the frontend owns workbook writing.

| Column | Field |
|--------|-------|
| A | seq (generated 1…N) |
| B | check_number |
| C | invoice_number |
| D | invoice_date |
| E | po_number |
| F | vendor_name |
| G | ff1 |
| H | ff2 |
| I | "AP" (constant) |
| J | je_number |
| K | distribution_description |
| L | last_name |
| M | first_name |
| N | home_state |
| O | episode |
| P | nights |
| Q | address |
| R | city |
| S | state |
| T | zip |
| U | "US" (constant) |
| V | amount |
| W | non_qualified |
| X | =V{r}-W{r} (formula) |
| Y | payment_method |
| Z | proof_of_payment |
| AA | payment_entity |
| AB | received_invoice |
| AC | loan_out |
| AD | loan_out_individual_name |
| AE | sales_tax_on_invoice |
| AF | active_sales_tax_account |
| AG | withholding |
| AH | w9 |
| AI | business_license |
| AJ | aicp_code |
| AK | AICP Category Description — `=IFERROR(VLOOKUP(AJ{r},Legend!$E$2:$F$26,2,0),"")` (formula) |
| AL | qualified |
| AM | notes |
| AN | website_address |
| AO–AR | TPC review columns — always blank |

## Endpoint Inputs

```
POST /extract-ga-ap
Content-Type: multipart/form-data

files[]          one or more PDF uploads
prodco_name      string (scalar fallback if payer_entities not provided)
prodco_address   string
agency_name      string
work_state       string (default "GA")
payer_entities   JSON array of { role, name, address }
                 roles: "client" | "agency" | "prodco" | "sub_prodco"
```

`payer_entities` is preferred over the scalar fields.  The LLM matches each
invoice's Bill-To against this list to populate `payment_entity`.

## Response Envelope

```json
{
  "rows": [ { ...row fields above... }, ... ],
  "issues": [ "filename.pdf: message", ... ],
  "files": [
    { "file": "filename.pdf", "rows": 3, "issues": [] },
    ...
  ]
}
```

`issues` is a flat list of all errors across all files.  `files` gives
per-file row counts and per-file error lists for the progress UI.
