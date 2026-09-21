# TX Petty Cash + ProdCC — wire up two new upload zones

## Backend status: done and pushed to `dev` (commit `446db18`)

Two new endpoints, already live once you're on `dev`:

- `POST /extract-tx-petty-cash`
- `POST /extract-tx-prodcc`

Both take the same request shape: `files` (multiple PDFs), `prodco_name` (the production company name, same source you already use elsewhere for Payment Entity). No `work_state` param — these two endpoints don't need one.

Both return the same shape every other extractor uses:

```json
{
  "rows":  [ {...}, {...} ],
  "issues": ["optional strings"],
  "files":  [ {"filename": "...", "company": "...", "rows": 1, "issues": []} ]
}
```

## The big architectural difference from GA/IL Petty Cash

GA and IL's existing Petty Cash/ProdCC tabs are **one row per receipt/line-item**. TX is deliberately different: **one PDF = one row = one total.** Every uploaded file — no matter how many receipts are inside it — produces exactly one row. There's no cross-file grouping either: two files like `"Maxfield, Terry - Petty Cash (1 of 2).pdf"` and `"(2 of 2).pdf"` are NOT merged into one row — each is processed independently and can legitimately end up as two separate envelopes/rows.

## Row fields (identical shape from both endpoints)

```json
{
  "name":           "Bonnie Cook",
  "id_number":      "1 of 1",
  "pymt_method":    "Petty Cash",
  "payment_entity": "Autozone Productions",
  "amount":         144.68,
  "vendor":         "",
  "receipt_date":   "",
  "description":    "",
  "address":        "",
  "city":           "",
  "state":          "",
  "zip":            "",
  "notes":          "",
  "sourceFile":     "Cook, Bonnie - Petty Cash.pdf"
}
```

- `pymt_method` is always the literal string `"Petty Cash"` from `/extract-tx-petty-cash`, or `"CC Reimb"` from `/extract-tx-prodcc`. Nothing else ever comes through in that field.
- `id_number` is whatever envelope number or PO number was found (falls back to a `"(N of M)"` pattern in the filename, or empty string). Same field name from both endpoints — see column mapping below for where each one lands.

### `amount` can be a live Excel formula — this is the part that needs care

When a document had no readable receipts at all, or exactly one stated total (a Petty Cash cover sheet's envelope total, or a ProdCC PO's Sub-Total), `amount` is a plain number, e.g. `144.68`. Write it as a normal value.

When a document had **no stated total but more than one receipt**, `amount` comes back as a **string starting with `=`**, e.g.:

```
"=36.69+45.89+25.03+8.99+25.03+32.44+98.99+11.37+44.88+5.77-13.73"
```

This needs to be written into the cell as a **live formula** (same as you already do for any other formula cell in these workbooks), not as text and not pre-evaluated. Check with `typeof row.amount === "string" && row.amount.startsWith("=")` before deciding value-write vs formula-write.

In that same case, `notes` carries a matching "show your work" breakdown, e.g.:

```
"Jalisco's Mexican Restaurant $36.69, Trader Joe's $45.89, ... Central Market -$13.73"
```

so a reviewer can tell which number in the formula belongs to which receipt without opening the source PDF.

### The single-receipt case

When a document had no stated total and exactly one receipt, `vendor`, `receipt_date`, `description`, `address`, `city`, `state`, `zip` are all filled in directly from that one receipt (no ambiguity, no formula, no notes needed). Otherwise all of those stay empty strings.

## Column mapping — TX Petty Cash tab

Existing sheet in `texas.xlsx` (no changes needed to the workbook itself), header row 6, data starting row 7:

```js
const TX_PETTY_CASH_MAP = [
  ["id_number",      "D",  "text"],   // Env #
  ["receipt_date",   "F",  "text"],
  ["amount",         "G",  "formula_or_value"],  // see note above
  ["name",           "J",  "text"],
  ["pymt_method",    "K",  "text"],
  ["payment_entity", "L",  "text"],
  ["vendor",         "M",  "text"],
  ["description",    "O",  "text"],
  ["address",        "P",  "text"],
  ["city",           "Q",  "text"],
  ["state",          "R",  "text"],
  ["zip",            "S",  "text"],
  ["notes",          "V",  "text"],
];
```

Column **C (PO #)** stays blank on every Petty Cash row.

## Column mapping — TX ProdCC tab

Same physical sheet, same header row, just a different id column:

```js
const TX_PRODCC_MAP = [
  ["id_number",      "C",  "text"],   // PO #
  ["receipt_date",   "F",  "text"],
  ["amount",         "G",  "formula_or_value"],
  ["name",           "J",  "text"],
  ["pymt_method",    "K",  "text"],
  ["payment_entity", "L",  "text"],
  ["vendor",         "M",  "text"],
  ["description",    "O",  "text"],
  ["address",        "P",  "text"],
  ["city",           "Q",  "text"],
  ["state",          "R",  "text"],
  ["zip",            "S",  "text"],
  ["notes",          "V",  "text"],
];
```

Column **D (Env #)** stays blank on every ProdCC row.

## H and I need a fixed value/formula on every row you write

Checked directly against the current `texas.xlsx` — columns H and I are blank in the template rows (no pre-existing formula), so these two DO need to be written on every row you populate (both Petty Cash and ProdCC):

- **H (Total Ineligible)** — always write `0`. It's the reviewer's job to determine what's ineligible later; automation never guesses at this.
- **I (Qualified Total)** — write it as a live formula referencing that row's own G and H cells, e.g. `=G7-H7` (adjust the row number per row). Since H is always 0 this evaluates to the same as G today, but wiring it as a formula (not a copy of the value) keeps it correct if a reviewer later changes H by hand.

## Columns you should NEVER write on either tab

- **E (Line #)** — not extracted at all, leave whatever the template row already has (blank/manual).
- **N (Type)** — always blank.
- **T (Contact#)** — never extracted.
- **U (Qualify?)** — manual reviewer field, same convention as every other tab's Qualify?/Valid DTR/TX Resident columns.

## Row layout: Petty Cash block, one blank row, then ProdCC block

Write every Petty Cash row first (starting at row 7), then leave **exactly one blank row**, then write every ProdCC row after that. `Item #` (column B) is one simple sequential counter running across the WHOLE combined block (Petty Cash rows, then the ProdCC rows after the blank row) — it does not reset at the blank row.

The sheet has 100 pre-built rows, 7–106 (`B106` reads `100`), then row 107 is a spare blank row, then row 108 is `Subtotal: =SUBTOTAL(9,G7:G107)` and row 111 is `SUBMITTED TOTAL =SUM(G7:G107)`. If Petty Cash rows + 1 blank + ProdCC rows exceeds the 100 pre-built rows, use whatever row-insertion mechanism you already use elsewhere in this workbook for a tab that grows past its pre-built rows (Talent Payroll already does this today) so those formulas and the reconciliation/summary blocks below still land in the right place and still cover every row you added.

## Two new upload zones

Create two new folders/zones, each accepting multiple PDFs:

- **Petty Cash** → calls `/extract-tx-petty-cash` with `files` + `prodco_name`
- **Prod CC** → calls `/extract-tx-prodcc` with `files` + `prodco_name`

Both write into the same Petty Cash tab, per the layout above.

## What NOT to build

- No reconciliation/matching logic between Petty Cash and ProdCC files — they're two independent sets of rows.
- No cross-file envelope grouping — every PDF produces exactly one row, full stop.
- Nothing new needed in `texas.xlsx` itself — the tab, its 100 pre-built rows, and its footer/reconciliation/summary blocks already exist and are unchanged.
