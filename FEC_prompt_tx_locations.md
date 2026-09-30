# TX Locations — new upload zone + tab wiring

## Backend status: done and pushed to `dev` (commit `691f3fc`)

New endpoint: `POST /extract-tx-locations`. Takes `files` (one or more call sheet PDFs) — no `prodco_name`, no `work_state`, this endpoint doesn't need either.

## What changed in `texas.xlsx`

The Locations tab's old pre-built rows (`DAY 1 of 4` through `DAY 4 of 4`, rows 5–8) are gone — the user deleted them from the master template because a hardcoded day count doesn't generalize across productions with a different number of shoot days. I re-synced the embedded workbook (commit `a3c5376`). **The tab is now a genuine blank slate**: header row 4, and nothing at all below it in columns A–J. No pre-built rows, no footer formulas, nothing to preserve — every row you write needs to be inserted fresh, starting at row 5, using whatever row-insertion mechanism you already use elsewhere in this workbook for a tab that grows (Talent Payroll/Petty Cash already do this).

Column L (rows 4–21) is a separate, unrelated reference list — the Rural Uplift county-exclusion list. Don't touch it; it's informational only, not something the extraction writes to.

Confirmed header row 4: `B=Location Name`, `C=Street Number`, `D=Street Name`, `E=City`, `F=State`, `G=Zip Code`, `H=Production Date/Range`, `I=Notes (Basecamp, Non-TX)`, `J=TX County`, and `A=Shoot Day`.

## Row shape returned by the endpoint

```json
{
  "rows": [
    {
      "shoot_day":     "DAY 1 of 2",
      "location_name": "Auto Zone",
      "street_number": "225",
      "street_name":   "Whitestone Blvd",
      "city":          "Cedar Park",
      "state":         "TX",
      "zip":           "78613",
      "date":          "09/03/2025",
      "notes":         "",
      "county":        "Williamson",
      "sourceFile":    "Call Sheet.pdf"
    }
  ],
  "issues": ["optional strings"],
  "files":  [{"filename": "...", "company": "2 day(s)", "rows": 3, "issues": []}]
}
```

**One important shape difference from every other TX extractor you've wired so far**: one uploaded call sheet PDF does NOT become one row. A single call sheet PDF is usually a multi-day document (one page per shoot day), and a single day can have more than one location. So one file can produce anywhere from 1 to a dozen+ rows. Don't apply the "one PDF = one row" placeholder-on-failure pattern from Petty Cash/ProdCC here — that doesn't apply to this tab.

## Column mapping

```js
const TX_LOCATIONS_MAP = [
  ["shoot_day",     "A", "text"],
  ["location_name", "B", "text"],
  ["street_number", "C", "text"],
  ["street_name",   "D", "text"],
  ["city",          "E", "text"],
  ["state",         "F", "text"],
  ["zip",           "G", "text"],
  ["date",          "H", "text"],
  ["notes",         "I", "text"],
  ["county",        "J", "text"],
];
```

All plain text values, no formulas, no currency, nothing special — every field just writes as-is (empty string clears the cell, same convention as everywhere else).

## Some fields will legitimately be empty on a given row — that's expected, not a bug

- `street_number`/`street_name`/`city`/`zip` can all be blank together when the location has no real postal address — in that case `street_name` instead holds a raw GPS-coordinate string (e.g. `"Tahitian Road / 30°06'02.3\"N 97°17'31.5\"W"`), or (if there's neither an address nor coordinates) all four stay blank and `notes` carries whatever descriptive text was available instead (e.g. `"Downtown driving / Park at basecamp"`).
- `county` is blank whenever `zip` is blank (nothing to look up against) or the ZIP isn't in our TX reference table.
- `notes` can also carry a non-shoot-day flag (e.g. `"Tech Scout D1"`) when the row came from a prep/scout day rather than an actual shoot day — that's intentional; the row still gets written like any other, just flagged so a reviewer knows it wasn't a real shoot day.

## New upload zone

Create a "Call Sheet" (or similar) upload zone accepting multiple PDFs, posting to `/extract-tx-locations` with just `files`. Writes into the Locations tab per the mapping above, inserting rows starting at row 5.

## What NOT to build

- No per-file placeholder/fallback row logic (see above — this isn't a one-PDF-one-row tab).
- No formula columns, no currency formatting.
- Nothing touching column L (the Rural Uplift reference list).
