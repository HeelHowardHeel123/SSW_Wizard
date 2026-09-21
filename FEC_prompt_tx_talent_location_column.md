# TX Talent Payroll — wire up the new `location` field

## Confirmed live (SONI 005 Talent Test 003)

The backend push already landed — Title is populating correctly in a real run ("Talent Coordinator", "Extra Buyout, Principal", etc.). Location (column K) is still blank, exactly as expected, since this is the one piece still needed from you.

## What changed on the backend

CMS Productions rows now return a `location` field (a new addition to the shared talent row model — every payroll company uses the same row shape, but only CMS populates this one for now). It's derived from the daily work-backup pages in a CMS invoice PDF, which print a line like `WORKED WEEK ENDING 07/22/2025 AUSTIN LOCATION IN TEXAS`. When a person worked under more than one shoot location across their invoices, the value is a comma-joined list of the distinct ones (e.g. `"TX, CA"`), and — only in that multi-location case — the Notes field gets an extra breakdown line naming which invoice had which location (e.g. `"Location by invoice: Invoice 604123: TX, Invoice 604205: CA"`), so it's not just a merged list with no way to tell which is which.

`title` (already wired) works the same way — comma-joined unique values across invoices (e.g. `"Extra Buyout, Principal"`). No frontend change needed there, this is just a heads-up on the pattern in case it's useful context.

## What needs to change in `source/Production Binder Wizard.dc.html`

1. Add `location` → column **K** to `TX_TALENT_PAYROLL_MAP`:
   ```js
   ["location", "K", "text"],
   ```
2. Remove `"K"` from `TX_TALENT_PAYROLL_MANUAL` (it's currently in the always-blank reviewer-field list — `J`, `K`, `L`, `X`, `Z`, `AD`). It should become `["J", "L", "X", "Z", "AD"]`.

That's the whole change — `location` is just another mapped field like `title`, `wages`, etc. No new upload zone, no new endpoint param, nothing else touches this.

## Note for testing

Every payroll company except CMS will send `location: ""` (or the key may simply be absent) for now, so column K stays blank for Highland/ER/Teams rows — that's expected, not a bug. Only CMS populates it today.
