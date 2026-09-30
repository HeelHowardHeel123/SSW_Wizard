# TX Petty Cash/ProdCC — map the new `line_number` field to column E

## Backend status: done and pushed to `dev` (commit `719c85b`)

Small, targeted change to the row shape you're already writing from `/extract-tx-petty-cash` and `/extract-tx-prodcc`.

## What changed

Every row now carries a new `line_number` field, alongside the fields you already map. It's always one of three values: `"1"`, `"Various"`, or `""`.

Also, the existing fields `vendor`, `receipt_date`, `description`, `address`, `city`, `state`, `zip` can now come back as the literal string `"Various"` (not just real values or empty strings like before). No shape change there — just a new possible value, written the same way you already write any other text value.

## The rule, in case it's useful context

- A document with exactly one receipt behind it (whether or not it also had a stated cover/PO total) → `line_number: "1"`, and `vendor`/`receipt_date`/`description`/`address`/`city`/`state`/`zip` are that one receipt's real values.
- A document with more than one receipt → `line_number: "Various"`, and all seven of those same fields are also `"Various"`.
- Nothing at all found (rare, unreadable file) → `line_number: ""`, same as everything else on that fallback row.

You don't need to do anything with this rule — just write `line_number` wherever it lands and keep writing the other seven fields exactly as you already do (they're still plain text, "Various" included).

## What needs to change in `source/Production Binder Wizard.dc.html`

One line in each of `TX_PETTY_CASH_MAP` and `TX_PRODCC_MAP` (or wherever you consolidated them, e.g. `TX_PETTY_COMMON_MAP`):

```js
["line_number", "E", "text"],
```

And remove `"E"` from wherever it's currently being force-cleared/never-touched (this was previously intentional — the field didn't exist yet). Every other column mapping and the H/I/N/T/U handling from the original prompt is unchanged.

## What NOT to build

- No new logic needed for the "Various" values themselves — they're just text, written the same way as any other string field.
- No change to how `amount` is written (still a plain number or a `=...` formula string, same as before).
- No change to the Petty Cash / ProdCC block layout, blank separator row, or Item # counter.
