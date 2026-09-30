# TX Petty Cash/ProdCC — `id_number` is gone, replaced by `env_number` + `po_number`

## Backend status: done and pushed to `dev` (commit `31237ad`)

This supersedes the `id_number` field from the original Petty Cash/ProdCC prompt (`FEC_prompt_tx_petty_cash_prodcc.md`). If you haven't wired that field yet, just build straight to this version. If you already have `id_number → D` (Petty Cash) / `id_number → C` (ProdCC) wired, this replaces it.

## Why

Confirmed a real bug on live SONI 005 data: `PO 21917 - Petty Cash.pdf` is a genuine Purchase Order used internally to authorize a petty cash advance (its "Company Name" field literally reads "Petty Cash") -- it has a real `PO # 21917` printed on it and **no envelope number at all**. The old design blindly routed whatever single `id_number` value came back into column D (Env#) for any file uploaded through the Petty Cash zone -- so a PO number ended up sitting in the Env# column, which is wrong regardless of which zone the file was uploaded through.

Env # and PO # are two independent, mutually exclusive fields on a real document -- it only ever has one or the other, never both. The backend now reads them separately and always tells you which one it found (or neither), rather than you inferring which column a single `id_number` belongs in based on which zone the upload came from.

## What changed in the row shape

Both `/extract-tx-petty-cash` and `/extract-tx-prodcc` now return two fields instead of `id_number`:

```json
{
  "env_number": "1 of 1",
  "po_number": "",
  ...
}
```

A row has at most one of these non-empty -- never both. Both are plain strings; empty string means that kind of number wasn't found on this document.

## What needs to change in `source/Production Binder Wizard.dc.html`

Add both fields to the shared map (`TX_PETTY_COMMON_MAP`, or wherever `TX_PETTY_CASH_MAP`/`TX_PRODCC_MAP` draw their common fields from) -- **the same two lines for BOTH zones**, not one-per-zone like before:

```js
["env_number", "D", "text"],
["po_number",  "C", "text"],
```

Then **delete** whatever zone-conditional logic currently does something like:

```js
// DELETE this -- no longer needed
put(map === this.TX_PETTY_CASH_MAP ? "C" : "D", "text", null);
```

That line existed to blank out "the other zone's id column" under the old single-`id_number` design. It's not needed anymore: since `env_number` and `po_number` now come back independently and correctly reflect what's actually on each document, just write both fields plainly on every row (empty string writes as a cleared cell, same as any other empty field) -- no zone-based branching required at all. A Petty Cash row will almost always end up with `env_number` filled and `po_number` empty, and a ProdCC row the reverse, but write them exactly as the data says rather than assuming that by zone.

## Also included in this push (from the earlier prompt, unchanged)

The backend also broadened envelope-number recognition to catch `PAGE 1 OF 1` / `PAGE 2 OF 3` style cover-sheet phrasing (not just `Envelope No:`) -- no frontend change needed for that part, it's purely how `env_number` gets populated on the backend.

## What NOT to build

- No new logic to decide which column a number belongs in -- that's fully decided by which field (`env_number` vs `po_number`) it comes back in.
- No change to `line_number`, the Various-value behavior, amount/formula handling, or the Petty Cash/ProdCC block layout -- all unchanged from the prior prompts.
