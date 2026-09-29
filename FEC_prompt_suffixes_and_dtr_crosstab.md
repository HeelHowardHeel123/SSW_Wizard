# Two more TX fixes — needs a bundle rebuild

## What's new

1. **Name suffixes and company entities weren't handled by the Last,
   First reorder.** "Robert D Burns Jr" became "Jr, Robert D Burns"
   (suffix treated as the surname), and "Afed Films Inc" (a vendor that
   ends up in the crew list on some real Production Reports) became
   "Inc, Afed Films".
2. **A DTR form for someone who's genuinely Talent was also creating a
   spurious blank row on Crew Payroll.** Confirmed on real NOVP 006
   data: several background "Extra Buyout" people had a correct Talent
   Payroll match *and* a second bare orphan row on Crew Payroll with no
   role, just because their DTR form (correctly) didn't match anyone on
   Crew Payroll's own roster — nothing was checking whether they'd
   already matched elsewhere.

## What I already built (source only, pushed to `dev` as `409362c` and `5e2ce12`)

`source/Production Binder Wizard.dc.html`:

1. `_properLastFirst()` now detects two cases before reordering:
   - Personal-name suffixes (Jr, Sr, II, III, IV) — kept with the actual
     surname instead of being treated as one ("Burns Jr, Robert D").
   - Company/entity words (Inc, LLC, Corp, Co, Ltd, LP, LLP, PLLC, PC) —
     the whole name is left in original order, just title-cased, since
     there's no surname to split on ("Afed Films Inc").
2. `matchDtrToRoster()` takes an optional shared `Set` (`matchedOut`)
   and records every name it successfully matches into it. Talent
   Payroll and the Crew IC/Talent IC loop now pass a shared
   `dtrMatchedElsewhere` Set declared once in `buildTexasBlob()`.
   **Crew Payroll — the only tab that turns unmatched DTR names into new
   rows — was moved to process LAST among the 4 TX tabs** (it used to run
   first) specifically so that Set is fully populated by the other 3
   tabs before Crew Payroll decides which DTR names are genuine orphans.
   A DTR name only becomes a new Crew Payroll row now if it matched
   nobody on any of the 4 tabs, not just Crew Payroll's own roster.

Side effect worth knowing about, not a bug: the run log will now show
Talent Payroll and Crew/Talent IC extraction steps *before* Crew
Payroll's own "Extracting payroll" step, since that whole block moved
later in the function. Purely cosmetic — same data lands in the same
cells either way.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` and send it back.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

No new dc.html edits needed — both fixes are already written and pushed.
This is purely: rebuild the compiled artifact from current source and
confirm the two markers.
