# Production Binder Wizard — project notes

## Templates (Wizard 04)
- The workbook templates the wizards populate live at `assets/workbooks/*.xlsx`
  (`illinois-local.xlsx` is the only one wired to an active wizard today).
- That bundled file is the **default every browser uses out of the box** — a
  first-time user does NOT need to upload anything.
- Wizard 04 lets a user replace a template **in their own browser** (stored in
  IndexedDB, key = workbook id). An override persists across sessions and is
  reused on every run until they Replace or Revert it. It is per-browser (there
  is no shared server), so it does not propagate to other machines.
- When a template change should apply to EVERYONE by default, the bundled file
  in `assets/workbooks/` must be updated (the code's column mapping must match
  it). Current default = **Template IL - Local - State Submission Workbook 2.2**
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

## TODO
- [ ] Write new-user instructions (how the tool works, that the correct template
      ships as the default so no setup is needed, and how/when to use Wizard 04
      to update a template).
