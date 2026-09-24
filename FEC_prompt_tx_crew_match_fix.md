# Call-sheet crew matching bug fix — needs a bundle rebuild

## What was wrong

First real-data test (SONIC 005: 92-person call sheet vs ~90-person Crew
Payroll roster) showed every single match failing — every existing payroll
row read `AQ=NO` even for people obviously also on the call sheet (e.g.
"Elins, Philip" on the roster, "Phil Elins" added as a brand-new duplicate
row two rows later instead of being matched to it).

Root cause: `POST /match-names`'s Claude call was hardcoded to
`max_tokens=4096`. Its response has one JSON key per input name regardless
of match quality, and a two-list match this size (92 + ~90 names)
plausibly exceeded that, truncating the response and silently falling back
to an empty mapping with no error — the same failure mode already found
once this session on crew-roster extraction's own `max_tokens`.

## What I already fixed (source only, pushed to `dev` as `9d03a46`)

- Backend (`backend/main.py`): `/match-names` raised to `max_tokens=16000`;
  now returns an explicit `error` field when it had to fall back to `{}`
  after a parse failure, instead of an indistinguishable "nobody matched."
- Frontend (`source/Production Binder Wizard.dc.html`):
  `matchCallSheetToRoster()` now surfaces a non-ok response, a network
  failure, or a backend `error` as a visible run issue.
- Also fixed a separate, purely cosmetic bug found while reading this run's
  log: the generic per-crew detail logger in `postExtract()` (~line 1148)
  assumed the older `{name, positions[], dates[]}` shape from the
  pre-existing AP-description call-sheet feature, so it printed "(no
  position)" for every single entry from the new `{name, title}` shape —
  even though titles were actually extracted correctly (confirmed by
  reading the real output workbook's Crew Position column directly). It
  now shows whichever field an entry actually has.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` (has this fix) and send it
   back, same pattern as prior rebuild handoffs.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present
3. The backend fix (`max_tokens`) is already live on Railway independent of
   this rebuild — only the two frontend-side error-surfacing fixes need the
   new bundle.

## Not asking you to change anything else

No new dc.html edits needed — the fix is already written and pushed. This
is purely: rebuild the compiled artifact from current source and confirm
the two markers.
