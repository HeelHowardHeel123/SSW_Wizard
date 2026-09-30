# "Start over" file leak fix — needs a bundle rebuild

## What's new

Root-caused a real test mix-up: Steven had loaded files for TMS 032,
clicked "Start over" intending to switch to a fresh SONI 005 run, but
6 Crew Payroll invoice PDFs and a Production Report from TMS 032
silently survived into the SONI 005 run — invoice numbers with no
matching PDF anywhere in SONI 005's own folder, scrambling its Crew
Payroll totals.

## What I already built (source only, pushed to `dev` as `3d3a665`)

`source/Production Binder Wizard.dc.html`, `resetEverything()` (the
handler behind the "Start over" button): it only ever reset React
state — `uploads: {}` clears the UI's displayed upload list, but the
actual uploaded File objects live on a separate instance property,
`this.files`, which `resetEverything()` never touched. The smaller,
per-wizard `restart()` button already clears `this.files` correctly —
`resetEverything()` was simply missing the same line. Added
`this.files = {}` and `this._downloaded = false` to match.

## What I need from you

1. **Rebuild `backend/frontend/index.html`** from the current
   `source/Production Binder Wizard.dc.html` and send it back.
2. **Re-verify the two post-rebuild injection markers**:
   - `__OM_EMBEDDED_WORKBOOKS__` — 4 keys present
   - `__OM_EMBEDDED_MODULE_SOURCE__` — 1 key present

## Not asking you to change anything else

No new dc.html edits needed — the fix is already written and pushed.
This is purely: rebuild the compiled artifact from current source and
confirm the two markers.
