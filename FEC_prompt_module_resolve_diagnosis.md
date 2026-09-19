# Generate Workbook is failing in your Design preview — need your diagnosis

## The error

Testing a TX Talent Payroll run (AZNE 008, Extreme Reach — 2 invoice PDFs + 1 PTIP file) from inside the Claude Design window, Generate Workbook fails immediately:

```
Result: FAILED — Failed to resolve module specifier './workbook-engine.js'

TypeError: Failed to resolve module specifier './workbook-engine.js'
  at Component.loadEngine (eval at evalDcLogic (blob:https://c55d6c2c-f048-44f3-94ef-96a793e74b0c.claudeusercontent.com/...):1:1), <anonymous>:535:33)
  at Component.buildTexasBlob (...):2486:28)
  at async Component.buildBlob (...):6028:27)
  at async Component.generate (...):6579:17)
```

Full run log: `AZNE 008 Talent Test 001 - run log 2026-09-18-23-48.txt`, in the Debug folder.

## What I've already ruled out on my end

- `loadEngine()` and the whole `__OM_EMBEDDED_MODULES__` fallback mechanism are **not new** — I traced it with `git log -S` and it goes back to `db5a993` ("GA Active, Highland Talent, Loan Out Withholding..."), long before this session's TX work. Nothing in the Wizard023/024 dc.html changes touched this function.
- The compiled `backend/frontend/index.html` bundle we pushed has the injection wired correctly — I decompressed it and confirmed `window.__OM_EMBEDDED_MODULE_SOURCE__` is set with the exact key `loadEngine()` looks for (`"./workbook-engine.js"`), and the wrapper script that turns it into `__OM_EMBEDDED_MODULES__` blob URLs is intact. That part isn't the problem — and it isn't even in play here anyway, since the stack trace shows this running from a `claudeusercontent.com` blob (`evalDcLogic`), i.e. straight from source inside your Design session, not from that compiled bundle.

So this looks like a limitation of running `Generate` directly inside the live Design preview — `loadEngine()` falls back to a bare `import('./workbook-engine.js')` when `__OM_EMBEDDED_MODULES__` isn't present, and that relative specifier apparently can't resolve in that hosting context.

## What I need you to figure out

You're connected to the actual environment the user is testing in — I'm not, so I can't reproduce or inspect it directly. Please diagnose:

1. **Has `Generate Workbook` ever worked when run directly inside your Design preview**, or has every real test this session (TMS 032/033, MCD 062, ABBVIE 022, NIS 006, SONI 005, etc.) actually been run against something else — a deployed copy, a different preview mode, anything other than clicking Generate live in this same window? If there's a different link/mode that does work, that's the simplest fix: point the user at it.
2. **If Generate has worked from directly inside Design before**, what's different now? Does your environment normally provide `workbook-engine.js` as a resolvable sibling module some other way (an import map, a dev-server route, something Design itself injects) that might have stopped working, or that this dc.html round accidentally bypassed?
3. Whichever it is, tell me what (if anything) needs to change in `source/Production Binder Wizard.dc.html` — e.g. should `loadEngine()` gain a further fallback for this specific hosting context, similar to how it already falls back from `__OM_EMBEDDED_MODULES__` to a bare relative import?

I don't want to guess at a fix without knowing which of these is actually true — over to you.
