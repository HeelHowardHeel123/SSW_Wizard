# TPC Invoice Extraction Service — deploy guide

A tiny HTTP service that the Production Binder Wizard calls to read invoice
PDFs with GPT-4o vision and return clean JSON rows. The wizard builds the
Excel workbook itself — this service only extracts.

```
backend/
  main.py                         FastAPI app (extraction + normalization)
  requirements.txt                Python dependencies
  invoice_extraction_prompt.txt   the GPT-4o system prompt (REPLACE with yours)
```

---

## What you need before you start

- An **OpenAI API key** (the same one from your `OpenAI Key.txt`).
- A **GitHub account** (free).
- A **Railway account** (recommended) or Render account.
- Your real **`invoice_extraction_prompt.txt`** — drop it in this folder,
  replacing the placeholder. Keep the `{prodco_names}` token; the service
  fills it in per request.

---

## Step by step (Railway — recommended)

**1. Put this `backend/` folder in a GitHub repo.**
   - Create a new repo on github.com (e.g. `tpc-invoice-service`), Private.
   - Upload the four files above to the repo root (drag-and-drop in the
     GitHub web UI works, or `git push` if you use git).

**2. Create the service on Railway.**
   - Go to railway.app → log in with GitHub.
   - **New Project → Deploy from GitHub repo →** pick your repo.
   - Railway detects Python and installs `requirements.txt` automatically.

**3. Tell Railway how to start it.**
   - Project → **Settings → Deploy → Start Command**, set:
     ```
     uvicorn main:app --host 0.0.0.0 --port $PORT
     ```

**4. Add your secrets.**
   - Project → **Variables → New Variable**:
     - `OPENAI_API_KEY` = your OpenAI key
     - `APP_SHARED_SECRET` = any random string you invent (e.g. a long password).
       This stops strangers from running up your OpenAI bill. You'll paste the
       same value into the wizard.
     - `ALLOWED_ORIGINS` = `*` for now (we can lock this to the wizard's URL later)

**5. Expose a public URL.**
   - Project → **Settings → Networking → Generate Domain**.
   - You'll get something like `https://tpc-invoice-service.up.railway.app`.

**6. Test it.**
   - Open `https://<your-domain>/health` in a browser. You should see
     `{"ok": true, "has_key": true}`. If `has_key` is false, the key
     variable didn't save — recheck step 4.

**7. Send me two things** and I'll wire the wizard to it:
   - the public URL (e.g. `https://tpc-invoice-service.up.railway.app`)
   - the `APP_SHARED_SECRET` value you chose

That's it. Railway doesn't sleep, so there's no cold-start delay.

---

## Alternative: Render (free, but sleeps)

Same as above, with these differences:
- render.com → **New → Web Service →** connect the repo.
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
- Add the same environment variables under **Environment**.
- Free instances sleep after ~15 min idle → the first request after a lull
  takes ~30–50s. Upgrade to the $7/mo instance to remove that.

---

## Run it locally first (optional sanity check)

```bash
cd backend
pip install -r requirements.txt
export OPENAI_API_KEY=sk-...        # Windows: set OPENAI_API_KEY=sk-...
uvicorn main:app --reload
```
Then open http://127.0.0.1:8000/health

---

## Notes

- **Cost:** hosting is ~$5/mo on Railway (free on Render with the sleep caveat).
  The only other cost is your normal OpenAI usage per extraction.
- **Security:** the `APP_SHARED_SECRET` is a simple shared password the wizard
  sends on every call. It's enough to keep a public URL from being abused for
  an internal tool. If you outgrow it we can add real auth later.
- **The prompt file is the quality lever.** The placeholder works, but your
  tuned `invoice_extraction_prompt.txt` is what makes extraction accurate —
  swap it in.
