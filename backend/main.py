"""
TPC Extraction Service
──────────────────────
Two extractors the browser wizard POSTs documents to:

  • /extract-invoices  — GPT-4o vision + normalization for freelance invoices
  • /extract-payroll   — fringe report parser for Wrapbook and CAPS PDFs
                         (auto-detects company; falls back to GPT-4o vision
                          for image-only Wrapbook fringe PDFs)

It never builds the workbook — the wizard assembles the final .xlsx client-side.

Endpoints
  GET  /health            → {"ok": true, "has_key": bool}
  POST /extract-invoices  → multipart: files[]=<pdf/png/jpg>, prodco_names="A, B"
                            returns {"invoices": [...], "issues": [...]}
  POST /extract-payroll   → multipart: files[]=<pdf>
                            returns {"rows": [...], "issues": [...]}

Environment variables
  OPENAI_API_KEY     (required for invoices + image-based fringe) your OpenAI key
  APP_SHARED_SECRET  (optional) if set, callers must send header X-App-Secret
  ALLOWED_ORIGINS    (optional) comma-separated CORS origins; default "*"
"""

import os
import re
import io
import json
import base64

import fitz  # PyMuPDF
import pdfplumber
from openai import OpenAI
from fastapi import FastAPI, UploadFile, File, Form, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware

from parsers.base import FRINGE_FIELDS
from parsers.caps.fringe import extract as caps_extract
from parsers.wrapbook.fringe import extract as wb_extract

# ── Config ──────────────────────────────────────────────────────────────────

OPENAI_API_KEY    = os.environ.get("OPENAI_API_KEY", "")
APP_SHARED_SECRET = os.environ.get("APP_SHARED_SECRET", "")
ALLOWED_ORIGINS   = [o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",") if o.strip()]

_PROMPT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "invoice_extraction_prompt.txt")

app = FastAPI(title="TPC Extraction Service")
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS or ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _client():
    if not OPENAI_API_KEY:
        raise HTTPException(500, "Server is missing OPENAI_API_KEY.")
    return OpenAI(api_key=OPENAI_API_KEY)


def _load_prompt(prodco_names):
    with open(_PROMPT_PATH, "r", encoding="utf-8") as f:
        template = f.read()
    return template.replace("{prodco_names}", ", ".join(prodco_names))


# ── PDF / image → page images (base64 PNG) ────────────────────────────────────

def _file_to_images_b64(filename, data, dpi_scale=2.0):
    """Render every page of a PDF (or a single image file) to base64 PNGs."""
    lower = filename.lower()
    if lower.endswith((".png", ".jpg", ".jpeg")):
        doc = fitz.open(stream=data, filetype="png" if lower.endswith(".png") else "jpg")
        data = doc.convert_to_pdf()
        doc.close()
    doc = fitz.open(stream=data, filetype="pdf")
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale))
        images.append(base64.b64encode(pix.tobytes("png")).decode())
    doc.close()
    return images


# ── GPT-4o vision call ────────────────────────────────────────────────────────

def _call_gpt(images_b64, system_prompt, client):
    content = [{"type": "text", "text": "Extract all invoices from these document pages."}]
    for img in images_b64:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{img}", "detail": "high"},
        })
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content},
        ],
        temperature=0,
        max_tokens=4096,
    )
    raw = resp.choices[0].message.content.strip()
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"\[.*\]", raw, re.DOTALL)
        if m:
            try:
                return json.loads(m.group())
            except Exception:
                return []
    return []


def _extract_from_file(filename, data, system_prompt, client):
    images = _file_to_images_b64(filename, data)
    MAX_BYTES = 45 * 1024 * 1024
    batches, cur, cur_size = [], [], 0
    for img in images:
        approx = len(img) * 3 // 4
        if cur and cur_size + approx > MAX_BYTES:
            batches.append(cur); cur, cur_size = [img], approx
        else:
            cur.append(img); cur_size += approx
    if cur:
        batches.append(cur)

    invoices = []
    for batch in batches:
        invoices.extend(_call_gpt(batch, system_prompt, client))
    return invoices


# ── Normalization (invoice extractor) ─────────────────────────────────────────

_STATE_ABBR = {
    "alabama":"AL","alaska":"AK","arizona":"AZ","arkansas":"AR","california":"CA",
    "colorado":"CO","connecticut":"CT","delaware":"DE","florida":"FL","georgia":"GA",
    "hawaii":"HI","idaho":"ID","illinois":"IL","indiana":"IN","iowa":"IA","kansas":"KS",
    "kentucky":"KY","louisiana":"LA","maine":"ME","maryland":"MD","massachusetts":"MA",
    "michigan":"MI","minnesota":"MN","mississippi":"MS","missouri":"MO","montana":"MT",
    "nebraska":"NE","nevada":"NV","new hampshire":"NH","new jersey":"NJ","new mexico":"NM",
    "new york":"NY","north carolina":"NC","north dakota":"ND","ohio":"OH","oklahoma":"OK",
    "oregon":"OR","pennsylvania":"PA","rhode island":"RI","south carolina":"SC",
    "south dakota":"SD","tennessee":"TN","texas":"TX","utah":"UT","vermont":"VT",
    "virginia":"VA","washington":"WA","west virginia":"WV","wisconsin":"WI",
    "wyoming":"WY","district of columbia":"DC",
}

_PYMT_MAP = {
    "check":"Check","cheque":"Check","p-card":"P-Card","pcard":"P-Card","p card":"P-Card",
    "purchasing card":"P-Card","credit card":"Credit Card","credit":"Credit Card",
    "visa":"Credit Card","mastercard":"Credit Card","amex":"Credit Card","cash":"Cash",
    "eft/wire":"EFT/WIRE","eft":"EFT/WIRE","wire":"EFT/WIRE","wire transfer":"EFT/WIRE",
    "ach":"EFT/WIRE","direct deposit":"EFT/WIRE","payroll company":"Payroll Company",
    "payroll":"Payroll Company","internal":"Internal","zero balance":"Zero Balance",
}

_CARD_ABBR = {
    "american express":"AMEX","amex":"AMEX","visa":"VISA",
    "mastercard":"MC","master card":"MC","mc":"MC","discover":"DISC",
}


def normalize_amount(val):
    if not val:
        return 0
    s = str(val).replace("$", "").replace(",", "").strip()
    try:
        return round(float(s), 2)
    except Exception:
        return 0


def clean_state(val):
    if not val:
        return ""
    s = str(val).strip()
    return _STATE_ABBR.get(s.lower(), s.upper()[:2])


def clean_zip(val):
    return str(val).strip()[:5] if val else ""


def normalize_pymt_method(val):
    return _PYMT_MAP.get(str(val).strip().lower(), "") if val else ""


def normalize_pymt_number(method, val):
    if not val:
        return ""
    s = str(val).strip()
    if method == "EFT/WIRE":
        return s if s.lower().startswith("on ") else "On " + s
    if method in ("Credit Card", "P-Card"):
        if "*" in s:
            return s.upper()
        if s.isdigit() and len(s) == 4:
            return s
        lower = s.lower()
        for name, abbr in _CARD_ABBR.items():
            if name in lower:
                digits = re.search(r"\d{4}", s)
                return f"{abbr}*{digits.group()}" if digits else abbr
        return s
    return s


def clean_name(val):
    if not val:
        return ""
    s = str(val).replace(".", "").strip()
    return " ".join(s.split()).title().rstrip(",").strip()


def clean_address(val):
    if not val:
        return ""
    s = str(val).strip()
    s = re.sub(r",?\s*#\s*\S+", "", s)
    s = re.sub(r",?\s*(Apt|Apartment|Suite|Ste|Unit|Room|Fl|Floor)(\s+\S+)?", "", s, flags=re.IGNORECASE)
    s = " ".join(s.replace(".", "").split()).title()
    directions = {"North":"N","South":"S","East":"E","West":"W"}
    words = s.split()
    if len(words) >= 2 and words[0][0].isdigit() and words[1] in directions:
        words[1] = directions[words[1]]
        s = " ".join(words)
    street_types = {
        "Avenue":"Ave","Boulevard":"Blvd","Circle":"Cir","Court":"Ct","Drive":"Dr",
        "Highway":"Hwy","Lane":"Ln","Place":"Pl","Road":"Rd","Street":"St",
        "Terrace":"Ter","Trail":"Trl",
    }
    for full, abbr in street_types.items():
        s = re.sub(r"\b" + full + r"\b", abbr, s)
    return s.rstrip(",").strip()


def normalize_invoice(inv):
    method = normalize_pymt_method(inv.get("pymt_method", ""))
    return {
        "po_number":      inv.get("po_number", ""),
        "vendor_name":    clean_name(inv.get("vendor_name", "")),
        "description":    clean_name(inv.get("description", "")),
        "address":        clean_address(inv.get("address", "")),
        "city":           clean_name(inv.get("city", "")),
        "state":          clean_state(inv.get("state", "")),
        "zip":            clean_zip(inv.get("zip", "")),
        "invoice_date":   str(inv.get("invoice_date", "")).strip(),
        "invoice_number": inv.get("invoice_number", ""),
        "invoice_amount": normalize_amount(inv.get("invoice_amount", "")),
        "pymt_method":    method,
        "pymt_number":    normalize_pymt_number(method, inv.get("pymt_number", "")),
        "notes":          inv.get("notes", ""),
    }


# ── Payroll company detection ─────────────────────────────────────────────────

def _detect_company(pdf_bytes: bytes) -> str:
    """Return 'caps' or 'wrapbook' based on text markers in the PDF.
    If no text is found (image-only PDF), defaults to 'wrapbook'."""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pg in pdf.pages[:15]:
                text = (pg.extract_text() or "").lower()
                if "fringe recap report" in text:
                    return "caps"
                if "fringe report" in text:
                    return "wrapbook"
    except Exception:
        pass
    return "wrapbook"


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return {"ok": True, "has_key": bool(OPENAI_API_KEY)}


@app.post("/extract-invoices")
async def extract_invoices(
    files: list[UploadFile] = File(...),
    prodco_names: str = Form(""),
    x_app_secret: str = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    client = _client()
    names = [n.strip() for n in prodco_names.split(",") if n.strip()]
    system_prompt = _load_prompt(names)

    invoices, issues = [], []
    for uf in files:
        data = await uf.read()
        try:
            raw = _extract_from_file(uf.filename, data, system_prompt, client)
        except Exception as e:
            issues.append(f"{uf.filename}: {e}")
            raw = []
        if not raw:
            invoices.append(normalize_invoice({
                "vendor_name": os.path.splitext(uf.filename)[0],
                "notes": "Invoice could not be read - please review manually",
            }))
            issues.append(f"{uf.filename}: no invoice data extracted")
            continue
        for inv in raw:
            invoices.append(normalize_invoice(inv))

    return {"invoices": invoices, "issues": issues}


async def _run_extract(files, x_app_secret):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    rows, issues = [], []
    for uf in files:
        data    = await uf.read()
        company = _detect_company(data)

        if company == "caps":
            extracted, errs = caps_extract(data)
        else:
            extracted, errs = wb_extract(data, openai_key=OPENAI_API_KEY)

        for r in extracted:
            r["sourceFile"] = uf.filename

        rows.extend(extracted)
        for e in errs:
            issues.append(f"{uf.filename}: {e}")

    return {"rows": rows, "issues": issues, "columns": FRINGE_FIELDS}


@app.post("/extract-fringe")
async def extract_fringe(
    files: list[UploadFile] = File(...),
    x_app_secret: str = Header(default=""),
):
    return await _run_extract(files, x_app_secret)


@app.post("/extract-payroll")
async def extract_payroll(
    files: list[UploadFile] = File(...),
    x_app_secret: str = Header(default=""),
):
    return await _run_extract(files, x_app_secret)
