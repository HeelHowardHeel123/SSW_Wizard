"""
TPC Extraction Service
──────────────────────
Two extractors the browser wizard POSTs documents to:

  • /extract-invoices          — GPT-4o vision + normalization for freelance invoices
  • /extract-payroll           — fringe report parser for Wrapbook and CAPS PDFs
                                 (auto-detects company; falls back to GPT-4o vision
                                  for image-only Wrapbook fringe PDFs)
  • /extract-billings          — GPT-4o vision for Agency, ProdCo, and Sub-ProdCo
                                 billing invoices; writes to the Billings tab
  • /extract-agency-subvendors — GPT-4o vision for sub-vendor invoices billed to the
                                 ad agency; writes to the Agency Sub-Vendors tab

It never builds the workbook — the wizard assembles the final .xlsx client-side.

Endpoints
  GET  /health                    → {"ok": true, "has_key": bool}
  POST /extract-invoices          → multipart: files[]=<pdf/png/jpg>, prodco_names="A, B"
                                    returns {"invoices": [...], "issues": [...]}
  POST /extract-payroll           → multipart: files[]=<pdf>
                                    returns {"rows": [...], "issues": [...]}
  POST /extract-billings          → multipart: files[]=<pdf>, vendor_type, vendor_name,
                                    vendor_address, vendor_city, vendor_state, vendor_zip,
                                    prodco_names, work_state
                                    returns {"rows": [...], "issues": [...]}
  POST /extract-agency-subvendors → multipart: files[]=<pdf>, agency_name, agency_address
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
from pydantic import BaseModel

from parsers.base import FRINGE_FIELDS
from parsers.wrapbook.fringe_001 import enrich_from_register
from parsers.ai_fringe import extract_unknown, make_exec_parser
from parsers import registry
from notify import send_parser_alert, send_run_summary, ALERT_EMAIL
from talent_extractor import extract_talent, extract_teams_talent

# ── Config ──────────────────────────────────────────────────────────────────

OPENAI_API_KEY    = os.environ.get("OPENAI_API_KEY", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
APP_SHARED_SECRET = os.environ.get("APP_SHARED_SECRET", "")
ALLOWED_ORIGINS   = [o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "*").split(",") if o.strip()]

_PROMPT_PATH                    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "invoice_extraction_prompt.txt")
_CREW_FREELANCE_PROMPT_PATH     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "crew_freelance_prompt.txt")
_TALENT_FREELANCE_PROMPT_PATH   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "talent_freelance_prompt.txt")
_HOURS_LETTER_PROMPT_PATH       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hours_letter_prompt.txt")
_BILLING_PROMPT_PATH            = os.path.join(os.path.dirname(os.path.abspath(__file__)), "billing_extraction_prompt.txt")
_AGENCY_SUBVENDORS_PROMPT_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "agency_subvendors_extraction_prompt.txt")
_AGENCY_HOURS_PROMPT_PATH       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "agency_hours_prompt.txt")
_RESIDENCY_DOCS_PROMPT_PATH     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "residency_docs_prompt.txt")

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


def _load_crew_freelance_prompt(prodco_names):
    with open(_CREW_FREELANCE_PROMPT_PATH, "r", encoding="utf-8") as f:
        template = f.read()
    label = ", ".join(prodco_names) if prodco_names else "the production company"
    return template.replace("{prodco_names}", label)


def _load_talent_freelance_prompt(prodco_names):
    with open(_TALENT_FREELANCE_PROMPT_PATH, "r", encoding="utf-8") as f:
        template = f.read()
    label = ", ".join(prodco_names) if prodco_names else "the production company"
    return template.replace("{prodco_names}", label)


def _load_hours_letter_prompt():
    with open(_HOURS_LETTER_PROMPT_PATH, "r", encoding="utf-8") as f:
        return f.read()


_VENDOR_TYPE_LABELS = {
    "sub_prodco": "Sub-ProdCo",
    "prodco":     "ProdCo",
    "agency":     "Agency",
}


def _load_billing_prompt(vendor_name: str, vendor_type: str, prodco_names: list) -> str:
    with open(_BILLING_PROMPT_PATH, "r", encoding="utf-8") as f:
        template = f.read()
    type_label   = _VENDOR_TYPE_LABELS.get(vendor_type.lower(), vendor_type)
    prodco_label = ", ".join(prodco_names) if prodco_names else "the production company"
    return (
        template
        .replace("{vendor_name}", vendor_name)
        .replace("{vendor_type}", type_label)
        .replace("{prodco_names}", prodco_label)
    )


def _load_residency_docs_prompt() -> str:
    with open(_RESIDENCY_DOCS_PROMPT_PATH, "r", encoding="utf-8") as f:
        return f.read()


def _load_agency_hours_prompt(agency_name: str = "") -> str:
    with open(_AGENCY_HOURS_PROMPT_PATH, "r", encoding="utf-8") as f:
        template = f.read()
    if agency_name.strip():
        context = (
            f"AGENCY: {agency_name.strip()} — this is the company that wrote these letters. "
            f"Use this name for agency_name on each row."
        )
    else:
        context = (
            "Extract the agency name from the letter header or signature block. "
            "Include it as agency_name on each row."
        )
    return template.replace("{agency_context}", context)


def _load_agency_subvendors_prompt(agency_name: str, agency_address: str = "") -> str:
    with open(_AGENCY_SUBVENDORS_PROMPT_PATH, "r", encoding="utf-8") as f:
        template = f.read()
    name_label = agency_name.strip() or "the ad agency"
    addr_label = f" ({agency_address.strip()})" if agency_address.strip() else ""
    return template.replace("{agency_name}", f"{name_label}{addr_label}")


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

def _call_gpt(images_b64, system_prompt, client, user_text="Extract all invoices from these document pages."):
    content = [{"type": "text", "text": user_text}]
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


def _is_handwritten(filename: str, data: bytes, client) -> bool:
    """Render first page only at low res and ask GPT if the form is handwritten."""
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
        img_b64 = base64.b64encode(pix.tobytes("png")).decode()
        doc.close()
    except Exception:
        return False
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": "Look at the data entry fields on this form (name, address, dates, etc.) — NOT the signature lines. Are those data fields filled in by hand (handwritten) or typed/digital? Reply with exactly one word: 'handwritten' or 'typed'."},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}", "detail": "low"}},
            ],
        }],
        max_tokens=10,
        temperature=0,
    )
    return "handwritten" in resp.choices[0].message.content.strip().lower()


def _extract_from_file(filename, data, system_prompt, client, user_text="Extract all invoices from these document pages.", dpi_scale=2.0):
    images = _file_to_images_b64(filename, data, dpi_scale=dpi_scale)
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
        invoices.extend(_call_gpt(batch, system_prompt, client, user_text=user_text))
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

_STATE_SET = frozenset(_STATE_ABBR.values())

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


def normalize_date(val: str) -> str:
    """Convert YYYY-MM-DD to MM/DD/YYYY; pass through anything else."""
    if not val:
        return ""
    s = str(val).strip()
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', s)
    if m:
        return f"{m.group(2)}/{m.group(3)}/{m.group(1)}"
    return s


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


_ADDR_STREET_TYPES = {
    "avenue":"Ave","boulevard":"Blvd","circle":"Cir","court":"Ct","drive":"Dr",
    "highway":"Hwy","lane":"Ln","parkway":"Pkwy","place":"Pl","road":"Rd",
    "street":"St","terrace":"Ter","trail":"Trl",
}
_ADDR_SINGLE_DIRS   = {"north":"N","south":"S","east":"E","west":"W"}
_ADDR_COMPOUND_DIRS = frozenset({"ne","nw","se","sw"})

def clean_address(val):
    if not val:
        return ""
    s = str(val).strip()
    s = re.sub(r",?\s*#\s*\S+", "", s)
    s = re.sub(r",?\s*(Apt|Apartment|Suite|Ste|Unit|Room|Fl|Floor)(\s+\S+)?", "", s, flags=re.IGNORECASE)
    result = []
    for word in s.split():
        lw = word.lower().rstrip(".")
        if lw in _ADDR_COMPOUND_DIRS:
            result.append(lw.upper())
        elif lw in _ADDR_SINGLE_DIRS:
            result.append(_ADDR_SINGLE_DIRS[lw])
        elif lw in _ADDR_STREET_TYPES:
            result.append(_ADDR_STREET_TYPES[lw])
        else:
            result.append(word.capitalize().rstrip("."))
    return " ".join(result).rstrip(",").strip()


def _parse_vendor_address(full_address: str) -> dict:
    """Split 'Street, City, ST Zip' into components working right-to-left."""
    if not full_address:
        return {"address": "", "city": "", "state": "", "zip": ""}
    rest = full_address.strip()

    zip_ = ""
    m = re.search(r'\b(\d{5}(?:-\d{4})?)\s*$', rest)
    if m:
        zip_ = m.group(1)
        rest = rest[:m.start()].strip().rstrip(",").strip()

    state = ""
    m = re.search(r'\b([A-Za-z]{2})\s*$', rest)
    if m and m.group(1).upper() in _STATE_SET:
        state = m.group(1).upper()
        rest = rest[:m.start()].strip().rstrip(",").strip()

    city = ""
    if "," in rest:
        idx  = rest.rfind(",")
        city = rest[idx + 1:].strip()
        rest = rest[:idx].strip()

    return {
        "address": clean_address(rest),
        "city":    clean_name(city),
        "state":   state,
        "zip":     zip_,
    }


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


def normalize_billing_row(
    raw: dict,
    vendor_type: str,
    vendor_name: str,
    vendor_address: str,
    vendor_city: str,
    vendor_state: str,
    vendor_zip: str,
    work_state: str,
    filename: str,
) -> dict:
    # If the wizard sent split fields, use them; otherwise parse the full address string.
    if vendor_city or vendor_state or vendor_zip:
        addr  = clean_address(vendor_address or raw.get("address", ""))
        city  = clean_name(vendor_city or raw.get("city", ""))
        state = clean_state(vendor_state or raw.get("state", ""))
        zip_  = clean_zip(vendor_zip or raw.get("zip", ""))
    else:
        parsed = _parse_vendor_address(vendor_address)
        addr  = parsed["address"] or clean_address(raw.get("address", ""))
        city  = parsed["city"]    or clean_name(raw.get("city", ""))
        state = parsed["state"]   or clean_state(raw.get("state", ""))
        zip_  = parsed["zip"]     or clean_zip(raw.get("zip", ""))

    local     = bool(state) and state.upper() == work_state.upper()
    geo_state = "Local" if local else "OOS"

    proj_fee_raw = raw.get("project_fee")
    proj_fee     = normalize_amount(proj_fee_raw) if proj_fee_raw not in (None, "", 0, 0.0) else None

    return {
        "qualify":         "",
        "vendorName":      vendor_name or clean_name(raw.get("vendor_name", "")),
        "vendorType":      _VENDOR_TYPE_LABELS.get(vendor_type.lower(), vendor_type),
        "jobType":         str(raw.get("job_type", "")).strip(),
        "state":           geo_state,
        "description":     str(raw.get("description", "")).strip(),
        "details":         str(raw.get("details", "")).strip(),
        "invoiceDate":     normalize_date(raw.get("invoice_date", "")),
        "invoiceNo":       str(raw.get("invoice_number", "")).strip(),
        "eligibleTotal":   normalize_amount(raw.get("invoice_amount", 0)),
        "projectFee":      proj_fee,
        "address":         addr,
        "city":            city,
        "vendorState":     state,
        "zip":             zip_,
        "receivedInvoice": "Yes",
        "pop":             "Yes" if raw.get("pop") else None,
        "paymentDate":     normalize_date(raw.get("payment_date", "")),
        "paymentMethod":   str(raw.get("payment_method", "")).strip(),
        "jobNumber":       str(raw.get("job_number", "")).strip(),
        "notes":           str(raw.get("notes", "")).strip(),
        "sourceFile":      filename,
    }


def normalize_agency_subvendor_row(raw: dict, agency_name: str, filename: str) -> dict:
    method = normalize_pymt_method(raw.get("payment_method", ""))
    return {
        "qualify":          "",
        "vendorName":       clean_name(raw.get("vendor_name", "")),
        "description":      str(raw.get("description", "")).strip(),
        "address":          clean_address(raw.get("address", "")),
        "city":             clean_name(raw.get("city", "")),
        "zip":              clean_zip(raw.get("zip", "")),
        "vendorState":      clean_state(raw.get("state", "")),
        "invoiceDate":      normalize_date(raw.get("invoice_date", "")),
        "invoiceNo":        str(raw.get("invoice_number", "")).strip(),
        "invoiceAmount":    normalize_amount(raw.get("invoice_amount", 0)),
        "receivedInvoice":  "Yes",
        "w9ValidDate":      normalize_date(raw.get("w9_date", "")),
        "pop":              "Yes" if raw.get("pop") else None,
        "paymentMethod":    method,
        "paymentNo":        normalize_pymt_number(method, raw.get("payment_number", "")),
        "paymentEntity":    agency_name.strip(),
        "jobNo":            str(raw.get("job_number", "")).strip(),
        "clientBillingNo":  str(raw.get("client_billing_number", "")).strip(),
        "notes":            str(raw.get("notes", "")).strip(),
        "sourceFile":       filename,
    }


def normalize_agency_hours_row(raw: dict, agency_name: str, filename: str) -> dict:
    hours = raw.get("agency_hours")
    try:
        hours = round(float(hours), 2) if hours is not None else None
    except (ValueError, TypeError):
        hours = None

    name = agency_name.strip() or clean_name(raw.get("agency_name", ""))

    return {
        "hoursLetter":        "Hours Letter",
        "invoiceDate":        normalize_date(raw.get("invoice_date", "")),
        "qualify":            "",
        "crewName":           clean_name(raw.get("crew_name", "")),
        "jobDescription":     str(raw.get("job_description", "")).strip(),
        "positionCategory":   str(raw.get("position_category", "")).strip(),
        "address":            clean_address(raw.get("address", "")),
        "city":               clean_name(raw.get("city", "")),
        "zip":                clean_zip(raw.get("zip", "")),
        "state":              clean_state(raw.get("state", "")),
        "agencyHours":        hours,
        "agencyHoursAmount":  normalize_amount(raw.get("agency_hours_amount", 0)),
        "hoursLetterType":    "Agency Hours Letter",
        "datesWorked":        str(raw.get("dates_worked", "")).strip(),
        "agencyName":         name,
        "sourceFile":         filename,
    }


def normalize_retainer_billing_row(raw: dict, agency_name: str, filename: str) -> dict:
    return {
        "invoiceNo":   str(raw.get("invoice_number", "")).strip(),
        "invoiceDate": normalize_date(raw.get("invoice_date", "")),
        "amount":      normalize_amount(raw.get("invoice_amount", 0)),
        "vendorName":  agency_name.strip() or clean_name(raw.get("vendor_name", "")),
        "sourceFile":  filename,
    }


def _i9_notes(raw: dict, handwritten: bool, shoot_date: str) -> str:
    notes = []

    if handwritten:
        notes.append("Handwritten - verify accuracy")

    if str(raw.get("document_type", "")).strip() != "I-9":
        return "; ".join(notes)

    if raw.get("employee_signed") is False:
        notes.append("Missing employee signature")

    sig_date = normalize_date(str(raw.get("employee_signature_date") or ""))
    if sig_date and shoot_date:
        try:
            from datetime import datetime
            diff = abs((datetime.strptime(sig_date, "%m/%d/%Y") - datetime.strptime(shoot_date, "%m/%d/%Y")).days)
            if diff > 90:
                notes.append(f"Signature date ({sig_date}) is not close to shoot date ({shoot_date})")
        except ValueError:
            pass

    exp_date = normalize_date(str(raw.get("expiration_date") or ""))
    dob      = normalize_date(str(raw.get("date_of_birth") or ""))
    if exp_date and dob:
        try:
            from datetime import datetime
            exp_dt = datetime.strptime(exp_date, "%m/%d/%Y")
            dob_dt = datetime.strptime(dob,      "%m/%d/%Y")
            if exp_dt.month != dob_dt.month or exp_dt.day != dob_dt.day:
                notes.append(f"Work auth expiration ({exp_date}) month/day does not match birth date ({dob})")
        except ValueError:
            pass

    missing = []
    if raw.get("employer_signed") is False:
        missing.append("employer signature")
    if not str(raw.get("employer_signature_date") or "").strip():
        missing.append("employer date")
    if not str(raw.get("first_day_of_employment") or "").strip():
        missing.append("first day of employment")
    if not str(raw.get("employer_name") or "").strip():
        missing.append("employer name")
    if not str(raw.get("employer_address") or "").strip():
        missing.append("employer address")
    if missing:
        notes.append("Missing employer section: " + ", ".join(missing))

    return "; ".join(notes)


def normalize_residency_row(raw: dict, filename: str, handwritten: bool = False, shoot_date: str = "") -> dict:
    return {
        "documentName":   str(raw.get("document_name", "")).strip(),
        "documentType":   str(raw.get("document_type", "")).strip(),
        "issueDate":      normalize_date(raw.get("issue_date", "")),
        "expirationDate": normalize_date(raw.get("expiration_date", "")),
        "address":        clean_address(raw.get("address", "")),
        "city":           clean_name(raw.get("city", "")),
        "zip":            clean_zip(raw.get("zip", "")),
        "state":          clean_state(raw.get("state", "")),
        "notes":          _i9_notes(raw, handwritten, shoot_date),
        "sourceFile":     filename,
    }


# ── Payroll company detection ─────────────────────────────────────────────────

def _is_wrapbook_register_only(pdf_bytes: bytes) -> bool:
    """Return True if this PDF is a standalone Wrapbook Payroll Register (NIS 007 style).

    A standalone register has 'Payroll Register' pages but no 'Fringe Report' pages.
    These are uploaded alongside fringe PDFs so their IL withholding data can enrich
    the project-level fringe rows (which have no invoice number).
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            has_register = has_fringe = False
            for pg in pdf.pages[:10]:
                text = (pg.extract_text() or "").lower()
                if "payroll register" in text and "xxx-xx-" in text:
                    has_register = True
                if "fringe report" in text or "fringe recap" in text:
                    has_fringe = True
                    break
            return has_register and not has_fringe
    except Exception:
        return False


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

    files = sorted(files, key=lambda f: (f.filename or "").lower())

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


async def _run_extract(files, x_app_secret, payroll_hints: str = ""):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    # Read all files up front so we can classify before extracting
    loaded = []
    for uf in files:
        data = await uf.read()
        loaded.append((uf.filename, data))
    loaded.sort(key=lambda x: (x[0] or "").lower())

    # Classify each file: fringe PDF vs standalone Wrapbook register
    fringe_files   = []   # (filename, bytes)
    register_files = []   # bytes of standalone Wrapbook register PDFs

    for filename, data in loaded:
        if _is_wrapbook_register_only(data):
            register_files.append(data)
        else:
            fringe_files.append((filename, data))

    rows, issues, file_summaries = [], [], []
    wb_rows: list[dict] = []   # Wrapbook fringe rows that may need register enrichment

    # Track project-level Wrapbook fringe sources (invoiceNo always blank — one fringe
    # report consolidates many invoices). Stored as (filename, row_list) so we can
    # assign "Fringe Report 001 / 002 / …" labels after enrichment is complete.
    wb_project_sources: list[tuple[str, list[dict]]] = []

    # Ordered list of email alerts — one entry per generation event.
    # [{company_name, code, file_count, row_count, is_update}]
    alert_queue: list[dict] = []

    def _latest_alert(company: str) -> dict | None:
        for entry in reversed(alert_queue):
            if entry["company_name"] == company:
                return entry
        return None

    for filename, data in fringe_files:
        # ── Find all matching parsers (static + runtime) sorted by priority ────
        candidates = registry.find_parsers(data)

        extracted: list[dict] = []
        errs:      list[str]  = []
        company:   str | None = None

        # Try each candidate in order; stop on first that returns rows
        for candidate in candidates:
            try:
                extracted, errs = candidate.extract(data, openai_key=OPENAI_API_KEY)
            except Exception as e:
                extracted, errs = [], [f"{candidate.COMPANY} parser error: {e}"]
            if extracted:
                company = candidate.COMPANY
                break

        # ── Wrapbook-specific post-processing ─────────────────────────────────
        if company == "wrapbook":
            wb_rows.extend(extracted)
            if extracted and all(r.get("invoiceNo", "") == "" for r in extracted):
                wb_project_sources.append((filename, extracted))

        # ── Update alert counts for successful runtime-parser hits ─────────────
        if extracted and company:
            entry = _latest_alert(company)
            if entry:
                entry["file_count"] += 1
                entry["row_count"]  += len(extracted)

        # ── Phase 3: all known layout versions failed → generate new variant ──
        if not extracted:
            known_company = candidates[0].COMPANY if candidates else None
            hints = payroll_hints
            if known_company:
                hints = (
                    f"Company: {known_company} — all {len(candidates)} known layout(s) failed, likely a new variant\n"
                    + hints
                )

            ai_rows, ai_errs, ai_company, markers, code = extract_unknown(
                data, OPENAI_API_KEY, hints=hints, anthropic_key=ANTHROPIC_API_KEY
            )
            extracted, errs = ai_rows, ai_errs
            if ai_company:
                exec_fn = make_exec_parser(code)
                if exec_fn:
                    registry.register_parser(ai_company, markers, exec_fn)
                existing = _latest_alert(ai_company)
                if existing:
                    existing["file_count"] += 1
                    existing["row_count"]  += len(extracted)
                else:
                    alert_queue.append({
                        "company_name": ai_company,
                        "code":         code,
                        "file_count":   1,
                        "row_count":    len(extracted),
                        "is_update":    bool(known_company),
                    })
                company = ai_company
            else:
                errs.append(f"Could not identify payroll company in {filename}")
                company = known_company or "unknown"

        for r in extracted:
            r["sourceFile"] = filename

        rows.extend(extracted)
        for e in errs:
            issues.append(f"{filename}: {e}")

        file_summaries.append({
            "filename": filename,
            "company":  company or "unknown",
            "rows":     len(extracted),
            "issues":   errs,
        })

    # Enrich Wrapbook fringe rows with any standalone register PDFs (NIS 007 style).
    # Must happen BEFORE "Fringe Report" label assignment — enrichment joins on invoiceNo == "".
    for register_bytes in register_files:
        enrich_from_register(wb_rows, register_bytes)

    if register_files and not wb_rows:
        issues.append(
            f"{len(register_files)} Wrapbook register file(s) uploaded but no Wrapbook fringe "
            "rows found to enrich. Upload the fringe PDF alongside the register."
        )

    # Assign "Fringe Report" invoice labels for project-level fringe rows.
    # Single source → "Fringe Report"; multiple sources → "Fringe Report 001", "002", …
    if len(wb_project_sources) == 1:
        for r in wb_project_sources[0][1]:
            r["invoiceNo"] = "Fringe Report"
    elif len(wb_project_sources) > 1:
        for idx, (_, src_rows) in enumerate(wb_project_sources, 1):
            label = f"Fringe Report {idx:03d}"
            for r in src_rows:
                r["invoiceNo"] = label

    # Send one email per generation event (initial discovery or retry)
    for info in alert_queue:
        company_name = info["company_name"]
        sent = send_parser_alert(
            company_name,
            info["file_count"],
            info["row_count"],
            info["code"],
            is_update=info.get("is_update", False),
        )
        verb      = "Updated parser" if info.get("is_update") else "New parser"
        alert_msg = (
            f"{verb} for '{company_name}' — "
            f"{info['file_count']} file(s) processed with AI extraction"
        )
        alert_msg += (
            f". Parser code emailed to {ALERT_EMAIL} for review." if sent
            else ". (Email alert failed — check SENDGRID_API_KEY)"
        )
        issues.append(alert_msg)

    return {"rows": rows, "issues": issues, "columns": FRINGE_FIELDS, "files": file_summaries}


@app.post("/extract-fringe")
async def extract_fringe(
    files: list[UploadFile] = File(...),
    payroll_hints: str = Form(default=""),
    x_app_secret: str = Header(default=""),
):
    return await _run_extract(files, x_app_secret, payroll_hints=payroll_hints)


@app.post("/extract-payroll")
async def extract_payroll(
    files: list[UploadFile] = File(...),
    payroll_hints: str = Form(default=""),
    x_app_secret: str = Header(default=""),
):
    return await _run_extract(files, x_app_secret, payroll_hints=payroll_hints)


# ── Crew freelance invoice extractor ─────────────────────────────────────────

_CREW_REQUIRED = {"worker", "invoiceNo", "invoiceDate", "wages"}

def normalize_crew_freelance_row(raw: dict, filename: str) -> dict:
    worker       = clean_name(raw.get("worker", "")) or "[missing information]"
    invoice_no   = str(raw.get("invoiceNo", "")).strip() or "[missing information]"
    invoice_date = str(raw.get("invoiceDate", "")).strip() or "[missing information]"
    wages        = normalize_amount(raw.get("wages", 0))
    if wages == 0:
        wages = "[missing information]"

    method = normalize_pymt_method(raw.get("pymtMethod", ""))

    days = raw.get("daysWorked")
    if isinstance(days, float):
        days = int(days)
    elif not isinstance(days, int):
        days = None

    return {
        "worker":        worker,
        "jobTitle":      clean_name(raw.get("jobTitle", "")),
        "invoiceNo":     invoice_no,
        "invoiceDate":   invoice_date,
        "workDates":     str(raw.get("workDates", "")).strip(),
        "daysWorked":    days,
        "wages":         wages,
        "kitRental":     normalize_amount(raw.get("kitRental", 0)) or None,
        "mileage":       normalize_amount(raw.get("mileage", 0)) or None,
        "reimbursement": normalize_amount(raw.get("reimbursement", 0)) or None,
        "other":         normalize_amount(raw.get("other", 0)) or None,
        "invoiceTotal":  normalize_amount(raw.get("invoiceTotal", 0)) or None,
        "poNo":          str(raw.get("poNo", "")).strip(),
        "pymtMethod":    method,
        "pymtNo":        normalize_pymt_number(method, raw.get("pymtNo", "")),
        "street":        clean_address(raw.get("street", "")),
        "city":          clean_name(raw.get("city", "")),
        "state":         clean_state(raw.get("state", "")),
        "zip":           clean_zip(raw.get("zip", "")),
        "sourceFile":    filename,
    }


@app.post("/extract-crew-freelance")
async def extract_crew_freelance(
    files: list[UploadFile] = File(...),
    prodco_names: str = Form(""),
    x_app_secret: str = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    files = sorted(files, key=lambda f: (f.filename or "").lower())

    client = _client()
    names         = [n.strip() for n in prodco_names.split(",") if n.strip()]
    system_prompt = _load_crew_freelance_prompt(names)
    user_text     = "Extract crew freelance invoice data from these document pages."

    rows, issues, file_summaries = [], [], []

    for uf in files:
        data = await uf.read()
        errs: list[str] = []

        try:
            raw_list = _extract_from_file(uf.filename, data, system_prompt, client, user_text=user_text)
        except Exception as e:
            errs.append(str(e))
            issues.append(f"{uf.filename}: {e}")
            raw_list = []

        file_rows: list[dict] = []
        if not raw_list:
            errs.append("no crew freelance data extracted — review manually")
            issues.append(f"{uf.filename}: no crew freelance data extracted")
        else:
            for raw in raw_list:
                try:
                    file_rows.append(normalize_crew_freelance_row(raw, uf.filename))
                except Exception as e:
                    errs.append(f"row normalization error: {e}")
                    issues.append(f"{uf.filename}: row normalization error: {e}")

        rows.extend(file_rows)
        worker_label = file_rows[0]["worker"] if file_rows else "unknown"
        file_summaries.append({
            "filename": uf.filename,
            "company":  worker_label,
            "rows":     len(file_rows),
            "issues":   errs,
        })

    return {"rows": rows, "issues": issues, "files": file_summaries}


# ── Talent freelance invoice extractor ───────────────────────────────────────

def normalize_talent_freelance_invoice(raw: dict, filename: str) -> list[dict]:
    talent_name    = clean_name(raw.get("talentName", "")) or "[missing information]"
    agency_name    = str(raw.get("agencyName", "")).strip()
    invoice_no     = str(raw.get("invoiceNo", "")).strip() or "[missing information]"
    invoice_date   = str(raw.get("invoiceDate", "")).strip() or "[missing information]"
    work_dates     = str(raw.get("workDates", "")).strip()
    payment_entity = str(raw.get("paymentEntity", "")).strip()

    days = raw.get("daysWorked")
    if isinstance(days, float):
        days = int(days)
    elif not isinstance(days, int):
        days = None

    talent_wages    = normalize_amount(raw.get("talentWages", 0))
    agency_fee      = normalize_amount(raw.get("agencyFee", 0))
    agency_expenses = normalize_amount(raw.get("agencyExpenses", 0))
    misc_pymt       = round(agency_fee + agency_expenses, 2)
    work_state      = clean_state(raw.get("workState", ""))

    method = normalize_pymt_method(raw.get("pymtMethod", ""))
    pymt_no = normalize_pymt_number(method, raw.get("pymtNo", ""))

    talent_row = {
        "talentName":      talent_name,
        "title":           "Talent",
        "rowType":         "talent",
        "invoiceNo":       invoice_no,
        "invoiceDate":     invoice_date,
        "workDates":       work_dates,
        "daysWorked":      days,
        "wages":           talent_wages if talent_wages else "[missing information]",
        "miscPymt":        0,
        "qualify":         "",
        "includedOnPtip":  "NO",
        "workState":       work_state,
        "receivedInvoice": "YES",
        "paymentEntity":   payment_entity,
        "pymtMethod":      method,
        "pymtNo":          pymt_no,
        "street":          clean_address(raw.get("talentStreet", "")),
        "city":            clean_name(raw.get("talentCity", "")),
        "state":           clean_state(raw.get("talentState", "")),
        "zip":             clean_zip(raw.get("talentZip", "")),
        "sourceFile":      filename,
    }

    rows = [talent_row]

    if agency_name:
        agency_row = {
            "talentName":      clean_name(agency_name),
            "title":           "Agency Fee",
            "rowType":         "agency",
            "invoiceNo":       invoice_no,
            "invoiceDate":     invoice_date,
            "workDates":       work_dates,
            "daysWorked":      days,
            "wages":           0,
            "miscPymt":        misc_pymt if misc_pymt else "[missing information]",
            "qualify":         "",
            "includedOnPtip":  "NO",
            "workState":       "",
            "receivedInvoice": "YES",
            "paymentEntity":   payment_entity,
            "pymtMethod":      method,
            "pymtNo":          pymt_no,
            "street":          "",
            "city":            "",
            "state":           "",
            "zip":             "",
            "sourceFile":      filename,
        }
        rows.append(agency_row)

    return rows


@app.post("/extract-talent-freelance")
async def extract_talent_freelance(
    files:        list[UploadFile] = File(...),
    prodco_names: str              = Form(""),
    x_app_secret: str              = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    files = sorted(files, key=lambda f: (f.filename or "").lower())

    client        = _client()
    names         = [n.strip() for n in prodco_names.split(",") if n.strip()]
    system_prompt = _load_talent_freelance_prompt(names)
    user_text     = "Extract talent freelance invoice data from these document pages."

    rows, issues, file_summaries = [], [], []

    for uf in files:
        data = await uf.read()
        errs: list[str] = []

        try:
            raw_list = _extract_from_file(uf.filename, data, system_prompt, client, user_text=user_text)
        except Exception as e:
            errs.append(str(e))
            issues.append(f"{uf.filename}: {e}")
            raw_list = []

        file_rows: list[dict] = []
        if not raw_list:
            errs.append("no talent freelance data extracted — review manually")
            issues.append(f"{uf.filename}: no talent freelance data extracted")
        else:
            for raw in raw_list:
                try:
                    file_rows.extend(normalize_talent_freelance_invoice(raw, uf.filename))
                except Exception as e:
                    errs.append(f"row normalization error: {e}")
                    issues.append(f"{uf.filename}: row normalization error: {e}")

        rows.extend(file_rows)
        talent_label = file_rows[0]["talentName"] if file_rows else "unknown"
        file_summaries.append({
            "filename": uf.filename,
            "talent":   talent_label,
            "rows":     len(file_rows),
            "issues":   errs,
        })

    return {"rows": rows, "issues": issues, "files": file_summaries}


# ── Hours letter extractor ────────────────────────────────────────────────────

def normalize_hours_letter_row(raw: dict, filename: str) -> dict:
    worker = clean_name(raw.get("worker", "")) or "[missing information]"
    wages  = normalize_amount(raw.get("wages", 0))
    if wages == 0:
        wages = "[missing information]"

    days = raw.get("daysWorked")
    if isinstance(days, float):
        days = int(days)
    elif not isinstance(days, int):
        days = None

    return {
        "worker":       worker,
        "jobTitle":     clean_name(raw.get("jobTitle", "")),
        "daysWorked":   days,
        "wages":        wages,
        "invoiceTotal": wages,
        "invoiceDate":  str(raw.get("invoiceDate", "")).strip(),
        "company":      str(raw.get("company", "")).strip(),
        "sourceFile":   filename,
    }


@app.post("/extract-hours-letters")
async def extract_hours_letters(
    files: list[UploadFile] = File(...),
    x_app_secret: str = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    files = sorted(files, key=lambda f: (f.filename or "").lower())

    client        = _client()
    system_prompt = _load_hours_letter_prompt()
    user_text     = "Extract crew hours and billable amounts from this hours confirmation letter."

    rows, issues, file_summaries = [], [], []
    sources: list[tuple[str, list[dict]]] = []   # (filename, file_rows) for labeling

    for uf in files:
        data = await uf.read()
        errs: list[str] = []

        try:
            raw_list = _extract_from_file(uf.filename, data, system_prompt, client, user_text=user_text)
        except Exception as e:
            errs.append(str(e))
            issues.append(f"{uf.filename}: {e}")
            raw_list = []

        file_rows: list[dict] = []
        if not raw_list:
            errs.append("no hours letter data extracted — review manually")
            issues.append(f"{uf.filename}: no hours letter data extracted")
        else:
            for raw in raw_list:
                try:
                    file_rows.append(normalize_hours_letter_row(raw, uf.filename))
                except Exception as e:
                    errs.append(f"row normalization error: {e}")
                    issues.append(f"{uf.filename}: row normalization error: {e}")

        rows.extend(file_rows)
        if file_rows:
            sources.append((uf.filename, file_rows))
        company_label = file_rows[0]["company"] if file_rows and file_rows[0]["company"] else "unknown"
        file_summaries.append({
            "filename": uf.filename,
            "company":  company_label,
            "rows":     len(file_rows),
            "issues":   errs,
        })

    # Assign invoice labels — mirrors Wrapbook "Fringe Report" pattern
    if len(sources) == 1:
        for row in sources[0][1]:
            row["invoiceNo"] = "Hours Letter"
    elif len(sources) > 1:
        for idx, (_, src_rows) in enumerate(sources, 1):
            for row in src_rows:
                row["invoiceNo"] = f"Hours Letter {idx:03d}"

    return {"rows": rows, "issues": issues, "files": file_summaries}


# ── Billings extractor ───────────────────────────────────────────────────────

@app.post("/extract-billings")
async def extract_billings(
    files:          list[UploadFile] = File(...),
    vendor_type:    str              = Form(""),
    vendor_name:    str              = Form(""),
    vendor_address: str              = Form(""),
    vendor_city:    str              = Form(""),
    vendor_state:   str              = Form(""),
    vendor_zip:     str              = Form(""),
    prodco_names:   str              = Form(""),
    work_state:     str              = Form("IL"),
    x_app_secret:   str              = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    files = sorted(files, key=lambda f: (f.filename or "").lower())

    client        = _client()
    names         = [n.strip() for n in prodco_names.split(",") if n.strip()]
    system_prompt = _load_billing_prompt(vendor_name, vendor_type, names)
    user_text     = "Extract billing invoice data from these document pages."

    rows, issues, file_summaries = [], [], []

    for uf in files:
        data = await uf.read()
        errs: list[str] = []

        try:
            raw_list = _extract_from_file(uf.filename, data, system_prompt, client, user_text=user_text)
        except Exception as e:
            errs.append(str(e))
            issues.append(f"{uf.filename}: {e}")
            raw_list = []

        file_rows: list[dict] = []
        if not raw_list:
            errs.append("no billing data extracted — review manually")
            issues.append(f"{uf.filename}: no billing data extracted")
        else:
            for raw in raw_list:
                try:
                    file_rows.append(normalize_billing_row(
                        raw,
                        vendor_type=vendor_type,
                        vendor_name=vendor_name,
                        vendor_address=vendor_address,
                        vendor_city=vendor_city,
                        vendor_state=vendor_state,
                        vendor_zip=vendor_zip,
                        work_state=work_state,
                        filename=uf.filename,
                    ))
                except Exception as e:
                    errs.append(f"row normalization error: {e}")
                    issues.append(f"{uf.filename}: row normalization error: {e}")

        rows.extend(file_rows)
        file_summaries.append({
            "filename": uf.filename,
            "company":  vendor_name or "unknown",
            "rows":     len(file_rows),
            "issues":   errs,
        })

    return {"rows": rows, "issues": issues, "files": file_summaries}


# ── Agency Sub-Vendors extractor ─────────────────────────────────────────────

@app.post("/extract-agency-subvendors")
async def extract_agency_subvendors(
    files:          list[UploadFile] = File(...),
    agency_name:    str              = Form(""),
    agency_address: str              = Form(""),
    x_app_secret:   str              = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    files = sorted(files, key=lambda f: (f.filename or "").lower())

    client        = _client()
    system_prompt = _load_agency_subvendors_prompt(agency_name, agency_address)
    user_text     = "Extract invoice data from these document pages."

    rows, issues, file_summaries = [], [], []

    for uf in files:
        data = await uf.read()
        errs: list[str] = []

        try:
            raw_list = _extract_from_file(uf.filename, data, system_prompt, client, user_text=user_text)
        except Exception as e:
            errs.append(str(e))
            issues.append(f"{uf.filename}: {e}")
            raw_list = []

        file_rows: list[dict] = []
        if not raw_list:
            errs.append("no agency sub-vendor data extracted — review manually")
            issues.append(f"{uf.filename}: no agency sub-vendor data extracted")
        else:
            for raw in raw_list:
                try:
                    file_rows.append(normalize_agency_subvendor_row(raw, agency_name, uf.filename))
                except Exception as e:
                    errs.append(f"row normalization error: {e}")
                    issues.append(f"{uf.filename}: row normalization error: {e}")

        rows.extend(file_rows)
        vendor_label = file_rows[0]["vendorName"] if file_rows else "unknown"
        file_summaries.append({
            "filename": uf.filename,
            "company":  vendor_label,
            "rows":     len(file_rows),
            "issues":   errs,
        })

    return {"rows": rows, "issues": issues, "files": file_summaries}


# ── Agency Hours extractor ───────────────────────────────────────────────────

@app.post("/extract-agency-hours")
async def extract_agency_hours(
    files:        list[UploadFile] = File(...),
    agency_name:  str              = Form(""),
    x_app_secret: str              = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    files = sorted(files, key=lambda f: (f.filename or "").lower())

    client        = _client()
    system_prompt = _load_agency_hours_prompt(agency_name)
    user_text     = "Extract crew hours data from these agency hours letter pages."

    rows, issues, file_summaries = [], [], []

    for uf in files:
        data = await uf.read()
        errs: list[str] = []

        try:
            raw_list = _extract_from_file(uf.filename, data, system_prompt, client, user_text=user_text)
        except Exception as e:
            errs.append(str(e))
            issues.append(f"{uf.filename}: {e}")
            raw_list = []

        file_rows: list[dict] = []
        if not raw_list:
            errs.append("no agency hours data extracted — review manually")
            issues.append(f"{uf.filename}: no agency hours data extracted")
        else:
            for raw in raw_list:
                try:
                    file_rows.append(normalize_agency_hours_row(raw, agency_name, uf.filename))
                except Exception as e:
                    errs.append(f"row normalization error: {e}")
                    issues.append(f"{uf.filename}: row normalization error: {e}")

        rows.extend(file_rows)
        agency_label = agency_name or (file_rows[0]["agencyName"] if file_rows else "unknown")
        file_summaries.append({
            "filename": uf.filename,
            "company":  agency_label,
            "rows":     len(file_rows),
            "issues":   errs,
        })

    return {"rows": rows, "issues": issues, "files": file_summaries}


# ── Retainer Billings extractor ───────────────────────────────────────────────

@app.post("/extract-retainer-billings")
async def extract_retainer_billings(
    files:        list[UploadFile] = File(...),
    agency_name:  str              = Form(""),
    prodco_names: str              = Form(""),
    x_app_secret: str              = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    files = sorted(files, key=lambda f: (f.filename or "").lower())

    client        = _client()
    names         = [n.strip() for n in prodco_names.split(",") if n.strip()]
    system_prompt = _load_billing_prompt(agency_name, "agency", names)
    user_text     = "Extract retainer billing invoice data from these document pages."

    rows, issues, file_summaries = [], [], []

    for uf in files:
        data = await uf.read()
        errs: list[str] = []

        try:
            raw_list = _extract_from_file(uf.filename, data, system_prompt, client, user_text=user_text)
        except Exception as e:
            errs.append(str(e))
            issues.append(f"{uf.filename}: {e}")
            raw_list = []

        file_rows: list[dict] = []
        if not raw_list:
            errs.append("no retainer billing data extracted — review manually")
            issues.append(f"{uf.filename}: no retainer billing data extracted")
        else:
            for raw in raw_list:
                try:
                    file_rows.append(normalize_retainer_billing_row(raw, agency_name, uf.filename))
                except Exception as e:
                    errs.append(f"row normalization error: {e}")
                    issues.append(f"{uf.filename}: row normalization error: {e}")

        rows.extend(file_rows)
        file_summaries.append({
            "filename": uf.filename,
            "company":  agency_name or "unknown",
            "rows":     len(file_rows),
            "issues":   errs,
        })

    return {"rows": rows, "issues": issues, "files": file_summaries}


# ── Residency documents extractor ────────────────────────────────────────────

@app.post("/extract-residency-docs")
async def extract_residency_docs(
    files:        list[UploadFile] = File(...),
    shoot_date:   str              = Form(default=""),
    x_app_secret: str              = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    files = sorted(files, key=lambda f: (f.filename or "").lower())

    client        = _client()
    system_prompt = _load_residency_docs_prompt()
    user_text     = "Extract personal information from these residency verification documents."

    rows, issues, file_summaries = [], [], []

    for uf in files:
        data = await uf.read()
        errs: list[str] = []

        handwritten = _is_handwritten(uf.filename, data, client)
        dpi_scale   = 4.17 if handwritten else 2.0  # ~300 DPI for handwritten, ~144 for typed

        try:
            raw_list = _extract_from_file(uf.filename, data, system_prompt, client, user_text=user_text, dpi_scale=dpi_scale)
        except Exception as e:
            errs.append(str(e))
            issues.append(f"{uf.filename}: {e}")
            raw_list = []

        file_rows: list[dict] = []
        if not raw_list:
            errs.append("no residency document data extracted — review manually")
            issues.append(f"{uf.filename}: no residency document data extracted")
        else:
            for raw in raw_list:
                try:
                    file_rows.append(normalize_residency_row(raw, uf.filename, handwritten=handwritten, shoot_date=shoot_date))
                except Exception as e:
                    errs.append(f"row normalization error: {e}")
                    issues.append(f"{uf.filename}: row normalization error: {e}")

        rows.extend(file_rows)
        name_label = file_rows[0]["documentName"] if file_rows else "unknown"
        file_summaries.append({
            "filename": uf.filename,
            "company":  name_label,
            "rows":     len(file_rows),
            "issues":   errs,
        })

    return {"rows": rows, "issues": issues, "files": file_summaries}


# ── Talent & Extras extractor ────────────────────────────────────────────────

@app.post("/extract-talent")
async def extract_talent_endpoint(
    pdf_files:      list[UploadFile] = File(default=[]),
    ptip_file:      UploadFile       = File(default=None),   # ER: single PTIP file
    ptip_files:     list[UploadFile] = File(default=[]),     # Teams: multiple PTIP files
    project_title:  str              = Form(default=""),
    workbook_type:  str              = Form(default=""),
    payroll_company: str             = Form(default="er"),   # "er" or "teams"
    x_app_secret:   str              = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    pdf_bytes_list: list[tuple[str, bytes]] = []
    for uf in sorted(pdf_files or [], key=lambda f: (f.filename or "").lower()):
        data = await uf.read()
        if data:
            pdf_bytes_list.append((uf.filename, data))

    if payroll_company.lower() == 'teams':
        ptip_bytes_list: list[bytes] = []
        for uf in (ptip_files or []):
            data = await uf.read()
            if data:
                ptip_bytes_list.append(data)
        # Also accept single ptip_file upload for Teams if ptip_files is empty
        if not ptip_bytes_list and ptip_file:
            data = await ptip_file.read()
            if data:
                ptip_bytes_list.append(data)

        if not pdf_bytes_list and not ptip_bytes_list:
            raise HTTPException(400, "Provide at least one PDF or PTIP file.")

        return extract_teams_talent(
            pdf_files=pdf_bytes_list,
            ptip_bytes_list=ptip_bytes_list,
            project_title=project_title,
            workbook_type=workbook_type,
        )

    # Default: Extreme Reach
    ptip_bytes: bytes | None = None
    if ptip_file:
        ptip_bytes = await ptip_file.read()
        if not ptip_bytes:
            ptip_bytes = None

    if not pdf_bytes_list and not ptip_bytes:
        raise HTTPException(400, "Provide at least one PDF or a PTIP file.")

    return extract_talent(
        pdf_files=pdf_bytes_list,
        ptip_bytes=ptip_bytes,
        project_title=project_title,
        workbook_type=workbook_type,
    )


# ── Consolidated run summary email ───────────────────────────────────────────

class _FileSummaryIn(BaseModel):
    filename: str = ""
    company:  str = ""
    rows:     int | None = 0
    issues:   list[str] = []

class _RunDataIn(BaseModel):
    endpoint: str = ""
    files:    list[_FileSummaryIn] = []
    issues:   list[str] = []

class _RunSummaryRequest(BaseModel):
    project_title: str = ""
    workbook_type: str = ""
    runs:          list[_RunDataIn] = []


@app.post("/send-run-summary")
async def send_run_summary_endpoint(
    payload: _RunSummaryRequest,
    x_app_secret: str = Header(default=""),
):
    if APP_SHARED_SECRET and x_app_secret != APP_SHARED_SECRET:
        raise HTTPException(401, "Bad or missing X-App-Secret header.")

    runs_dicts = [
        {
            "endpoint": r.endpoint,
            "files":    [f.dict() for f in r.files],
            "issues":   r.issues,
        }
        for r in payload.runs
    ]
    send_run_summary(
        project_title=payload.project_title,
        workbook_type=payload.workbook_type,
        runs=runs_dicts,
    )
    return {"ok": True}
