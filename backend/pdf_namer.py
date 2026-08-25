"""
Wizard 01 -- PDF classify / extract / name logic.

Pure logic module: no FastAPI, no job/storage/filesystem concerns (mirrors
how talent_extractor.py is separated from main.py's endpoints). The job
layer (wizard01_jobs.py) is responsible for everything about where bytes
live on disk; this module only ever deals with bytes in memory and returns
a DocResult describing what to do with them.

Naming conventions, prompt wording, and all string/PDF helpers below are
ported near-verbatim from the standalone Tkinter tool this replaces
(`Rename Residency and Diversity/rename_tool.py`), which already handles
these exact document types correctly -- the only genuinely new pieces here
are:
  - the type-classifier step (that tool always knew the type already,
    since a human pre-sorted into per-type folders; here we don't)
  - BatchContext (Client/Agency/ProdCo names+addresses, received_from) and
    the sender-mismatch check on Vendor documents
  - closed reason codes instead of free-text-only reasons, for a UI that
    needs to group/count/label these
"""

import os
import io
import re
import json
import base64
from dataclasses import dataclass, field

import fitz  # PyMuPDF
from PIL import Image, ImageOps

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_BUSINESS_WORDS = {
    "llc", "inc", "incorporated", "corp", "corporation", "dba", "productions",
    "production", "company", "co", "ltd", "llp", "enterprises", "group",
    "studios", "studio", "media", "films", "film", "pictures", "entertainment",
}
_UPPERCASE_EXCEPTIONS = {"llc", "inc"}
_NAME_SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v", "2nd", "3rd", "4th"}


# ── Reason codes ──────────────────────────────────────────────────────────────
# Closed set so a UI can group/count/label distinctly. reason_detail carries
# the free-text specifics alongside every code below (may be "" for codes
# that are self-explanatory, e.g. REASON_OK).

REASON_OK               = ""                  # success, no reason needed
REASON_UNCLASSIFIED     = "unclassified"      # readable, but not a known type
REASON_UNREADABLE       = "unreadable"        # corrupt/blank, both providers failed
REASON_MISSING_NAME     = "missing_name"
REASON_MISSING_DATE     = "missing_date"
REASON_MISSING_STATE    = "missing_state"
REASON_MISSING_COMPANY  = "missing_company"
REASON_SENDER_MISMATCH  = "sender_mismatch"
REASON_LOW_CONFIDENCE   = "low_confidence"
REASON_PROVIDER_ERROR   = "provider_error"

BUCKET_RENAMED          = "renamed"
BUCKET_UNABLE_TO_RENAME = "unable_to_rename"
BUCKET_NOT_READABLE     = "not_readable"

FOLDER_RESIDENCY  = "Residency"
FOLDER_DIVERSITY  = "Diversity Form"
FOLDER_VENDOR     = "Vendor"
FOLDER_NOT_READABLE = "Not Readable"


@dataclass
class BatchContext:
    """One per job, filled once from the Overview-style intake form and
    threaded into every per-file call. received_from is exactly one of
    "prodco" / "agency" / "client" -- a batch is never mixed-source."""
    client_name: str = ""
    client_address: str = ""
    agency_name: str = ""
    agency_address: str = ""
    prodco_name: str = ""
    prodco_address: str = ""
    received_from: str = ""


@dataclass
class DocResult:
    bucket: str            # BUCKET_* constant
    doc_type: str          # "residency" | "diversity" | "vendor" | "unknown"
    subfolder: str         # FOLDER_* constant -- where this lands in the zip
    output_pdf_bytes: bytes  # possibly image->PDF-converted bytes to write
    new_filename: str = "" # set only when bucket == BUCKET_RENAMED
    reason_code: str = REASON_OK
    reason_detail: str = ""
    provider: str = ""     # "openai" | "anthropic" | ""


# ── Prompts ────────────────────────────────────────────────────────────────────

TYPE_CLASSIFIER_PROMPT = """You are looking at a single scanned or photographed document submitted as part of a film/TV production crew paperwork batch. Determine which broad family it belongs to:

- "residency" -- a Driver's License, a State ID card, or a USCIS Form I-9 (Employment Eligibility Verification).
- "diversity" -- an Illinois Department of Commerce & Economic Opportunity (DCEO) "Illinois Film Tax Credit Tracking Sheet."
- "vendor" -- a hotel guest folio, a purchase receipt/order confirmation, or a formal invoice (vendor or freelance) billing the production for goods, services, or labor.
- "unknown" -- anything else, or if you cannot confidently tell from what's visible.

Return ONLY a JSON object with exactly this key: {"document_family": "residency" | "diversity" | "vendor" | "unknown"}
No explanation. No markdown. No code fences. JSON object only."""

ORIENTATION_PROMPT = """You are looking at page images from a single PDF document -- a scan or phone photo submitted as part of a film/TV production crew paperwork batch. For each page, determine how many degrees it needs to be rotated CLOCKWISE for its content (text, photos) to appear upright/right-side-up. Common causes: a phone photo taken in landscape when the document is portrait (or vice versa), or a scanner fed the page sideways or upside-down.

Return one entry per page, in the same order given (page 1 first).
- rotation: one of 0, 90, 180, 270 -- how many degrees clockwise this page needs to be rotated to become upright. 0 if it's already upright.

Return ONLY a JSON object with exactly this key: {"pages": [{"rotation": 0 | 90 | 180 | 270}, ...]}
No explanation. No markdown. No code fences. JSON object only."""

RESIDENCY_SYSTEM_PROMPT = """You are analyzing a scanned identification or employment eligibility document submitted by a film production crew member. It will be one of: a Driver's License, a State ID card, or a USCIS Form I-9 (Employment Eligibility Verification). If a second page is included, it is almost always just the BACK of the same card (blood type, restrictions, barcodes) -- treat all pages together as ONE document, never two.

Determine which type this is:
- "drivers_license" -- explicitly says DRIVER'S LICENSE
- "state_id" -- a state-issued ID card that is NOT a driver's license (says ID CARD / IDENTIFICATION CARD, no driving class/privileges)
- "i9" -- USCIS Form I-9, Employment Eligibility Verification
- "unknown" -- if you cannot confidently tell from what's visible

Then extract:
- last_name: the person's last/family name, exactly as printed, title case. For a driver's license/state ID this is field "1" (the first name line, e.g. "ALLEN"). For an I-9 this is "Last Name (Family Name)" in Section 1. If a generational suffix (Jr, Sr, II, III, 3rd, etc.) is printed as part of that same name, keep it attached to last_name (e.g. "Sanders Jr") -- never treat the suffix as first_name or as the whole last_name by itself.
- first_name: the person's FIRST/given name ONLY -- never include a middle name. For a driver's license/state ID, field "2" often has first + middle name together (e.g. "CAROLINE ROSE") -- use only the first word ("Caroline"). For an I-9 use "First Name (Given Name)" in Section 1.
- state: ONLY for drivers_license or state_id, the 2-letter issuing state abbreviation exactly as printed on the card (e.g. "IL", "CA") -- this is usually part of the header ("ILLINOIS DRIVER LICENSE") or printed near the state seal. Empty string for i9 or unknown.
- expiration_date: ONLY for drivers_license or state_id, in exactly MM/DD/YYYY as printed next to "EXP" or field "4b". Empty string for i9 or unknown.

Return ONLY a JSON object with exactly these keys: {"document_type": "...", "last_name": "...", "first_name": "...", "state": "...", "expiration_date": "..."}
No explanation. No markdown. No code fences. JSON object only."""

RESIDENCY_MULTI_SYSTEM_PROMPT = """You are analyzing a multi-page PDF that was batch-scanned by a film production crew coordinator. It contains one or more separate people's identification documents -- each is a Driver's License, a State ID card, or a USCIS Form I-9 (Employment Eligibility Verification). A single person's card is sometimes 2 consecutive pages (front then back -- the back has no photo, just barcodes/restrictions/small print) and sometimes just 1 page (front only, no back scanned). Different people's documents never share a page.

Look at every page, in the order given, and group them into documents -- one group per person. For each document, report which pages belong to it and extract its fields.

For each document:
- pages: a list of the page numbers belonging to this document (1-indexed, in the order given -- e.g. [1, 2] for a front+back pair, or [3] for a front-only page).
- document_type: "drivers_license" / "state_id" / "i9" / "unknown".
- last_name: the person's last/family name, exactly as printed, title case. If a generational suffix (Jr, Sr, II, III, 3rd, etc.) is printed as part of that same name, keep it attached to last_name -- never treat the suffix as first_name or as the whole last_name by itself.
- first_name: the person's FIRST/given name ONLY -- never include a middle name.
- state: ONLY for drivers_license or state_id, the 2-letter issuing state abbreviation exactly as printed on the card (e.g. "IL", "CA"). Empty string for i9 or unknown.
- expiration_date: ONLY for drivers_license or state_id, in exactly MM/DD/YYYY as printed next to "EXP" or field "4b". Empty string for i9 or unknown.

Every page must belong to exactly one document's "pages" list -- never omit a page and never assign one page to two documents.

Return ONLY a JSON object with exactly this key: {"documents": [{"pages": [...], "document_type": "...", "last_name": "...", "first_name": "...", "state": "...", "expiration_date": "..."}, ...]}
No explanation. No markdown. No code fences. JSON object only."""

DIVERSITY_SYSTEM_PROMPT = """You are analyzing an Illinois Department of Commerce & Economic Opportunity (DCEO) "Illinois Film Tax Credit Tracking Sheet" submitted by a film production crew member.

Extract:
- employee_name_field: the exact text written in the "Employee Name" field, verbatim -- even if it looks like a business/loan-out company name rather than an individual person.
- looks_like_person: true if employee_name_field is clearly one individual's name (e.g. "First Last"), false if it looks like a business name (contains words like LLC, Inc, Corp, dba, Productions, Company, etc.), is blank, or is illegible.
- signature_name_guess: your best-effort reading of the person's name from the handwritten cursive signature near the "Signature" line, as First and Last name. Handwriting is often imperfect -- give your best guess. Empty string only if truly illegible/blank.

Return ONLY a JSON object with exactly these keys: {"employee_name_field": "...", "looks_like_person": true or false, "signature_name_guess": "..."}
No explanation. No markdown. No code fences. JSON object only."""

VENDOR_SYSTEM_PROMPT_TEMPLATE = """You are analyzing a vendor document submitted as part of a film/TV production's Accounts Payable packet. It is one of: a hotel guest folio, a purchase receipt/order confirmation, or a formal vendor invoice.
{prodco_line}
Some of these PDFs begin with an internal Purchase Order cover sheet (a page showing "Purchase Order: PO-XXXXXXX", a Subsidiary line, and a Vendor/Accounts/Description/Amount summary) before the vendor's own actual invoice/receipt/folio appears on a later page. When this cover sheet is present, completely ignore its "Purchase Order" number and its summary line items for every field below -- always read the ACTUAL vendor document (the page(s) that look like a real invoice, receipt, or folio) for every field, never the internal PO cover sheet.

FIRST, determine which of these three this document is:

1. HOTEL FOLIO -- a hotel guest bill/folio: shows a hotel name, Room Number, Arrival/Departure dates, and a running list of room/tax charges for one guest's stay. Set is_hotel_folio=true.

2. RECEIPT -- a purchase receipt or order confirmation for a straightforward purchase (a restaurant/catering order, a retail purchase, a delivery confirmation email) that has NO formal invoice structure (no "Bill To" block, no invoice number). Set is_receipt=true (and is_hotel_folio=false).

3. VENDOR INVOICE -- a formal invoice from a business or individual billing the production, with a "Bill To"/client block and (usually) its own invoice number. Set both is_hotel_folio=false and is_receipt=false.

Then extract the fields relevant to whichever type you picked:

FOR A HOTEL FOLIO ONLY:
- hotel_name: the hotel's name as printed. Required.
- folio_number: the number printed specifically next to a "Folio No." / "Folio #" label. Empty string if that field is blank or not shown -- do NOT substitute a confirmation number, reservation number, or room number instead, even if one is visible nearby.

FOR A RECEIPT OR VENDOR INVOICE, identify the BILLER -- the individual and/or company actually issuing this document and being paid, from its own header/logo/contact-info area. NEVER use the "Bill To"/client/production-company block for these fields, no matter what it's labeled, and never use the internal PO cover sheet's "Subsidiary" or "Vendor" fields if a more specific header appears on the actual document itself.
- has_person_name: true ONLY if an individual person's own name is presented as who this document is FROM (e.g. it's headed with a person's own name, made out personally, or says "make checks payable to [person]"). false otherwise.
- last_name / first_name: that person's name, split, ONLY if has_person_name is true. Empty strings otherwise. A generational suffix (Jr, Sr, II, III, etc.) stays attached to last_name.
- has_company_name: true if a business name is shown as the billing entity/merchant.
- company_name: that company's name, exactly as printed, ONLY if has_company_name is true. Empty string otherwise.
- bill_to_name: whoever this document is addressed TO -- the "Bill To" / client / recipient name, exactly as printed, regardless of who that turns out to be. This is the party being billed, never the biller. Empty string if no Bill To block is present (common on a receipt).

FOR A VENDOR INVOICE ONLY (skip these two for a receipt):
- invoice_number: the vendor's OWN invoice number, exactly as printed on the actual invoice itself (e.g. "1080" from "Invoice no.: 1080", "CHIC-26-000208" from "Invoice CHIC-26-000208"). This is never the internal "PO-XXXXXXX" purchase order number from a cover sheet. Empty string if the actual invoice has no invoice number printed anywhere.
- has_labor_line_item: true ONLY if the invoice shows a distinct billed line item whose own printed label literally is (or very closely matches) "Labor", "Delivery Labor", "Install Labor", "Labor Fee", or equivalent wording for a labor/delivery-labor charge. This is a narrow, literal check -- a flat day-rate professional service fee (e.g. "Casting Session $1,500/day") or an hourly service charge described some other way (e.g. security guard hours billed by shift) does NOT count unless a line is actually labeled "Labor" (or equivalent). false otherwise.

Return ONLY a JSON object with exactly these keys: {"is_hotel_folio": true or false, "hotel_name": "...", "folio_number": "...", "is_receipt": true or false, "has_person_name": true or false, "last_name": "...", "first_name": "...", "has_company_name": true or false, "company_name": "...", "bill_to_name": "...", "invoice_number": "...", "has_labor_line_item": true or false}
No explanation. No markdown. No code fences. JSON object only."""


def build_vendor_system_prompt(batch: BatchContext) -> str:
    """Inserts an explicit "this is the client, never the biller" hint for
    every named entity in the batch context -- without it, a ProdCo/Agency/
    Client name that happens to look distinctive can get mistaken for the
    vendor itself, especially on a receipt or a folio with no clear Bill To
    block to signal "this is the client, not the biller." This is also the
    mechanism the older tool lacked, which is why it sometimes renamed an
    invoice after who was being billed instead of who was doing the billing."""
    names = [
        n for n in (batch.client_name, batch.agency_name, batch.prodco_name)
        if n and n.strip()
    ]
    if not names:
        prodco_line = ""
    else:
        joined = ", ".join(f'"{n.strip()}"' for n in names)
        prodco_line = (
            f'\nThe production company / agency / client on this job may appear on these documents as '
            f'{joined} (or an LLC/dba variant of one of these names, e.g. "{names[0]} Productions, LLC"). '
            f'These are ALWAYS potential CLIENT / "Bill To" parties -- never the vendor issuing the document, '
            f'no matter where a name appears on the page or how prominently. Never let any of them become '
            f'has_person_name/last_name/first_name, has_company_name/company_name, or hotel_name below.\n'
        )
    return VENDOR_SYSTEM_PROMPT_TEMPLATE.replace("{prodco_line}", prodco_line)


# ── PDF / image / vision helpers ──────────────────────────────────────────────

try:
    import pillow_heif
    pillow_heif.register_heif_opener()  # lets Image.open() read iPhone HEIC/HEIF photos too
except ImportError:
    pass


def _flatten_to_rgb(img):
    """PDF has no real transparency model -- a transparent PNG/GIF/etc.
    composited straight in would render with an undefined (often black)
    background in some viewers, so anything with real alpha gets flattened
    onto white first. Everything else just gets a normal mode conversion."""
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        return background
    if img.mode != "RGB":
        return img.convert("RGB")
    return img


def load_source_as_pdf_bytes(raw_bytes: bytes, filename: str) -> bytes:
    """Returns valid single-document PDF bytes regardless of source format.
    A PDF (by extension or by content, in case someone renamed/mislabeled
    it) passes through untouched. Anything else is handed to Pillow as an
    image -- covers PNG/JPG same as before, plus BMP/GIF/TIFF/WEBP/HEIC-HEIF
    (iPhone photos) and effectively any other format Pillow understands,
    regardless of the file's own extension. A multi-page/multi-frame image
    (e.g. a scanner's multi-page TIFF) becomes a multi-page PDF, feeding the
    same downstream pipeline (including Residency's page-splitting) as a
    native multi-page PDF would. Raises if the file is neither a valid PDF
    nor an image Pillow can open -- the caller already treats that as an
    unreadable file, so no separate extension whitelist is needed here."""
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf" or raw_bytes[:5] == b"%PDF-":
        return raw_bytes

    try:
        img = Image.open(io.BytesIO(raw_bytes))
    except Exception:
        # PIL's own exception message embeds a raw BytesIO object repr,
        # which is meaningless in a user-facing reason_detail -- replace it
        # with a clean one instead of letting that string surface.
        raise ValueError(f"not a PDF and not a recognizable image format ({ext or 'no extension'})")
    n_frames = getattr(img, "n_frames", 1)
    frames = []
    for i in range(n_frames):
        img.seek(i)
        # exif_transpose bakes in a phone photo's stored orientation flag
        # before it ever reaches the vision-based rotation check, so that
        # check only has to catch genuine scanner/photo mistakes, not every
        # normally-oriented phone photo.
        frame = ImageOps.exif_transpose(img.copy())
        frames.append(_flatten_to_rgb(frame))

    buf = io.BytesIO()
    if len(frames) == 1:
        frames[0].save(buf, format="PDF")
    else:
        frames[0].save(buf, format="PDF", save_all=True, append_images=frames[1:])
    return buf.getvalue()


def render_pdf_bytes_to_images_b64(pdf_bytes: bytes, dpi_scale: float = 2.0, max_pages: int = 3) -> list:
    """Renders up to max_pages (default 3, the widest need across the three
    starting types -- vendor documents sometimes lead with a PO cover sheet
    pushing the real invoice to page 2-3). Rendered once per file and reused
    for both the classify call and the type-specific extract call."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for i, page in enumerate(doc):
        if max_pages and i >= max_pages:
            break
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale))
        images.append(base64.b64encode(pix.tobytes("png")).decode())
    doc.close()
    return images


MAX_PAGES_TO_INSPECT = 30


def extract_pages_as_pdf(pdf_bytes: bytes, page_numbers_1indexed: list) -> bytes:
    """Slices specific pages (1-indexed, in the given order) out of a PDF
    into a new standalone PDF -- used when one uploaded file contains
    multiple people's residency documents (e.g. a batch-scanned stack of
    licenses), so each person ends up as their own output file."""
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()
    for pn in page_numbers_1indexed:
        out.insert_pdf(src, from_page=pn - 1, to_page=pn - 1)
    result = out.tobytes()
    out.close()
    src.close()
    return result


def looks_like_refusal(raw: str) -> bool:
    s = raw.strip()
    if not s:
        return True
    if "{" in s[:20] or s.startswith("```"):
        return False
    refusal_openers = ("i'm sorry", "i am sorry", "i cannot", "i can't", "sorry,", "i apologize")
    return s.lower().startswith(refusal_openers) or "{" not in s


def parse_json_object(raw: str):
    try:
        return json.loads(raw)
    except Exception:
        pass
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if m:
        try:
            return json.loads(m.group())
        except Exception:
            pass
    return None


def call_gpt_json(client, images_b64, system_prompt, user_text, max_tokens=500, detail="high"):
    content = [{"type": "text", "text": user_text}]
    for img in images_b64:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{img}", "detail": detail},
        })
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content},
        ],
        temperature=0,
        max_tokens=max_tokens,
    )
    return resp.choices[0].message.content.strip()


def call_claude_json(client, images_b64, system_prompt, user_text, max_tokens=500):
    content = []
    for img in images_b64:
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": img},
        })
    content.append({"type": "text", "text": user_text})
    resp = client.messages.create(
        model="claude-sonnet-5",
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": "user", "content": content}],
    )
    text_block = next((b for b in resp.content if b.type == "text"), None)
    return text_block.text.strip() if text_block else ""


async def classify_document(images_b64, user_text, system_prompt, openai_client, anthropic_client, loop, detail="high"):
    """Tries GPT-4o first; falls back to Claude if GPT refuses/errors/comes
    back empty. Returns (parsed_dict_or_None, provider_used). Runs the
    (synchronous) SDK calls in the executor so this can be awaited alongside
    other files' calls under a semaphore. detail="low" is worth passing for
    a coarse judgment call (e.g. orientation) made over many images at
    once, where GPT-4o's high-detail mode would multiply token cost for no
    real accuracy benefit."""
    import functools
    try:
        raw = await loop.run_in_executor(
            None, functools.partial(call_gpt_json, openai_client, images_b64, system_prompt, user_text, detail=detail),
        )
        if looks_like_refusal(raw):
            raise ValueError("refusal")
        parsed = parse_json_object(raw)
        if parsed is None:
            raise ValueError("unparseable")
        return parsed, "openai"
    except Exception:
        if anthropic_client is None:
            return None, "openai"
        try:
            raw = await loop.run_in_executor(
                None, functools.partial(call_claude_json, anthropic_client, images_b64, system_prompt, user_text),
            )
            parsed = parse_json_object(raw)
            return (parsed, "anthropic") if parsed is not None else (None, "anthropic")
        except Exception:
            return None, "anthropic"


async def correct_page_orientations(pdf_bytes: bytes, openai_client, anthropic_client, loop) -> bytes:
    """Detects and fixes any page that's sideways or upside-down (a phone
    photo taken in the wrong orientation, or a scanner fed the page wrong)
    before any classification/extraction happens -- everything downstream
    just sees an upright document, no matter how it was actually scanned.

    Sets the PDF's own per-page /Rotate attribute rather than baking in a
    pixel-level transform: every standard PDF viewer honors it, and so does
    our own get_pixmap() rendering used everywhere else in this file, so
    the fix propagates for free through classification, extraction, and
    (for Residency) the page-splitting step -- none of that code needs to
    know this ever happened."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = min(doc.page_count, MAX_PAGES_TO_INSPECT)
    if n == 0:
        doc.close()
        return pdf_bytes

    images = []
    for i in range(n):
        pix = doc[i].get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        images.append(base64.b64encode(pix.tobytes("png")).decode())

    user_text = "Determine the needed rotation for each page, in order, as described in the system prompt."
    parsed, _provider = await classify_document(
        images, user_text, ORIENTATION_PROMPT, openai_client, anthropic_client, loop, detail="low",
    )
    pages_info = (parsed or {}).get("pages") if parsed else None
    if not isinstance(pages_info, list):
        doc.close()
        return pdf_bytes

    changed = False
    for i, info in enumerate(pages_info[:n]):
        if not isinstance(info, dict):
            continue
        try:
            rotation = int(info.get("rotation", 0))
        except (TypeError, ValueError):
            rotation = 0
        if rotation in (90, 180, 270):
            page = doc[i]
            page.set_rotation((page.rotation + rotation) % 360)
            changed = True

    result = doc.tobytes() if changed else pdf_bytes
    doc.close()
    return result


# ── Name / filename helpers ───────────────────────────────────────────────────

def sanitize_component(s: str) -> str:
    s = INVALID_FILENAME_CHARS.sub("", s or "").strip()
    return s.rstrip(". ")


def proper_case(s: str) -> str:
    """Word-by-word Proper/Title Case, with two exceptions: LLC and INC
    (however punctuated) always render fully uppercase rather than
    title-cased, and a single letter directly after an apostrophe stays
    lowercase (a possessive/contraction suffix -- "Joe's", "McDonald's" --
    rather than a second name part, unlike "O'Brien"/"D'Angelo" where the
    run after the apostrophe is longer and does get title-cased).
    Capitalizes each contiguous run of letters independently:
    "REAGAN-BARNES" -> "Reagan-Barnes", "O'BRIEN" -> "O'Brien",
    "JOE'S BAKERY" -> "Joe's Bakery"."""
    def cap_word(word):
        core = re.sub(r"[^a-zA-Z]", "", word).lower()
        if core in _UPPERCASE_EXCEPTIONS:
            return word.upper()

        def cap_run(m):
            run = m.group(0)
            preceded_by_apostrophe = m.start() > 0 and word[m.start() - 1] == "'"
            if preceded_by_apostrophe and len(run) == 1:
                return run.lower()
            return run[:1].upper() + run[1:].lower()

        return re.sub(r"[A-Za-z]+", cap_run, word)

    s = (s or "").strip()
    return " ".join(cap_word(w) for w in s.split(" ")) if s else s


def tokenize(s: str) -> set:
    return set(re.findall(r"[a-z]+", (s or "").lower()))


def names_overlap(a: str, b: str) -> bool:
    ta, tb = tokenize(a), tokenize(b)
    return bool(ta and tb and (ta & tb))


def name_looks_like_person(text: str) -> bool:
    if not text or not text.strip():
        return False
    words = re.findall(r"[a-zA-Z']+", text.lower())
    if not words or any(w in _BUSINESS_WORDS for w in words):
        return False
    return 1 <= len(words) <= 4


def reformat_date_for_filename(mmddyyyy: str) -> str:
    """MM/DD/YYYY -> YYYY_MM_DD, for the Driver's License / State ID
    filename pattern: "Last, First (ST) YYYY_MM_DD.pdf"."""
    s = (mmddyyyy or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        mm, dd, yyyy = m.groups()
        return f"{yyyy}_{int(mm):02d}_{int(dd):02d}"
    return s.replace("/", "_")


def split_person_name(full_name: str):
    """'First Last' (no comma) -> (last, first); 'Last, First' -> (last, first).
    A trailing generational suffix stays attached to the last name."""
    full_name = (full_name or "").strip()
    if "," in full_name:
        last, _, first = full_name.partition(",")
        return last.strip(), first.strip()
    parts = full_name.split()
    if len(parts) >= 3 and parts[-1].strip(".").lower() in _NAME_SUFFIXES:
        return f"{parts[-2]} {parts[-1]}", " ".join(parts[:-2])
    if len(parts) >= 2:
        return parts[-1], " ".join(parts[:-1])
    return full_name, ""


def build_residency_filename(doc_type: str, last: str, first: str, state: str, exp_date: str):
    last_s, first_s = sanitize_component(proper_case(last)), sanitize_component(proper_case(first))
    if not last_s or not first_s:
        return None
    if doc_type in ("drivers_license", "state_id"):
        state_s = sanitize_component((state or "").strip().upper())
        exp_s = sanitize_component(reformat_date_for_filename(exp_date))
        if not state_s or not exp_s:
            return None
        return f"{last_s}, {first_s} ({state_s}) {exp_s}.pdf"
    if doc_type == "i9":
        return f"{last_s}, {first_s} - I9.pdf"
    return None


def build_diversity_filename(last: str, first: str):
    last_s, first_s = sanitize_component(proper_case(last)), sanitize_component(proper_case(first))
    if not last_s or not first_s:
        return None
    return f"{last_s}, {first_s} - Diversity Form.pdf"


def build_freelance_filename(has_person: bool, last: str, first: str, has_company: bool, company: str, invoice_number: str):
    last_s = sanitize_component(proper_case(last)) if has_person else ""
    first_s = sanitize_component(proper_case(first)) if has_person else ""
    company_s = sanitize_component(proper_case(company)) if has_company else ""
    num_s = sanitize_component((invoice_number or "").strip())
    invoice_suffix = f"Invoice {num_s}" if num_s else "Invoice"

    if has_person and not (last_s and first_s):
        return None
    if has_company and not company_s:
        return None
    if has_person and has_company:
        return f"{last_s}, {first_s} ({company_s}) - {invoice_suffix}.pdf"
    if has_person:
        return f"{last_s}, {first_s} - {invoice_suffix}.pdf"
    if has_company:
        return f"{company_s} - {invoice_suffix}.pdf"
    return None


def build_vendor_invoice_filename(company: str, invoice_number: str):
    company_s = sanitize_component(proper_case(company))
    if not company_s:
        return None
    num_s = sanitize_component((invoice_number or "").strip())
    invoice_suffix = f"Invoice {num_s}" if num_s else "Invoice"
    return f"{company_s} - {invoice_suffix}.pdf"


def build_vendor_receipt_filename(company: str):
    company_s = sanitize_component(proper_case(company))
    return f"{company_s} - receipt.pdf" if company_s else None


def build_hotel_folio_filename(hotel_name: str, folio_number: str):
    hotel_s = sanitize_component(proper_case(hotel_name))
    if not hotel_s:
        return None
    num_s = sanitize_component((folio_number or "").strip())
    suffix = f"Folio {num_s}" if num_s else "Folio"
    return f"{hotel_s} - {suffix}.pdf"


# ── Sender-mismatch detection (Vendor only) ───────────────────────────────────

def _expected_biller_name(batch: BatchContext) -> str:
    return {
        "prodco": batch.prodco_name,
        "agency": batch.agency_name,
        "client": batch.client_name,
    }.get(batch.received_from, "")


def check_sender_mismatch(bill_to_name: str, batch: BatchContext) -> str:
    """Returns a human-readable mismatch detail if bill_to_name confidently
    matches a DIFFERENT named entity than received_from implies, else "".
    Only flags a confident mismatch against one of the batch's OTHER two
    named entities -- never merely "didn't match the expected one", since
    that alone is too weak a signal (could just be phrased differently)."""
    bill_to_name = (bill_to_name or "").strip()
    if not bill_to_name or not batch.received_from:
        return ""
    expected = _expected_biller_name(batch)
    if expected and names_overlap(bill_to_name, expected):
        return ""
    others = {"client": batch.client_name, "agency": batch.agency_name, "prodco": batch.prodco_name}
    for key, name in others.items():
        if key == batch.received_from or not name:
            continue
        if names_overlap(bill_to_name, name):
            return (
                f"Bill-To reads '{bill_to_name}', but batch says received_from={batch.received_from} "
                f"({expected or 'not provided'})"
            )
    return ""


# ── Per-family extraction ─────────────────────────────────────────────────────

async def _extract_residency_multi(pdf_bytes, openai_client, anthropic_client, loop):
    """Residency is the one family that can't assume "one uploaded file = one
    person" -- a coordinator batch-scanning a stack of physical IDs produces a
    single multi-page PDF containing several different people's cards, each
    spanning 1 page (front only) or 2 (front+back). This renders every page
    (up to MAX_PAGES_TO_INSPECT) and asks the model to group them by person in
    one call, then physically splits the PDF so each person becomes their own
    output file -- one uploaded file can produce many DocResults here, unlike
    every other family."""
    images = render_pdf_bytes_to_images_b64(pdf_bytes, max_pages=MAX_PAGES_TO_INSPECT)
    if not images:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
            reason_detail="No renderable pages found",
        )]

    user_text = (
        "Identify every person's ID document across all pages of this PDF and group "
        "pages accordingly, as described in the system prompt."
    )
    parsed, provider = await classify_document(
        images, user_text, RESIDENCY_MULTI_SYSTEM_PROMPT, openai_client, anthropic_client, loop,
    )
    documents = (parsed or {}).get("documents") if parsed else None
    if not isinstance(documents, list) or not documents:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )]

    total_pages = len(images)
    results = []
    for doc in documents:
        page_nums = [p for p in (doc.get("pages") or []) if isinstance(p, int) and 1 <= p <= total_pages]
        if not page_nums:
            continue  # nothing usable to split out for this group

        try:
            sub_pdf_bytes = extract_pages_as_pdf(pdf_bytes, page_nums)
        except Exception as e:
            results.append(DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
                reason_detail=f"Could not split page(s) {page_nums} into their own file: {e}", provider=provider,
            ))
            continue

        doc_type = str(doc.get("document_type", "unknown")).strip().lower()
        if doc_type not in ("drivers_license", "state_id", "i9"):
            results.append(DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=sub_pdf_bytes, reason_code=REASON_UNCLASSIFIED,
                reason_detail=f"Unrecognized document type ({doc_type!r}) on page(s) {page_nums}", provider=provider,
            ))
            continue

        last, first = doc.get("last_name", ""), doc.get("first_name", "")
        state = doc.get("state", "")
        exp_date = doc.get("expiration_date", "")
        new_name = build_residency_filename(doc_type, last, first, state, exp_date)
        if not new_name:
            if not last or not first:
                missing = REASON_MISSING_NAME
            elif doc_type in ("drivers_license", "state_id") and not state:
                missing = REASON_MISSING_STATE
            else:
                missing = REASON_MISSING_DATE
            results.append(DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="residency", subfolder=FOLDER_RESIDENCY,
                output_pdf_bytes=sub_pdf_bytes, reason_code=missing,
                reason_detail=(f"Missing/unreadable name, state, or expiration date on page(s) {page_nums} "
                                f"(type={doc_type}, last={last!r}, first={first!r}, state={state!r}, exp={exp_date!r})"),
                provider=provider,
            ))
            continue

        results.append(DocResult(
            bucket=BUCKET_RENAMED, doc_type="residency", subfolder=FOLDER_RESIDENCY,
            output_pdf_bytes=sub_pdf_bytes, new_filename=new_name, provider=provider,
        ))

    if not results:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNCLASSIFIED,
            reason_detail="Model returned no usable page groupings", provider=provider,
        )]
    return results


async def _extract_diversity(images_b64, openai_client, anthropic_client, loop, orig_filename_hint):
    user_text = "Extract the fields described in the system prompt from this DCEO tracking sheet."
    parsed, provider = await classify_document(
        images_b64, user_text, DIVERSITY_SYSTEM_PROMPT, openai_client, anthropic_client, loop,
    )
    if parsed is None:
        return DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="diversity", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=b"", reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )

    employee_name = str(parsed.get("employee_name_field", "")).strip()
    looks_like_person = bool(parsed.get("looks_like_person")) and name_looks_like_person(employee_name)
    signature_guess = str(parsed.get("signature_name_guess", "")).strip()

    if looks_like_person:
        last, first = split_person_name(employee_name)
    else:
        filename_guess = _guess_name_from_filename(orig_filename_hint)
        if signature_guess and names_overlap(signature_guess, filename_guess):
            last, first = split_person_name(signature_guess)
        else:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="diversity", subfolder=FOLDER_DIVERSITY,
                output_pdf_bytes=b"", reason_code=REASON_LOW_CONFIDENCE,
                reason_detail=(f"Employee Name field looks like a company ({employee_name!r}) and signature "
                               f"({signature_guess!r}) didn't confirm the existing filename"),
                provider=provider,
            )

    new_name = build_diversity_filename(last, first)
    if not new_name:
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="diversity", subfolder=FOLDER_DIVERSITY,
            output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
            reason_detail=f"Could not derive a usable name (last={last!r}, first={first!r})", provider=provider,
        )
    return DocResult(
        bucket=BUCKET_RENAMED, doc_type="diversity", subfolder=FOLDER_DIVERSITY,
        output_pdf_bytes=b"", new_filename=new_name, provider=provider,
    )


async def _extract_vendor(images_b64, openai_client, anthropic_client, loop, batch: BatchContext):
    user_text = "Extract the fields described in the system prompt from this vendor document."
    system_prompt = build_vendor_system_prompt(batch)
    parsed, provider = await classify_document(
        images_b64, user_text, system_prompt, openai_client, anthropic_client, loop,
    )
    if parsed is None:
        return DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="vendor", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=b"", reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )

    is_hotel_folio = bool(parsed.get("is_hotel_folio"))
    is_receipt = bool(parsed.get("is_receipt"))
    has_labor = bool(parsed.get("has_labor_line_item"))
    hotel_name = parsed.get("hotel_name", "")
    folio_number = parsed.get("folio_number", "")
    has_person = bool(parsed.get("has_person_name"))
    has_company = bool(parsed.get("has_company_name"))
    last, first = parsed.get("last_name", ""), parsed.get("first_name", "")
    company = parsed.get("company_name", "")
    bill_to_name = parsed.get("bill_to_name", "")
    invoice_number = parsed.get("invoice_number", "")

    mismatch_detail = check_sender_mismatch(bill_to_name, batch)

    if is_hotel_folio:
        new_name = build_hotel_folio_filename(hotel_name, folio_number)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_COMPANY,
                reason_detail=f"Hotel folio but missing hotel name (hotel_name={hotel_name!r})", provider=provider,
            )
    elif is_receipt:
        new_name = build_vendor_receipt_filename(company)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_COMPANY,
                reason_detail=f"Receipt but missing company name (company={company!r})", provider=provider,
            )
    elif has_labor:
        if not has_person and not has_company:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
                reason_detail="Labor line item found but could not identify a billing person or company", provider=provider,
            )
        new_name = build_freelance_filename(has_person, last, first, has_company, company, invoice_number)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
                reason_detail=(f"Labor line item found but missing usable name/company (last={last!r}, "
                               f"first={first!r}, company={company!r})"),
                provider=provider,
            )
    else:
        new_name = build_vendor_invoice_filename(company, invoice_number)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_COMPANY,
                reason_detail=f"Missing company name (company={company!r})", provider=provider,
            )

    if mismatch_detail:
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
            output_pdf_bytes=b"", reason_code=REASON_SENDER_MISMATCH,
            reason_detail=mismatch_detail, provider=provider,
        )

    return DocResult(
        bucket=BUCKET_RENAMED, doc_type="vendor", subfolder=FOLDER_VENDOR,
        output_pdf_bytes=b"", new_filename=new_name, provider=provider,
    )


def _guess_name_from_filename(filename: str) -> str:
    base = os.path.splitext(filename)[0]
    base = re.sub(r"\([^)]*\)", " ", base)
    base = re.sub(r"_.*$", "", base)
    base = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", base)
    base = base.replace(",", " ")
    return re.sub(r"\s+", " ", base).strip()


# ── Top-level dispatcher ──────────────────────────────────────────────────────

async def process_one_document(
    raw_bytes: bytes,
    filename: str,
    batch: BatchContext,
    openai_client,
    anthropic_client,
    loop,
) -> list:
    """The single entry point the job layer calls per uploaded file. Converts
    to PDF if needed (local, no LLM cost), classifies the document family
    from a quick look at the first few pages, then dispatches to that
    family's extraction. Any unexpected exception is caught here so one bad
    file never kills a job.

    Returns a LIST of DocResult, not one -- for every family except
    Residency this is always a single-element list (one uploaded file, one
    output file), but Residency can legitimately split one uploaded file
    into several (a batch-scanned stack of different people's IDs)."""
    try:
        pdf_bytes = load_source_as_pdf_bytes(raw_bytes, filename)
    except Exception as e:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="unknown", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=raw_bytes, reason_code=REASON_UNREADABLE,
            reason_detail=f"Could not open/convert file: {e}",
        )]

    try:
        # Fix orientation FIRST, before anything else looks at this file --
        # a sideways/upside-down scan (phone photo, misfed scanner page)
        # would otherwise confuse classification/extraction too. A failure
        # here is never fatal to the rest of the pipeline; just proceed with
        # whatever orientation the file already had.
        pdf_bytes = await correct_page_orientations(pdf_bytes, openai_client, anthropic_client, loop)
    except Exception:
        pass

    try:
        # A quick, cheap look (first few pages) is enough to tell the family
        # -- Residency's own extraction re-renders up to MAX_PAGES_TO_INSPECT
        # separately once we know that's what we're dealing with, since a
        # batch-scanned stack can run well past this classification cap.
        classify_images = render_pdf_bytes_to_images_b64(pdf_bytes, max_pages=3)
        if not classify_images:
            return [DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="unknown", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
                reason_detail="No renderable pages found",
            )]

        parsed, provider = await classify_document(
            classify_images, "Identify the broad family of this document as described in the system prompt.",
            TYPE_CLASSIFIER_PROMPT, openai_client, anthropic_client, loop,
        )
        family = str((parsed or {}).get("document_family", "unknown")).strip().lower()

        if family == "residency":
            results = await _extract_residency_multi(pdf_bytes, openai_client, anthropic_client, loop)
        elif family == "diversity":
            result = await _extract_diversity(classify_images, openai_client, anthropic_client, loop, filename)
            result.output_pdf_bytes = pdf_bytes
            results = [result]
        elif family == "vendor":
            result = await _extract_vendor(classify_images, openai_client, anthropic_client, loop, batch)
            result.output_pdf_bytes = pdf_bytes
            results = [result]
        else:
            results = [DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="unknown", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=pdf_bytes, reason_code=REASON_UNCLASSIFIED,
                reason_detail="Could not confidently classify this document into a known type" if parsed is not None
                              else "Could not read document (both providers failed)",
                provider=provider,
            )]
        return results
    except Exception as e:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="unknown", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_PROVIDER_ERROR,
            reason_detail=f"Unexpected error: {e}",
        )]


# ── Naming conventions, exposed as data (GET /wizard01/conventions) ──────────

NAMING_CONVENTIONS = [
    {
        "type_id": "residency",
        "label": "Residency (Driver's License / State ID / I-9)",
        "patterns": [
            {"pattern": "{Last}, {First} ({State}) {YYYY_MM_DD}.pdf", "description": "Driver's License / State ID, named by issuing state + expiration date"},
            {"pattern": "{Last}, {First} - I9.pdf", "description": "USCIS Form I-9"},
        ],
        "notes": (
            "If one uploaded PDF contains several different people's ID documents "
            "(a batch-scanned stack), each person is split out into their own "
            "renamed output file -- one upload can produce multiple results here."
        ),
    },
    {
        "type_id": "diversity",
        "label": "Diversity Form",
        "patterns": [
            {"pattern": "{Last}, {First} - Diversity Form.pdf", "description": "Illinois DCEO Film Tax Credit Tracking Sheet"},
        ],
    },
    {
        "type_id": "vendor",
        "label": "Vendor (Invoice / Receipt / Hotel Folio)",
        "patterns": [
            {"pattern": "{Company} - Invoice {Number}.pdf", "description": "Company invoice (number omitted if not printed)"},
            {"pattern": "{Company} - receipt.pdf", "description": "Purchase receipt / order confirmation"},
            {"pattern": "{Hotel Name} - Folio {Number}.pdf", "description": "Hotel guest folio (number omitted if blank)"},
            {"pattern": "{Last}, {First} - Invoice {Number}.pdf", "description": "Invoice with a Labor/Delivery Labor line item, billed by a person"},
            {"pattern": "{Last}, {First} ({Company}) - Invoice {Number}.pdf", "description": "Same, when both a person and a company are named"},
        ],
    },
]
