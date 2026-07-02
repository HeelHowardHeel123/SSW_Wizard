"""
Wrapbook Fringe Report parser — coordinate-based (pdfplumber extract_words).

Handles two header format variants automatically:
  Format 1 (bundled per-invoice PDF):  first col = "Name",   state cols = "Work State" / "Res State"
  Format 2 (project-level fringe PDF): first col = "Worker", state cols = "Work" / "Res"
                                       with a "State  State" sub-line on the next visual row

Falls back to GPT-4o vision for image-only PDFs (no text layer detected).
"""

import re
import io
import json
import os

import pdfplumber
from openai import OpenAI

from parsers.base import empty_row, parse_amount, pdf_to_images_b64, has_text_layer
from parsers.wrapbook.register import extract_register

# ─── Registry metadata ────────────────────────────────────────────────────────
COMPANY  = "wrapbook"
MARKERS  = ["Fringe Report"]
PRIORITY = 10

# ─── Column label definitions ─────────────────────────────────────────────────
# Each entry: (header_text, canonical_field_name)
# Multi-token labels are matched as consecutive words in the header line.
# "Worker"/"Name" are both tried for the first column (handled in _detect_columns).

_LABELS = [
    ("Worker",     "worker"),
    ("SSN",        "ssn"),
    ("Work Dates", "workDates"),
    ("Type",       "type"),
    ("Dept",       "dept"),
    ("Union",      "union"),
    ("Work",       "workState"),   # Format 1: "Work State"; Format 2: "Work"
    ("Res",        "resState"),    # Format 1: "Res State";  Format 2: "Res"
    ("Wages",      "wages"),
    ("Reimb/Rent", "reimbRent"),
    ("Soc Sec",    "socSec"),
    ("Med",        "med"),
    ("FUTA",       "futa"),
    ("SUI",        "sui"),
    ("W/C",        "wc"),
    ("PH&W",       "phw"),
    ("Other",      "other"),
    ("Benefits",   "benefits"),
    ("Plat Fee",   "platFee"),
    ("EFF Rate %", "effRate"),
    ("Total",      "total"),
]

_NUMERIC_FIELDS = {
    "wages", "reimbRent", "socSec", "med", "futa", "sui", "wc", "phw",
    "other", "benefits", "platFee", "effRate", "total",
}

_MONEY_RE = re.compile(r"\d\.\d{2}")


# ─── Coordinate utilities ─────────────────────────────────────────────────────

def _cluster_lines(words, tol=3.0):
    """Cluster extract_words() output into visual lines sorted top→down, left→right."""
    ws = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    result = []
    cur = None
    for w in ws:
        if cur is None or abs(cur["top"] - w["top"]) > tol:
            cur = {"top": w["top"], "words": [w]}
            result.append(cur)
        else:
            cur["words"].append(w)
    for ln in result:
        ln["words"].sort(key=lambda w: w["x0"])
    return result


def _detect_columns(all_lines):
    """Find the header line and build midpoint x-boundaries for each column.

    Accepts 'Worker' (Format 2) or 'Name' (Format 1) as the first column token.
    Returns None if no valid header is found.
    """
    hdr = None
    for ln in all_lines:
        texts = [w["text"] for w in ln["words"]]
        has_first_col = "Worker" in texts or "Name" in texts
        if has_first_col and "Total" in texts:
            hdr = ln
            break
    if hdr is None:
        return None

    words = hdr["words"]
    texts = [w["text"] for w in words]

    # Swap first label to "Name" if that's what the PDF uses
    labels = list(_LABELS)
    if "Name" in texts and "Worker" not in texts:
        labels[0] = ("Name", "worker")

    anchors = []
    i = 0
    for label, field in labels:
        toks = label.split()
        pos  = None
        for k in range(i, len(texts) - len(toks) + 1):
            if all(texts[k + j].lower() == toks[j].lower() for j in range(len(toks))):
                pos = k
                break
        if pos is not None:
            anchors.append({"field": field, "x0": words[pos]["x0"]})
            i = pos + len(toks)

    anchors.sort(key=lambda a: a["x0"])
    for idx, a in enumerate(anchors):
        a["left"]  = -1e9 if idx == 0           else (anchors[idx - 1]["x0"] + a["x0"]) / 2.0
        a["right"] =  1e9 if idx == len(anchors) - 1 else (a["x0"] + anchors[idx + 1]["x0"]) / 2.0

    return {"top": hdr["top"], "anchors": anchors}


def _col_for(cols, x):
    for a in cols["anchors"]:
        if a["left"] <= x < a["right"]:
            return a["field"]
    return None


# ─── Text fragment merging ────────────────────────────────────────────────────

def _merge_text(field, frags):
    """Reconstruct a text column value from word fragments that may span wrapped lines."""
    if field == "worker":
        # "Last, First M": last name may wrap mid-word (no hyphen), so join prefix
        # up to comma without a space; join given-name fragments with spaces;
        # drop trailing middle initial.
        ci = next((k for k, f in enumerate(frags) if "," in f), -1)
        if ci < 0:
            return re.sub(r"\s+", " ", " ".join(frags)).strip()
        comma_frag  = frags[ci]
        before      = comma_frag[: comma_frag.index(",") + 1]
        after       = comma_frag[comma_frag.index(",") + 1:].strip()
        last        = "".join(frags[:ci]) + before
        first_frags = ([after] if after else []) + frags[ci + 1:]
        first       = re.sub(r"\s+", " ", " ".join(first_frags)).strip()
        first       = re.sub(r"\s+[A-Z]\.?$", "", first).strip()
        return (last.rstrip(",") + ", " + first).strip()

    if field == "ssn":
        return "".join(frags)

    out = ""
    for idx, fr in enumerate(frags):
        if idx == 0:
            out = fr
        elif re.search(r"\S-$", out):   # hyphenated mid-word wrap
            out += fr
        else:
            out += " " + fr
    return re.sub(r"\s+", " ", out).strip()


# ─── Page row extraction ──────────────────────────────────────────────────────

def _rows_from_page(all_lines, cols, has_header):
    """Extract employee rows from one page of a Fringe Report."""
    header_top = cols["top"] if has_header else -1e9
    data = [
        ln for ln in all_lines
        if ln["top"] > header_top + 2
        and not re.search(r"Page \d+ of \d+", " ".join(w["text"] for w in ln["words"]))
    ]

    wages_anchor = next(a for a in cols["anchors"] if a["field"] == "wages")
    recs = []
    cur  = None
    for ln in data:
        # Primary row: has a money value at or to the right of the Wages column
        is_primary = any(
            w["x0"] >= wages_anchor["left"] and _MONEY_RE.search(w["text"])
            for w in ln["words"]
        )
        if is_primary:
            if cur is not None:
                recs.append(cur)
            cur = {}
        if cur is None:
            continue
        for w in ln["words"]:
            f = _col_for(cols, w["x0"])
            if f:
                cur.setdefault(f, []).append(w["text"])
    if cur is not None:
        recs.append(cur)

    rows = []
    for r in recs:
        row = empty_row()
        row["payrollCompany"] = "wrapbook"
        for _, field in _LABELS:
            frags = r.get(field, [])
            if field in _NUMERIC_FIELDS:
                row[field] = parse_amount(frags[0]) if frags else None
            else:
                row[field] = _merge_text(field, frags) if frags else ""

        # Skip grand-total / summary rows (no real worker name)
        w_val = row.get("worker", "")
        if not w_val or len(w_val) <= 1 or "GRAND" in w_val.upper() or "TOTAL" in w_val.upper():
            continue

        # Normalize work dates: "04/07/2026 04/09/2026" → "04/07/2026 - 04/09/2026"
        # (the "-" token sometimes falls outside the Work Dates column x-range)
        wd = row.get("workDates", "")
        if wd:
            row["workDates"] = re.sub(
                r"(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})", r"\1 - \2", wd
            )

        # Detect loan-out from Type column
        if "loan" in row.get("type", "").lower():
            row["loanOut"] = True

        rows.append(row)
    return rows


def _parse_meta(text):
    """Extract invoice #, date, and work-date range from a Fringe Report page header.

    Handles two formats:
      Format 1 (bundled per-invoice): "Invoice #00715136" + "Invoice date MM/DD/YYYY"
                                       + "Work dates Mon D, YYYY - Mon D, YYYY"
      Format 2 (project-level):       "NNNNNNNN MM/DD/YYYY Mon D, YYYY - Mon D, YYYY"
                                       (legacy inline format; no invoice # on project fringe)
    """
    invoice_no = invoice_date = work_dates = ""

    # Format 1 — labeled fields in page header
    m_inv = re.search(r"Invoice\s+#?0*(\d+)", text, re.IGNORECASE)
    if m_inv:
        invoice_no = m_inv.group(1)

    m_date = re.search(r"Invoice\s+date\s+(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    if m_date:
        invoice_date = m_date.group(1)

    m_wd = re.search(
        r"Work\s+dates?\s+([A-Z][a-z]{2} \d{1,2}, \d{4}\s*-\s*[A-Z][a-z]{2} \d{1,2}, \d{4})",
        text, re.IGNORECASE,
    )
    if m_wd:
        work_dates = re.sub(r"\s+", " ", m_wd.group(1))

    # Format 2 — inline "NNNNNN MM/DD/YYYY Mon D, YYYY - Mon D, YYYY" (NIS 007 style)
    if not invoice_no and not work_dates:
        m = re.search(
            r"(\d{6,8})\s+(\d{2}/\d{2}/\d{4})\s+"
            r"([A-Z][a-z]{2} \d{1,2}, \d{4}\s*-\s*[A-Z][a-z]{2} \d{1,2}, \d{4})",
            text,
        )
        if m:
            invoice_no, invoice_date = m.group(1), m.group(2)
            work_dates = re.sub(r"\s+", " ", m.group(3))

    return invoice_no, invoice_date, work_dates


# ─── Text-based extraction ────────────────────────────────────────────────────

def _extract_text(pdf_bytes: bytes) -> tuple[list[dict], list[str]]:
    rows, issues = [], []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n           = len(pdf.pages)
        page_words  = []
        page_text   = []
        for pg in pdf.pages:
            page_words.append(
                pg.extract_words(x_tolerance=1.5, keep_blank_chars=False, use_text_flow=False)
            )
            page_text.append(pg.extract_text() or "")

        found_any = False
        i = 0
        while i < n:
            if "Fringe Report" in page_text[i]:
                found_any = True
                inv_no, inv_date, wp = _parse_meta(page_text[i])
                m_span = re.search(r"Page \d+ of (\d+)", page_text[i])
                span   = int(m_span.group(1)) if m_span else 1
                cols   = _detect_columns(_cluster_lines(page_words[i]))
                if cols is None:
                    issues.append(
                        f"Page {i + 1}: could not read the fringe table header — "
                        "column positions not detected. Rows on this page were skipped."
                    )
                    i += span
                    continue
                for fp in range(i, min(i + span, n)):
                    for row in _rows_from_page(_cluster_lines(page_words[fp]), cols, fp == i):
                        row["invoiceNo"]        = inv_no
                        row["invoiceDate"]      = inv_date
                        row["invoiceWorkDates"] = wp
                        row["sourcePage"]       = fp + 1
                        rows.append(row)
                i += span
            else:
                i += 1

        if not found_any:
            issues.append(
                "No Fringe Report page found — verify this is a Wrapbook payroll PDF. "
                "If it contains only invoices or payroll register pages, it may not include fringe data."
            )

        # For project-level fringe (no invoice date in header), fall back to the first
        # "Check Date:" found anywhere in the PDF (typically on the invoice cover page).
        if rows and not rows[0].get("invoiceDate"):
            for pt in page_text:
                m = re.search(r"Check Date:\s*([A-Za-z]+ \d{1,2},\s*\d{4})", pt)
                if m:
                    check_date = m.group(1)
                    for row in rows:
                        if not row.get("invoiceDate"):
                            row["invoiceDate"] = check_date
                    break

        # Auto-enrich when the same PDF also contains a Payroll Register (GM 004 style)
        has_register = any("Payroll Register" in t for t in page_text)
        if has_register and rows:
            register_data = extract_register(pdf_bytes)
            if register_data:
                _enrich_rows(rows, register_data)

    return rows, issues


# ─── Vision fallback (image-only PDFs) ───────────────────────────────────────

_VISION_SYSTEM = """You are extracting data from a Wrapbook Fringe Report table.

For EVERY employee data row, extract exactly these fields:
  worker        — name as "Last, First" (drop middle initials)
  ssn           — as shown (may be masked)
  workDates     — date range string
  type          — W2 / Loan Out / etc.
  dept          — department code
  union         — union code
  workState     — work state abbreviation
  resState      — residence state abbreviation
  wages         — Wages (number or null)
  reimbRent     — Reimb/Rent (number or null)
  socSec        — Soc Sec (number or null)
  med           — Med (number or null)
  futa          — FUTA (number or null)
  sui           — SUI (number or null)
  wc            — W/C (number or null)
  phw           — PH&W (number or null)
  other         — Other (number or null)
  benefits      — Benefits (number or null)
  platFee       — Plat Fee (number or null)
  effRate       — EFF Rate % (number or null)
  total         — Total (number or null)
  invoiceNo     — invoice number from page header, or ""
  invoiceDate   — invoice date from page header, or ""
  invoiceWorkDates — work date range from page header, or ""

SKIP any GRAND TOTAL or summary rows.
Return ONLY a JSON array — no explanation, no markdown fences."""


def _extract_vision(pdf_bytes: bytes, openai_key: str) -> tuple[list[dict], list[str]]:
    rows, issues = [], []
    client  = OpenAI(api_key=openai_key)
    images  = pdf_to_images_b64(pdf_bytes)

    # Batch pages to stay under OpenAI's ~45 MB decoded image limit
    MAX_BYTES = 45 * 1024 * 1024
    batches, cur, cur_size = [], [], 0
    for img in images:
        approx = len(img) * 3 // 4
        if cur and cur_size + approx > MAX_BYTES:
            batches.append(cur)
            cur, cur_size = [img], approx
        else:
            cur.append(img)
            cur_size += approx
    if cur:
        batches.append(cur)

    source_page = 1
    for batch in batches:
        content = [
            {"type": "text", "text": "Extract all employee fringe rows from these Wrapbook Fringe Report pages."}
        ]
        for img in batch:
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img}", "detail": "high"},
            })
        try:
            resp = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": _VISION_SYSTEM},
                    {"role": "user",   "content": content},
                ],
                temperature=0,
                max_tokens=4096,
            )
            raw = resp.choices[0].message.content.strip()
            try:
                extracted = json.loads(raw)
            except Exception:
                m = re.search(r"\[.*\]", raw, re.DOTALL)
                extracted = json.loads(m.group()) if m else []

            for r in extracted:
                row = empty_row()
                row["payrollCompany"] = "wrapbook"
                for k, v in r.items():
                    if k in row:
                        row[k] = v
                row["sourcePage"] = source_page
                if "loan" in str(row.get("type", "")).lower():
                    row["loanOut"] = True
                rows.append(row)
        except Exception as e:
            issues.append(f"Vision batch starting page {source_page}: {e}")
        source_page += len(batch)

    return rows, issues


# ─── Register enrichment ──────────────────────────────────────────────────────

def _ssn_last4_wb(ssn: str) -> str:
    """Extract last 4 digits from any Wrapbook SSN format: 'xxx-xx-7510' → '7510'."""
    digits = re.sub(r"[^0-9]", "", ssn)
    return digits[-4:] if len(digits) >= 4 else digits


def _enrich_rows(rows: list[dict], register_data: dict) -> None:
    """Enrich fringe rows in-place with Payroll Register data.

    Join logic:
      - row.invoiceNo != "" → per-invoice join via by_ssn_invoice (GM 004 style)
      - row.invoiceNo == "" → project-level join via by_ssn (NIS 007 style)
    """
    by_ssn         = register_data.get("by_ssn", {})
    by_ssn_invoice = register_data.get("by_ssn_invoice", {})

    for row in rows:
        ssn    = _ssn_last4_wb(row.get("ssn", ""))
        inv_no = row.get("invoiceNo", "")

        if inv_no:
            norm_inv   = inv_no.lstrip("0") or "0"
            enrichment = by_ssn_invoice.get((ssn, norm_inv))
        else:
            enrichment = by_ssn.get(ssn)

        if not enrichment:
            continue

        row["jobTitle"]       = enrichment.get("jobTitle", "")
        row["daysWorked"]     = enrichment.get("daysWorked")
        row["withholdingsIL"] = enrichment.get("withholdingsIL")
        row["street"]         = enrichment.get("street", "")
        row["city"]           = enrichment.get("city", "")
        row["zip"]            = enrichment.get("zip", "")
        if enrichment.get("resState"):
            row["resState"] = enrichment["resState"]


def enrich_from_register(rows: list[dict], register_bytes: bytes) -> None:
    """Parse a standalone Wrapbook register PDF and enrich fringe rows in-place.

    Called by main.py when the user uploads a separate register file (NIS 007 style).
    """
    register_data = extract_register(register_bytes)
    if register_data:
        _enrich_rows(rows, register_data)


# ─── Public API ───────────────────────────────────────────────────────────────

def extract(pdf_bytes: bytes, openai_key: str = "", **_) -> tuple[list[dict], list[str]]:
    """Extract Wrapbook fringe rows from a PDF.

    Automatically detects whether the PDF is text-based or image-only.
    Image-only PDFs fall back to GPT-4o vision (requires openai_key).
    Returns (rows, issues).
    """
    if not has_text_layer(pdf_bytes):
        if not openai_key:
            return [], [
                "Scanned/image-based PDF — no extractable text found. "
                "This fringe report cannot be read automatically and must be entered manually."
            ]
        return _extract_vision(pdf_bytes, openai_key)
    try:
        return _extract_text(pdf_bytes)
    except Exception as e:
        return [], [f"Unexpected error reading PDF: {e}"]
