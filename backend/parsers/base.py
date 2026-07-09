"""
Shared utilities and canonical schema for all fringe parsers.
"""

import re
import io
import base64

import fitz  # PyMuPDF

# ─── Canonical fringe output schema ───────────────────────────────────────────
# Every parser returns a list of dicts with exactly these keys.
# Fields not present in a given payroll company's report stay at their zero value.

FRINGE_FIELDS = [
    # Identity
    "worker",       # "Last, First"
    "ssn",          # last-4 or masked string
    "workDates",    # date range string
    "type",         # W2 / Loan Out / etc.  (Wrapbook only)
    "dept",         # department code       (Wrapbook only)
    "union",        # union code
    "workState",    # work state abbrev
    "resState",     # residence state abbrev
    # Wages
    "wages",        # taxable wages  → Excel col O
    "reimbRent",    # non-taxable    → Excel col P (kit/mileage/rental — categorized later)
    "corporate",    # loan-out wages → Excel col T  (CAPS only; Wrapbook loan-outs use wages)
    # Fringes
    "socSec",       # FICA / Soc Sec
    "med",          # Medicare / Medi
    "futa",         # FUTA / FUI
    "sui",          # SUI
    "wc",           # W/C
    "phw",          # PH&W
    "vacHol",       # Vac/Hol  (CAPS only)
    "adv",          # Adv      (CAPS only)
    "other",        # Other
    "benefits",     # Benefits (Wrapbook only)
    "platFee",      # Plat Fee (Wrapbook only)
    "effRate",      # EFF Rate % (Wrapbook only)
    "hand",         # Hand     (CAPS only)
    "total",        # Total fringe
    # Loan-out metadata
    "loanOut",        # bool: True if this is a loan-out row
    "loanOutCompany", # company name (CAPS: from sub-line; Wrapbook: type="Loan Out")
    # Payroll Register enrichment (populated by caps/register.py)
    "jobTitle",       # job title / crew position
    "daysWorked",     # number of days worked on this invoice
    "withholdingsIL", # Illinois state income tax withheld
    "street",         # home street address
    "city",           # home city
    "zip",            # home zip (5-digit)
    # Provenance
    "invoiceNo",
    "invoiceDate",
    "invoiceWorkDates",
    "sourcePage",
    "sourceFile",
    "payrollCompany",  # "wrapbook" | "caps"
]

_NUMERIC_FIELDS = {
    "wages", "reimbRent", "corporate",
    "socSec", "med", "futa", "sui", "wc", "phw",
    "vacHol", "adv", "other", "benefits", "platFee", "effRate", "hand",
    "total",
    "daysWorked", "withholdingsIL",
}


def empty_row() -> dict:
    """Return a dict with all fringe fields at their zero/empty values."""
    row = {}
    for f in FRINGE_FIELDS:
        if f in _NUMERIC_FIELDS:
            row[f] = None
        elif f == "loanOut":
            row[f] = False
        else:
            row[f] = ""
    return row


# ─── Amount parsing ────────────────────────────────────────────────────────────

def parse_amount(val) -> float | None:
    """Parse a monetary string like '$1,234.56' or '1234.56%' to float."""
    if val is None:
        return None
    s = re.sub(r"[$,%\s]", "", str(val))
    if not s:
        return None
    try:
        return round(float(s), 2)
    except Exception:
        return None


# ─── Address formatting ───────────────────────────────────────────────────────

_ADDR_STREET_TYPES = {
    'AVENUE': 'Ave', 'STREET': 'St', 'DRIVE': 'Dr', 'ROAD': 'Rd',
    'BOULEVARD': 'Blvd', 'PLACE': 'Pl', 'COURT': 'Ct', 'LANE': 'Ln',
    'CIRCLE': 'Cir', 'HIGHWAY': 'Hwy', 'PARKWAY': 'Pkwy',
    'TERRACE': 'Ter', 'TRAIL': 'Trl',
}
_ADDR_COMPOUND_DIRS = frozenset({'NE', 'NW', 'SE', 'SW'})
_ADDR_SINGLE_DIRS   = {'NORTH': 'N', 'SOUTH': 'S', 'EAST': 'E', 'WEST': 'W'}


def fmt_street_address(addr: str) -> str:
    """Title-case a street address: strip unit numbers, abbreviate street types
    and directionals, remove trailing periods from directionals."""
    if not addr:
        return ''
    addr = re.sub(r',?\s*#\s*\S+', '', addr.strip())
    addr = re.sub(r',?\s*(APT|APARTMENT|UNIT|SUITE|STE|FL|FLOOR|ROOM)\s+\S+',
                  '', addr.strip(), flags=re.IGNORECASE)
    result = []
    for word in addr.split():
        up = word.upper().rstrip('.')
        if up in _ADDR_COMPOUND_DIRS:
            result.append(up)
        elif up in _ADDR_SINGLE_DIRS:
            result.append(_ADDR_SINGLE_DIRS[up])
        elif up in _ADDR_STREET_TYPES:
            result.append(_ADDR_STREET_TYPES[up])
        else:
            result.append(word.capitalize().rstrip('.'))
    return ' '.join(result).strip()


# ─── Name cleaning ─────────────────────────────────────────────────────────────

def clean_fringe_name(val: str, from_caps: bool = False) -> str:
    """
    Normalize a person's name from fringe report text.

    from_caps=True: input is ALL CAPS (CAPS payroll company format).
    Handles:
      - Asterisk removal:  "MICHAEL F*"  → "Michael F"
      - Leading initial:   "R. SCOTT"    → "Scott"
      - Trailing middle initial after comma: "Smith, John M" → "Smith, John"
    """
    if not val:
        return ""
    s = str(val).strip().replace("*", "")
    if from_caps:
        s = s.title()
    # Drop leading single-letter initial: "R. Smith, ..." → "Smith, ..."
    s = re.sub(r"^[A-Za-z]\.\s+", "", s)
    # Drop trailing middle initial when name has "Last, First M" format
    s = re.sub(r"(,\s*\S.*?)\s+[A-Z]\.?\s*$", lambda m: m.group(1), s)
    return s.strip()


# ─── PDF utilities ─────────────────────────────────────────────────────────────

def has_text_layer(pdf_bytes: bytes) -> bool:
    """Return True if the PDF has any extractable text words (not purely image-based)."""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pg in pdf.pages:
                if pg.extract_words():
                    return True
    except Exception:
        pass
    return False


def pdf_to_images_b64(pdf_bytes: bytes, dpi_scale: float = 2.0) -> list[str]:
    """Render every PDF page to a base64-encoded PNG string."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale))
        images.append(base64.b64encode(pix.tobytes("png")).decode())
    doc.close()
    return images
