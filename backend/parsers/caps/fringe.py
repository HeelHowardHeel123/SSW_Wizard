"""
CAPS Fringe Recap Report parser — text-based (pdfplumber extract_text).

Layout per invoice section:
  Invoice Group: XXXXXX Invoice: XXXXXXXXXX Batch: XXX Pay Date: MM/DD/YYYY Work Dates: MM/DD/YYYY - MM/DD/YYYY
  Employee Name  SSN  Work Dates  Union  Taxable  Non Tax.  Corporate  FICA  Medi  FUI  SUI  W/C  PH&W  Vac/Hol  Adv  Other  Hand  Total
  LAST, FIRST M  x-3223  02/23 - 02/27/26  476  0.00  0.00  3,603.49  ...    ← W2 or loan-out
  LOAN OUT CO NAME  xxxxx5484                                                  ← sub-line: company name (may wrap to next line)

Multiple invoice sections can appear in one PDF.
"""

import re
import io

import pdfplumber

from parsers.base import empty_row, parse_amount, clean_fringe_name

# ─── Regexes ──────────────────────────────────────────────────────────────────

_INVOICE_RE = re.compile(
    r"Invoice:\s*(\d+).*?Pay Date:\s*(\d{2}/\d{2}/\d{4}).*?Work Dates:\s*"
    r"(\d{2}/\d{2}/\d{4}\s*-\s*\d{2}/\d{2}/\d{4})",
    re.IGNORECASE | re.DOTALL,
)

# Employee rows: "MM/DD - MM/DD/YY" (year only on end date) or full "MM/DD/YY - MM/DD/YY"
# The year portion is optional on either side to handle both formats.
_WORK_DATES_RE = re.compile(
    r"\d{2}/\d{2}(?:/\d{2,4})?\s*-\s*\d{2}/\d{2}(?:/\d{2,4})?"
)

# SSN patterns seen in CAPS: "x-3223", "x-1234", "XXXX", plain "1234"
_SSN_RE = re.compile(
    r"^(?:\d{4}|[Xx]{4}|[Xx]-\d{4}|[Xx]{3}(?:-[Xx]{2}-)?\d{4}|[Xx]{5}\d{4})$",
    re.IGNORECASE,
)

# Monetary column order (left → right)
_AMOUNT_FIELDS = [
    "wages",     # Taxable
    "reimbRent", # Non Tax.
    "corporate", # Corporate
    "socSec",    # FICA
    "med",       # Medi
    "futa",      # FUI
    "sui",       # SUI
    "wc",        # W/C
    "phw",       # PH&W
    "vacHol",    # Vac/Hol
    "adv",       # Adv
    "other",     # Other
    "hand",      # Hand
    "total",     # Total
]


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _parse_company_line(line: str) -> str:
    """Extract company name from a CAPS loan-out sub-line.
    Strips corporate EIN patterns (xxxxx1234) whether trailing or mid-joined."""
    s = re.sub(r"\s+[Xx]{5}\d{4}", "", line)            # strip EIN in middle (joined lines)
    s = re.sub(r"\s+(?:\d{4}|[Xx]-\d{4}|[Xx]{5}\d{4})\s*$", "", s)  # strip trailing
    return s.strip().title()


def _is_data_line(line: str) -> bool:
    """Return True if line looks like an employee or summary row (not a company name)."""
    if _WORK_DATES_RE.search(line):
        return True
    if re.match(r"^(Total|Employee Name|Invoice|Grand Total|Fringe Recap)", line, re.IGNORECASE):
        return True
    if len(re.findall(r"\d[\d,]+\.\d{2}", line)) >= 4:   # many money values = data row
        return True
    return False


# ─── Row parsing ──────────────────────────────────────────────────────────────

def _parse_employee_line(line: str) -> dict | None:
    """Parse one CAPS employee data line. Returns None if not an employee row."""
    m = _WORK_DATES_RE.search(line)
    if not m:
        return None

    head       = line[:m.start()].strip()
    work_dates = m.group(0).strip()
    tail       = line[m.end():].strip()

    # Head ends with SSN (last-4 digits or masked variant like "x-3223")
    head_parts = head.rsplit(None, 1)
    if len(head_parts) == 2 and _SSN_RE.match(head_parts[1]):
        name_raw  = head_parts[0]
        ssn_last4 = head_parts[1]
    else:
        name_raw  = head
        ssn_last4 = ""

    # Tail: union code + monetary values
    tail_parts = tail.split()
    if not tail_parts:
        return None

    union   = tail_parts[0]
    amounts = []
    for tok in tail_parts[1:]:
        val = parse_amount(tok)
        if val is not None:
            amounts.append(val)

    row = empty_row()
    row["payrollCompany"] = "caps"
    row["worker"]         = clean_fringe_name(name_raw, from_caps=True)
    row["ssn"]            = ssn_last4
    row["workDates"]      = work_dates
    row["union"]          = union

    for idx, field in enumerate(_AMOUNT_FIELDS):
        row[field] = amounts[idx] if idx < len(amounts) else None

    return row


# ─── Page-level parsing ───────────────────────────────────────────────────────

def _parse_page(text: str, page_num: int) -> tuple[list[dict], list[str]]:
    rows   = []
    issues = []
    lines  = [ln.strip() for ln in text.split("\n") if ln.strip()]

    invoice_no = invoice_date = work_dates = ""
    i = 0
    while i < len(lines):
        line = lines[i]

        # Invoice section header
        m = _INVOICE_RE.search(line)
        if m:
            invoice_no   = m.group(1)
            invoice_date = m.group(2)
            work_dates   = m.group(3)
            i += 1
            continue

        # Column header and other fixed header/summary lines — skip
        if re.match(
            r"^(Employee Name|Fringe Recap|Total|Grand Total|Subtotal|Invoice Totals)",
            line, re.IGNORECASE,
        ):
            i += 1
            continue

        # Employee data row
        row = _parse_employee_line(line)
        if row:
            row["invoiceNo"]        = invoice_no
            row["invoiceDate"]      = invoice_date
            row["invoiceWorkDates"] = work_dates
            row["sourcePage"]       = page_num
            rows.append(row)

            # Loan-out: corporate > 0 → collect following non-data lines as company name
            if row.get("corporate") and row["corporate"] > 0:
                row["loanOut"] = True
                row["type"]    = "Loan Out"
                j = i + 1
                company_frags = []
                while j < len(lines):
                    peek = lines[j].strip()
                    if not peek:
                        j += 1
                        continue
                    if _is_data_line(peek):
                        break
                    company_frags.append(peek)
                    j += 1
                if company_frags:
                    row["loanOutCompany"] = _parse_company_line(" ".join(company_frags))
                    i = j  # skip past all company name lines
                    continue

        i += 1

    return rows, issues


# ─── Public API ───────────────────────────────────────────────────────────────

def extract(pdf_bytes: bytes) -> tuple[list[dict], list[str]]:
    """Extract all CAPS Fringe Recap Report rows from a PDF.

    Returns (rows, issues) where rows is a list of canonical fringe dicts.
    """
    rows   = []
    issues = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            found_any = False
            for pg_idx, pg in enumerate(pdf.pages):
                text = pg.extract_text() or ""
                if "Fringe Recap Report" not in text:
                    continue
                found_any = True
                page_rows, page_issues = _parse_page(text, pg_idx + 1)
                rows.extend(page_rows)
                issues.extend(page_issues)
        if not found_any:
            issues.append("No CAPS Fringe Recap Report page found")
    except Exception as e:
        issues.append(str(e))
    return rows, issues
