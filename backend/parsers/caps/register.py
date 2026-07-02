"""
CAPS Invoice Wage/Payroll Register parser.

Extracts per-employee enrichment data from "Invoice Wage/Payroll Register"
pages inside a CAPS invoice PDF. The results are joined to fringe rows
by (ssn_last4, invoice_no) in caps/fringe.py.

Fields extracted per employee:
  jobTitle      - crew position (e.g. "Key Grip")
  daysWorked    - count of individual work days
  workDates     - formatted date string (overrides fringe abbreviated form)
  street        - home street address
  city          - home city
  zip           - 5-digit home zip
  resState      - residence state abbreviation
  withholdingsIL - Illinois State Tax withheld (null for loan-outs)
"""

import re
import io

import pdfplumber


# ─── Regexes ──────────────────────────────────────────────────────────────────

# Detect an employee name line: split on the SSN token "xxx-xx-DDDD"
_SSN_TOKEN_RE = re.compile(r"\s+(xxx-xx-(\d{4}))\s+", re.IGNORECASE)

# Address tail after SSN: "STREET, CITY, ST ZIPEXT US"
_ADDR_RE = re.compile(
    r"^(.+?),\s+(.+?),\s+([A-Z]{2})\s+(\d{5})(?:-\d{4})?\s+US$",
    re.IGNORECASE,
)

# Invoice number in page header: "Invoice 1001385632"
_INV_HEADER_RE = re.compile(r"\bInvoice\s+(\d{8,12})\b")

# Per-employee work dates line: "Work Dates: 10/ 14, 15, 16, 17/2025 Check Date:"
_WORK_LINE_RE = re.compile(r"Work Dates:\s*(.+?)\s+Check Date:", re.IGNORECASE)

# Parse the dates content: "10/ 14, 15, 16, 17/2025" or "10/ 15/2025"
_DATES_CONTENT_RE = re.compile(r"(\d{1,2})/\s*([\d,\s]+)/(\d{4})")

# Job title is everything before "Federal Filing Status:"
_JOB_TITLE_RE = re.compile(r"^(.+?)\s+Federal Filing Status:", re.IGNORECASE)

# IL state tax line: "Illinois State Tax 83.77"
_IL_TAX_RE = re.compile(r"Illinois State Tax\s+([\d,]+\.\d{2})", re.IGNORECASE)

# Explicit res state override: "Res State: IL"
_RES_STATE_RE = re.compile(r"Res State:\s*([A-Z]{2})", re.IGNORECASE)

# Page stamp to filter: "11/05/2025 12:27 PM 2"
_PAGE_STAMP_RE = re.compile(r"^\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}\s+[AP]M\s+\d+$")

# Production Totals marks end of individual employee records
_TOTALS_RE = re.compile(r"^(Production Totals|State Totals|Total number)", re.IGNORECASE)

# Page header lines to drop when joining multi-page registers
_HEADER_LINE_RE = re.compile(
    r"^(LA OFFICE|2300 EMPIRE|BURBANK CA|\(310\)|User:|"
    r"Batch \d+ Invoice \d+ Type|"
    r"Work Dates \d{2}/\d{2}/\d{4}|"  # invoice-level date header (no colon)
    r"Payment Type Location|"
    r"Invoice Wage/Payroll Register)",
    re.IGNORECASE,
)


# ─── Date helpers ──────────────────────────────────────────────────────────────

def _parse_work_dates(dates_str: str) -> tuple[str, int]:
    """Parse 'Work Dates:' content into (formatted_string, day_count).

    '10/ 15/2025'               → ('10/15/2025', 1)
    '10/ 14, 15, 16, 17/2025'  → ('10/14/2025 - 10/17/2025', 4)
    '10/ 23, 24/2025'           → ('10/23/2025 - 10/24/2025', 2)
    Non-consecutive days        → ('10/22/2025, 10/24/2025', 2)
    """
    m = _DATES_CONTENT_RE.search(dates_str)
    if not m:
        return "", 0

    month = m.group(1).zfill(2)
    days_raw = m.group(2)
    year = m.group(3)

    days = sorted(int(d) for d in re.findall(r"\d+", days_raw))
    if not days:
        return "", 0

    count = len(days)
    if count == 1:
        return f"{month}/{days[0]:02d}/{year}", 1

    consecutive = all(days[i + 1] - days[i] == 1 for i in range(count - 1))
    if consecutive:
        return f"{month}/{days[0]:02d}/{year} - {month}/{days[-1]:02d}/{year}", count
    else:
        return ", ".join(f"{month}/{d:02d}/{year}" for d in days), count


def _to_title(s: str) -> str:
    """Title-case that handles mixed alphanumerics: '2ND PROPS' → '2nd Props'."""
    return re.sub(r"\b([A-Za-z]+)\b", lambda m: m.group(1).capitalize(), s.lower())


# ─── Core parsing ─────────────────────────────────────────────────────────────

def _parse_employee_block(emp_lines: list[str], addr_state: str) -> dict:
    """Extract enrichment fields from one employee's line block."""
    job_title = ""
    work_dates = ""
    days_worked = 0
    il_tax = None
    res_state = addr_state

    for line in emp_lines:
        if not work_dates:
            wm = _WORK_LINE_RE.search(line)
            if wm:
                work_dates, days_worked = _parse_work_dates(wm.group(1))

        if not job_title:
            jm = _JOB_TITLE_RE.match(line)
            if jm:
                job_title = _to_title(jm.group(1).strip())

        tm = _IL_TAX_RE.search(line)
        if tm:
            try:
                il_tax = round(float(tm.group(1).replace(",", "")), 2)
            except ValueError:
                pass

        rm = _RES_STATE_RE.search(line)
        if rm:
            res_state = rm.group(1).upper()

    return {
        "jobTitle":       job_title,
        "workDates":      work_dates,
        "daysWorked":     days_worked if days_worked else None,
        "resState":       res_state,
        "withholdingsIL": il_tax,
    }


def _parse_invoice_block(lines: list[str], invoice_no: str, result: dict) -> None:
    """Walk lines for one invoice's register and populate result dict."""
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]

        # Stop at production totals
        if _TOTALS_RE.match(line):
            break

        # Employee start: line contains "xxx-xx-DDDD"
        sm = _SSN_TOKEN_RE.search(line)
        if not sm:
            i += 1
            continue

        ssn_last4 = sm.group(2)

        # Parse address from the portion after the SSN token
        after_ssn = line[sm.end():]
        addr_m = _ADDR_RE.match(after_ssn)
        if addr_m:
            street   = _to_title(addr_m.group(1).strip())
            city     = _to_title(addr_m.group(2).strip())
            zip5     = addr_m.group(4)
            st       = addr_m.group(3).upper()
        else:
            street = city = zip5 = st = ""

        # Collect this employee's block until next employee, totals, or end
        j = i + 1
        emp_lines = []
        while j < n:
            peek = lines[j]
            if _TOTALS_RE.match(peek):
                break
            if _SSN_TOKEN_RE.search(peek):
                break
            emp_lines.append(peek)
            j += 1

        enrichment = _parse_employee_block(emp_lines, st)
        enrichment["street"] = street
        enrichment["city"]   = city
        enrichment["zip"]    = zip5

        key = (ssn_last4, invoice_no)
        # If same (ssn, invoice) appears twice (split checks), keep first
        if key not in result:
            result[key] = enrichment

        i = j


# ─── Public API ───────────────────────────────────────────────────────────────

def extract_register(pdf_bytes: bytes) -> dict:
    """Parse all Payroll Register pages from a CAPS PDF.

    Returns a dict keyed by (ssn_last4, invoice_no) mapping to an enrichment
    dict with keys: jobTitle, workDates, daysWorked, resState, withholdingsIL,
    street, city, zip.
    """
    result: dict = {}

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            # Group register pages by invoice number
            by_invoice: dict[str, list[str]] = {}
            current_invoice = ""

            for pg in pdf.pages:
                text = pg.extract_text() or ""
                lines = [l.strip() for l in text.split("\n") if l.strip()]
                if not lines:
                    continue

                # Check for register page (first few lines)
                is_register = any(
                    "Invoice Wage/Payroll Register" in l for l in lines[:6]
                )
                if not is_register:
                    continue

                # Extract invoice number from this page's header
                for line in lines[:10]:
                    m = _INV_HEADER_RE.search(line)
                    if m and "Invoice Wage/Payroll Register" not in line:
                        current_invoice = m.group(1)
                        break

                if not current_invoice:
                    continue

                if current_invoice not in by_invoice:
                    by_invoice[current_invoice] = []
                by_invoice[current_invoice].extend(lines)

            # Parse each invoice's collected lines
            for invoice_no, raw_lines in by_invoice.items():
                # Filter page headers and stamps
                clean = [
                    l for l in raw_lines
                    if not _PAGE_STAMP_RE.match(l)
                    and not _HEADER_LINE_RE.match(l)
                    and l
                ]
                _parse_invoice_block(clean, invoice_no, result)

    except Exception:
        pass

    return result
