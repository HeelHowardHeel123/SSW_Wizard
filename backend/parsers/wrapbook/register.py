"""
Wrapbook Payroll Register parser.

Handles two delivery formats:
  - Per-invoice (GM 004 style): register pages inside the same invoice PDF.
    Each employee has exactly one check; invoice ID from "Batch ID #NNNNNN".
  - Project-level (NIS 007 style): standalone register PDF covering all invoices.
    Each employee may have multiple checks; invoice ID from "Payroll ID #NNNNNN".
    Reimbursement-only checks (Kit Rental, Mileage) carry no IL tax.

Returns
-------
{
  "by_ssn": {
      ssn_last4: {jobTitle, street, city, zip, resState, daysWorked,
                  withholdingsIL, withholdingsGA}   ← formula string "=a+b" when multi-check
  },
  "by_ssn_invoice": {
      (ssn_last4, invoice_id): {same fields, withholdingsIL/withholdingsGA as plain numbers}
  }
}

Join logic in fringe.py:
  - invoiceNo != "" → look up by_ssn_invoice (strip leading zeros from invoice no.)
  - invoiceNo == "" → look up by_ssn (aggregate totals)
"""

import re
import io

import pdfplumber


# ─── Regexes ──────────────────────────────────────────────────────────────────

# Employee start: "Last, First [Middle] $1,234.56" at end of line
# First char is [A-Za-z] to handle lowercase-prefixed surnames like "van Sauter"
_EMP_NAME_RE = re.compile(
    r"^([A-Za-z][A-Za-z\-\'\s]+,\s+[A-Za-z][A-Za-z\-\'\s\.]+?)\s+\$([\d,]+\.\d{2})$"
)

# SSN line: "XXX-XX-7510"
_SSN_LINE_RE = re.compile(r"^XXX-XX-(\d{4})$", re.IGNORECASE)

# Total days worked from employee header line 3:
# "Gaffer, Grip/Electric 3521 N Kostner Ave Honey Bear Films Days Worked: 3 Total Hours Worked: 33.0"
_TOTAL_DAYS_RE = re.compile(r"Days Worked:\s*(\d+)", re.IGNORECASE)

# Check line: "Check Date: Feb 02, 2026 Payroll ID #684317 ... Days Worked: 2 Date Range: ... $amount"
#          or "Check Date: Apr 20, 2026 Batch ID #715136 ... Days Worked: 3 ..."
_CHECK_LINE_RE = re.compile(
    r"Check Date:.*?(?:Payroll|Batch)\s+ID\s+#(\d+).*?Days Worked:\s*(\d+)",
    re.IGNORECASE,
)

# Reimbursement-only check (no wages): "Check Date: ... Payroll ID #NNN Payroll Type: Regular $47.85"
# followed only by Mileage/Kit Rental lines — detected by absence of IL tax in block
# (no separate regex needed; handled by checking whether IL tax was found in the block)

# IL state tax: "Illinois State Tax $83.77" or "Illinois State Tax Credit $79.31"
_IL_TAX_RE = re.compile(
    r"Illinois State Tax(?:\s+Credit)?\s+\$([\d,]+\.\d{2})",
    re.IGNORECASE,
)

# GA state tax: "Georgia State Tax $60.33" or "Georgia State Tax Credit $..." -- same
# format/position as the IL line above
_GA_TAX_RE = re.compile(
    r"Georgia State Tax(?:\s+Credit)?\s+\$([\d,]+\.\d{2})",
    re.IGNORECASE,
)

# City, State ZIP from line 4:
# "Non-Union Chicago, IL 60641-3808 XX-XXX4292 Jan 22, 2026"
_CITY_STATE_ZIP_RE = re.compile(
    r"([A-Za-z][A-Za-z ]+),\s+([A-Z]{2})\s+(\d{5})(?:-\d{4})?"
)

# Street suffix markers
_STREET_SUFFIX_RE = re.compile(
    r"\b(Ave(?:nue)?|St(?:reet)?|Dr(?:ive)?|Blvd|Boulevard|Rd|Road|Ln|Lane|"
    r"Way|Ct|Court|Pl(?:ace)?|Cir(?:cle)?|Ter(?:race)?|Hwy|Highway|"
    r"Pkwy|Parkway|Trl|Trail)\b",
    re.IGNORECASE,
)

# Page stamp / header lines to filter when joining pages
_PAGE_STAMP_RE = re.compile(r"^\d{2}/\d{2}/\d{4}\s+-\s+\d{2}:\d{2}\s+\|\s+Page")
_HEADER_RE = re.compile(
    r"^(Nissan Payroll Register|Betty Crocker.*Payroll Register|"
    r".*Payroll Register$|"        # generic project name + Payroll Register
    r"TakeOne Network|228 Park|1 \(833\)|"
    r"Payment Type Location)",
    re.IGNORECASE,
)


# ─── Address helpers ───────────────────────────────────────────────────────────

def _parse_street(text: str) -> str:
    """Extract the street address from a mixed text string.
    Stops after the last street suffix (Ave, St, Dr, etc.) plus optional apt.
    """
    m_num = re.search(r"\b(\d{2,6})\s+", text)
    if not m_num:
        return ""
    segment = text[m_num.start():]
    # Find the last street suffix in the segment
    last_suffix = None
    for m in _STREET_SUFFIX_RE.finditer(segment):
        last_suffix = m
    if not last_suffix:
        return segment.strip()
    end = last_suffix.end()
    # Optionally capture apartment/unit after the suffix
    apt = re.match(r"\s+(?:Apt|Unit|Suite|#|No\.?)\s*\S+", segment[end:], re.IGNORECASE)
    if apt:
        end += apt.end()
    return segment[:end].strip()


def _parse_city_state_zip(line: str) -> tuple[str, str, str]:
    """Return (city, state, zip5) from a line like 'Non-Union Chicago, IL 60641-3808 ...'

    Strips the leading union designation ("Non-Union", "Union") before searching,
    so "Non-Union Chicago, IL" → "Chicago" not "Union Chicago".
    """
    clean = re.sub(r"^Non-Union\s+|^Union\s+", "", line)
    m = _CITY_STATE_ZIP_RE.search(clean)
    if not m:
        return "", "", ""
    return m.group(1).strip().title(), m.group(2).upper(), m.group(3)


def _parse_job_title(text: str) -> str:
    """Extract job title from line 3 text (before the house number).
    Strips the department suffix (everything after the first ', ').
    """
    m_num = re.search(r"\b\d{2,6}\s+", text)
    raw = text[: m_num.start()].strip() if m_num else text.strip()
    # Drop department (after first ', ')
    if ", " in raw:
        raw = raw.split(", ", 1)[0]
    return raw.strip()


def _format_withholdings(amounts: list[float]) -> float | str | None:
    """Return a number if single value, formula string if multiple, None if empty."""
    amounts = [a for a in amounts if a is not None]
    if not amounts:
        return None
    if len(amounts) == 1:
        return amounts[0]
    parts = "+".join(str(a) for a in amounts)
    return f"={parts}"


# ─── Employee block parser ─────────────────────────────────────────────────────

def _parse_employee(emp_lines: list[str]) -> dict | None:
    """Parse one employee's block of lines. Returns enrichment dict or None."""
    if len(emp_lines) < 4:
        return None

    # Line 0: name + total net (already confirmed as employee start)
    # Line 1: SSN
    m_ssn = _SSN_LINE_RE.match(emp_lines[1]) if len(emp_lines) > 1 else None
    if not m_ssn:
        return None
    ssn_last4 = m_ssn.group(1)

    # Line 2: job title, street, optional company, Days Worked total
    line3 = emp_lines[2] if len(emp_lines) > 2 else ""
    job_title = _parse_job_title(line3)
    street = _parse_street(line3)
    m_total_days = _TOTAL_DAYS_RE.search(line3)
    total_days = int(m_total_days.group(1)) if m_total_days else None

    # Line 3: union, city/state/zip, date range
    line4 = emp_lines[3] if len(emp_lines) > 3 else ""
    city, state, zip5 = _parse_city_state_zip(line4)

    # Scan all remaining lines for check blocks
    checks: list[dict] = []  # {invoice_id, days, il_tax, ga_tax}
    current_check: dict | None = None

    for line in emp_lines[4:]:
        # New check block
        m_check = _CHECK_LINE_RE.search(line)
        if m_check:
            if current_check is not None:
                checks.append(current_check)
            current_check = {
                "invoice_id": m_check.group(1),
                "days": int(m_check.group(2)),
                "il_tax": None,
                "ga_tax": None,
            }
            continue

        # IL state tax within current check block
        m_tax = _IL_TAX_RE.search(line)
        if m_tax and current_check is not None:
            try:
                current_check["il_tax"] = round(float(m_tax.group(1).replace(",", "")), 2)
            except ValueError:
                pass

        # GA state tax within current check block
        m_ga_tax = _GA_TAX_RE.search(line)
        if m_ga_tax and current_check is not None:
            try:
                current_check["ga_tax"] = round(float(m_ga_tax.group(1).replace(",", "")), 2)
            except ValueError:
                pass

    if current_check is not None:
        checks.append(current_check)

    # Build aggregate (for project-level fringe with no invoice number)
    all_il_taxes = [c["il_tax"] for c in checks if c.get("il_tax") is not None]
    aggregate_withholdings = _format_withholdings(all_il_taxes)
    all_ga_taxes = [c["ga_tax"] for c in checks if c.get("ga_tax") is not None]
    aggregate_withholdings_ga = _format_withholdings(all_ga_taxes)

    base = {
        "jobTitle":  job_title,
        "street":    street,
        "city":      city,
        "zip":       zip5,
        "resState":  state,
        "daysWorked": total_days,
    }

    return {
        "ssn_last4":  ssn_last4,
        "aggregate":  {**base, "withholdingsIL": aggregate_withholdings,
                       "withholdingsGA": aggregate_withholdings_ga},
        "checks":     [{**base, "invoice_id": c["invoice_id"],
                        "daysWorked": c["days"],
                        "withholdingsIL": c["il_tax"],
                        "withholdingsGA": c["ga_tax"]} for c in checks],
    }


# ─── Public API ───────────────────────────────────────────────────────────────

def extract_register(pdf_bytes: bytes) -> dict:
    """Parse a Wrapbook Payroll Register PDF (standalone or embedded).

    Returns {"by_ssn": {...}, "by_ssn_invoice": {...}} — see module docstring.
    """
    by_ssn: dict = {}
    by_ssn_invoice: dict = {}

    try:
        # Collect and filter all page lines
        all_lines: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pg in pdf.pages:
                text = pg.extract_text() or ""
                for line in text.split("\n"):
                    line = line.strip()
                    if not line:
                        continue
                    if _PAGE_STAMP_RE.match(line):
                        continue
                    if _HEADER_RE.match(line):
                        continue
                    all_lines.append(line)

        # Split into employee blocks
        # Boundary: line matching _EMP_NAME_RE AND next line matching _SSN_LINE_RE
        employee_starts: list[int] = []
        for idx in range(len(all_lines) - 1):
            if (_EMP_NAME_RE.match(all_lines[idx])
                    and _SSN_LINE_RE.match(all_lines[idx + 1])):
                employee_starts.append(idx)

        for i, start in enumerate(employee_starts):
            end = employee_starts[i + 1] if i + 1 < len(employee_starts) else len(all_lines)
            emp_lines = all_lines[start:end]
            result = _parse_employee(emp_lines)
            if not result:
                continue

            ssn = result["ssn_last4"]
            by_ssn[ssn] = result["aggregate"]

            for check in result["checks"]:
                inv_id = check["invoice_id"].lstrip("0") or "0"
                by_ssn_invoice[(ssn, inv_id)] = check

    except Exception:
        pass

    return {"by_ssn": by_ssn, "by_ssn_invoice": by_ssn_invoice}
