"""
Revolution Entertainment Services fringe parser.

Two invoice types found in the wild:
  A23720A-series: single-employee (Durable Goods / BIDDING projects), 7 pages
  A23730A-series: multi-employee  (Concrete Images / Hot Wins),       21-77 pages

Both use the same 'Payroll Fringe Recap' page layout; large invoices span
multiple Fringe Recap pages (each with its own header line).

Fringe Recap column order (values after SSN/Name):
  Check #, Non-Tax Wages, Taxable Wages, Vac/Hol, SS & MEDI (combined),
  BCR FUTA, FUTA, SUTA, Work Comp, Admin, Other,
  [EE Deds, ER Benefits, Local Tax — count varies 0-3 zeroes],
  Total Fringe, Total Cost

SS & MEDI is employer SS+Medicare combined; we split by 6.2 / 1.45 ratio.
BCR FUTA + FUTA are both mapped to futa (BCR FUTA is the primary charge).

Loan-out employees have two consecutive SSN lines before the amounts:
  *****XXXX COMPANY NAME      <- loan-out company (FEIN last-4)
  *****YYYY EMPLOYEE NAME     <- actual employee   (SSN last-4)
  CHECK# amounts...

Enrichment (address, job title, days worked) is read from the
'Employee Information' / 'Current Payee Information' pages, matched
by check number.
"""

import re
import io

import pdfplumber

from parsers.base import empty_row, parse_amount, clean_fringe_name

COMPANY  = "revolution_entertainment_services"
MARKERS  = ["Payroll Fringe Recap", "1210 W Burbank Blvd"]
PRIORITY = 10

_SS_RATIO  = 6.2  / 7.65
_MED_RATIO = 1.45 / 7.65

_SSN_RE     = re.compile(r"^\*+(\d{4})\s+(.+)$")
_INVOICE_RE = re.compile(r"Invoice\s*#:\s*([A-Z0-9]+)", re.IGNORECASE)
_PERIOD_RE  = re.compile(r"Period:\s*(\d{2}/\d{2}/\d{4})\s+to\s+(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
_POSTED_RE  = re.compile(r"Posted:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)


# ─── Fringe Recap parsing ─────────────────────────────────────────────────────

def _parse_fringe_page(text: str, invoice_no: str, period: str, inv_date: str) -> list[dict]:
    rows  = []
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

    i = 0
    while i < len(lines):
        line = lines[i]

        if re.match(
            r"(Non-Tax|SSN\s|Wages|Payroll Fringe|Client:|Project:|Payroll ID|"
            r"Invoice\s*#:|Period:|Posted:|Report Totals?:|Page \d)",
            line, re.IGNORECASE,
        ):
            i += 1
            continue

        m = _SSN_RE.match(line)
        if not m:
            i += 1
            continue

        ssn_last4   = m.group(1)
        name_part   = m.group(2).strip()
        loan_out_co = ""

        # Loan-out: the next line is also an SSN line
        if i + 1 < len(lines) and _SSN_RE.match(lines[i + 1]):
            loan_out_co = name_part
            i += 1
            m2        = _SSN_RE.match(lines[i])
            ssn_last4 = m2.group(1)
            name_part = m2.group(2).strip()

        i += 1
        if i >= len(lines):
            break

        amounts = lines[i].split()
        if not (amounts and re.match(r"^\d+$", amounts[0])):
            continue

        def tok(idx):
            pos = idx if idx >= 0 else len(amounts) + idx
            return parse_amount(amounts[pos]) if 0 <= pos < len(amounts) else None

        ss_medi  = tok(4) or 0.0
        bcr_futa = tok(5) or 0.0
        futa_reg = tok(6) or 0.0

        row = empty_row()
        row["payrollCompany"] = "Revolution Entertainment Services"
        row["worker"]         = clean_fringe_name(name_part)
        row["ssn"]            = ssn_last4
        row["invoiceNo"]      = invoice_no
        row["invoiceDate"]    = inv_date
        row["workDates"]      = period
        row["reimbRent"]      = tok(1)
        row["wages"]          = tok(2)
        row["vacHol"]         = tok(3)
        row["socSec"]         = round(ss_medi * _SS_RATIO,  2) if ss_medi else None
        row["med"]            = round(ss_medi * _MED_RATIO, 2) if ss_medi else None
        row["futa"]           = round(bcr_futa + futa_reg,  2) if (bcr_futa or futa_reg) else None
        row["sui"]            = tok(7)
        row["wc"]             = tok(8)
        row["adv"]            = tok(9)
        row["other"]          = tok(10)
        row["total"]          = tok(-1)
        row["_check_no"]      = amounts[0]   # temporary key for enrichment matching

        if loan_out_co:
            row["loanOut"]        = True
            row["loanOutCompany"] = loan_out_co.title()

        rows.append(row)
        i += 1

    return rows


# ─── Enrichment from Employee / Payee Info pages ─────────────────────────────

def _extract_register(pdf_bytes: bytes) -> dict:
    """Return {check_number: {street, city, zip, jobTitle, daysWorked}} from info pages."""
    reg = {}
    _CHECK_RE = re.compile(r"Check Number:\s*(\d+)", re.IGNORECASE)
    _ADDR_RE  = re.compile(r"^Address:\s*(.+),\s*([^,]+),\s*([A-Z]{2}),\s*([\d-]+)", re.IGNORECASE)
    _TITLE_RE = re.compile(r"Job Title:\s*([^\n]+?)(?:\s{2,}|$)", re.IGNORECASE)
    _DAYS_RE  = re.compile(r"Number of Days Worked:\s*(\d+)", re.IGNORECASE)

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pg in pdf.pages:
                text = pg.extract_text() or ""
                if "Employee Information:" not in text and "Current Payee Information:" not in text:
                    continue

                m_check = _CHECK_RE.search(text)
                if not m_check:
                    continue
                check_no = m_check.group(1)

                entry = {}

                for line in text.split("\n"):
                    line = line.strip()
                    m_addr = _ADDR_RE.match(line)
                    if m_addr and "street" not in entry:
                        entry["street"] = m_addr.group(1).strip()
                        entry["city"]   = m_addr.group(2).strip()
                        entry["zip"]    = m_addr.group(4).strip()[:5]
                        break

                m_title = _TITLE_RE.search(text)
                if m_title:
                    entry["jobTitle"] = m_title.group(1).strip().rstrip(",")

                m_days = _DAYS_RE.search(text)
                if m_days:
                    entry["daysWorked"] = int(m_days.group(1))

                if entry:
                    reg[check_no] = entry
    except Exception:
        pass

    return reg


# ─── Public API ───────────────────────────────────────────────────────────────

def extract(pdf_bytes: bytes, **_) -> tuple[list[dict], list[str]]:
    rows   = []
    issues = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            invoice_no = ""
            period     = ""
            inv_date   = ""

            for pg in pdf.pages:
                text = pg.extract_text() or ""
                if "Payroll Fringe Recap" not in text:
                    continue

                if not invoice_no:
                    m = _INVOICE_RE.search(text)
                    if m:
                        invoice_no = m.group(1)

                if not period:
                    m = _PERIOD_RE.search(text)
                    if m:
                        period = f"{m.group(1)} to {m.group(2)}"

                if not inv_date:
                    m = _POSTED_RE.search(text)
                    if m:
                        inv_date = m.group(1)

                rows.extend(_parse_fringe_page(text, invoice_no, period, inv_date))

        if not rows:
            issues.append(
                "No Fringe Recap rows found — verify this is a Revolution "
                "Entertainment Services PDF (1210 W Burbank Blvd)."
            )
            return rows, issues

        # Enrich with address / job title / days worked
        reg = _extract_register(pdf_bytes)
        if reg:
            for row in rows:
                info = reg.get(row.get("_check_no", ""))
                if info:
                    row["street"]     = info.get("street", "")
                    row["city"]       = info.get("city", "")
                    row["zip"]        = info.get("zip", "")
                    row["jobTitle"]   = info.get("jobTitle", "")
                    row["daysWorked"] = info.get("daysWorked")

        # Remove temporary internal key
        for row in rows:
            row.pop("_check_no", None)

    except Exception as e:
        issues.append(str(e))

    return rows, issues
