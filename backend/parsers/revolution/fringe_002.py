"""
Revolution Business Services, LLC fringe parser.

Same address as Revolution Entertainment Services (1210 W Burbank Blvd,
Burbank CA) and the same general "Fringe Recap" + "Payroll Wage Register"
report shape as fringe_001 -- almost certainly the same real vendor under a
different DBA/invoice generation -- but the column layout genuinely
differs, so this is a separate variant rather than a shared code path:

  - fringe_001's Fringe Recap page literally says "Payroll Fringe Recap"
    and has a split "BCR FUTA" + "FUTA" pair; this one says just "Fringe
    Recap" and has a single combined "FUTA" column.
  - Fringe Recap column order here (confirmed against a real invoice,
    A23930A1269): Non-Tax Wages, Taxable Wages, VAC HOL, SS & Medi
    (combined), FUTA, SUTA, Work Comp, Admin, EE Benefits, Ded's
    Credited, ER Local Tax, Other, Total Fringe, Total Cost -- 14 values
    after the check number. Cross-checked against the invoice cover
    page's own line items (Admin=Administration Fees, EE Benefits=
    Client Billing for Employee Benefits, FUTA/SUTA/Work Comp all match
    exactly, Other=Postage & Delivery) and the Report Totals row, which
    equals the cover page's TOTAL AMOUNT DUE to the penny.
  - Enrichment (address, job title, exact FICA/Medicare split) comes from
    "Payroll Wage Register" pages, keyed by check number -- a different
    page layout than fringe_001's "Employee/Current Payee Information"
    pages.

The invoice number is ONLY the cover page's "Invoice No." (repeated as
"Invoice #:" on every later page) -- it's the same for every person on
this PDF. The per-person "Check#"/"Check No" column on the Fringe Recap
and Wage Register pages identifies one paycheck, not the invoice/batch;
an earlier AI-generated draft of this parser mistook it for the invoice
number.

Ded's Credited and ER Local Tax have no dedicated fringe field and are
always 0 on every real row seen so far -- folded into "other" along with
the real "Other" column rather than dropped. Total Fringe (a redundant
subtotal) isn't stored; "total" comes from Total Cost, always the last
value on the line regardless of what's ahead of it.
"""

import re
import io

import pdfplumber

from parsers.base import empty_row, parse_amount, clean_fringe_name

COMPANY  = "revolution_business_services"
MARKERS  = ["Revolution Business Services"]
PRIORITY = 10

_SS_RATIO  = 6.2  / 7.65
_MED_RATIO = 1.45 / 7.65

_SSN_RE      = re.compile(r"^\*+(\d{4})\s+(.+)$")
_INVOICE_RE  = re.compile(r"Invoice\s*(?:No\.?|#:)\s*(\S+)", re.IGNORECASE)
_PERIOD_RE   = re.compile(r"Period:?\s*(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})", re.IGNORECASE)
_INVDATE_RE  = re.compile(r"Invoice Date\s+(\S+)", re.IGNORECASE)


# ─── Fringe Recap parsing ─────────────────────────────────────────────────────

def _parse_fringe_page(text: str, invoice_no: str, period: str, inv_date: str) -> list[dict]:
    rows  = []
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

    i = 0
    while i < len(lines):
        line = lines[i]
        m = _SSN_RE.match(line)
        if not m:
            i += 1
            continue

        ssn_last4 = m.group(1)
        name_part = m.group(2).strip()
        # This vendor's Fringe Recap prints "First, Last" -- backwards from
        # the usual "Last, First" -- confirmed on every name checked across
        # multiple invoices (e.g. "SIERRA, BARTON", "WESLEY, DIXON" -- Sierra
        # and Wesley are the first names). Reverse it before formatting.
        if "," in name_part:
            first, _, last = name_part.partition(",")
            name_part = f"{last.strip()}, {first.strip()}"

        i += 1
        if i >= len(lines):
            break

        amounts = lines[i].split()
        if not (amounts and re.match(r"^\d{4,7}$", amounts[0])):
            continue

        def tok(idx):
            pos = idx if idx >= 0 else len(amounts) + idx
            return parse_amount(amounts[pos]) if 0 <= pos < len(amounts) else None

        row = empty_row()
        row["payrollCompany"] = "Revolution Business Services"
        row["worker"]      = clean_fringe_name(name_part, from_caps=name_part.isupper())
        row["ssn"]         = ssn_last4
        row["invoiceNo"]   = invoice_no
        row["invoiceDate"] = inv_date
        row["workDates"]   = period
        row["reimbRent"]   = tok(1)
        row["wages"]       = tok(2)
        row["vacHol"]      = tok(3)
        row["_ss_medi"]    = tok(4) or 0.0
        row["futa"]        = tok(5)
        row["sui"]         = tok(6)
        row["wc"]          = tok(7)
        row["hand"]        = tok(8)
        row["phw"]         = tok(9)
        row["adv"]         = tok(10)
        other_tail = [v for v in (tok(11), tok(12)) if v]
        row["other"]  = round(sum(other_tail), 2) if other_tail else None
        row["total"]  = tok(-1)
        row["_check_no"] = amounts[0]

        rows.append(row)
        i += 1

    return rows


# ─── Enrichment from Payroll Wage Register pages ─────────────────────────────

def _extract_register(pdf_bytes: bytes) -> dict:
    """Return {check_number: {street, city, zip, jobTitle, socSec, med}}."""
    reg = {}
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            wage_text = ""
            for pg in pdf.pages:
                text = pg.extract_text() or ""
                if "Payroll Wage Register" in text or "Occ Code-Desc" in text:
                    wage_text += text + "\n"
    except Exception:
        return reg

    blocks = re.split(r"(?=L/O Name:.*?Check No:)", wage_text)
    for block in blocks:
        m_check = re.search(r"Check No:\s*(\d+)", block)
        if not m_check:
            continue
        check_no = m_check.group(1).strip()

        m_job = re.search(r"Occ Code-Desc:\s*(.+?)(?:\s+Contract Occ Code)", block)
        job_title = m_job.group(1).strip() if m_job else ""

        fica_amounts = re.findall(r"FICA\s+([\d,]+\.\d{2})", block)
        med_amounts  = re.findall(r"Medicare\s+([\d,]+\.\d{2})", block)
        fica = parse_amount(fica_amounts[0]) if fica_amounts else None
        med  = parse_amount(med_amounts[0])  if med_amounts  else None

        m_addr = re.search(
            r"Mail To Address:\s*(.+?)\s*,\s*([^,]+?),?\s+[A-Z]{2}\s+([\d-]+)",
            block,
        )
        street = city = zip_code = ""
        if m_addr:
            street   = m_addr.group(1).strip()
            city     = m_addr.group(2).strip()
            zip_code = m_addr.group(3).strip()

        reg[check_no] = {
            "jobTitle": job_title, "socSec": fica, "med": med,
            "street": street, "city": city, "zip": zip_code,
        }

    return reg


# ─── Public API ───────────────────────────────────────────────────────────────

def extract(pdf_bytes: bytes, **_) -> tuple[list[dict], list[str]]:
    rows   = []
    issues = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            invoice_no = invoice_date = period = ""

            if pdf.pages:
                p1 = pdf.pages[0].extract_text() or ""
                m = _INVOICE_RE.search(p1)
                if m:
                    invoice_no = m.group(1).strip()
                m = _INVDATE_RE.search(p1)
                if m:
                    invoice_date = m.group(1).strip()

            for pg in pdf.pages:
                text = pg.extract_text() or ""
                if "Fringe Recap" not in text:
                    continue
                if not invoice_no:
                    m = _INVOICE_RE.search(text)
                    if m:
                        invoice_no = m.group(1).strip()
                if not period:
                    m = _PERIOD_RE.search(text)
                    if m:
                        period = f"{m.group(1)} to {m.group(2)}"
                rows.extend(_parse_fringe_page(text, invoice_no, invoice_date, period))

        if not rows:
            issues.append(
                "No Fringe Recap rows found — verify this is a Revolution "
                "Business Services PDF (1210 W Burbank Blvd)."
            )
            return rows, issues

        reg = _extract_register(pdf_bytes)
        for row in rows:
            info = reg.get(row.pop("_check_no", ""), {})
            ss_medi = row.pop("_ss_medi", 0.0)
            if info.get("socSec") is not None or info.get("med") is not None:
                row["socSec"] = info.get("socSec") or 0.0
                row["med"]    = info.get("med") or 0.0
            elif ss_medi:
                row["socSec"] = round(ss_medi * _SS_RATIO, 2)
                row["med"]    = round(ss_medi * _MED_RATIO, 2)
            row["jobTitle"] = info.get("jobTitle", "")
            row["street"]   = info.get("street", "")
            row["city"]     = info.get("city", "")
            row["zip"]      = info.get("zip", "")

    except Exception as e:
        issues.append(f"Error parsing PDF: {e}")

    return rows, issues
