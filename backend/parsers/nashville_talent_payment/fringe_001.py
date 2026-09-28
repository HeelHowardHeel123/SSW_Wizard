"""
Nashville Talent Payment, Inc. fringe report parser.

One page per invoice: "NAME corp JOB WAGES OTHER TAXES NTP TOTAL" columns,
up to ~25 pre-printed blank employee rows per page (only the filled ones
matter). "corp" is a literal "Y" for loan-out/corp-paid people.

A row only prints a number for OTHER/TAXES/NTP when it's nonzero, so most
rows show 2-3 numbers between WAGES and TOTAL rather than all 3 -- and
which of the three is missing varies row to row (confirmed against rows
where all 5 columns did print: e.g. TAXES is ~15% of wages and NTP ~2%,
consistently, so a "4-number" row is WAGES+TAXES+NTP with OTHER at zero,
not WAGES+OTHER+TAXES). Rather than guess the omitted column per row,
every present middle value is summed into `other`, which reconciles
wages+other=total exactly regardless of which column(s) were blank.
`corporate` is left at 0 -- this vendor doesn't report a separate loan-out
dollar amount -- and the corp flag is instead carried as `loanOut`.
"""

import re
import io
from datetime import datetime

import pdfplumber

from parsers.base import empty_row, parse_amount, clean_fringe_name

COMPANY  = "nashville_talent_payment"
MARKERS  = ["NASHVILLE TALENT PAYMENT", "THIS INVOICE IS FOR PRODUCTION SALARIES AS LISTED"]
PRIORITY = 10

_AMOUNT_RE = re.compile(r"[\d,]+\.\d{2}")
_INVOICE_HEADER_RE = re.compile(r"(\w+\s+\d+,\s+\d{4})\s+INV\s+(\S+)")
_WORK_DATES_RE = re.compile(r"TX DATES:\s*(.*)")
_COLUMN_HEADER_RE = re.compile(r"\bNAME\b.*\bJOB\b.*\bWAGES\b.*\bTOTAL\b")


def extract(pdf_bytes: bytes, **_) -> tuple[list[dict], list[str]]:
    rows: list[dict] = []
    issues: list[str] = []

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ""
    except Exception as e:
        return rows, [f"Error opening PDF: {e}"]

    invoice_no = ""
    invoice_date = ""
    m = _INVOICE_HEADER_RE.search(text)
    if m:
        invoice_no = m.group(2)
        try:
            invoice_date = datetime.strptime(m.group(1).strip(), "%B %d, %Y").strftime("%m/%d/%Y")
        except Exception:
            invoice_date = m.group(1).strip()

    work_dates = ""
    m = _WORK_DATES_RE.search(text)
    if m:
        work_dates = m.group(1).strip()

    lines = text.split("\n")
    header_idx = next((i for i, line in enumerate(lines) if _COLUMN_HEADER_RE.search(line)), None)
    if header_idx is None:
        return rows, ["Could not find header line with NAME/JOB/WAGES/TOTAL columns"]

    for line in lines[header_idx + 1:]:
        stripped = line.strip()
        if not stripped or re.match(r"^Total\b", stripped, re.IGNORECASE):
            continue
        if re.match(r"^[\d\s.,]+$", stripped):
            continue  # blank pre-printed row (all-zero columns, no name)

        amounts = _AMOUNT_RE.findall(line)
        if len(amounts) < 3:
            continue

        text_part = re.sub(r"\s+", " ", _AMOUNT_RE.sub("", line)).strip()
        if not text_part or not re.match(r"^[A-Z]", text_part):
            continue

        tokens = text_part.split()
        if len(tokens) < 2:
            continue
        name_tokens, rest_tokens = tokens[:2], tokens[2:]

        corp_flag = bool(rest_tokens) and rest_tokens[0] == "Y"
        if corp_flag:
            rest_tokens = rest_tokens[1:]
        job_title = " ".join(rest_tokens).strip()

        wages = parse_amount(amounts[0])
        total = parse_amount(amounts[-1])
        corporate = 0.0
        other_val = round(sum(parse_amount(a) or 0.0 for a in amounts[1:-1]), 2)

        last_name = name_tokens[1] if len(name_tokens) > 1 else ""
        worker = clean_fringe_name(f"{last_name}, {name_tokens[0]}" if last_name else name_tokens[0], from_caps=True)

        row = empty_row()
        row.update({
            "worker": worker,
            "invoiceNo": invoice_no,
            "invoiceDate": invoice_date,
            "workDates": work_dates,
            "wages": wages,
            "corporate": corporate,
            "other": other_val,
            "total": total,
            "jobTitle": job_title,
            "loanOut": corp_flag,
        })
        rows.append(row)

    return rows, issues
