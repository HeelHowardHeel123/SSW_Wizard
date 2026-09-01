"""
Wrapbook "EMPLOYER PAYROLL" thin invoice parser.

A second, structurally different Wrapbook invoice layout seen on some jobs
(e.g. PGC 005) -- these invoices have NO "Fringe Report" page at all (that's
fringe_001's layout, and its marker). Instead: page 1 is a plain invoice
summary, page 2 is an "EMPLOYER PAYROLL" invoice-level category-totals page,
and page 3 onward is a per-person wage/fringe breakdown table -- sometimes
one person (a single off-cycle payment), sometimes dozens (a full weekly
batch run), with no page break between people.

Deliberately thin: the per-person table only carries a LUMP fringe total,
not a FICA/Medicare/FUTA/SUI/W-C breakdown per person, and there's no
loan-out indicator anywhere in this layout (no company sub-line, no explicit
flag). Whenever this layout is in play, the job also always has a Production
Report, which is the real source for fringe/job-title detail and loan-out
status -- this parser only needs to establish presence and a per-row dollar
total for the Automation Total cross-check: worker, invoiceNo, invoiceDate,
total. Everything else (jobTitle included) is left blank for the Production
Report / reconciler's blank-field fallback to fill in, and loanOut is left
at its default False, same as anywhere else the Production Report is silent
on it.

The breakdown table has no ruled grid (pdfplumber's extract_tables() finds
nothing), column widths visibly shift from invoice to invoice (a "Job
Title"/"Hours" boundary seen at x=146 on one invoice sits past x=194 on
another), and long names wrap their first name onto its own line below --
e.g. "Aguilar, Craft Services ... $2,928.95" followed by a lone "Gabriela"
on the next line in the Name column position. Given that column-width
instability, this parser only trusts two things about each data line's
x-position: the Name column always starts at the far left (x0 < 83 in every
invoice seen so far), and the row's Wage Total is always the LAST dollar
amount on the line, regardless of which earlier columns (Expenses, Agent
Fee, ...) are blank for that person. Extracting by x-position rather than
treating each visual line as one row also lets a line with only Name-column
text and no hours/dollar tokens be recognized as a name-wrap continuation,
glued onto the previous row instead of read as its own row.
"""

import re
import io

import pdfplumber

from parsers.base import empty_row, clean_fringe_name

COMPANY  = "wrapbook"
MARKERS  = ["EMPLOYER PAYROLL", "TakeOne Network Corp"]
PRIORITY = 10

_PAYROLL_NO_RE = re.compile(r"Payroll (?:ID|#):\s*0*(\d+)")
_PAY_DATE_RE   = re.compile(r"Pay Date:\s*(\d{1,2}/\d{1,2}/\d{4})")
_DOLLAR_RE     = re.compile(r"^\$[\d,]+\.\d{2}$")

_NAME_MAX_X = 83


def _parse_amount(word: str):
    if not word:
        return None
    try:
        return round(float(word.replace("$", "").replace(",", "")), 2)
    except Exception:
        return None


def _group_lines(words: list[dict]) -> list[list[dict]]:
    """Group words into visual lines by their 'top' position."""
    lines: list[list[dict]] = []
    current: list[dict] = []
    current_top = None
    for w in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
        top = round(w["top"], 1)
        if current_top is None or abs(top - current_top) <= 2.5:
            current.append(w)
            current_top = current_top if current_top is not None else top
        else:
            lines.append(current)
            current = [w]
            current_top = top
    if current:
        lines.append(current)
    return lines


def _parse_breakdown_page(words: list[dict]) -> list[dict]:
    rows: list[dict] = []
    for line in _group_lines(words):
        name_tokens   = [w["text"] for w in line if w["x0"] < _NAME_MAX_X]
        dollar_tokens = [w["text"] for w in line if _DOLLAR_RE.match(w["text"])]
        has_hours     = any("hrs" in w["text"].lower() for w in line)

        if not dollar_tokens and not has_hours:
            # Name-only continuation line (a wrapped first/last name) --
            # glue it onto the previous row instead of treating it as its
            # own row.
            if rows and name_tokens:
                rows[-1]["name_frag"] = (rows[-1]["name_frag"] + " " + " ".join(name_tokens)).strip()
            continue

        name_text = " ".join(name_tokens).strip()
        if not name_text or name_text.lower() == "total":
            continue  # the invoice/section grand-total line, not a hire

        rows.append({
            "name_frag": name_text,
            "total":     _parse_amount(dollar_tokens[-1]) if dollar_tokens else None,
        })
    return rows


def extract(pdf_bytes: bytes, **kwargs) -> tuple[list[dict], list[str]]:
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if len(pdf.pages) < 3:
                return [], [
                    "This Wrapbook invoice doesn't have the expected per-person "
                    "breakdown page (page 3 onward) -- verify it's the EMPLOYER "
                    "PAYROLL layout, not a different Wrapbook format."
                ]

            header_text = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:2])
            m_no   = _PAYROLL_NO_RE.search(header_text)
            m_date = _PAY_DATE_RE.search(header_text)
            invoice_no   = m_no.group(1) if m_no else ""
            invoice_date = m_date.group(1) if m_date else ""

            rows: list[dict] = []
            for pg in pdf.pages[2:]:
                # A tight x_tolerance matters here -- the default merges the
                # end of a short Name (e.g. "Aaron") straight into the next
                # column's Job Title ("Production...") whenever the gap
                # between them is small, producing one bogus word like
                # "AaronProduction" that corrupts the matched-against name.
                for parsed in _parse_breakdown_page(pg.extract_words(x_tolerance=0.3)):
                    row = empty_row()
                    row["worker"]         = clean_fringe_name(parsed["name_frag"])
                    row["total"]          = parsed["total"]
                    row["invoiceNo"]      = invoice_no
                    row["invoiceDate"]    = invoice_date
                    row["payrollCompany"] = "wrapbook"
                    rows.append(row)
    except Exception as e:
        return [], [f"Failed to parse Wrapbook EMPLOYER PAYROLL invoice: {e}"]

    if not rows:
        return [], ["No per-person breakdown rows found on this Wrapbook invoice."]
    return rows, []
