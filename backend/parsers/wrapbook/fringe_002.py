"""
Wrapbook "EMPLOYER PAYROLL" thin invoice parser.

A second, structurally different Wrapbook invoice layout seen on some jobs
(e.g. PGC 005) -- these invoices have NO "Fringe Report" page at all (that's
fringe_001's layout, and its marker). Instead: page 1 is a plain invoice
summary, page 2 is an "EMPLOYER PAYROLL" invoice-level category-totals page,
and page 3 onward is a per-person wage/fringe breakdown table -- sometimes
one person (a single off-cycle payment), sometimes dozens (a full weekly
batch run), with no page break between people.

Some invoices (confirmed real, e.g. ABBVIE 022) tack on further pages after
the breakdown table ends -- an "Invoice Fee Summary" page seen so far, which
lists FEE CATEGORIES (Employer FICA, Federal Unemployment, Workers'
Compensation, ...) in the same "short text, then a dollar amount" shape a
person row has. Naively reading every page from 3 onward produced fake
"employees" named after fee categories. Fixed by stopping the page loop
entirely the moment a page's own heading matches a known trailing-section
title (_TRAILING_SECTION_TITLES below) -- not a per-page allowlist the way
caps/fringe_001.py's "Fringe Recap Report" marker works, because a
continuation page of the real breakdown table carries NO marker of its own
(confirmed real, a 56-person invoice: page 4 continues straight into more
names with no repeated column header) -- only the boundary INTO a trailing
section is reliably self-labeled, so that's the only thing checked for.
Extend the title list as new trailing sections turn up on real invoices.

Deliberately thin: the per-person table only carries a LUMP fringe total,
not a FICA/Medicare/FUTA/SUI/W-C breakdown per person, and there's no
loan-out indicator anywhere in this layout (no company sub-line, no explicit
flag). Whenever this layout is in play, the job also always has a Production
Report, which is the real source for fringe detail and loan-out status --
this parser only needs to establish presence and a per-row dollar total for
the Automation Total cross-check: worker, invoiceNo, invoiceDate, total.
Everything else besides jobTitle is left blank for the Production Report /
reconciler's blank-field fallback to fill in, and loanOut is left at its
default False, same as anywhere else the Production Report is silent on it.

jobTitle IS extracted here, from the same breakdown table (a "Job Title"
column sits between Name and Hours) -- confirmed real, e.g. ABBVIE 022,
where this layout's Production Report has no job-title column of its own at
all, so the "Production Report fills it in" assumption above doesn't hold
for every job that uses this invoice layout. See _job_title_tokens for how
the column boundary is found without trusting the unstable x-position the
module docstring above warns about elsewhere.

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

_PAYROLL_NO_RE  = re.compile(r"Payroll (?:ID|#):\s*0*(\d+)")
_PAY_DATE_RE    = re.compile(r"Pay Date:\s*(\d{1,2}/\d{1,2}/\d{4})")
_DOLLAR_RE      = re.compile(r"^\$[\d,]+\.\d{2}$")
_PAGE_FOOTER_RE = re.compile(r"^Page\s+\d+\s+of\s+\d+$", re.IGNORECASE)

_NAME_MAX_X = 83

# Pages seen after the real per-person breakdown table ends -- once one of
# these titles appears, every page from there on is some other report
# section, never more people. See the module docstring for why this is a
# stop condition rather than a per-page allowlist check.
_TRAILING_SECTION_TITLES = ("invoice fee summary",)


def _is_trailing_section(text: str) -> bool:
    norm = " ".join(text.lower().split())
    return any(title in norm for title in _TRAILING_SECTION_TITLES)


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


_LONE_INITIAL_RE = re.compile(r"^[A-Z]\.$")


def _job_title_tokens(rest: list[dict]) -> list[str]:
    """Tokens in the Job Title column: everything in the non-Name region of a
    line, UP TO (not including) the Hours/dollar columns that follow it.
    `rest` is already left-to-right ordered (see _group_lines' sort), so this
    just walks it and stops at the first boundary marker. Hours are rendered
    as two separate words ("12", "hrs"), so a bare integer immediately
    followed by an "hrs" word is excluded too -- otherwise the leading digits
    of the Hours column would glue onto the end of the title. A lone leading
    "K."-style token is a middle initial that spilled past the fixed name
    boundary on a long name (confirmed real, "Fox-Mills, Finn K." -- the
    initial's x0 lands just past _NAME_MAX_X on that row), not the start of
    a real job title, so it's dropped rather than prefixed onto the title."""
    out = []
    for idx, w in enumerate(rest):
        t = w["text"]
        if _DOLLAR_RE.match(t) or "hrs" in t.lower():
            break
        if re.fullmatch(r"\d+", t) and idx + 1 < len(rest) and "hrs" in rest[idx + 1]["text"].lower():
            break
        if idx == 0 and _LONE_INITIAL_RE.match(t):
            continue
        out.append(t)
    return out


def _parse_breakdown_page(words: list[dict]) -> list[dict]:
    rows: list[dict] = []
    for line in _group_lines(words):
        line_text = " ".join(w["text"] for w in line).strip()
        if _PAGE_FOOTER_RE.match(line_text):
            # A page footer ("Page 3 of 3") has no $/hrs tokens of its own, so
            # without this it reads as a continuation line and glues onto
            # whichever row happens to be last on the page (confirmed real:
            # "Production Assistant Page 3 of 3" on the last row of a page).
            continue

        name_tokens   = [w["text"] for w in line if w["x0"] < _NAME_MAX_X]
        rest          = [w for w in line if w["x0"] >= _NAME_MAX_X]
        dollar_tokens = [w["text"] for w in rest if _DOLLAR_RE.match(w["text"])]
        has_hours     = any("hrs" in w["text"].lower() for w in rest)
        job_tokens    = _job_title_tokens(rest)

        if not dollar_tokens and not has_hours:
            # Name/title continuation line (a wrapped first/last name and/or
            # a wrapped job title) -- glue onto the previous row.
            if rows and (name_tokens or job_tokens):
                if name_tokens:
                    rows[-1]["name_frag"] = (rows[-1]["name_frag"] + " " + " ".join(name_tokens)).strip()
                if job_tokens:
                    rows[-1]["job_frag"] = (rows[-1]["job_frag"] + " " + " ".join(job_tokens)).strip()
            continue

        name_text = " ".join(name_tokens).strip()
        if not name_text or name_text.lower() == "total":
            continue  # the invoice/section grand-total line, not a hire

        rows.append({
            "name_frag": name_text,
            "job_frag":  " ".join(job_tokens).strip(),
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
                # Stop entirely once a page turns out to be a trailing report
                # section (e.g. "Invoice Fee Summary") rather than more of the
                # breakdown table -- see _TRAILING_SECTION_TITLES and the
                # module docstring. A continuation page of the real table has
                # no marker of its own, so this can only ever fire on a page
                # that's genuinely something else; it never cuts the real
                # table short.
                if _is_trailing_section(pg.extract_text() or ""):
                    break
                # A tight x_tolerance matters here -- the default merges the
                # end of a short Name (e.g. "Aaron") straight into the next
                # column's Job Title ("Production...") whenever the gap
                # between them is small, producing one bogus word like
                # "AaronProduction" that corrupts the matched-against name.
                for parsed in _parse_breakdown_page(pg.extract_words(x_tolerance=0.3)):
                    row = empty_row()
                    row["worker"]         = clean_fringe_name(parsed["name_frag"])
                    row["jobTitle"]       = " ".join(parsed["job_frag"].split())
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
