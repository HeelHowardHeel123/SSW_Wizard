"""
The TEAM Companies — Crew Payroll Fringe parser.

Primary source: "Invoice Recap - Employer Cost" page — one data row per payee.

Enrichment pages (one type per invoice):
  A. Crew Pre Check Report          → job title, work dates, street/city/zip
  B. Extended Payment Breakdown     → job title, work dates (no address)
  C. Payroll History Report (PAY26) → job title, W/E date (order-based join, no address)
  D. Historical Payroll Register    → job title, work dates (order-based join, no address)

Name rule: prefer the enrichment page name when it produces a longer cleaned last name
(Invoice Recap truncates at ~20 chars; enrichment pages sometimes carry the full form).
"""

import re
import io

import pdfplumber

from parsers.base import (
    empty_row, parse_amount, clean_fringe_name, fmt_street_address,
)

# ── Registry metadata ─────────────────────────────────────────────────────────
COMPANY  = "the_team_companies"
MARKERS  = ["Invoice Recap - Employer Cost"]
PRIORITY = 10


# ── Name cleaning ─────────────────────────────────────────────────────────────

_SUFFIX_RE   = re.compile(r'\s+(Jr\.?|Sr\.?|II|III|IV|V)\s*$', re.IGNORECASE)
_PAY_TYPE_RE = re.compile(
    r'\s+(?:STRAIGHT\s+TIME|TIME\s*[&]\s*1/2|TIME\s*AND\s*1/2|BOX\s+(?:KIT\s+)?RENTAL'
    r'|MILEAGE|REIMB|PER\s+DIEM|OVERTIME|DOUBLE\s+TIME|PREP\s+TIME|WRAP\s+TIME'
    r'|EXP\s+NON|EXPENSE)\b',
    re.IGNORECASE,
)


def _clean_name(raw: str, is_loan_out: bool = False) -> str:
    if not raw:
        return ""
    if is_loan_out:
        return raw.strip().title()
    name = clean_fringe_name(raw.strip(), from_caps=True)
    return _SUFFIX_RE.sub("", name).strip()


def _resolve_name(recap_raw: str, enrich_raw: str | None, is_loan_out: bool = False) -> str:
    """Use whichever source yields the longer cleaned name (handles truncation on either side).
    For loan-outs the Invoice Recap name is the company name — always keep it."""
    recap_clean = _clean_name(recap_raw, is_loan_out)
    if not enrich_raw or is_loan_out:
        return recap_clean
    enrich_clean = _clean_name(enrich_raw, is_loan_out)
    return enrich_clean if len(enrich_clean) >= len(recap_clean) else recap_clean


def _strip_pay_cols(line: str) -> str:
    """Extract the name/occupation token from a Pre Check line that includes pay columns."""
    m = _PAY_TYPE_RE.search(line)
    return line[:m.start()].strip() if m else line.strip()


# ── Page classification ───────────────────────────────────────────────────────

def _classify(text: str) -> str:
    if "Invoice Recap - Employer Cost" in text:
        return "recap"
    if "Crew Pre Check Report" in text:
        return "precheck"
    if "Extended Payment Breakdown" in text:
        return "extended"
    if "PAY26" in text or "Payroll History Report" in text:
        return "payrollhist"
    if "Historical Payroll Register" in text:
        return "histregister"
    if "Work Start:" in text and "Balance Due" in text:
        return "cover"
    return "other"


# ── Cover page ────────────────────────────────────────────────────────────────

def _parse_cover(text: str) -> dict:
    def _find(pat):
        m = re.search(pat, text)
        return m.group(1) if m else ""
    return {
        "invoice_no": _find(r'Invoice\s*#:\s*(\d+)'),
        "inv_date":   _find(r'Inv\.\s*Date:\s*(\d{2}/\d{2}/\d{2,4})'),
        "work_start": _find(r'Work\s*Start:\s*(\d{2}/\d{2}/\d{2,4})'),
        "work_end":   _find(r'Work\s*End:\s*(\d{2}/\d{2}/\d{2,4})'),
    }


# ── Invoice Recap ─────────────────────────────────────────────────────────────

_RECAP_INV_RE  = re.compile(r'Invoice[\s.]+(\d{7,9})')
_RECAP_DATE_RE = re.compile(r'Inv\s*Date\s*[.\s]+(\d{2}/\d{2}/\d{2,4})')
_SSN5_RE       = re.compile(r'\*{5}(\d{4})')


def _parse_recap_row(line: str, page_no: int) -> dict | None:
    m = _SSN5_RE.search(line)
    if not m:
        return None
    ssn4     = m.group(1)
    name_raw = line[:m.start()].strip()
    if not name_raw or re.match(r'^TOTAL', name_raw, re.IGNORECASE):
        return None

    rest   = line[m.end():].strip()
    tokens = rest.split()
    if len(tokens) < 2:
        return None

    ws        = tokens[0]
    remaining = tokens[1:]

    # Split remaining on the Res State code (2-letter uppercase, after ≥10 numeric tokens)
    before_rs, rs, after_rs = [], "", []
    for tok in remaining:
        if not rs and re.match(r'^[A-Z]{2}$', tok) and len(before_rs) >= 10:
            rs = tok
        elif not rs:
            before_rs.append(tok)
        else:
            after_rs.append(tok)

    def _a(lst, i):
        return parse_amount(lst[i]) if i < len(lst) else None

    row = empty_row()
    row["payrollCompany"] = "the_team_companies"
    row["ssn"]       = ssn4
    row["workState"] = ws
    row["resState"]  = rs
    # Columns: TAXABLE, CORP, REIMB, FICA, MEDICARE, FUTA, SUI, SDI(skip), OTHER(skip), WC
    row["wages"]     = _a(before_rs, 0)
    row["corporate"] = _a(before_rs, 1)
    row["reimbRent"] = _a(before_rs, 2)
    row["socSec"]    = _a(before_rs, 3)
    row["med"]       = _a(before_rs, 4)
    row["futa"]      = _a(before_rs, 5)
    row["sui"]       = _a(before_rs, 6)
    row["wc"]        = _a(before_rs, 9)
    # After RS: PH&W, HANDLING, DEDUCTIONS(skip), TOTAL
    row["phw"]   = _a(after_rs, 0)
    row["hand"]  = _a(after_rs, 1)
    row["total"] = (
        _a(after_rs, 3) if len(after_rs) >= 4 else
        parse_amount(after_rs[-1]) if after_rs else None
    )

    corp       = row.get("corporate") or 0
    is_loan_out = corp > 0 or "," not in name_raw
    row["loanOut"] = is_loan_out
    if is_loan_out:
        row["loanOutCompany"] = _clean_name(name_raw, is_loan_out=True)
    row["worker"]     = _clean_name(name_raw, is_loan_out)
    row["_name_raw"]  = name_raw
    row["sourcePage"] = page_no
    return row


def _parse_recap_page(text: str, page_no: int) -> tuple[str, str, list[dict]]:
    inv_no   = (m.group(1) if (m := _RECAP_INV_RE.search(text))  else "")
    inv_date = (m.group(1) if (m := _RECAP_DATE_RE.search(text)) else "")
    rows = []
    for line in text.split("\n"):
        row = _parse_recap_row(line.strip(), page_no)
        if row:
            rows.append(row)
    return inv_no, inv_date, rows


# ── Enrichment A: Crew Pre Check Report ──────────────────────────────────────

_FROM_TO_RE  = re.compile(r'From\s+(\d{2}/\d{2}/\d{2,4})\s+To\s+(\d{2}/\d{2}/\d{2,4})', re.IGNORECASE)
_TOTAL_HR_RE = re.compile(r'^Total\s+Hours', re.IGNORECASE)
_STARS_SEP   = re.compile(r'^\*{20,}')


def _parse_precheck_page(text: str) -> tuple[str, list[dict]]:
    """Returns (invoice_no, [ordered employee dicts keyed by ssn4])."""
    inv_m  = re.search(r'Invoice[\s.]+(\d{7,9})', text)
    inv_no = inv_m.group(1) if inv_m else ""

    lines     = text.split("\n")
    employees = []
    i         = 0

    while i < len(lines):
        ssn_m = _SSN5_RE.search(lines[i])
        # SSN line: not the separator row (which is all stars)
        if ssn_m and not _STARS_SEP.match(lines[i].strip()):
            ssn4 = ssn_m.group(1)

            # Name: 2 lines above; occupation: 1 line above — strip trailing pay columns
            name_raw   = _strip_pay_cols(lines[i - 2]) if i >= 2 else ""
            occupation = _strip_pay_cols(lines[i - 1]).title() if i >= 1 else ""

            # Work dates: first "From ... To ..." after SSN line
            work_dates = ""
            j = i + 1
            while j < len(lines) and not _FROM_TO_RE.search(lines[j]):
                if _STARS_SEP.match(lines[j].strip()):
                    break
                j += 1
            if j < len(lines):
                fm = _FROM_TO_RE.search(lines[j])
                if fm:
                    work_dates = f"{fm.group(1)} - {fm.group(2)}"

            # Address: after the "NON-UNION" line (the all-caps one), before "Total Hours"
            street, city, zip_code = "", "", ""
            k        = j + 1
            in_addr  = False
            addr_acc = []
            while k < len(lines):
                ln = lines[k].strip()
                if re.match(r'^NON-UNION$', ln):
                    in_addr = True
                    k += 1
                    continue
                if _TOTAL_HR_RE.match(ln) or _STARS_SEP.match(ln):
                    break
                if in_addr and ln and not re.match(
                    r'^(NON\s*UNION|Marital\s*Status|MARITAL)', ln, re.IGNORECASE
                ):
                    addr_acc.append(ln)
                k += 1

            if addr_acc:
                street, city, zip_code = _parse_addr_lines(addr_acc)

            employees.append({
                "ssn4":       ssn4,
                "name_raw":   name_raw,
                "occupation": occupation,
                "work_dates": work_dates,
                "street":     street,
                "city":       city,
                "zip":        zip_code,
            })
        i += 1

    return inv_no, employees


def _parse_addr_lines(addr_lines: list[str]) -> tuple[str, str, str]:
    street_parts, city, zip_code = [], "", ""
    for ln in addr_lines:
        ln = ln.strip()
        if not ln:
            continue
        if re.match(r'^\d{5}$', ln):
            zip_code = ln
        elif re.search(r',\s*[A-Z]{2}\s*$', ln):
            parts = ln.rsplit(',', 1)
            city  = parts[0].strip().title() + ', ' + parts[1].strip()
        elif re.match(r'^[\dA-Z]{1,5}$', ln):
            pass  # standalone unit (e.g. "3", "2E", "446") — skip
        else:
            street_parts.append(ln)

    raw_street = ' '.join(street_parts)
    street     = fmt_street_address(raw_street)
    # Strip trailing ", UNIT" artifact (e.g. ", 7K" left after fmt)
    street = re.sub(r',\s*\w{1,5}$', '', street).strip()
    return street, city, zip_code


# ── Enrichment B: Extended Payment Breakdown ─────────────────────────────────

_EXT_SSN_RE = re.compile(r'^\*{3}-\*{2}-(\d{4})\s+(.+?)(?:\s*\|\s*(.+))?$')


def _parse_extended_page(text: str) -> tuple[str, list[dict]]:
    """Returns (invoice_no, [ordered employee dicts])."""
    inv_m  = re.search(r'Invoice#:\s*(\d+)', text)
    inv_no = inv_m.group(1) if inv_m else ""

    employees = []
    current   = None

    for line in text.split("\n"):
        line = line.strip()
        em   = _EXT_SSN_RE.match(line)
        if em:
            ssn4       = em.group(1)
            name_raw   = em.group(2).strip()
            occupation = (em.group(3) or "").strip().title()
            current    = {
                "ssn4": ssn4, "name_raw": name_raw,
                "occupation": occupation, "work_dates": "",
            }
            employees.append(current)
            continue

        # "Total NAME" line → done with this employee
        if re.match(r'^Total\s+\S', line, re.IGNORECASE) and current:
            current = None
            continue

        if current and not current["work_dates"]:
            fm = _FROM_TO_RE.search(line)
            if fm:
                current["work_dates"] = f"{fm.group(1)} - {fm.group(2)}"

    return inv_no, employees


# ── Enrichment C: Payroll History Report (PAY26) ─────────────────────────────

def _parse_payrollhist(text: str) -> tuple[str, list[dict]]:
    """Returns (invoice_no, [ordered employee dicts])."""
    inv_m  = re.search(r'Invoice[^\n\d]{0,30}(\d{7,9})', text)
    inv_no = inv_m.group(1) if inv_m else ""

    employees = []
    blocks    = re.split(r'-{30,}', text)

    for block in blocks:
        lines = [ln.strip() for ln in block.split("\n") if ln.strip()]
        if not lines:
            continue
        first = lines[0]
        if "M/S:" not in first:
            continue
        name_raw = first[:first.index("M/S:")].strip()
        if not name_raw or re.match(r'^GRAND\s*TOTAL', name_raw, re.IGNORECASE):
            continue

        occupation = ""

        for ln in lines:
            if "DM#:" in ln:
                # Line: COMPANY PROJECT UNION OCCUPATION DM#: N
                # pdfplumber collapses to single spaces — scan backwards for NONE or union number
                prefix = ln[:ln.index("DM#:")].strip()
                tokens = prefix.split()
                occ_start = None
                for i in range(len(tokens) - 1, -1, -1):
                    if tokens[i].upper() == "NONE" or re.match(r'^\d{3,4}$', tokens[i]):
                        occ_start = i + 1
                        break
                if occ_start is not None and occ_start < len(tokens):
                    occupation = " ".join(tokens[occ_start:]).strip().title()
                break

        employees.append({
            "name_raw":   name_raw,
            "occupation": occupation,
            "work_dates": "",  # W/E date alone is not a range; caller uses inv_work_dates
        })

    return inv_no, employees


# ── Enrichment D: Historical Payroll Register ─────────────────────────────────

_HIST_SSN_RE = re.compile(r'\*{2}-\*{3}(\d{4})')


def _parse_histregister(text: str) -> tuple[str, list[dict]]:
    """Returns (invoice_no, [ordered employee dicts])."""
    inv_m  = re.search(r'Invoice[^\n\d]{0,30}(\d{7,9})', text)
    inv_no = inv_m.group(1) if inv_m else ""

    employees = []
    current   = None
    lines     = [ln.strip() for ln in text.split("\n") if ln.strip()]

    for line in lines:
        ssn_m = _HIST_SSN_RE.search(line)
        if ssn_m and "Occupation:" not in line:
            name_raw = line[:ssn_m.start()].strip()
            if name_raw and not re.match(r'^(Employee Name|TOTAL)', name_raw, re.IGNORECASE):
                current = {"name_raw": name_raw, "occupation": "", "work_dates": ""}
                employees.append(current)

        if current:
            om = re.search(r'Occupation:\s*([A-Z][A-Z\s/\-]+?)(?:\s+Check#:|$)', line, re.IGNORECASE)
            if om:
                current["occupation"] = om.group(1).strip().title()

            wm = re.search(r'Work\s*Dates?:([\d,/]+)', line, re.IGNORECASE)
            if wm:
                parts = [d.strip() for d in wm.group(1).split(',') if d.strip()]
                if parts:
                    current["work_dates"] = f"{parts[0]} - {parts[-1]}"

            if not current.get("work_dates"):
                cm = re.search(r'Check\s*Date:\s*(\d{2}/\d{2}/\d{4})', line)
                if cm:
                    current["work_dates"] = cm.group(1)

    return inv_no, employees


# ── Invoice number extraction helper ─────────────────────────────────────────

def _inv_no_from(text: str, pat: str) -> str:
    m = re.search(pat, text)
    return m.group(1) if m else ""


# ── Public API ────────────────────────────────────────────────────────────────

def extract(pdf_bytes: bytes, source_file: str = "", **_) -> tuple[list[dict], list[str]]:
    rows, issues = [], []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            cover_meta = {}   # inv_no → {inv_date, work_start, work_end}
            recap_pages = []  # (inv_no, inv_date, page_no, [rows])

            # Accumulate enrichment page text per invoice (pages may span across boundaries)
            precheck_texts    = {}  # inv_no → [page_text, ...]
            extended_texts    = {}  # inv_no → [page_text, ...]
            payrollhist_texts = {}  # inv_no → [page_text, ...]
            histregister_texts= {}  # inv_no → [page_text, ...]

            for pg_idx, pg in enumerate(pdf.pages):
                text      = pg.extract_text() or ""
                page_no   = pg_idx + 1
                page_type = _classify(text)

                if page_type == "cover":
                    meta = _parse_cover(text)
                    if meta["invoice_no"]:
                        cover_meta[meta["invoice_no"]] = meta

                elif page_type == "recap":
                    inv_no, inv_date, page_rows = _parse_recap_page(text, page_no)
                    if inv_no:
                        recap_pages.append((inv_no, inv_date, page_no, page_rows))

                elif page_type == "precheck":
                    inv_no = _inv_no_from(text, r'Invoice[\s.]+(\d{7,9})')
                    if inv_no:
                        precheck_texts.setdefault(inv_no, []).append(text)

                elif page_type == "extended":
                    inv_no = _inv_no_from(text, r'Invoice#:\s*(\d+)')
                    if inv_no:
                        extended_texts.setdefault(inv_no, []).append(text)

                elif page_type in ("payrollhist", "histregister"):
                    inv_no = _inv_no_from(text, r'Invoice[^\n\d]{0,30}(\d{7,9})')
                    if inv_no:
                        bucket = (payrollhist_texts if page_type == "payrollhist"
                                  else histregister_texts)
                        bucket.setdefault(inv_no, []).append(text)

            # Parse enrichment (concatenate multi-page blocks first)
            ordered_enr = {}  # inv_no → [emp dicts] (order-based join)
            ssn_enr     = {}  # inv_no → {ssn4: [emp dicts]} (SSN-based join)

            for inv_no, texts in precheck_texts.items():
                for text in texts:
                    _, emps = _parse_precheck_page(text)
                    bucket  = ssn_enr.setdefault(inv_no, {})
                    for emp in emps:
                        bucket.setdefault(emp["ssn4"], []).append(emp)

            for inv_no, texts in extended_texts.items():
                _, emps = _parse_extended_page("\n".join(texts))
                bucket  = ssn_enr.setdefault(inv_no, {})
                for emp in emps:
                    bucket.setdefault(emp["ssn4"], []).append(emp)

            for inv_no, texts in payrollhist_texts.items():
                _, emps = _parse_payrollhist("\n".join(texts))
                ordered_enr.setdefault(inv_no, []).extend(emps)

            for inv_no, texts in histregister_texts.items():
                _, emps = _parse_histregister("\n".join(texts))
                ordered_enr.setdefault(inv_no, []).extend(emps)

            # Assemble final output rows
            for inv_no, inv_date, page_no, recap_rows in recap_pages:
                meta = cover_meta.get(inv_no, {})
                ws   = meta.get("work_start", "")
                we   = meta.get("work_end", "")
                inv_work_dates = f"{ws} - {we}" if ws else ""
                if not inv_date:
                    inv_date = meta.get("inv_date", "")

                has_ordered = inv_no in ordered_enr
                ssn_bucket  = ssn_enr.get(inv_no, {})
                ssn_seen    = {}

                for emp_idx, row in enumerate(recap_rows):
                    ssn4     = row.get("ssn", "")
                    is_lo    = row.get("loanOut", False)
                    name_raw = row.pop("_name_raw", "")

                    enrich = {}
                    if has_ordered:
                        ordered = ordered_enr[inv_no]
                        if emp_idx < len(ordered):
                            enrich = ordered[emp_idx]
                    else:
                        occ_num  = ssn_seen.get(ssn4, 0)
                        ssn_list = ssn_bucket.get(ssn4, [])
                        if occ_num < len(ssn_list):
                            enrich = ssn_list[occ_num]
                        ssn_seen[ssn4] = occ_num + 1

                    row["worker"]           = _resolve_name(
                        name_raw, enrich.get("name_raw", ""), is_lo
                    )
                    row["jobTitle"]         = enrich.get("occupation", "")
                    row["workDates"]        = enrich.get("work_dates", "") or inv_work_dates
                    row["street"]           = enrich.get("street", "")
                    row["city"]             = enrich.get("city", "")
                    row["zip"]              = enrich.get("zip", "")
                    row["invoiceNo"]        = inv_no
                    row["invoiceDate"]      = inv_date
                    row["invoiceWorkDates"] = inv_work_dates
                    row["sourceFile"]       = source_file
                    rows.append(row)

    except Exception as e:
        issues.append(str(e))

    if not rows:
        issues.append(
            "No Invoice Recap rows found — verify this is a Teams crew payroll PDF."
        )

    return rows, issues
