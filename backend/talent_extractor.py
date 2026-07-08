"""
Talent & Extras extraction for TPC Production Binder Wizard.

Converts Extreme Reach talent invoice PDFs + PTIP Excel into workbook rows
for the Talent & Extras tab.  Three intake scenarios:
  A  PDFs only (no PTIP)
  B  PTIP only (no PDFs)
  C  Both PDFs and PTIP (primary path)

The PTIP is authoritative for financial data; PDFs win on wages/misc placement
and supply work dates.  The backend returns fully-assembled rows in workbook
column order.  It never touches the .xlsx.
"""

import io
import re
import base64
from collections import defaultdict
from datetime import datetime

import openpyxl
import pdfplumber


# ── Small helpers ────────────────────────────────────────────────────────────

def _to_float(v) -> float:
    if v is None:
        return 0.0
    try:
        return round(float(str(v).replace(',', '').replace('$', '').strip()), 2)
    except Exception:
        return 0.0


def _fmt_date(v) -> str:
    """Normalize any date value to MM/DD/YYYY string."""
    if isinstance(v, datetime):
        return v.strftime('%m/%d/%Y')
    s = str(v).strip()
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}"
    return s


def _expand_date_yy(d: str) -> str:
    """Convert MM/DD/YY → MM/DD/YYYY (assumes 2000s)."""
    parts = d.split('/')
    if len(parts) == 3 and len(parts[2]) == 2:
        parts[2] = '20' + parts[2]
    return '/'.join(parts)


# ── Name utilities ───────────────────────────────────────────────────────────

_COMPANY_TOKENS = {'inc', 'llc', 'corp', 'ltd', 'agency', 'talent', 'agents',
                   'artists', 'models', 'group', 'management', 'entertainment'}


def _is_company_name(name: str) -> bool:
    """Heuristic: name is a company, not a person."""
    if not name:
        return False
    tokens = set(name.lower().replace(',', '').replace('.', '').split())
    return bool(tokens & _COMPANY_TOKENS) or '&' in name or name.endswith(('Inc', 'LLC', 'Corp', 'Ltd'))


def _ptip_name_to_last_first(name: str, is_agent: bool = False) -> str:
    """Convert 'First [Middle] Last' PTIP name → 'Last, First' workbook format.

    Company/agency names are returned unchanged.
    Single-letter middle initials are dropped; full middle names are kept.
    """
    if not name:
        return ''
    name = name.strip()
    if is_agent or _is_company_name(name) or ',' in name:
        return name
    parts = name.split()
    if len(parts) <= 1:
        return name
    last = parts[-1]
    first = parts[0]
    middle = [p for p in parts[1:-1] if len(p.rstrip('.')) > 1]
    if middle:
        return f"{last}, {first} {' '.join(middle)}"
    return f"{last}, {first}"


def _normalize_for_match(name: str) -> tuple[str, str]:
    """Return (full_norm, first_last_norm) for PDF↔PTIP name matching."""
    clean = re.sub(r"[.,']", '', name.lower()).strip()
    parts = clean.split()
    full = ' '.join(parts)
    first_last = f"{parts[0]} {parts[-1]}" if len(parts) >= 2 else full
    return full, first_last


# ── Address parsing ───────────────────────────────────────────────────────────

def _parse_ptip_address(addr_str) -> tuple[str, str, str, str]:
    """Parse multi-line PTIP address → (street, city, state, zip)."""
    if not addr_str:
        return '', '', '', ''
    lines = [ln.strip() for ln in str(addr_str).split('\n') if ln.strip()]
    if not lines:
        return '', '', '', ''

    city = state = zip_code = ''
    street_lines = lines

    m = re.match(r'^(.+?),\s+([A-Z]{2})\s+(\d{5})', lines[-1])
    if m:
        city, state, zip_code = m.group(1).strip(), m.group(2), m.group(3)
        street_lines = lines[:-1]

    street = street_lines[0] if street_lines else ''
    return street, city, state, zip_code


# ── PTIP Excel parsing ────────────────────────────────────────────────────────

def _find_ptip_sheet(wb):
    """Find the worksheet that contains the PTIP talent data.

    Supports both the ER native format (Customer Name as first column, no '#')
    and the manually-enhanced format ('#' prepended as first column).
    """
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row_vals in ws.iter_rows(max_row=10, values_only=True):
            row_strs = [str(v).strip() if v is not None else '' for v in row_vals]
            if 'Invoice Number' in row_strs and 'Wages' in row_strs and 'Talent Name' in row_strs:
                return ws, row_strs
    return wb.active, []


def parse_ptip_xlsx(xlsx_bytes: bytes) -> tuple[list[dict], list[str], list[str]]:
    """
    Parse a PTIP Excel file.

    Returns (rows, issues, header_cols).
    Each row dict has keys matching the header column names.
    header_cols is the original ordered list of column names (for re-export).
    """
    rows: list[dict] = []
    issues: list[str] = []
    header_cols: list[str] = []

    try:
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
        ws, header_cols = _find_ptip_sheet(wb)

        if not header_cols:
            issues.append("PTIP: could not locate header row with Invoice Number, Wages, Talent Name")
            return rows, issues, header_cols

        # Find the row number of the header so we can start data below it
        header_found_at = None
        for i, row_vals in enumerate(ws.iter_rows(values_only=True), start=1):
            row_strs = [str(v).strip() if v is not None else '' for v in row_vals]
            if 'Invoice Number' in row_strs and 'Wages' in row_strs and 'Talent Name' in row_strs:
                header_found_at = i
                break

        if header_found_at is None:
            issues.append("PTIP: header row not found")
            return rows, issues, header_cols

        inv_col_idx = next(
            (i for i, c in enumerate(header_cols) if 'Invoice' in c and 'Number' in c),
            -1,
        )

        for row_vals in ws.iter_rows(min_row=header_found_at + 1, values_only=True):
            if all(v is None for v in row_vals[:5]):
                break

            row = dict(zip(header_cols, row_vals))
            # Skip rows with no invoice number (totals / footer lines)
            if inv_col_idx >= 0:
                inv_val = str(row_vals[inv_col_idx] or '').strip()
                if not inv_val:
                    continue
            rows.append(row)

    except Exception as e:
        issues.append(f"PTIP parse error: {e}")

    return rows, issues, header_cols


# ── ER Invoice PDF parsing ───────────────────────────────────────────────────

_AMOUNT_RE = re.compile(r'^\d[\d,]*\.\d{2}$')
_DATE_YY_RE = re.compile(r'^\d{2}/\d{2}/\d{2}$')


def _parse_talent_row(line: str) -> dict | None:
    """
    Parse one row from an Extreme Reach talent invoice table.

    Handles both union (5 amounts, 2 dates, cam code) and non-union (3 amounts,
    1 date, no cam code) formats.  Agent rows have 'Off' before their role token.

    Returns dict with keys: row_no, name, is_agent, work_date, gross_wages,
    misc_pmt, pah.  Returns None if the line is not a talent data row.
    """
    parts = line.split()
    if len(parts) < 7:
        return None
    if not (parts[0].isdigit() and parts[1].isdigit()):
        return None
    if parts[-1] != 'USD':
        return None

    row_no = int(parts[0])
    tail = parts[2:-1]  # strip row#, TID, USD

    # Pop trailing amount tokens
    amounts: list[float] = []
    while tail and _AMOUNT_RE.match(tail[-1]):
        amounts.insert(0, float(tail.pop().replace(',', '')))

    if len(amounts) not in (3, 5):
        return None

    # Pop trailing date tokens (MM/DD/YY)
    dates: list[str] = []
    while tail and _DATE_YY_RE.match(tail[-1]):
        dates.insert(0, tail.pop())

    if not dates:
        return None

    work_date = _expand_date_yy(dates[0])

    # Union: [Yr, GrossWages, MiscPmt, ApplyAmt, P&HSub]
    # Non-union: [Yr, GrossWages, MiscPmt]
    if len(amounts) == 5:
        gross_wages, misc_pmt, pah = amounts[1], amounts[2], amounts[4]
    else:
        gross_wages, misc_pmt, pah = amounts[1], amounts[2], 0.0

    # Remaining tail: [name...] [cam_code?] [On|Off] [role_or_agent_code] [OS1] [tags...]
    on_off_idx = next((i for i, t in enumerate(tail) if t.lower() in ('on', 'off')), None)
    if on_off_idx is None:
        return None

    on_off = tail[on_off_idx].lower()
    role_code = tail[on_off_idx + 1] if on_off_idx + 1 < len(tail) else ''
    is_agent = (on_off == 'off') or (role_code.lower() == 'agent')

    # Name is everything before On/Off, minus a trailing cam-code number
    before = tail[:on_off_idx]
    if before and before[-1].isdigit():
        before = before[:-1]
    name = ' '.join(before)

    return {
        'row_no':     row_no,
        'name':       name,
        'is_agent':   is_agent,
        'work_date':  work_date,
        'gross_wages': round(gross_wages, 2),
        'misc_pmt':   round(misc_pmt, 2),
        'pah':        round(pah, 2),
    }


def parse_er_invoice_pdf(pdf_bytes: bytes) -> dict | None:
    """
    Parse an Extreme Reach talent invoice PDF.

    Returns dict with: invoice_no, invoice_date, cycle_dates, union_type,
    total_wages, total_misc, total_pah, total_er_tax, total_handling,
    talent_rows (list of _parse_talent_row dicts).
    Returns None if this does not look like an ER talent invoice.
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = '\n'.join(pg.extract_text() or '' for pg in pdf.pages)
    except Exception:
        return None

    if 'Extreme Reach Talent, Inc' not in text:
        return None

    inv_no = ''
    m = re.search(r'Invoice\s*#\s+(\d+)', text)
    if m:
        inv_no = m.group(1).strip()

    invoice_date = ''
    m = re.search(r'Invoice Date\s+(\d{1,2}/\d{1,2}/\d{4})', text)
    if m:
        invoice_date = _fmt_date(m.group(1))

    cycle_dates = ''
    m = re.search(r'Cycle Dates\s+([\d/]+ - [\d/]+)', text)
    if m:
        cycle_dates = m.group(1).strip()
    else:
        m = re.search(r'Cycle Dates\s+([\d/]+)\s*-\s*$', text, re.MULTILINE)
        if m:
            cycle_dates = m.group(1).strip()

    union_type = ''
    m = re.search(r'Union\s+(SAG|Non-Union|AFTRA)', text, re.IGNORECASE)
    if m:
        raw = m.group(1)
        union_type = 'Non-union' if 'non' in raw.lower() else raw.upper()

    total_wages    = _to_float(re.search(r'Wages\s+\$([\d,.]+)', text) and
                                re.search(r'Wages\s+\$([\d,.]+)', text).group(1))
    total_misc     = _to_float(re.search(r'Misc Payments\s+\$([\d,.]+)', text) and
                                re.search(r'Misc Payments\s+\$([\d,.]+)', text).group(1))
    total_pah      = _to_float(re.search(r'SAG Pension.*?\$([\d,.]+)', text) and
                                re.search(r'SAG Pension.*?\$([\d,.]+)', text).group(1))
    total_er_tax   = _to_float(re.search(r'Payroll Taxes\s+\$([\d,.]+)', text) and
                                re.search(r'Payroll Taxes\s+\$([\d,.]+)', text).group(1))
    total_handling = _to_float(re.search(r'Handling\s+\$([\d,.]+)', text) and
                                re.search(r'Handling\s+\$([\d,.]+)', text).group(1))

    talent_rows: list[dict] = []
    pdf_lines = text.split('\n')
    i = 0
    while i < len(pdf_lines):
        row = _parse_talent_row(pdf_lines[i].strip())
        if row:
            if not row['name'] and i + 1 < len(pdf_lines):
                next_line = pdf_lines[i + 1].strip()
                if next_line and not _parse_talent_row(next_line):
                    row['name'] = next_line
                    if _is_company_name(next_line):
                        row['is_agent'] = True
                    i += 1
            talent_rows.append(row)
        i += 1

    return {
        'invoice_no':      inv_no,
        'invoice_date':    invoice_date,
        'cycle_dates':     cycle_dates,
        'union_type':      union_type,
        'total_wages':     total_wages,
        'total_misc':      total_misc,
        'total_pah':       total_pah,
        'total_er_tax':    total_er_tax,
        'total_handling':  total_handling,
        'talent_rows':     talent_rows,
    }


# ── Duplicate detection ───────────────────────────────────────────────────────

def _detect_duplicates(ptip_rows: list[dict], pdf_name_counts: dict) -> set[int]:
    """
    Return the set of ptip_rows indices that are duplicates.

    Duplicate rule: same (invoice_no, talent_name, wages, er_tax) in the PTIP,
    where the PDF shows fewer occurrences of that name on that invoice than the
    PTIP does.  Agent rows (cast_category == 'Agent') follow the same rule.

    pdf_name_counts: {(invoice_no, full_norm_name): count_in_pdf}
    """
    duplicates: set[int] = set()
    # Group PTIP rows by (invoice_no, norm_name, wages, er_tax)
    ptip_groups: dict[tuple, list[int]] = defaultdict(list)

    for idx, row in enumerate(ptip_rows):
        inv_no = str(row.get('Invoice Number', '')).strip()
        name_raw = str(row.get('Talent Name', '')).strip()
        wages = _to_float(row.get('Wages'))
        er_tax = _to_float(row.get('Employer Taxes'))
        full_norm, _ = _normalize_for_match(name_raw)
        key = (inv_no, full_norm, wages, er_tax)
        ptip_groups[key].append(idx)

    for (inv_no, full_norm, wages, er_tax), indices in ptip_groups.items():
        if len(indices) <= 1:
            continue
        pdf_count = pdf_name_counts.get((inv_no, full_norm), len(indices))
        for extra_idx in indices[pdf_count:]:
            duplicates.add(extra_idx)

    return duplicates


# ── Workbook row builder ──────────────────────────────────────────────────────

def _build_row(
    *,
    item_no: int,
    ptip: dict | None,
    pdf_talent: dict | None,
    pdf_invoice: dict | None,
    received_invoice: bool,
    is_duplicate: bool,
    first_tax_row: bool,
    scenario: str,  # 'A', 'B', or 'C'
    ptip_row_no: int | None,
) -> dict:
    """Assemble one workbook row from PTIP and/or PDF data."""

    is_agent = False
    cast_cat = ''
    if ptip:
        cast_cat = str(ptip.get('Cast Category', '') or '').strip()
        # Strip multiline cell artifacts (e.g. "Dream Team Talent\n(Dream Team Talent Agency LLC)")
        name_for_agent_check = str(ptip.get('Talent Name', '') or '').split('\n')[0].strip()
        # Some PTIP files label agent rows as "Model" instead of "Agent"
        # so use company-name heuristic as a fallback
        is_agent = cast_cat.lower() == 'agent' or _is_company_name(name_for_agent_check)
    elif pdf_talent:
        is_agent = pdf_talent.get('is_agent', False)

    # ── Name ────────────────────────────────────────────────────────────────
    if ptip:
        name_raw = str(ptip.get('Talent Name', '') or '').split('\n')[0].strip()
        talent_name = _ptip_name_to_last_first(name_raw, is_agent)
    elif pdf_talent:
        name_raw = pdf_talent.get('name', '')
        talent_name = _ptip_name_to_last_first(name_raw, is_agent)
    else:
        talent_name = ''

    # ── Title ────────────────────────────────────────────────────────────────
    if is_agent:
        title = 'Agency fee'
    elif cast_cat:
        title = cast_cat
    elif pdf_talent and pdf_invoice:
        # For PDF-only: try to extract role from invoice (non-union uses role codes)
        title = ''
    else:
        title = ''

    # ── Address ──────────────────────────────────────────────────────────────
    street = city = state = zip_code = ''
    if ptip:
        street, city, state, zip_code = _parse_ptip_address(ptip.get('Talent Address'))

    # ── Work state & qualify ─────────────────────────────────────────────────
    if ptip:
        work_state = str(ptip.get('Work State', '') or '').strip().upper()
    elif pdf_invoice:
        work_state = 'IL'  # default; frontend overrides from project
    else:
        work_state = ''

    if state and work_state and state.upper() != work_state:
        qualify = 'NO-OOS'
    else:
        qualify = ''

    # ── PTIP financial columns ───────────────────────────────────────────────
    if ptip:
        # Try explicit Total column (manually-enhanced PTIP); otherwise compute it.
        # ER native PTIP excludes State Tax Withheld from the total (it is employee-side).
        _RAW_TOTAL = _to_float(ptip.get('Total') or ptip.get('TOTAL'))
        if _RAW_TOTAL:
            ptip_amount = _RAW_TOTAL
        else:
            _TOTAL_KEYS = [
                'Wages', 'P&H', 'Non Resident Corp Tax Amount',
                'Local Tax Withheld', 'State Disability Withheld',
                'Employer Taxes', 'Workers Compensation', 'Handling Fee', 'Signatory Fee',
            ]
            ptip_amount = round(sum(_to_float(ptip.get(k)) for k in _TOTAL_KEYS), 2)
        er_tax_ptip    = _to_float(ptip.get('Employer Taxes'))
        wc_ptip        = _to_float(ptip.get('Workers Compensation'))
        handling_ptip  = _to_float(ptip.get('Handling Fee'))
        sag_ptip       = _to_float(ptip.get('P&H'))
        signatory_ptip = _to_float(ptip.get('Signatory Fee'))
        check_number   = str(ptip.get('Check Number', '') or '').strip()
        commercial_id  = str(ptip.get('Commercial Id', '') or '').strip()
        ssn_fein       = (str(ptip.get('SSN', '') or '').strip() or
                          str(ptip.get('FEIN', '') or '').strip())
        on_ptip        = True
    else:
        ptip_amount = er_tax_ptip = wc_ptip = handling_ptip = sag_ptip = signatory_ptip = None
        check_number = commercial_id = ssn_fein = ''
        on_ptip = False

    # ── Wages and Misc Pmt ───────────────────────────────────────────────────
    if scenario in ('A',) and pdf_talent:
        # PDF-only: wages and misc from PDF talent row
        wages    = pdf_talent.get('gross_wages', 0.0)
        misc_pmt = pdf_talent.get('misc_pmt', 0.0)
        # Taxes: first talent row only (agent rows always get 0 taxes in PDF-only)
        if first_tax_row and not is_agent and pdf_invoice:
            er_tax   = pdf_invoice.get('total_er_tax', 0.0)
            sag      = pdf_invoice.get('total_pah', 0.0)
            handling = pdf_invoice.get('total_handling', 0.0)
        else:
            er_tax = sag = handling = 0.0
        wc = 0.0
        signatory_fee = 0.0
    elif scenario == 'B' and ptip:
        # PTIP-only: talent rows go to wages; agent rows go to misc (we know the cast)
        ptip_wages = _to_float(ptip.get('Wages'))
        if is_agent:
            wages    = 0.0
            misc_pmt = ptip_wages
        else:
            wages    = ptip_wages
            misc_pmt = 0.0
        er_tax   = er_tax_ptip or 0.0
        wc       = wc_ptip or 0.0
        handling = handling_ptip or 0.0
        sag      = sag_ptip or 0.0
        signatory_fee = signatory_ptip or 0.0
    elif scenario == 'C':
        # Both: PDF wins on wages/misc placement; PTIP for taxes
        if pdf_talent and received_invoice:
            wages    = pdf_talent.get('gross_wages', 0.0)
            misc_pmt = pdf_talent.get('misc_pmt', 0.0)
        else:
            # No PDF for this row (PTIP-only invoice in a mixed batch)
            ptip_wages = _to_float(ptip.get('Wages')) if ptip else 0.0
            if is_agent:
                wages    = 0.0
                misc_pmt = ptip_wages
            else:
                wages    = ptip_wages
                misc_pmt = 0.0
        er_tax   = er_tax_ptip or 0.0
        wc       = wc_ptip or 0.0
        handling = handling_ptip or 0.0
        sag      = sag_ptip or 0.0
        signatory_fee = signatory_ptip or 0.0
    else:
        wages = misc_pmt = er_tax = wc = handling = sag = signatory_fee = 0.0

    # ── Dates ────────────────────────────────────────────────────────────────
    if pdf_talent:
        work_dates   = pdf_talent.get('work_date', '')
        work_days    = 1
    else:
        work_dates = ''
        work_days  = None

    if pdf_invoice:
        invoice_date = pdf_invoice.get('invoice_date', '')
    elif ptip:
        invoice_date = _fmt_date(ptip.get('Check Date') or '')
    else:
        invoice_date = ''

    # ── Invoice info ─────────────────────────────────────────────────────────
    if ptip:
        invoice_no = str(ptip.get('Invoice Number', '') or '').strip()
    elif pdf_invoice:
        invoice_no = pdf_invoice.get('invoice_no', '')
    else:
        invoice_no = ''

    # ── Union type ───────────────────────────────────────────────────────────
    if pdf_invoice:
        union_type = pdf_invoice.get('union_type', '')
    else:
        union_type = ''

    # ── Total (static) ───────────────────────────────────────────────────────
    if ptip_amount is not None:
        total = ptip_amount
    else:
        total = round(wages + misc_pmt + er_tax + wc + handling + sag + signatory_fee, 2)

    return {
        'item_no':        item_no,
        'qualify':        qualify,
        'on_ptip':        on_ptip,
        'ptip_amount':    ptip_amount,
        'work_state':     work_state,
        'talent_name':    talent_name,
        'title':          title,
        'work_days':      work_days,
        'work_dates':     work_dates,
        'invoice_no':     invoice_no,
        'invoice_date':   invoice_date,
        'wages':          round(wages, 2),
        'misc_pymt':      round(misc_pmt, 2),
        'er_tax':         round(er_tax, 2),
        'wc':             round(wc, 2),
        'handling':       round(handling, 2),
        'sag':            round(sag, 2),
        'signatory_fee':  round(signatory_fee, 2),
        'total':          total,
        'check_number':   check_number,
        'received_invoice': received_invoice,
        'payment_entity': 'Extreme Reach Talent, Inc',
        'type':           union_type,
        'home_address':   street,
        'city':           city,
        'state':          state,
        'zip':            zip_code,
        'ssn_fein':       ssn_fein,
        'commercial_id':  commercial_id,
        'notes':          '',
        'ptip_row_no':    ptip_row_no,
        'is_duplicate':   is_duplicate,
    }


# ── Organized PTIP builder ───────────────────────────────────────────────────

def build_organized_ptip_xlsx(
    ptip_rows: list[dict],
    header_cols: list[str],
    received_invoice_nos: set[str],
    duplicate_indices: set[int],
    xlsx_bytes: bytes,
) -> bytes:
    """
    Return a new .xlsx that is the PTIP sorted by invoice number with an
    updated '#' column:
      - sequential int (resets per invoice group)
      - blank   : no matching PDF received for this invoice
      - 'duplicate' : row is a duplicate, excluded from workbook
    """
    # Sort rows by invoice number (stable), keeping duplicates in place
    invoice_col = 'Invoice Number'
    if invoice_col not in header_cols:
        # Find close match
        invoice_col = next((c for c in header_cols if 'Invoice' in c), header_cols[0])

    # Sort while preserving original indices for duplicate lookup
    sorted_with_orig = sorted(enumerate(ptip_rows),
                              key=lambda t: str(t[1].get(invoice_col, '') or ''))

    has_hash_col = bool(header_cols) and header_cols[0] == '#'
    out_header = list(header_cols) if has_hash_col else ['#'] + list(header_cols)

    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = 'Organized PTIP'
    ws.append(out_header)

    last_invoice = None
    seq = 0
    for orig_idx, row in sorted_with_orig:
        if has_hash_col:
            row_vals = [row.get(col) for col in header_cols]
        else:
            row_vals = [None] + [row.get(col) for col in header_cols]

        inv_no = str(row.get(invoice_col, '') or '').strip()
        if inv_no != last_invoice:
            last_invoice = inv_no
            seq = 0

        if orig_idx in duplicate_indices:
            hash_val = 'duplicate'
        elif inv_no not in received_invoice_nos:
            hash_val = ''
        else:
            seq += 1
            hash_val = seq

        row_vals[0] = hash_val
        ws.append(row_vals)

    buf = io.BytesIO()
    wb_out.save(buf)
    return buf.getvalue()


# ── Main extraction entry point ───────────────────────────────────────────────

def extract_talent(
    pdf_files: list[tuple[str, bytes]],  # [(filename, bytes), ...]
    ptip_bytes: bytes | None,
    project_title: str,
    workbook_type: str,
) -> dict:
    """
    Run the full talent extraction.

    Returns the /extract-talent response dict:
    {rows, ptip_excel_b64, summary: {total_rows, invoices_with_pdf,
     invoices_ptip_only, duplicates_found, issues}}
    """
    issues: list[str] = []

    # ── Step 1: Parse PDFs ────────────────────────────────────────────────────
    pdf_invoices: dict[str, dict] = {}  # invoice_no → parse result

    for filename, data in pdf_files:
        result = parse_er_invoice_pdf(data)
        if result is None:
            issues.append(
                f"{filename}: not recognized as an Extreme Reach talent invoice "
                "(AI fallback not yet implemented for talent — review manually)"
            )
            continue
        inv_no = result['invoice_no']
        if not inv_no:
            issues.append(f"{filename}: could not extract invoice number")
            continue
        if inv_no in pdf_invoices:
            issues.append(f"{filename}: duplicate invoice number {inv_no} — skipped")
            continue
        pdf_invoices[inv_no] = result

    received_invoice_nos: set[str] = set(pdf_invoices.keys())

    # ── Step 2: Parse PTIP ────────────────────────────────────────────────────
    ptip_rows: list[dict] = []
    header_cols: list[str] = []
    original_ptip_bytes = ptip_bytes  # keep for organized PTIP

    if ptip_bytes:
        ptip_rows, ptip_issues, header_cols = parse_ptip_xlsx(ptip_bytes)
        issues.extend(ptip_issues)

    # ── Step 3: Determine scenario ────────────────────────────────────────────
    has_pdf  = bool(pdf_invoices)
    has_ptip = bool(ptip_rows)

    if has_pdf and has_ptip:
        scenario = 'C'
    elif has_pdf:
        scenario = 'A'
    elif has_ptip:
        scenario = 'B'
    else:
        return {
            'rows': [],
            'ptip_excel_b64': None,
            'summary': {
                'total_rows': 0,
                'invoices_with_pdf': [],
                'invoices_ptip_only': [],
                'duplicates_found': 0,
                'issues': issues + ['No PDFs or PTIP file provided.'],
            },
        }

    # ── Step 4: Build PDF name count map (for duplicate detection) ────────────
    pdf_name_counts: dict[tuple, int] = defaultdict(int)
    if scenario == 'C':
        for inv_no, inv_data in pdf_invoices.items():
            for tr in inv_data['talent_rows']:
                full_norm, _ = _normalize_for_match(tr['name'])
                pdf_name_counts[(inv_no, full_norm)] += 1

    # ── Step 5: Detect duplicates in PTIP ────────────────────────────────────
    duplicate_indices: set[int] = set()
    if has_ptip:
        duplicate_indices = _detect_duplicates(ptip_rows, pdf_name_counts)

    # ── Step 6: Group PTIP rows by invoice for ordered assembly ─────────────
    ptip_by_invoice: dict[str, list[dict]] = defaultdict(list)
    ptip_orig_by_invoice: dict[str, list[int]] = defaultdict(list)
    for idx, pr in enumerate(ptip_rows):
        inv_no = str(pr.get('Invoice Number', '') or '').strip()
        ptip_by_invoice[inv_no].append(pr)
        ptip_orig_by_invoice[inv_no].append(idx)

    # ── Step 7: Assemble workbook rows (invoice-sorted, PDF order within each) ─
    workbook_rows: list[dict] = []
    item_no = 1

    all_invoice_nos = sorted(
        set(received_invoice_nos) | set(ptip_by_invoice.keys()),
        key=lambda x: x.zfill(20),
    )

    for inv_no in all_invoice_nos:
        inv_data = pdf_invoices.get(inv_no)
        inv_ptip_rows = ptip_by_invoice.get(inv_no, [])
        inv_ptip_orig = ptip_orig_by_invoice.get(inv_no, [])
        received = (inv_no in received_invoice_nos)
        inv_has_ptip = bool(inv_ptip_rows)

        if inv_data:
            # PDF exists — iterate talent rows in PDF order; match each to a PTIP row
            ptip_consumed: set[int] = set()
            first_talent_seen = False

            for tr in inv_data['talent_rows']:
                ptip_match: dict | None = None
                ptip_match_local_i: int | None = None
                tr_full_norm, _ = _normalize_for_match(tr['name'])

                for local_i, pr in enumerate(inv_ptip_rows):
                    if local_i in ptip_consumed:
                        continue
                    if inv_ptip_orig[local_i] in duplicate_indices:
                        continue
                    pr_name = str(pr.get('Talent Name', '') or '').split('\n')[0].strip()
                    pr_full_norm, _ = _normalize_for_match(pr_name)
                    if pr_full_norm == tr_full_norm:
                        ptip_match = pr
                        ptip_match_local_i = local_i
                        ptip_consumed.add(local_i)
                        break

                # Invoice-level tax stacking only when no PTIP exists for this invoice
                is_first_tax = (not inv_has_ptip) and (not tr['is_agent']) and (not first_talent_seen)
                if not tr['is_agent']:
                    first_talent_seen = True
                eff_scenario = 'A' if not inv_has_ptip else scenario

                row = _build_row(
                    item_no=item_no,
                    ptip=ptip_match,
                    pdf_talent=tr,
                    pdf_invoice=inv_data,
                    received_invoice=True,
                    is_duplicate=False,
                    first_tax_row=is_first_tax,
                    scenario=eff_scenario,
                    ptip_row_no=(inv_ptip_orig[ptip_match_local_i] + 1) if ptip_match_local_i is not None else None,
                )
                workbook_rows.append(row)
                item_no += 1

            # Append any PTIP rows for this invoice not matched to a PDF row
            for local_i, (pr, orig_idx) in enumerate(zip(inv_ptip_rows, inv_ptip_orig)):
                if local_i in ptip_consumed:
                    continue
                is_dup = orig_idx in duplicate_indices
                row = _build_row(
                    item_no=item_no if not is_dup else 0,
                    ptip=pr,
                    pdf_talent=None,
                    pdf_invoice=inv_data,
                    received_invoice=True,
                    is_duplicate=is_dup,
                    first_tax_row=False,
                    scenario=scenario,
                    ptip_row_no=orig_idx + 1,
                )
                workbook_rows.append(row)
                if not is_dup:
                    item_no += 1

        else:
            # PTIP-only invoice (no PDF received)
            for pr, orig_idx in zip(inv_ptip_rows, inv_ptip_orig):
                is_dup = orig_idx in duplicate_indices
                row = _build_row(
                    item_no=item_no if not is_dup else 0,
                    ptip=pr,
                    pdf_talent=None,
                    pdf_invoice=None,
                    received_invoice=False,
                    is_duplicate=is_dup,
                    first_tax_row=False,
                    scenario=scenario,
                    ptip_row_no=orig_idx + 1,
                )
                workbook_rows.append(row)
                if not is_dup:
                    item_no += 1

    # ── Step 8: Note invoices with PDF but no PTIP match ─────────────────────
    if has_ptip and has_pdf:
        ptip_inv_nos = {str(r.get('Invoice Number', '') or '').strip() for r in ptip_rows}
        for inv_no in sorted(received_invoice_nos):
            if inv_no not in ptip_inv_nos:
                issues.append(
                    f"Invoice {inv_no}: PDF received, not in PTIP "
                    "— rows included with 'Included in PTIP report' = NO"
                )

    # ── Step 9: Build organized PTIP ─────────────────────────────────────────
    ptip_excel_b64 = None
    if has_ptip and original_ptip_bytes and header_cols:
        # Re-map duplicate indices: original PTIP row order vs sorted order
        # build_organized_ptip_xlsx expects original indices
        try:
            ptip_excel_bytes = build_organized_ptip_xlsx(
                ptip_rows,
                header_cols,
                received_invoice_nos,
                duplicate_indices,
                original_ptip_bytes,
            )
            ptip_excel_b64 = base64.b64encode(ptip_excel_bytes).decode()
        except Exception as e:
            issues.append(f"Organized PTIP build error: {e}")

    # ── Step 10: Build summary ────────────────────────────────────────────────
    ptip_inv_nos = {str(r.get('Invoice Number', '') or '').strip() for r in ptip_rows}
    invoices_with_pdf  = sorted(received_invoice_nos & ptip_inv_nos) if has_ptip else sorted(received_invoice_nos)
    invoices_ptip_only = sorted(ptip_inv_nos - received_invoice_nos) if has_ptip else []
    duplicates_found   = len(duplicate_indices)
    non_dup_rows       = [r for r in workbook_rows if not r['is_duplicate']]

    return {
        'rows': workbook_rows,
        'ptip_excel_b64': ptip_excel_b64,
        'summary': {
            'total_rows':         len(non_dup_rows),
            'invoices_with_pdf':  invoices_with_pdf,
            'invoices_ptip_only': invoices_ptip_only,
            'duplicates_found':   duplicates_found,
            'issues':             issues,
        },
    }
