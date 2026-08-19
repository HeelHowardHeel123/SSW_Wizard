"""
Crew Timecard reconciliation for TPC Production Binder Wizard.

Cross-checks already-reconciled Payroll Report rows (from
payroll_reconciler.reconcile_payroll) against already-extracted timecard
rows (from timecard_parser.extract_timecards), by SSN/name + date-range
overlap -- there's no shared ID between a Production Report row and a
timecard (a Production Report row's invoice number belongs to a different
system than a timecard's own Batch/TC number, confirmed against real PJ 004
data), so date-range overlap is the only reliable match key.

A single Production Report row's date range sometimes spans MORE than one
timecard week (e.g. a report row covering 05/20-05/30 when the actual
timecards are split into a 05/18-05/24 week and a 05/25-05/31 week) --
every timecard for that person whose week overlaps the row's range gets
summed together into one timecardTotal, mirroring the same
sum-the-matched-group pattern used elsewhere in this project (Consolidated
Production Reports' person-level totals, the original per-invoice
Automation Total).
"""

import re
from collections import defaultdict
from datetime import datetime as _dt

from payroll_reconciler import _ssn_last4, _normalize_name

_WORKDATES_RE = re.compile(r"(\d{2}/\d{2}/\d{4})\s*(?:-\s*(\d{2}/\d{2}/\d{4}))?")


def _parse_mmddyyyy(s):
    try:
        return _dt.strptime(str(s).strip(), "%m/%d/%Y").date()
    except Exception:
        return None


def _row_date_range(row: dict):
    """Prefers the row's own actual work-date fields (startDate/endDate) --
    workDates may show the pay-period week-ending label instead when one is
    available, which isn't necessarily the same range a timecard's week
    actually falls in. Only parses workDates as a fallback for a row that
    never had real start/end dates at all."""
    start = _parse_mmddyyyy(row.get("startDate"))
    end   = _parse_mmddyyyy(row.get("endDate")) or start
    if start and end:
        return start, end

    m = _WORKDATES_RE.search(row.get("workDates") or "")
    if not m:
        return None, None
    start = _parse_mmddyyyy(m.group(1))
    end = _parse_mmddyyyy(m.group(2)) if m.group(2) else start
    return start, end


def match_timecards(rows: list[dict], timecard_rows: list[dict]) -> list[dict]:
    """Mutates and returns rows, adding hasTimecard (bool), timecardTotal
    (float|None), and appending a "Timecards: ..." breakdown to notes when
    at least one timecard matched."""
    tc_groups: dict[str, list[dict]] = defaultdict(list)
    for tc in timecard_rows:
        key = _ssn_last4(tc.get("ssn")) or _normalize_name(tc.get("worker"))
        if key:
            tc_groups[key].append(tc)

    for row in rows:
        key = _ssn_last4(row.get("ssn")) or _normalize_name(row.get("worker"))
        candidates = tc_groups.get(key, [])
        r_start, r_end = _row_date_range(row)

        matched = []
        if candidates and r_start and r_end:
            for tc in candidates:
                tc_start = _parse_mmddyyyy(tc.get("weekStart"))
                tc_end   = _parse_mmddyyyy(tc.get("weekEnd"))
                if tc_start and tc_end and tc_start <= r_end and tc_end >= r_start:
                    matched.append(tc)

        row["hasTimecard"] = bool(matched)
        if matched:
            row["timecardTotal"] = round(sum((t.get("total") or 0) for t in matched), 2)
            parts = [
                f"{t.get('weekStart')}-{t.get('weekEnd')} (${(t.get('total') or 0):,.2f})"
                for t in matched
            ]
            breakdown = "Timecards: " + ", ".join(parts)
            existing = (row.get("notes") or "").strip()
            row["notes"] = f"{existing} | {breakdown}" if existing else breakdown
        else:
            row["timecardTotal"] = None

    return rows
