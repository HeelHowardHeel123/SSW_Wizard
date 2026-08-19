"""
Crew Timecard extraction for TPC Production Binder Wizard.

Timecards are a third Crew Payroll source alongside PDF invoices and the
Production Report -- some productions (e.g. PJ 004) have no invoice PDFs at
all, only a "Payroll Batches" folder of these. Each batch PDF is a stack of
individual CREW TIME CARD pages (Cast & Crew/CAPS's own rendering), one page
per person per work week (Sun-Sat), signed off before the Production Report
is even generated.

Deliberately thin, same philosophy as the Wrapbook fringe_002 parser: pulls
only what's needed to establish presence and a per-week dollar total for
matching against the Production Report (worker, ssn, weekStart, weekEnd,
total, batchNo, tcId) -- not the full day-by-day hours/fringe breakdown,
since the Production Report is the authoritative source for that.

Extraction is regex-first. Timecard layouts can vary by payroll company or
even by production, so any page the deterministic regex can't read falls
back to ONE Anthropic API call covering every unrecognized page in that
file at once (not a call per page), asking it to generate a single regex
that covers them. The generated pattern is applied locally -- no further AI
calls -- and is always surfaced in the returned issues list so it can be
reviewed and folded into _DETERMINISTIC_PATTERNS permanently if this layout
shows up again.
"""

import io
import re
import json
from datetime import datetime as _dt

import pdfplumber

from parsers.base import clean_fringe_name, parse_amount

_SSN_RE   = re.compile(r"\b(xxx-xx-\d{4})\b", re.IGNORECASE)
_NAME_RE  = re.compile(r"EMPLOYEE NAME.*?LOCATION\s*\n([A-Z][A-Za-z,.' -]+?)\n", re.DOTALL)
_SUN_RE   = re.compile(r"Sun\s*\n?\s*(\d{2}/\d{2})")
_SAT_RE   = re.compile(r"Sat\s*\n?\s*(\d{2}/\d{2})")
_TOTAL_RE = re.compile(r"Total:\s*\$?([\d,]+\.\d{2})")
_BATCH_RE = re.compile(r"Batch\((\d+)\),\s*TC\((\d+)\)")

_REQUIRED_GROUPS = {"worker", "ssn", "weekStart", "weekEnd", "total"}


def _infer_year(filename: str) -> str:
    m = re.search(r"(\d{4})-\d{2}-\d{2}", filename or "")
    return m.group(1) if m else str(_dt.now().year)


def _mmdd_to_mmddyyyy(mmdd: str, year: str) -> str:
    return f"{mmdd}/{year}" if mmdd and "/" in mmdd and len(mmdd.split("/")) == 2 else mmdd


def _extract_page_deterministic(text: str, year: str) -> dict | None:
    """The known "CREW TIME CARD" (Cast & Crew/CAPS) layout, confirmed
    against real PJ 004 batch pages."""
    name_m  = _NAME_RE.search(text)
    ssn_m   = _SSN_RE.search(text)
    sun_m   = _SUN_RE.search(text)
    sat_m   = _SAT_RE.search(text)
    total_m = _TOTAL_RE.search(text)
    if not (name_m and ssn_m and sun_m and sat_m and total_m):
        return None

    batch_m = _BATCH_RE.search(text)
    raw_name = name_m.group(1).strip()
    return {
        "worker":    clean_fringe_name(raw_name, from_caps=raw_name.isupper()),
        "ssn":       ssn_m.group(1),
        "weekStart": _mmdd_to_mmddyyyy(sun_m.group(1), year),
        "weekEnd":   _mmdd_to_mmddyyyy(sat_m.group(1), year),
        "total":     parse_amount(total_m.group(1)),
        "batchNo":   batch_m.group(1) if batch_m else "",
        "tcId":      batch_m.group(2) if batch_m else "",
    }


def _generate_ai_regex(sample_texts: list[str], anthropic_client) -> str | None:
    """One Claude call covering every distinct unrecognized page in this
    file, asking for a single regex that covers them all. Best-effort:
    returns None on any failure so the caller can fall through to leaving
    those pages for manual entry."""
    prompt = (
        "These are pages of raw text extracted from CREW TIME CARD PDFs that "
        "an existing regex-based parser could not read -- likely a new or "
        "different timecard layout. Here are up to 5 example page texts:\n\n"
        + "\n\n---PAGE BREAK---\n\n".join(sample_texts[:5])
        + "\n\nReturn a single Python regular expression (for use with "
        "re.search, not necessarily anchored) with named groups that "
        "extracts:\n"
        "  (?P<worker>...)    -- the employee's full name\n"
        "  (?P<ssn>...)       -- their SSN, masked is fine (e.g. xxx-xx-1234)\n"
        "  (?P<weekStart>...) -- the Sunday date of the work week, MM/DD\n"
        "  (?P<weekEnd>...)   -- the Saturday date of the work week, MM/DD\n"
        "  (?P<total>...)     -- the final dollar total paid for the week "
        "(digits, commas, decimal point only -- no $ sign in the captured group)\n\n"
        'Return ONLY a JSON object: {"regex": "<the pattern as a string>"}\n'
        "No explanation. No markdown. No code fences."
    )
    try:
        resp = anthropic_client.messages.create(
            model="claude-sonnet-5",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}],
        )
        text_block = next((b for b in resp.content if b.type == "text"), None)
        if not text_block:
            return None
        raw = text_block.text.strip()
        try:
            data = json.loads(raw)
        except Exception:
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            if not m:
                return None
            data = json.loads(m.group())
        pattern_str = data.get("regex")
        pattern = re.compile(pattern_str)
        if not _REQUIRED_GROUPS <= set(pattern.groupindex):
            return None
        return pattern_str
    except Exception:
        return None


def _apply_ai_pattern(pattern_str: str, text: str, year: str) -> dict | None:
    try:
        pattern = re.compile(pattern_str)
        m = pattern.search(text)
        if not m:
            return None
        gd = m.groupdict()
        if not all(gd.get(k) for k in _REQUIRED_GROUPS):
            return None
        raw_name = gd["worker"].strip()
        return {
            "worker":    clean_fringe_name(raw_name, from_caps=raw_name.isupper()),
            "ssn":       gd["ssn"].strip(),
            "weekStart": _mmdd_to_mmddyyyy(gd["weekStart"].strip(), year),
            "weekEnd":   _mmdd_to_mmddyyyy(gd["weekEnd"].strip(), year),
            "total":     parse_amount(gd["total"]),
            "batchNo":   "",
            "tcId":      "",
        }
    except Exception:
        return None


def extract_timecards(
    pdf_bytes: bytes,
    filename: str,
    anthropic_client=None,
) -> tuple[list[dict], list[str]]:
    """Returns (rows, issues). Each row: worker, ssn, weekStart, weekEnd
    (MM/DD/YYYY), total, batchNo, tcId."""
    issues: list[str] = []
    year = _infer_year(filename)

    rows: list[dict] = []
    failed_pages: list[tuple[int, str]] = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for i, pg in enumerate(pdf.pages):
                text = pg.extract_text() or ""
                if not text.strip():
                    continue
                parsed = _extract_page_deterministic(text, year)
                if parsed:
                    rows.append(parsed)
                else:
                    failed_pages.append((i, text))
    except Exception as e:
        return rows, [f"{filename}: failed to read PDF -- {e}"]

    if not failed_pages:
        return rows, issues

    if anthropic_client is None:
        issues.append(
            f"{filename}: {len(failed_pages)} page(s) have an unrecognized "
            "timecard layout and no Anthropic key is configured for the AI "
            "fallback -- enter these manually."
        )
        return rows, issues

    sample_texts = [t for _, t in failed_pages[:5]]
    pattern_str = _generate_ai_regex(sample_texts, anthropic_client)
    if not pattern_str:
        issues.append(
            f"{filename}: {len(failed_pages)} page(s) have an unrecognized "
            "timecard layout and the AI fallback couldn't generate a usable "
            "regex -- enter these manually."
        )
        return rows, issues

    recovered, still_failed = 0, []
    for i, text in failed_pages:
        parsed = _apply_ai_pattern(pattern_str, text, year)
        if parsed:
            rows.append(parsed)
            recovered += 1
        else:
            still_failed.append(i + 1)

    issues.append(
        f"{filename}: {recovered} page(s) used an AI-generated fallback regex "
        f"(new/unrecognized timecard layout). Add this pattern permanently if "
        f"this layout shows up again: {pattern_str}"
    )
    if still_failed:
        issues.append(
            f"{filename}: {len(still_failed)} page(s) could not be parsed even "
            f"with the AI fallback (page(s) {still_failed}) -- enter manually."
        )

    return rows, issues
