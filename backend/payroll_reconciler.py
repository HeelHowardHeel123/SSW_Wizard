"""
GA Crew Payroll reconciliation for TPC Production Binder Wizard.

Takes already-extracted rows from the two independent Crew Payroll sources
(PDF payroll invoices via /extract-fringe, and the payroll company/accounting
Production Report via /extract-ga-production-report) and matches them against
each other, invoice-by-invoice, to populate the Payroll Report tab's
"On Production Report" / "On Invoice PDF" reconciliation columns.

Mirrors talent_extractor.py's approach (same three-scenario shape: PDFs only,
Production Report only, both), with two differences suited to crew data:
  - SSN last-4 is tried as the primary match key (crew data carries it
    reliably; Talent invoices/PTIP mostly don't), before falling back to
    normalized-name matching.
  - Financial precedence is simpler than Talent's split rule: the Production
    Report always wins on dollar values when a person matches both sides.
    The PDF's own total is only ever used for the Automation Total
    cross-check column, never written into the row's real financial columns.

This module never touches the .xlsx -- it returns fully-assembled row dicts
in the same FRINGE_FIELDS shape the two source endpoints already use, plus
three extra keys (onProductionReport, onInvoicePdf, automationTotal), in
final sorted order ready to write starting at row 4.
"""

import re
import json
from collections import defaultdict


# ── Matching keys ────────────────────────────────────────────────────────────

def _norm_invoice(v) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    return s.lstrip("0") or "0"


def _ssn_last4(v) -> str:
    digits = re.sub(r"\D", "", str(v or ""))
    return digits[-4:] if len(digits) >= 4 else ""


def _normalize_name(name: str) -> str:
    clean = re.sub(r"[.,']", "", str(name or "").lower()).strip()
    return re.sub(r"\s+", " ", clean)


# ── LLM fuzzy-match fallback ──────────────────────────────────────────────────

def _llm_fuzzy_match_payroll(
    pdf_names: list[str],
    report_names: list[str],
    openai_key: str,
) -> tuple[dict[str, str], dict[str, str]]:
    """Last-resort match for names that SSN and exact normalization both missed.

    Only ever called on names already known to be in the same invoice AND
    where at least one other name in that invoice already matched via SSN or
    exact name -- see the "only fire once there's a real basis for pairing
    this invoice's two files at all" comment at the call site.

    Returns (matches, unmatched_reasons):
      matches:           {pdf_name: report_name} for confident matches
      unmatched_reasons: {pdf_name: reason} for names that couldn't be matched
    """
    if not pdf_names or not report_names or not openai_key:
        return {}, {}
    try:
        from openai import OpenAI
        client = OpenAI(api_key=openai_key)
        prompt = (
            "You are reconciling a crew payroll invoice against a payroll company's "
            "Production Report, matching people who appear on both by name despite "
            "formatting differences.\n\n"
            "Invoice names (not yet matched):\n"
            f"{json.dumps(pdf_names, indent=2)}\n\n"
            "Production Report names (not yet matched):\n"
            f"{json.dumps(report_names, indent=2)}\n\n"
            "Common differences:\n"
            "- Name order: 'John Smith' (invoice) = 'Smith, John' (report)\n"
            "- Suffixes: 'Carl Johnson Jr' = 'Johnson, Carl, Jr.'\n"
            "- A middle name/initial present on only one side\n"
            "- Minor typos\n\n"
            "Rules:\n"
            "- Only match if you are CONFIDENT they are the same person.\n"
            "- Do NOT guess -- if uncertain, put the name in 'unmatched' with a brief reason.\n"
            "- Each invoice name maps to at most one report name and vice versa.\n"
            "- Every invoice name must appear in either 'matches' or 'unmatched'.\n\n"
            "Return ONLY a JSON object with exactly two keys:\n"
            '{"matches": {"invoice_name": "report_name", ...}, '
            '"unmatched": {"invoice_name": "brief reason why no confident match was found", ...}}'
        )
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=800,
            response_format={"type": "json_object"},
        )
        raw = json.loads(resp.choices[0].message.content)
        if not isinstance(raw, dict):
            return {}, {}
        pdf_set    = set(pdf_names)
        report_set = set(report_names)
        matches = {
            str(k): str(v)
            for k, v in (raw.get("matches") or {}).items()
            if str(k) in pdf_set and str(v) in report_set
        }
        reasons = {
            str(k): str(v)
            for k, v in (raw.get("unmatched") or {}).items()
            if str(k) in pdf_set
        }
        return matches, reasons
    except Exception:
        return {}, {}


# ── Per-invoice matching ──────────────────────────────────────────────────────

def _match_invoice(
    pdf_list: list[tuple[int, dict]],
    report_list: list[tuple[int, dict]],
    openai_key: str,
) -> tuple[list[tuple[int, int]], set[int], set[int], dict[str, str]]:
    """Match one invoice's PDF rows against its Production Report rows.

    pdf_list / report_list: [(original_index, row_dict), ...] for this invoice only.

    Returns (pairs, pdf_consumed, report_consumed, llm_reasons) where pairs is
    a list of (local_pdf_index, local_report_index) into pdf_list/report_list.
    """
    pdf_consumed:    set[int] = set()
    report_consumed: set[int] = set()
    pairs: list[tuple[int, int]] = []

    # Pass 1: SSN last-4 exact match -- the most reliable key crew data has.
    for pi, (_, p_row) in enumerate(pdf_list):
        p_ssn = _ssn_last4(p_row.get("ssn"))
        if not p_ssn:
            continue
        for ri, (_, r_row) in enumerate(report_list):
            if ri in report_consumed:
                continue
            if _ssn_last4(r_row.get("ssn")) == p_ssn:
                pairs.append((pi, ri))
                pdf_consumed.add(pi)
                report_consumed.add(ri)
                break

    # Pass 2: normalized-name exact match, for anything SSN didn't resolve
    # (missing/unmasked SSN on one side, etc.)
    for pi, (_, p_row) in enumerate(pdf_list):
        if pi in pdf_consumed:
            continue
        p_name = _normalize_name(p_row.get("worker"))
        if not p_name:
            continue
        for ri, (_, r_row) in enumerate(report_list):
            if ri in report_consumed:
                continue
            if _normalize_name(r_row.get("worker")) == p_name:
                pairs.append((pi, ri))
                pdf_consumed.add(pi)
                report_consumed.add(ri)
                break

    # Pass 3: LLM fuzzy fallback -- only once this invoice already has at
    # least one confirmed match, so there's a real basis for treating these
    # two files as covering the same invoice at all.
    llm_reasons: dict[str, str] = {}
    if openai_key and pairs:
        remaining_pdf    = [pi for pi in range(len(pdf_list)) if pi not in pdf_consumed]
        remaining_report = [ri for ri in range(len(report_list)) if ri not in report_consumed]
        if remaining_pdf and remaining_report:
            pdf_names    = [pdf_list[pi][1].get("worker", "") for pi in remaining_pdf]
            report_names = [report_list[ri][1].get("worker", "") for ri in remaining_report]
            llm_matches, llm_reasons = _llm_fuzzy_match_payroll(pdf_names, report_names, openai_key)
            if llm_matches:
                name_to_report_local = {
                    report_list[ri][1].get("worker", ""): ri for ri in remaining_report
                }
                for pi in remaining_pdf:
                    pdf_name = pdf_list[pi][1].get("worker", "")
                    matched_report_name = llm_matches.get(pdf_name)
                    if not matched_report_name:
                        continue
                    ri = name_to_report_local.get(matched_report_name)
                    if ri is not None and ri not in report_consumed:
                        pairs.append((pi, ri))
                        pdf_consumed.add(pi)
                        report_consumed.add(ri)

    return pairs, pdf_consumed, report_consumed, llm_reasons


# ── Sorting ────────────────────────────────────────────────────────────────────

_UNPLACED = 10 ** 9  # sort-key filler for a row with no position on that axis


def _sort_key_name_invoice(entry: dict) -> tuple:
    row = entry["row"]
    return (_normalize_name(row.get("worker")), _norm_invoice(row.get("invoiceNo")).zfill(20))


def _sort_key_invoice_pdf_layout(entry: dict) -> tuple:
    row = entry["row"]
    pdf_idx = entry["pdf_idx"] if entry["pdf_idx"] is not None else _UNPLACED
    return (_norm_invoice(row.get("invoiceNo")).zfill(20), pdf_idx)


def _sort_key_report_layout(entry: dict):
    return entry["report_idx"] if entry["report_idx"] is not None else _UNPLACED


def _resolve_sort_option(sort_option: str, has_pdf: bool, has_report: bool) -> str:
    """Silently fall back when the chosen option's required source is missing,
    rather than erroring or producing a degenerate sort."""
    if sort_option == "production_report_layout" and not has_report:
        return "invoice_pdf_layout"
    if sort_option == "invoice_pdf_layout" and not has_pdf:
        return "production_report_layout"
    return sort_option


# ── Public API ─────────────────────────────────────────────────────────────────

def reconcile_payroll(
    pdf_rows: list[dict],
    production_report_rows: list[dict],
    sort_option: str = "invoice_pdf_layout",
    openai_key: str = "",
) -> dict:
    """Match PDF-extracted and Production-Report rows against each other by
    invoice number, flag which source(s) each person came from, and return
    the combined row list in the requested order.

    Returns {"rows": [...], "issues": [...]}.
    """
    issues: list[str] = []

    pdf_by_invoice:    dict[str, list[tuple[int, dict]]] = defaultdict(list)
    report_by_invoice: dict[str, list[tuple[int, dict]]] = defaultdict(list)

    for idx, row in enumerate(pdf_rows):
        pdf_by_invoice[_norm_invoice(row.get("invoiceNo"))].append((idx, row))
    for idx, row in enumerate(production_report_rows):
        report_by_invoice[_norm_invoice(row.get("invoiceNo"))].append((idx, row))

    all_invoices = sorted(
        set(pdf_by_invoice) | set(report_by_invoice),
        key=lambda x: x.zfill(20),
    )

    matched_rows:   list[dict] = []
    unmatched_rows: list[dict] = []

    for inv_no in all_invoices:
        pdf_list    = list(pdf_by_invoice.get(inv_no, []))
        report_list = list(report_by_invoice.get(inv_no, []))

        pairs, pdf_consumed, report_consumed, llm_reasons = _match_invoice(
            pdf_list, report_list, openai_key,
        )

        for pi, ri in pairs:
            p_idx, p_row = pdf_list[pi]
            r_idx, r_row = report_list[ri]
            row = dict(r_row)  # Production Report wins on every financial/identity field
            row["onProductionReport"] = True
            row["onInvoicePdf"]       = True
            row["automationTotal"]    = p_row.get("total")
            matched_rows.append({"row": row, "pdf_idx": p_idx, "report_idx": r_idx})

        for ri, (r_idx, r_row) in enumerate(report_list):
            if ri in report_consumed:
                continue
            row = dict(r_row)
            row["onProductionReport"] = True
            row["onInvoicePdf"]       = False
            row["automationTotal"]    = None
            unmatched_rows.append({"row": row, "pdf_idx": None, "report_idx": r_idx})
            if pdf_list:
                issues.append(
                    f"Invoice {inv_no}: {r_row.get('worker', '(unnamed)')} is on the "
                    "Production Report but no matching PDF invoice row was found."
                )

        for pi, (p_idx, p_row) in enumerate(pdf_list):
            if pi in pdf_consumed:
                continue
            row = dict(p_row)
            row["onProductionReport"] = False
            row["onInvoicePdf"]       = True
            row["automationTotal"]    = p_row.get("total")
            reason = llm_reasons.get(p_row.get("worker", ""), "")
            if reason:
                row["notes"] = (row.get("notes") or "") or f"[no Production Report match] {reason}"
            unmatched_rows.append({"row": row, "pdf_idx": p_idx, "report_idx": None})
            if report_list:
                issues.append(
                    f"Invoice {inv_no}: {p_row.get('worker', '(unnamed)')} is on the "
                    "PDF invoice but no matching Production Report row was found."
                )

    # ── Sort ─────────────────────────────────────────────────────────────────
    has_pdf    = bool(pdf_rows)
    has_report = bool(production_report_rows)
    effective_sort = _resolve_sort_option(sort_option, has_pdf=has_pdf, has_report=has_report)

    def _sort_by_effective(entries: list[dict]) -> None:
        if effective_sort == "name_invoice":
            entries.sort(key=_sort_key_name_invoice)
        elif effective_sort == "production_report_layout":
            entries.sort(key=_sort_key_report_layout)
        else:  # "invoice_pdf_layout" (also the default/fallback)
            entries.sort(key=_sort_key_invoice_pdf_layout)

    if has_pdf and has_report:
        # Real reconciliation: sort the matched rows per the chosen layout,
        # then append the (hopefully few) unmatched stragglers at the end,
        # grouped together and sorted by name so they're easy to spot for
        # review -- this is the "something doesn't line up" case.
        _sort_by_effective(matched_rows)
        unmatched_rows.sort(key=lambda e: _normalize_name(e["row"].get("worker")))
        final_entries = matched_rows + unmatched_rows
    else:
        # Only one source was ever provided -- there's nothing to reconcile,
        # so every row ends up in unmatched_rows trivially. Treating that as
        # a "needs review" bucket would be wrong here: sort the whole list
        # per the resolved layout instead of forcing alphabetical order.
        final_entries = matched_rows + unmatched_rows
        _sort_by_effective(final_entries)

    final_rows = [e["row"] for e in final_entries]

    return {"rows": final_rows, "issues": issues}
