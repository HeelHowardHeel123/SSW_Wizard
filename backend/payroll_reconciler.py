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
    Report always wins on any field it actually has a value for, when a
    person matches both sides. The PDF only fills in fields the Production
    Report left blank -- it never overrides a value the Production Report
    did provide. The PDF's own total is separately used for the Automation
    Total cross-check column regardless.

This module never touches the .xlsx -- it returns fully-assembled row dicts
in the same FRINGE_FIELDS shape the two source endpoints already use, plus
four extra keys (onProductionReport, onInvoicePdf, automationTotal, aicpCode),
in final sorted order ready to write starting at row 4.

AICP classification: every row also gets an AICP category (1-25, same list
and descriptions as the GA AP prompt's aicp_code field) via one batched GPT
call over the whole reconciled row set, using each row's worker name and job
description as context. No category is off-limits on this tab -- 22/23/24
(Georgia Crew/Cast/Extras Hires) will be correct for the vast majority of
rows here, but that's steering, not a restriction, since an odd case (e.g. a
per diem paid through payroll instead of AP) is genuinely possible.

A second reconciliation mode exists for "consolidated" Production Reports
(one row per person for the WHOLE project, every pay period pre-summed, no
invoice-number column at all -- seen from some payroll companies alongside
their normal per-invoice "expanded" export of the same data). There's
nothing to group by invoice in that shape, so _reconcile_person_level matches
a person against every one of their PDF rows across every invoice instead of
one invoice at a time, and sums the whole group into a single Automation
Total. Dispatch between the two modes is automatic: if no Production Report
row anywhere has an invoice number, it's treated as consolidated.
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
    """Word-order-insensitive so "DeMunn, Kevin" (Last, First -- how PDFs
    format it) and "Kevin DeMunn" (First Last -- how some Production
    Reports, e.g. PGC 005's Expanded report, format it with no SSN column
    to fall back on) still compare equal."""
    clean = re.sub(r"[.,']", "", str(name or "").lower()).strip()
    clean = re.sub(r"\s+", " ", clean)
    return " ".join(sorted(clean.split()))


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


# ── AICP classification ───────────────────────────────────────────────────────
# Same 25-category list as the GA AP prompt's aicp_code field -- no category is
# off-limits on any tab, this is steering (22/23/24 will be right most of the
# time here) rather than a restriction.

_AICP_CATEGORIES_TEXT = """1 — Lodging (Hotels, Condos, etc.)
2 — Car Rental: an actual rental car from a rental company, rented for a specific named individual.
3 — Transportation/Truck Rentals/Gasoline/Car Services: everything else in ground transportation.
4 — Airfare Purchase.
5 — Catering/Crafty.
6 — Construction Hardware/Lumber/Supplies.
7 — Office/Production Equipment Rentals and Purchases.
8 — Camera: Package/Rentals/Expendables.
9 — Grip/Electric: Package/Rentals/Expendables.
10 — Sound: Package/Rentals/Walkies/Expendables.
11 — Set Dressing/Props: Rentals/Purchases/Expendables.
12 — Wardrobe: Rentals/Purchases/Dry Cleaning/Laundry.
13 — Makeup/Hair/Special Effects Purchases.
14 — Location Fees/Permits.
15 — Facility Rental: Office.
16 — Facility Rental: Stage/Warehouse.
17 — Post Editing in Georgia (post-production/editing services with a Georgia vendor address only).
18 — Original Music Scored.
19 — Other: catch-all when nothing else fits but you can tell what the payment is for.
20 — Off-Duty Government Personnel (Police/Fire): only when explicitly off-duty police/fire.
21 — Security Personnel: default for general security services when off-duty status isn't stated.
22 — Georgia Crew Hires: below-the-line crew labor (gaffer, grip, PA, etc.) -- the default for a normal payroll wage row.
23 — Georgia Cast Hires: a credited/principal performer being paid through payroll instead of the Talent pipeline.
24 — Georgia Extras Hires: a background/extra performer being paid through payroll instead of the Talent pipeline.
25 — Per Diem Payments Cast & Crew: a flat daily allowance not tied to hours/wages."""


def _classify_aicp_codes(rows: list[dict], openai_key: str) -> None:
    """Assigns row["aicpCode"] (int, or None if unclassifiable) to every row,
    in place, via one batched GPT call rather than one call per person.

    Best-effort: any failure (no key, API error, malformed response) leaves
    every row's aicpCode as None rather than raising -- a missing AICP code
    is something a reviewer can fill in by hand, not worth failing the whole
    reconciliation over.
    """
    for row in rows:
        row["aicpCode"] = None

    if not rows or not openai_key:
        return

    try:
        from openai import OpenAI
        client = OpenAI(api_key=openai_key)

        items = [
            {
                "index":       i,
                "worker":      row.get("worker", ""),
                "jobTitle":    row.get("jobTitle", ""),
                "wages":       row.get("wages"),
                "total":       row.get("total"),
            }
            for i, row in enumerate(rows)
        ]

        prompt = (
            "You are classifying rows on a Georgia film production's Crew Payroll Report "
            "by AICP category, for the state tax incentive submission.\n\n"
            "Pick the single best-matching AICP category number for each row below, using "
            "this list:\n"
            f"{_AICP_CATEGORIES_TEXT}\n\n"
            "On this tab, categories 22, 23, and 24 (Georgia Crew/Cast/Extras Hires) will be "
            "correct for the large majority of rows -- most rows are simply below-the-line "
            "crew being paid regular wages (22). But this is steering, not a restriction: if a "
            "row's job title clearly indicates something else (a flat per diem with no hours "
            "tied to it -- 25; or genuinely a different category if the data clearly shows it), "
            "use the correct one instead of defaulting to 22.\n\n"
            "Rows to classify:\n"
            f"{json.dumps(items, indent=2)}\n\n"
            "Return ONLY a JSON object mapping each row's index (as a string) to its AICP "
            'category number (an integer): {"0": 22, "1": 25, ...}\n'
            "Every index above must appear in your response. No explanation. No markdown. "
            "No code fences. JSON object only."
        )

        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=4096,
            response_format={"type": "json_object"},
        )
        raw = json.loads(resp.choices[0].message.content)
        if not isinstance(raw, dict):
            return

        for key, val in raw.items():
            try:
                idx = int(key)
                code = int(val)
            except (TypeError, ValueError):
                continue
            if 0 <= idx < len(rows) and 1 <= code <= 25:
                rows[idx]["aicpCode"] = code
    except Exception:
        pass  # leave every aicpCode as None -- reviewable by hand, not fatal


# ── Person-level reconciliation (no invoice numbers at all) ────────────────────
# Some payroll companies' "consolidated" Production Reports (e.g. PGC 005's
# "2) Consolidated") have exactly one row per person for the WHOLE project --
# every pay period already summed together, with no invoice-number-equivalent
# column at all. The per-invoice matcher above can't apply to that (there's
# nothing to group invoices by), so this is a separate mode: match a person
# against every PDF row of theirs across every invoice, and sum the whole
# group's totals into one Automation Total to compare against the
# Production Report's single already-aggregated figure.

def _reconcile_person_level(
    pdf_rows: list[dict],
    production_report_rows: list[dict],
    openai_key: str,
) -> dict:
    issues: list[str] = []

    pdf_groups: dict[str, list[dict]] = defaultdict(list)
    for row in pdf_rows:
        key = _ssn_last4(row.get("ssn")) or _normalize_name(row.get("worker"))
        if key:
            pdf_groups[key].append(row)

    consumed: set[str] = set()
    report_out: list[dict] = []

    for r_row in production_report_rows:
        r_ssn  = _ssn_last4(r_row.get("ssn"))
        r_name = _normalize_name(r_row.get("worker"))
        group_key = None
        if r_ssn and r_ssn in pdf_groups and r_ssn not in consumed:
            group_key = r_ssn
        elif r_name in pdf_groups and r_name not in consumed:
            group_key = r_name

        row = dict(r_row)
        if group_key:
            consumed.add(group_key)
            group = pdf_groups[group_key]
            for p_row in group:
                for field, value in p_row.items():
                    if field in ("worker", "invoiceNo"):
                        continue
                    if row.get(field) in (None, "") and value not in (None, ""):
                        row[field] = value
            row["onProductionReport"] = True
            row["onInvoicePdf"]       = True
            row["automationTotal"]    = round(sum((p.get("total") or 0) for p in group), 2)
        else:
            row["onProductionReport"] = True
            row["onInvoicePdf"]       = False
            row["automationTotal"]    = None
            issues.append(
                f"{r_row.get('worker', '(unnamed)')} is on the Production Report "
                "but no matching PDF invoice(s) were found across the whole batch."
            )
        report_out.append(row)

    # LLM fuzzy fallback for anything still unmatched -- safe to run globally
    # here (unlike the per-invoice matcher, which only fires once an invoice
    # already has a confirmed pairing) since there's only one scope, the
    # whole project, so there's no wrong-invoice cross-matching risk.
    if openai_key:
        remaining_report_idx = [
            i for i, r in enumerate(report_out) if not r["onInvoicePdf"]
        ]
        remaining_pdf_keys = [k for k in pdf_groups if k not in consumed]
        if remaining_report_idx and remaining_pdf_keys:
            report_names = [report_out[i].get("worker", "") for i in remaining_report_idx]
            pdf_names    = [pdf_groups[k][0].get("worker", "") for k in remaining_pdf_keys]
            matches, _ = _llm_fuzzy_match_payroll(pdf_names, report_names, openai_key)
            name_to_key   = {pdf_groups[k][0].get("worker", ""): k for k in remaining_pdf_keys}
            name_to_ridx  = {report_out[i].get("worker", ""): i for i in remaining_report_idx}
            for pdf_name, report_name in matches.items():
                key  = name_to_key.get(pdf_name)
                ridx = name_to_ridx.get(report_name)
                if key is None or ridx is None or key in consumed:
                    continue
                consumed.add(key)
                group = pdf_groups[key]
                row = report_out[ridx]
                for p_row in group:
                    for field, value in p_row.items():
                        if field in ("worker", "invoiceNo"):
                            continue
                        if row.get(field) in (None, "") and value not in (None, ""):
                            row[field] = value
                row["onInvoicePdf"]    = True
                row["automationTotal"] = round(sum((p.get("total") or 0) for p in group), 2)

    pdf_only_out: list[dict] = []
    for key, group in pdf_groups.items():
        if key in consumed:
            continue
        for p_row in group:
            row = dict(p_row)
            row["onProductionReport"] = False
            row["onInvoicePdf"]       = True
            row["automationTotal"]    = p_row.get("total")
            pdf_only_out.append(row)
        issues.append(
            f"{group[0].get('worker', '(unnamed)')} is on a PDF invoice but not "
            "on the Production Report."
        )

    # No invoice number exists anywhere in this mode, so none of the three
    # invoice-based sort options apply -- everything just sorts by name, with
    # the (hopefully rare) PDF-only stragglers grouped at the end.
    report_out.sort(key=lambda r: _normalize_name(r.get("worker")))
    pdf_only_out.sort(key=lambda r: _normalize_name(r.get("worker")))
    final_rows = report_out + pdf_only_out

    _classify_aicp_codes(final_rows, openai_key)

    return {"rows": final_rows, "issues": issues}


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
    if production_report_rows and not any(
        _norm_invoice(r.get("invoiceNo")) for r in production_report_rows
    ):
        # No row in this Production Report carries an invoice number at all
        # -- it's a "consolidated" one-row-per-person report, not one this
        # module can group by invoice. Dispatch to the person-level path
        # instead of silently bucketing every row under invoice "".
        return _reconcile_person_level(pdf_rows, production_report_rows, openai_key)

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

    report_rows:   list[dict] = []  # has a Production Report row: matched + report-only
    pdf_only_rows: list[dict] = []  # exists ONLY on a PDF, no Production Report row at all

    for inv_no in all_invoices:
        pdf_list    = list(pdf_by_invoice.get(inv_no, []))
        report_list = list(report_by_invoice.get(inv_no, []))

        pairs, pdf_consumed, report_consumed, llm_reasons = _match_invoice(
            pdf_list, report_list, openai_key,
        )

        for pi, ri in pairs:
            p_idx, p_row = pdf_list[pi]
            r_idx, r_row = report_list[ri]
            row = dict(r_row)  # Production Report wins on every field it actually has a value for
            for field, value in p_row.items():
                if field in ("worker", "invoiceNo"):
                    continue  # matching keys -- Production Report's own spelling wins outright
                if row.get(field) in (None, "") and value not in (None, ""):
                    row[field] = value  # PDF fills in whatever the Production Report left blank
            row["onProductionReport"] = True
            row["onInvoicePdf"]       = True
            row["automationTotal"]    = p_row.get("total")
            report_rows.append({"row": row, "pdf_idx": p_idx, "report_idx": r_idx})

        for ri, (r_idx, r_row) in enumerate(report_list):
            if ri in report_consumed:
                continue
            row = dict(r_row)
            row["onProductionReport"] = True
            row["onInvoicePdf"]       = False
            row["automationTotal"]    = None
            report_rows.append({"row": row, "pdf_idx": None, "report_idx": r_idx})
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
            pdf_only_rows.append({"row": row, "pdf_idx": p_idx, "report_idx": None})
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
        # Real reconciliation: everything with a Production Report row --
        # matched or not -- has a real name/invoice/position to sort by, so
        # it all sorts together per the chosen layout. Only rows that exist
        # EXCLUSIVELY on a PDF, with no Production Report row at all, get
        # pushed to the end as the "needs review" bucket, grouped together
        # and sorted by name so they're easy to spot.
        _sort_by_effective(report_rows)
        pdf_only_rows.sort(key=lambda e: _normalize_name(e["row"].get("worker")))
        final_entries = report_rows + pdf_only_rows
    else:
        # Only one source was ever provided -- there's nothing to reconcile,
        # so every row ends up in one list or the other trivially. Sort the
        # whole thing per the resolved layout instead of forcing alphabetical
        # order.
        final_entries = report_rows + pdf_only_rows
        _sort_by_effective(final_entries)

    final_rows = [e["row"] for e in final_entries]

    _classify_aicp_codes(final_rows, openai_key)

    return {"rows": final_rows, "issues": issues}
