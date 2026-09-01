"""
AI-assisted fringe extractor for unknown payroll company layouts.

Called when no static parser matches a PDF. Two GPT-4o calls:
  1. Extract fringe rows + identify company name + text markers
  2. Generate a Python parser for permanent deployment

Same-batch reuse: main.py exec()s the generated code and caches it keyed
by company name. Subsequent files with the same text markers skip the AI.
"""

import re
import io
import json

import pdfplumber

_MAX_TEXT_PAGES = 8


# ─── Prompts ──────────────────────────────────────────────────────────────────

_EXTRACT_SYSTEM = """
You are a payroll data extraction specialist. Analyze this fringe/payroll report PDF text, then extract every employee row.

Return ONLY valid JSON — no markdown, no explanation:
{
  "company_name": "Exact payroll company or software name (e.g. 'Revolution Entertainment Services', 'Cast & Crew')",
  "text_markers": ["2-3 short strings that uniquely identify this PDF layout — e.g. header titles or software names. Avoid generic terms like 'Payroll Report'."],
  "layout_description": "Concise technical description of this document for a developer writing a parser. Include: which page/section contains the fringe data, the exact column order left-to-right, how employee rows are identified, SSN and name format, and any summary/total rows to skip.",
  "rows": [
    {
      "worker": "LAST, FIRST",
      "ssn": "last 4 digits as string e.g. '3421'",
      "invoiceNo": "invoice or check number as string, or '' if consolidated",
      "invoiceDate": "MM/DD/YYYY or ''",
      "workDates": "date range string or ''",
      "union": "union code or ''",
      "wages": 0.00,
      "reimbRent": 0.00,
      "corporate": 0.00,
      "socSec": 0.00,
      "med": 0.00,
      "futa": 0.00,
      "sui": 0.00,
      "wc": 0.00,
      "phw": 0.00,
      "vacHol": 0.00,
      "adv": 0.00,
      "other": 0.00,
      "hand": 0.00,
      "total": 0.00,
      "withholdingsIL": null,
      "jobTitle": "",
      "daysWorked": null,
      "street": "",
      "city": "",
      "zip": ""
    }
  ]
}

Rules:
- worker must be "LAST, FIRST" format — convert if needed
- ssn: last 4 digits only (strip XXX-XX- or similar prefix)
- All monetary fields: numbers not strings; use null if the field is absent
- text_markers: must be UNIQUE to this company — not "Total", "Page 1", or "Payroll"
- layout_description: be specific and technical — column names, page markers, row patterns
"""

_CODEGEN_SYSTEM = """
You are a Python developer writing a production payroll fringe parser.

The module must export three registry-required names:
  COMPANY  = "company_slug"        # lowercase with underscores, e.g. "ep_financial"
  MARKERS  = ["text1", "text2"]    # 2-3 strings UNIQUE to this layout (not generic)
  PRIORITY = 20                    # AI-generated parsers use 20

And one public function:
  def extract(pdf_bytes: bytes, **_) -> tuple[list[dict], list[str]]

Requirements:
- Use pdfplumber to extract text
- Parse the specific layout of the company's fringe/payroll report
- Return (rows, issues); rows must use empty_row() as the base dict
- Handle multi-page PDFs; wrap body in try/except and append to issues on error

Canonical fields to populate per row:
  worker, ssn, invoiceNo, invoiceDate, workDates, union, wages, reimbRent,
  corporate, socSec, med, futa, sui, wc, phw, vacHol, adv, other, hand,
  total, withholdingsIL, jobTitle, daysWorked, street, city, zip

Allowed imports ONLY:
  import re, io
  import pdfplumber
  from parsers.base import empty_row, parse_amount, clean_fringe_name

Return ONLY the Python code — no markdown fences, no explanation.
"""


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _pdf_text(pdf_bytes: bytes, max_pages: int = _MAX_TEXT_PAGES) -> str:
    """Extract and join text from the first N pages of a PDF."""
    parts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pg in pdf.pages[:max_pages]:
                text = pg.extract_text() or ""
                if text.strip():
                    parts.append(f"--- Page {pg.page_number} ---\n{text}")
    except Exception:
        pass
    return "\n\n".join(parts)


def _gpt(client, system: str, user: str, max_tokens: int = 4096) -> str:
    """GPT-4o call — used for document reading / data extraction (Step 1)."""
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system},
            {"role": "user",   "content": user},
        ],
        temperature=0,
        max_tokens=max_tokens,
    )
    return resp.choices[0].message.content.strip()


def _claude(anthropic_key: str, system: str, user: str, max_tokens: int = 4096) -> str:
    """Claude call — used for code generation (Step 2)."""
    import anthropic
    client = anthropic.Anthropic(api_key=anthropic_key)
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return msg.content[0].text.strip()


def _parse_json(raw: str) -> dict | None:
    """Parse a GPT response as JSON, stripping markdown fences if present."""
    text = raw
    m = re.search(r"```(?:json)?\s*([\s\S]+?)```", text)
    if m:
        text = m.group(1)
    try:
        return json.loads(text)
    except Exception:
        return None


# ─── Public API ───────────────────────────────────────────────────────────────

def make_exec_parser(code: str):
    """exec() a generated parser module and return its extract function, or None.

    Logs the real failure reason instead of swallowing it -- a debug run
    against 13 real CAPS invoices (2026-09-01) found this failing on 4 of 7
    attempts, and with no visibility into why it was impossible to tell a
    syntax error apart from a disallowed import or a runtime error."""
    try:
        import pdfplumber as _pdf
        from parsers.base import empty_row, parse_amount, clean_fringe_name
        ns: dict = {
            "re": re,
            "io": io,
            "pdfplumber": _pdf,
            "empty_row": empty_row,
            "parse_amount": parse_amount,
            "clean_fringe_name": clean_fringe_name,
        }
        exec(compile(code, "<ai_generated>", "exec"), ns)  # noqa: S102
        fn = ns.get("extract")
        if not callable(fn):
            print("[make_exec_parser] generated module has no callable 'extract' function", flush=True)
            return None
        return fn
    except Exception as e:
        print(f"[make_exec_parser] generated code failed to compile/exec: {e!r}", flush=True)
        return None


def extract_unknown(
    pdf_bytes: bytes,
    openai_key: str,
    hints: str = "",
    anthropic_key: str = "",
) -> tuple[list[dict], list[str], str, list[str], str]:
    """Use AI to extract fringe rows from an unknown payroll company PDF.

    hints: optional free-text context from the user (company name, column descriptions, etc.)
           injected into the extraction prompt to improve accuracy.

    Returns:
        rows         — canonical fringe dicts
        issues       — error/warning strings
        company_name — detected company name ('' on failure)
        text_markers — strings to identify this company in future PDFs
        code         — generated Python parser ('' on failure)
    """
    from parsers.base import empty_row as _empty_row

    if not openai_key:
        return [], ["AI extraction requires OPENAI_API_KEY"], "", [], ""

    try:
        from openai import OpenAI
        client = OpenAI(api_key=openai_key)
    except Exception as e:
        return [], [f"OpenAI init error: {e}"], "", [], ""

    sample_text = _pdf_text(pdf_bytes)
    if not sample_text.strip():
        return [], ["No text layer found — AI text extraction not possible"], "", [], ""

    user_message = sample_text
    if hints and hints.strip():
        user_message = f"User-provided hints about this document:\n{hints.strip()}\n\n{sample_text}"

    # ── Step 1: extract rows + identify company ────────────────────────────────
    try:
        raw     = _gpt(client, _EXTRACT_SYSTEM, user_message)
        parsed  = _parse_json(raw)
    except Exception as e:
        return [], [f"AI extraction call failed: {e}"], "", [], ""

    if not parsed or not isinstance(parsed.get("rows"), list):
        return [], ["AI returned unexpected structure — could not parse rows"], "", [], ""

    company_name     = str(parsed.get("company_name", "Unknown Payroll Company")).strip()
    text_markers     = parsed.get("text_markers") or [company_name]
    layout_description = str(parsed.get("layout_description", "")).strip()

    rows = []
    for r in parsed.get("rows", []):
        row = _empty_row()
        row["payrollCompany"] = company_name
        for field in (
            "worker", "ssn", "invoiceNo", "invoiceDate", "workDates", "union",
            "wages", "reimbRent", "corporate", "socSec", "med", "futa", "sui",
            "wc", "phw", "vacHol", "adv", "other", "hand", "total",
            "withholdingsIL", "jobTitle", "daysWorked", "street", "city", "zip",
        ):
            val = r.get(field)
            if val is not None:
                row[field] = val
        rows.append(row)

    # ── Step 2: generate Python parser code (Claude) ──────────────────────────
    codegen_user_parts = [f"Company: {company_name}\n"]
    if layout_description:
        codegen_user_parts.append(f"Layout description (auto-detected):\n{layout_description}\n")
    if hints and hints.strip():
        codegen_user_parts.append(f"Additional user hints:\n{hints.strip()}\n")
    codegen_user_parts.append(
        f"Sample PDF text (first {_MAX_TEXT_PAGES} pages):\n{sample_text[:6000]}\n\n"
        f"Extracted rows for reference (first 3):\n"
        f"{json.dumps(parsed.get('rows', [])[:3], indent=2)}"
    )
    codegen_user = "\n".join(codegen_user_parts)
    try:
        if anthropic_key:
            code = _claude(anthropic_key, _CODEGEN_SYSTEM, codegen_user, max_tokens=3000)
        else:
            code = _gpt(client, _CODEGEN_SYSTEM, codegen_user, max_tokens=3000)
        m_fen = re.search(r"```(?:python)?\s*([\s\S]+?)```", code)
        if m_fen:
            code = m_fen.group(1).strip()
    except Exception as e:
        code = f"# Code generation failed: {e}\n"

    return rows, [], company_name, text_markers, code
