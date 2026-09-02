"""
Parser registry — auto-discovers layout modules and provides unified dispatch.

Each parser module in any parsers/ subdirectory must export:
  COMPANY  : str        — company identifier (e.g. "caps", "wrapbook", "ep_financial")
  MARKERS  : list[str]  — text strings unique to this layout; checked against first N pages
  PRIORITY : int        — optional, lower = tried first within a company (default 10)
  extract(pdf_bytes, **kwargs) -> tuple[list[dict], list[str]]

Optionally, a module may export:
  can_parse(pdf_bytes) -> bool  — replaces marker matching with explicit detection logic

AI-generated parsers are registered at runtime via register_parser() and persist
for the process lifetime — no restart or redeploy needed until GitHub commit.

Usage in main.py:
  candidates = registry.find_parsers(pdf_bytes)
  for candidate in candidates:
      rows, errs = candidate.extract(pdf_bytes, openai_key=key)
      if rows:
          break
  # If no candidate returned rows → Phase 3
"""

import io
import importlib
import pkgutil
import pdfplumber
import parsers as _parsers_pkg

_STATIC:  list = []
_RUNTIME: list = []
_loaded:  bool = False

_SKIP = {"base", "registry", "ai_fringe", "register"}


def _scan():
    global _STATIC, _loaded
    if _loaded:
        return
    _loaded = True
    found = []
    for _, modname, ispkg in pkgutil.walk_packages(
        path=_parsers_pkg.__path__,
        prefix="parsers.",
        onerror=lambda x: None,
    ):
        if ispkg:
            continue
        if modname.split(".")[-1] in _SKIP:
            continue
        try:
            mod = importlib.import_module(modname)
            if (hasattr(mod, "COMPANY")
                    and hasattr(mod, "MARKERS")
                    and hasattr(mod, "extract")):
                found.append(mod)
        except Exception:
            pass
    _STATIC = sorted(found, key=lambda m: (getattr(m, "PRIORITY", 10), m.__name__))


def _page_texts(pdf_bytes: bytes) -> list[str]:
    """Scans every page, not a fixed window -- a large multi-employee CAPS
    invoice's Payroll Register section (one entry per employee) can push its
    Fringe Recap Report marker page well past any small fixed limit. Confirmed
    on real Coke 012 invoices where that marker didn't appear until page 21 and
    25: a 15-page-limited scan found no candidate parser at all, and the file
    fell through to the generic AI-guess fallback instead of the CAPS parser
    that would have handled it correctly (same failure shape as the identical
    bug fixed in _is_wrapbook_register_only)."""
    texts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pg in pdf.pages:
                t = (pg.extract_text() or "").lower()
                if t.strip():
                    texts.append(t)
    except Exception:
        pass
    return texts


def find_parsers(pdf_bytes: bytes) -> list:
    """Return all matching parser objects in priority order (static first, then runtime).

    Static parsers have PRIORITY 10 (hand-written).
    Runtime AI-generated parsers have PRIORITY 20 (tried after all static layouts fail).

    The caller should try each in sequence and use the first that returns rows.
    An empty list means no known parser — trigger Phase 3 AI generation.
    """
    _scan()
    texts = _page_texts(pdf_bytes)
    if not texts:
        return []  # image-only — handled upstream

    # Whitespace-normalized copy for marker matching only -- a rotated page
    # commonly extracts with one word per line (e.g. a real CAPS "Fringe
    # Recap Report" page came back as "fringe\nrecap\nreport"), which a plain
    # substring check against the raw per-page text never matches even
    # though the marker is genuinely present. Confirmed on a real batch of 9
    # CAPS invoices that were misclassified as an unrecognized payroll
    # company for exactly this reason -- same failure shape as the identical
    # bug already fixed in _is_wrapbook_register_only. Collapsing whitespace
    # is a no-op for markers that were already on one line, so this only
    # ever makes matching more permissive, never less.
    normalized_texts = [" ".join(t.split()) for t in texts]

    result = []
    for mod in _STATIC + _RUNTIME:
        matched = False
        if hasattr(mod, "can_parse"):
            matched = mod.can_parse(pdf_bytes)
        else:
            for marker in getattr(mod, "MARKERS", []):
                marker_norm = " ".join(marker.lower().split())
                if any(marker_norm in t for t in normalized_texts):
                    matched = True
                    break
        if matched:
            result.append(mod)
    return result


def register_parser(
    company: str,
    markers: list[str],
    extract_fn,
    priority: int = 20,
):
    """Register or replace a runtime parser (AI-generated).

    Persists for the lifetime of the process — subsequent batches reuse it.
    Priority 20 means it is tried after all static layout files for the same company.
    """
    global _RUNTIME
    _RUNTIME = [m for m in _RUNTIME if getattr(m, "COMPANY", "") != company]

    mod           = type("_Runtime", (), {})()
    mod.COMPANY   = company
    mod.MARKERS   = markers
    mod.PRIORITY  = priority
    mod.__name__  = f"parsers.runtime.{company}"
    mod.extract   = extract_fn
    _RUNTIME.append(mod)
    return mod


