// wrapbook-fringe.js
// Deterministic, in-browser extractor for Wrapbook "Fringe Report" pages.
//
// A Wrapbook payroll PDF bundles several report types (Invoice, Employer
// Payroll, Fringe Report, Payroll Register). The Fringe Report page(s) carry
// the ~85% of data we need for the Crew Payroll tab. This module finds those
// pages, reconstructs the table from PDF text + coordinates (pdf.js has no
// table detection, so we cluster items by Y into rows and assign them to
// columns by X against the detected header), and returns clean canonical rows.
//
// Returns raw fringe fields ONLY — mapping to workbook columns happens in the
// wizard so it can evolve without touching the parser.
//
//   const { rows, sections, issues } = await extractFringe(arrayBuffer);
//
// Each row: { worker, ssn, workDates, type, dept, union, workState, resState,
//             wages, reimbRent, socSec, med, futa, sui, wc, phw, other,
//             benefits, platFee, effRate, total, invoiceNo, invoiceDate,
//             invoiceWorkDates, sourcePage }

const PDFJS_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.7.76/build/pdf.min.mjs";
const PDFJS_WORKER = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.7.76/build/pdf.worker.min.mjs";

let _pdfjs = null;
async function getPdfjs() {
  if (_pdfjs) return _pdfjs;
  _pdfjs = await import(PDFJS_URL);
  _pdfjs.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
  return _pdfjs;
}

// Header label → canonical field. Order matters (match "Work Dates" before "Work").
const COLS = [
  ["Worker", "worker"], ["SSN", "ssn"], ["Work Dates", "workDates"], ["Type", "type"],
  ["Dept", "dept"], ["Union", "union"], ["Work", "workState"], ["Res", "resState"],
  ["Wages", "wages"], ["Reimb/Rent", "reimbRent"], ["Soc Sec", "socSec"], ["Med", "med"],
  ["FUTA", "futa"], ["SUI", "sui"], ["W/C", "wc"], ["PH&W", "phw"], ["Other", "other"],
  ["Benefits", "benefits"], ["Plat Fee", "platFee"], ["EFF Rate %", "effRate"], ["Total", "total"],
];
const NUMERIC = new Set([
  "wages", "reimbRent", "socSec", "med", "futa", "sui", "wc", "phw",
  "other", "benefits", "platFee", "effRate", "total",
]);

// Cluster text items into visual lines (same baseline within 3px), sorted top→down.
function lineify(items) {
  const its = items
    .filter((i) => i.str.trim() !== "")
    .map((i) => ({ x: +i.transform[4].toFixed(1), y: +i.transform[5].toFixed(1), s: i.str.trim() }));
  its.sort((a, b) => b.y - a.y || a.x - b.x);
  const lines = [];
  let cur = null;
  for (const it of its) {
    if (!cur || Math.abs(cur.y - it.y) > 3) { cur = { y: it.y, its: [it] }; lines.push(cur); }
    else cur.its.push(it);
  }
  lines.forEach((l) => l.its.sort((a, b) => a.x - b.x));
  return lines;
}

// From the header line, read each column label's x-position and build boundaries.
function detectColumns(lines) {
  let hdr = null;
  for (const l of lines) {
    const s = l.its.map((i) => i.s);
    if (s.includes("Worker") && s.includes("Total")) { hdr = l; break; }
  }
  if (!hdr) return null;
  const anchors = [];
  for (const [label, field] of COLS) {
    const it = hdr.its.find((i) => i.s === label);
    if (it) anchors.push({ field, x: it.x });
  }
  anchors.sort((a, b) => a.x - b.x);
  for (let i = 0; i < anchors.length; i++) {
    anchors[i].left = i === 0 ? -1e9 : (anchors[i - 1].x + anchors[i].x) / 2;
    anchors[i].right = i === anchors.length - 1 ? 1e9 : (anchors[i].x + anchors[i + 1].x) / 2;
  }
  return { headerY: hdr.y, anchors };
}
function colFor(cols, x) {
  for (const a of cols.anchors) if (x >= a.left && x < a.right) return a.field;
  return null;
}

function parseMoney(s) {
  if (s == null || s === "") return null;
  const n = parseFloat(String(s).replace(/[$,%\s]/g, ""));
  return isNaN(n) ? null : n;
}

// Reconstruct a text column's value from fragments spread over wrapped lines.
function mergeText(field, frags) {
  if (field === "worker") {
    // Names are "Last, First M". Last name may wrap mid-word (no hyphen), so
    // concatenate everything up to the comma with NO space; join the given-name
    // fragments after the comma with spaces; drop a trailing middle initial.
    const ci = frags.findIndex((f) => f.includes(","));
    if (ci < 0) return frags.join(" ").replace(/\s+/g, " ").trim();
    const commaFrag = frags[ci];
    const beforeComma = commaFrag.slice(0, commaFrag.indexOf(",") + 1);
    const afterComma = commaFrag.slice(commaFrag.indexOf(",") + 1).trim();
    const last = frags.slice(0, ci).join("") + beforeComma;
    const firstFrags = [];
    if (afterComma) firstFrags.push(afterComma);
    firstFrags.push(...frags.slice(ci + 1));
    let first = firstFrags.join(" ").replace(/\s+/g, " ").trim();
    first = first.replace(/\s+[A-Z]\.?$/, "").trim(); // drop middle initial
    return (last.replace(/,$/, "") + ", " + first).trim();
  }
  if (field === "ssn") return frags.join("");
  // generic: mid-word wrap (fragment ends with "x-") joins with no space; else space
  let out = "";
  for (let i = 0; i < frags.length; i++) {
    if (i === 0) out = frags[i];
    else if (/\S-$/.test(out)) out += frags[i];
    else out += " " + frags[i];
  }
  return out.replace(/\s+/g, " ").trim();
}

// Extract employee rows from one fringe page.
function extractRowsFromPage(lines, cols, hasHeader) {
  const startY = hasHeader ? cols.headerY - 5 : 1e9;
  const data = lines.filter(
    (l) => l.y < startY && !/Page \d+ of \d+/.test(l.its.map((i) => i.s).join(" "))
  );
  const wagesAnchor = cols.anchors.find((a) => a.field === "wages");
  // A "primary" line carries the money values; continuation lines carry only
  // wrapped text (name/ssn/date/dept overflow). Each employee = 1 primary + its
  // trailing continuations.
  const isPrimary = (l) => l.its.some((i) => i.x >= wagesAnchor.left && /\d\.\d{2}/.test(i.s));
  const recs = [];
  let cur = null;
  for (const l of data) {
    if (isPrimary(l)) {
      if (cur) recs.push(cur);
      cur = { frags: {} };
    }
    if (!cur) continue;
    for (const it of l.its) {
      const f = colFor(cols, it.x);
      if (f) (cur.frags[f] = cur.frags[f] || []).push(it.s);
    }
  }
  if (cur) recs.push(cur);
  return recs
    .map((r) => {
      const row = {};
      for (const [, field] of COLS) {
        const frags = r.frags[field] || [];
        if (NUMERIC.has(field)) row[field] = frags.length ? parseMoney(frags[0]) : null;
        else row[field] = frags.length ? mergeText(field, frags) : "";
      }
      return row;
    })
    .filter((r) => r.worker && r.worker.length > 1); // skip grand-total row (no worker)
}

// Parse the meta header (Invoice # / Invoice Date / Work Dates) of a fringe section.
function parseMeta(lines) {
  const joined = lines.map((l) => l.its.map((i) => i.s).join(" ")).join("\n");
  let invoiceNo = "", invoiceDate = "", workDates = "";
  const m = joined.match(
    /Invoice #\s+Invoice Date\s+Work Dates[\s\S]*?\n\s*(\d{6,8})\s+(\d{2}\/\d{2}\/\d{4})\s+([A-Za-z]+ \d+, \d{4}\s*-\s*[A-Za-z]+ \d+, \d{4})/
  );
  if (m) { invoiceNo = m[1]; invoiceDate = m[2]; workDates = m[3].replace(/\s+/g, " "); }
  return { invoiceNo, invoiceDate, workDates };
}

// ── public API ──
export async function extractFringe(arrayBuffer) {
  const pdfjs = await getPdfjs();
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;

  const pageLines = [], pageText = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const pg = await pdf.getPage(p);
    const tc = await pg.getTextContent();
    pageLines[p] = lineify(tc.items);
    pageText[p] = tc.items.map((i) => i.str).join(" ");
  }

  const sections = [];
  const rows = [];
  const issues = [];
  let p = 1;
  while (p <= pdf.numPages) {
    if (pageText[p].includes("Fringe Report")) {
      const meta = parseMeta(pageLines[p]);
      const footer = pageText[p].match(/Page \d+ of (\d+)/);
      const span = footer ? +footer[1] : 1; // fringe report may run several pages
      const cols = detectColumns(pageLines[p]);
      if (!cols) {
        issues.push("Fringe Report page " + p + ": could not detect table header.");
        p += span;
        continue;
      }
      const secRows = [];
      for (let fp = p; fp < p + span && fp <= pdf.numPages; fp++) {
        const pageRows = extractRowsFromPage(pageLines[fp], cols, fp === p);
        for (const r of pageRows) {
          r.invoiceNo = meta.invoiceNo;
          r.invoiceDate = meta.invoiceDate;
          r.invoiceWorkDates = meta.workDates;
          r.sourcePage = fp;
          secRows.push(r);
        }
      }
      sections.push({ meta, pages: [p, p + span - 1], count: secRows.length });
      rows.push(...secRows);
      p += span;
    } else {
      p++;
    }
  }

  if (!sections.length) issues.push("No Fringe Report page found in this PDF.");
  return { rows, sections, issues };
}
