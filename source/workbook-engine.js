// workbook-engine.js
// Surgical XLSX cell-patcher. Opens an .xlsx (a ZIP of XML), rewrites only the
// target cells inside one worksheet, and re-zips with every other entry kept
// byte-for-byte (original compressed blocks copied verbatim). This preserves
// formulas, number formats, styles, drawings, comments, and external links.
//
// Runs in the browser (uses DecompressionStream for the few XML parts we read).

// ---------- CRC32 ----------
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ---------- inflate one entry ----------
async function inflateRaw(bytes, method) {
  if (method === 0) return bytes;
  const ds = new DecompressionStream("deflate-raw");
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

// ---------- parse ZIP central directory ----------
function parseZip(buf) {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("Not a valid .xlsx (no EOCD)");
  const count = dv.getUint16(eocd + 10, true);
  const cdOffset = dv.getUint32(eocd + 16, true);
  const td = new TextDecoder();
  const entries = [];
  let p = cdOffset;
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const crc = dv.getUint32(p + 16, true);
    const compSize = dv.getUint32(p + 20, true);
    const uncompSize = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const name = td.decode(buf.subarray(p + 46, p + 46 + nameLen));
    // locate the compressed data via the local header
    const lhNameLen = dv.getUint16(localOff + 26, true);
    const lhExtraLen = dv.getUint16(localOff + 28, true);
    const dataStart = localOff + 30 + lhNameLen + lhExtraLen;
    const comp = buf.subarray(dataStart, dataStart + compSize);
    entries.push({ name, method, crc, compSize, uncompSize, comp });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

async function readText(entries, name) {
  const e = entries.find((x) => x.name === name);
  if (!e) return null;
  return new TextDecoder().decode(await inflateRaw(e.comp, e.method));
}

// ---------- build a ZIP ----------
function buildZip(parts) {
  // parts: [{ name, bytes (Uint8Array, uncompressed), method, crc, comp (Uint8Array) }]
  // Modified parts pass bytes+stored(method 0). Copied parts pass comp+method+crc+uncompSize.
  const enc = new TextEncoder();
  const chunks = [];
  const central = [];
  let offset = 0;
  for (const part of parts) {
    const nameBytes = enc.encode(part.name);
    const lh = new Uint8Array(30 + nameBytes.length);
    const dv = new DataView(lh.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);          // version needed
    dv.setUint16(6, 0, true);           // flags
    dv.setUint16(8, part.method, true); // method
    dv.setUint16(10, 0, true);          // mod time
    dv.setUint16(12, 0x21, true);       // mod date (1980-01-01)
    dv.setUint32(14, part.crc, true);
    dv.setUint32(18, part.compSize, true);
    dv.setUint32(22, part.uncompSize, true);
    dv.setUint16(26, nameBytes.length, true);
    dv.setUint16(28, 0, true);          // extra len
    lh.set(nameBytes, 30);

    chunks.push(lh, part.data);
    const localOffset = offset;
    offset += lh.length + part.data.length;

    const ch = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0, true);
    cv.setUint16(10, part.method, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0x21, true);
    cv.setUint32(16, part.crc, true);
    cv.setUint32(20, part.compSize, true);
    cv.setUint32(24, part.uncompSize, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(42, 0, true);          // external attrs hi/lo skipped
    cv.setUint32(42, localOffset, true);
    ch.set(nameBytes, 46);
    central.push(ch);
  }
  const cdStart = offset;
  let cdSize = 0;
  for (const ch of central) { chunks.push(ch); cdSize += ch.length; }
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, central.length, true);
  ev.setUint16(10, central.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, cdStart, true);
  chunks.push(eocd);
  return new Blob(chunks, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

// ---------- cell helpers ----------
function escapeXml(s) {
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
function colToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n;
}
function dateToSerial(iso) {
  // iso = "YYYY-MM-DD" -> Excel serial (1900 date system, with the 1900 leap bug)
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!m) return null;
  const utc = Date.UTC(+m[1], +m[2] - 1, +m[3]);
  const epoch = Date.UTC(1899, 11, 30); // 1899-12-30
  return Math.round((utc - epoch) / 86400000);
}
function dateToSerialAny(v) {
  // Accepts "YYYY-MM-DD" or "MM/DD/YYYY"
  const s = String(v).trim();
  const md = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (md) {
    const utc = Date.UTC(+md[3], +md[1] - 1, +md[2]);
    return Math.round((utc - Date.UTC(1899, 11, 30)) / 86400000);
  }
  return dateToSerial(s);
}

// ---------- row insertion (append data rows, push existing rows down) ----------
// Shift every row reference >= `at` down by N inside formulas / shared-formula
// ranges. Cross-sheet references must be left completely alone: the insert
// happened on THIS sheet, so `Legend!$D$2:$E$26` still ends at row 26. Guarding
// only the char before the column letter isn't enough — in a qualified RANGE the
// second endpoint is preceded by ":", not "!", so it would still be bumped and
// the range would silently grow. Mask whole qualified refs out, bump, restore.
const QUALIFIED_REF = /(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g;
function bumpFormulaRefs(xml, at, N) {
  return xml.replace(/<f\b[^>]*>[\s\S]*?<\/f>|<f\b[^>]*\/>/g, (seg) => {
    const held = [];
    const masked = seg.replace(QUALIFIED_REF, (ref) => {
      held.push(ref);
      return "\u0000" + (held.length - 1) + "\u0000";
    });
    const bumped = masked.replace(/(?<![!A-Za-z0-9])(\$?[A-Z]{1,3}\$?)(\d+)/g, (m, c, d) => {
      const n = +d;
      return n >= at ? c + (n + N) : m;
    });
    return bumped.replace(/\u0000(\d+)\u0000/g, (m, i) => held[+i]);
  });
}

// Inverse of colToNum: 2 -> "B", 70 -> "BR".
function numToCol(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

// Columns whose formatting is copied onto every inserted row: B .. BR.
const FILL_FROM = colToNum("B");   // 2
const FILL_TO = colToNum("BR");    // 70

// Build one <row> of data. cells: [{ col, value, type }]; styleMap: { col: s }.
// Emits a cell for EVERY column B..BR so the archetype row's formatting (number
// formats, borders, fills) carries across the whole row — a value cell where
// data exists, an empty styled cell everywhere else. Data cells outside B..BR
// (should be none) are still written.
function buildDataRow(R, cells, styleMap) {
  const dataByCol = {};
  for (const c of cells) {
    if (c.value == null || String(c.value).trim() === "") continue;
    dataByCol[c.col] = c;
  }
  const valueXml = (col, c, s) => {
    if (c.type === "number" || c.type === "currency") {
      const n = typeof c.value === "number" ? c.value : parseFloat(String(c.value).replace(/[^0-9.\-]/g, ""));
      if (!isFinite(n)) return null;
      return `<c r="${col}${R}"${s}><v>${n}</v></c>`;
    }
    if (c.type === "date") {
      const ser = dateToSerialAny(c.value);
      if (ser == null) return null;
      return `<c r="${col}${R}"${s}><v>${ser}</v></c>`;
    }
    if (c.type === "formula") {
      // c.value is a formula string; "{R}" is replaced with this row's number
      const f = String(c.value).replace(/\{R\}/g, R);
      return `<c r="${col}${R}"${s}><f>${escapeXml(f)}</f></c>`;
    }
    return `<c r="${col}${R}"${s} t="inlineStr"><is><t xml:space="preserve">${escapeXml(c.value)}</t></is></c>`;
  };
  const parts = [];
  const emit = (col) => {
    const s = styleMap[col] != null ? ` s="${styleMap[col]}"` : "";
    const c = dataByCol[col];
    if (c) {
      const xml = valueXml(col, c, s);
      if (xml) { parts.push({ col, xml }); return; }
    }
    // empty but styled cell, so the archetype format is uniform across the row
    if (styleMap[col] != null) parts.push({ col, xml: `<c r="${col}${R}"${s}/>` });
  };
  for (let n = FILL_FROM; n <= FILL_TO; n++) emit(numToCol(n));
  for (const col in dataByCol) {  // defensive: any data cell outside B..BR
    const n = colToNum(col);
    if (n < FILL_FROM || n > FILL_TO) emit(col);
  }
  parts.sort((a, b) => colToNum(a.col) - colToNum(b.col));
  return `<row r="${R}" spans="1:70" ht="15.6">${parts.map((p) => p.xml).join("")}</row>`;
}

// Insert N data rows starting at row `at`, shifting all existing rows >= at
// down by N. New rows inherit cell styles from the template row currently at
// `at` (so number formats look right). Returns the modified sheet XML.
function insertRowsIntoSheet(sheetXml, at, rows) {
  const N = rows.length;
  if (!N) return sheetXml;
  // style map from the archetype row (the one currently at `at`)
  const arch = (sheetXml.match(new RegExp(`<row r="${at}"[\\s\\S]*?</row>`)) || [])[0] || "";
  const styleMap = {};
  for (const cm of arch.matchAll(/<c r="([A-Z]+)\d+"(?:\s+s="(\d+)")?/g)) styleMap[cm[1]] = cm[2] != null ? cm[2] : null;

  sheetXml = bumpFormulaRefs(sheetXml, at, N);
  const sd = sheetXml.match(/(<sheetData[^>]*>)([\s\S]*?)(<\/sheetData>)/);
  const rowStrs = sd[2].match(/<row\b[\s\S]*?<\/row>/g) || [];
  const out = [];
  for (const rs of rowStrs) {
    const oldR = +rs.match(/<row r="(\d+)"/)[1];
    if (oldR >= at) {
      const nr = oldR + N;
      out.push({ r: nr, xml: rs.replace(/\br="([A-Z]*)\d+"/g, (m, c) => `r="${c}${nr}"`) });
    } else {
      out.push({ r: oldR, xml: rs });
    }
  }
  for (let i = 0; i < N; i++) out.push({ r: at + i, xml: buildDataRow(at + i, rows[i], styleMap) });
  out.sort((a, b) => a.r - b.r);
  sheetXml = sheetXml.replace(sd[0], sd[1] + out.map((o) => o.xml).join("") + sd[3]);
  // extend the sheet dimension's end row
  sheetXml = sheetXml.replace(/(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)("\s*\/>)/, (m, a, d, b) => a + (+d + N) + b);
  return sheetXml;
}

// Style attr (s="..") of a cell, or "" if none.
function styleAttrOf(cellXml) { const m = cellXml.match(/\ss="(\d+)"/); return m ? ` s="${m[1]}"` : ""; }

// Insert N blank slots in columns 1..maxCol starting at row `at`, pushing those
// columns' cells (rows >= at) DOWN by N. Columns > maxCol stay on their original
// rows (so a side table to the right is undisturbed). New slot cells are seeded
// empty but keep the style of whatever occupied that position, so formatting is
// preserved and later value-patches inherit it. No formula bumping — verified
// that nothing (this sheet or others) references these columns positionally.
function insertPartialRows(sheetXml, at, N, maxCol) {
  if (!N) return sheetXml;
  const sd = sheetXml.match(/(<sheetData[^>]*>)([\s\S]*?)(<\/sheetData>)/);
  const orig = new Map();
  const rowRe = /<row r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g;
  let m;
  while ((m = rowRe.exec(sd[2]))) {
    const rnum = +m[1], ae = [], other = [];
    for (const cm of m[3].matchAll(/<c r="([A-Z]+)(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) {
      const col = cm[1], cn = colToNum(col);
      (cn <= maxCol ? ae : other).push({ col, cn, xml: cm[0] });
    }
    orig.set(rnum, { attrs: m[2], ae, other });
  }
  const out = new Map();
  const ensure = (rn, attrs) => { if (!out.has(rn)) out.set(rn, { attrs: attrs || "", cells: [] }); return out.get(rn); };
  for (const [rn, rec] of orig) {
    ensure(rn, rec.attrs);
    for (const c of rec.other) out.get(rn).cells.push(c);
    const target = rn >= at ? rn + N : rn;
    const t = ensure(target, orig.has(target) ? orig.get(target).attrs : rec.attrs);
    for (const c of rec.ae) t.cells.push({ cn: c.cn, xml: c.xml.replace(/r="[A-Z]+\d+"/, `r="${c.col}${target}"`) });
    if (rn >= at && rn < at + N) {  // this row becomes a slot: seed empty styled A–E cells
      const slot = out.get(rn);
      for (const c of rec.ae) slot.cells.push({ cn: c.cn, xml: `<c r="${c.col}${rn}"${styleAttrOf(c.xml)}/>` });
    }
  }
  for (let s = at; s < at + N; s++) if (!out.has(s)) out.set(s, { attrs: "", cells: [] });
  const body = [...out.keys()].sort((a, b) => a - b).map((rn) => {
    const rec = out.get(rn);
    rec.cells.sort((a, b) => a.cn - b.cn);
    return `<row r="${rn}"${rec.attrs}>` + rec.cells.map((c) => c.xml).join("") + `</row>`;
  }).join("");
  sheetXml = sheetXml.replace(sd[0], sd[1] + body + sd[3]);
  sheetXml = sheetXml.replace(/(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)("\s*\/>)/, (mm, a, d, b) => a + (+d + N) + b);
  return sheetXml;
}

// Build the inner of a populated <c>, preserving the existing style attr (s="..").
function buildCell(addr, attrs, value, type, sOverride) {
  // strip any existing t= from attrs, keep s= and anything else
  let keep = attrs.replace(/\s+t="[^"]*"/g, "");
  // An explicit style wins over whatever the template cell carried. Needed when
  // the value's type doesn't match the placeholder's format (a date written into
  // an accounting-formatted cell) or when the cell didn't exist at all.
  if (sOverride != null) keep = keep.replace(/\s+s="\d+"/g, "") + ` s="${sOverride}"`;
  if (type === "currency" || type === "number") {
    const num = parseFloat(String(value).replace(/[^0-9.\-]/g, ""));
    if (!isFinite(num)) return null;
    return `<c r="${addr}"${keep}><v>${num}</v></c>`;
  }
  if (type === "date") {
    const serial = dateToSerialAny(value);
    if (serial == null) return null;
    return `<c r="${addr}"${keep}><v>${serial}</v></c>`;
  }
  if (type === "formula") {
    // value is a formula string; "{R}" resolves to this cell's own row number.
    // No cached <v> is emitted, so Excel computes it (see calcPr fullCalcOnLoad).
    const R = (addr.match(/\d+/) || [""])[0];
    const f = String(value).replace(/\{R\}/g, R);
    return `<c r="${addr}"${keep}><f>${escapeXml(f)}</f></c>`;
  }
  // text -> inline string
  return `<c r="${addr}"${keep} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function patchCellInSheet(xml, addr, value, type, sOverride) {
  const re = new RegExp(`<c r="${addr}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`);
  const m = xml.match(re);
  if (m) {
    const cell = buildCell(addr, m[1], value, type, sOverride);
    if (cell == null) return xml;
    return xml.slice(0, m.index) + cell + xml.slice(m.index + m[0].length);
  }
  // cell not present -> insert into its row in column order
  const col = addr.match(/[A-Z]+/)[0];
  const row = addr.match(/\d+/)[0];
  const targetCol = colToNum(col);
  const cell = buildCell(addr, "", value, type, sOverride);
  if (cell == null) return xml;
  const rowRe = new RegExp(`(<row r="${row}"[^>]*>)([\\s\\S]*?)(</row>)`);
  const rm = xml.match(rowRe);
  if (!rm) return xml; // row missing; skip silently
  const inner = rm[2];
  const cells = [...inner.matchAll(/<c r="([A-Z]+)\d+"[\s\S]*?(?:\/>|<\/c>)/g)];
  let insertAt = inner.length;
  for (const cm of cells) {
    if (colToNum(cm[1]) > targetCol) { insertAt = cm.index; break; }
  }
  const newInner = inner.slice(0, insertAt) + cell + inner.slice(insertAt);
  return xml.slice(0, rm.index) + rm[1] + newInner + rm[3] + xml.slice(rm.index + rm[0].length);
}

function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
// XML-encode a sheet name the way it appears inside attribute values / formulas
function xmlName(s) { return s.replace(/&/g, "&amp;"); }
// Resolve an OPC relationship target (relative to a part's folder) to a package path
function resolvePath(baseDir, rel) {
  const stack = baseDir.replace(/\/$/, "").split("/");
  for (const seg of rel.split("/")) {
    if (seg === "..") stack.pop();
    else if (seg === "." || seg === "") continue;
    else stack.push(seg);
  }
  return stack.join("/");
}

// Find a sheet's { rId, path } by visible name.
function resolveSheet(wbXml, relsXml, name) {
  const enc = xmlName(name);
  const tag = (wbXml.match(new RegExp(`<sheet[^>]*name="${escapeRegex(enc)}"[^>]*/>`)) || [])[0];
  if (!tag) return null;
  const rid = (tag.match(/r:id="(rId\d+)"/) || [])[1];
  if (!rid) return null;
  const relMatch = relsXml.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]*)"`));
  if (!relMatch) return null;
  let target = relMatch[1].replace(/^\//, "");
  return { rid, tag, path: target.startsWith("xl/") ? target : "xl/" + target };
}

// ---------- public API ----------
// templateBuf:  Uint8Array of the .xlsx
// sheetName:    visible tab name to write values into
// values:       [{ addr, value, type }]  (type: "text" | "date" | "currency")
// deleteSheets: array of visible tab names to remove from the workbook
// opts.insert:  optional { sheetName, at, rows } to append data rows into
//               another tab (rows: [[{col,value,type}], ...]) pushing existing
//               rows down. Used to write the Crew Payroll fringe rows.
// opts.cellPatches: optional [{ sheetName, addr, value, type }] to patch single
//               cells on any sheet (e.g. 1st/Last Shoot Day on Crew Payroll).
export async function generateWorkbook(templateBuf, sheetName, values, deleteSheets = [], opts = {}) {
  const entries = parseZip(templateBuf);
  const enc = new TextEncoder();
  const overrides = {};      // name -> Uint8Array (stored verbatim)
  const dropped = new Set(); // part names removed from the package
  const setText = (name, str) => { overrides[name] = enc.encode(str); };
  const curText = async (name) =>
    overrides[name] ? new TextDecoder().decode(overrides[name]) : await readText(entries, name);

  let wbXml = await readText(entries, "xl/workbook.xml");
  let relsXml = await readText(entries, "xl/_rels/workbook.xml.rels");
  if (!wbXml || !relsXml) throw new Error("workbook.xml missing");

  // --- 1. patch the values into the chosen sheet ---
  const keep = resolveSheet(wbXml, relsXml, sheetName);
  if (!keep) throw new Error('Sheet not found: "' + sheetName + '"');
  const sheetPath = keep.path;
  let sheetXml = await readText(entries, sheetPath);
  if (sheetXml == null) throw new Error("sheet xml missing: " + sheetPath);
  // --- 0. open room in columns A–E for extra agencies / production companies ---
  // (applied before value patching so the shifted addresses line up). Only the
  // left form (cols A–E) moves; the calc table to the right stays put.
  if (opts.columnInserts && opts.columnInserts.length) {
    for (const ci of opts.columnInserts) sheetXml = insertPartialRows(sheetXml, ci.at, ci.count, ci.maxCol || 5);
  }
// Read the style index off an existing cell, so a caller can say "format this
// like that cell" instead of hardcoding a cellXfs index that a replacement
// template might not share.
function styleOfCell(xml, addr) {
  const m = xml.match(new RegExp(`<c r="${addr}"([^>]*?)(?:/>|>)`));
  if (!m) return null;
  const s = m[1].match(/\s+s="(\d+)"/);
  return s ? s[1] : null;
}

  let written = 0;
  for (const v of values) {
    if (v.value == null || String(v.value).trim() === "") continue;
    const before = sheetXml;
    // v.styleFrom: copy the number format from another cell on this sheet.
    const sOverride = v.styleFrom ? styleOfCell(sheetXml, v.styleFrom) : null;
    sheetXml = patchCellInSheet(sheetXml, v.addr, v.value, v.type, sOverride);
    if (sheetXml !== before) written++;
  }
  setText(sheetPath, sheetXml);

  // --- 1b. insert data rows into other sheets (Crew Payroll fringe, Talent &
  //         Extras). Supports opts.insert (single, legacy) + opts.inserts (array).
  //         Each job is an independent sheet, so their row shifts don't interact. ---
  const insertJobs = [opts.insert, ...(opts.inserts || [])].filter((j) => j && j.rows && j.rows.length);
  for (const job of insertJobs) {
    const ins = resolveSheet(wbXml, relsXml, job.sheetName);
    if (!ins) continue;
    let insXml = await curText(ins.path);
    if (insXml == null) continue;
    insXml = insertRowsIntoSheet(insXml, job.at || 10, job.rows);
    setText(ins.path, insXml);
    // stale calcChain would point at pre-shift cells; drop it to force recalc
    if (entries.some((e) => e.name === "xl/calcChain.xml")) dropped.add("xl/calcChain.xml");
  }

  // --- 1c. patch individual cells on arbitrary sheets (e.g. 1st/Last Shoot Day
  //         on Crew Payroll). Grouped by sheet; reads any override already made
  //         by the insert step above so both apply. ---
  if (opts.cellPatches && opts.cellPatches.length) {
    const bySheet = {};
    for (const cp of opts.cellPatches) (bySheet[cp.sheetName] = bySheet[cp.sheetName] || []).push(cp);
    for (const [sn, patches] of Object.entries(bySheet)) {
      const target = resolveSheet(wbXml, relsXml, sn);
      if (!target) continue;
      let px = await curText(target.path);
      if (px == null) continue;
      for (const p of patches) {
        if (p.clear) {
          // Blank the cell but keep its style (removes template placeholder values).
          px = px.replace(new RegExp(`<c r="${p.addr}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`), (full, attrs) => {
            const s = (attrs.match(/\s+s="\d+"/) || [""])[0];
            return `<c r="${p.addr}"${s}/>`;
          });
          continue;
        }
        if (p.value == null || String(p.value).trim() === "") continue;
        const sOverride = p.styleFrom ? styleOfCell(px, p.styleFrom) : null;
        px = patchCellInSheet(px, p.addr, p.value, p.type, sOverride);
        written++;
      }
      setText(target.path, px);
      if (entries.some((e) => e.name === "xl/calcChain.xml")) dropped.add("xl/calcChain.xml");
    }
  }
  let deleted = 0;
  let ctXml = await readText(entries, "[Content_Types].xml");
  for (const delName of deleteSheets) {
    const s = resolveSheet(wbXml, relsXml, delName);
    if (!s) continue;
    deleted++;
    wbXml = wbXml.replace(s.tag, "");
    relsXml = relsXml.replace(new RegExp(`<Relationship[^>]*Id="${s.rid}"[^>]*/>`), "");
    dropped.add(s.path);
    // its rels + the parts only it referenced (comments, vml, printer, threaded comments)
    const relName = s.path.replace(/(.*)\/([^/]+)$/, "$1/_rels/$2.rels");
    const sr = await readText(entries, relName);
    if (sr) {
      dropped.add(relName);
      const baseDir = s.path.replace(/\/[^/]+$/, "/"); // xl/worksheets/
      for (const m of sr.matchAll(/Target="([^"]*)"/g)) {
        dropped.add(resolvePath(baseDir, m[1]));
      }
    }
  }

  if (deleted) {
    // rebuild calc chain on open (avoids stale sheet-index references)
    if (entries.some((e) => e.name === "xl/calcChain.xml")) dropped.add("xl/calcChain.xml");
    // strip Content-Types Overrides for everything we removed
    for (const d of dropped) {
      ctXml = ctXml.replace(new RegExp(`<Override PartName="/${escapeRegex(d)}"[^>]*/>`), "");
    }
    setText("[Content_Types].xml", ctXml);
    // reset active/first sheet so we don't point past the new sheet count
    wbXml = wbXml.replace(/(<workbookView\b[^>]*?)\s+activeTab="\d+"/, "$1")
                 .replace(/(<workbookView\b[^>]*?)\s+firstSheet="\d+"/, "$1");
    setText("xl/workbook.xml", wbXml);
    setText("xl/_rels/workbook.xml.rels", relsXml);
    // rewrite any formula in a surviving sheet that referenced a deleted tab -> #REF!
    for (const e of entries) {
      if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(e.name) || dropped.has(e.name) || e.name === sheetPath) continue;
      let xml = await curText(e.name);
      let changed = false;
      for (const delName of deleteSheets) {
        const re = new RegExp(`'${escapeRegex(xmlName(delName))}'!(\\$?[A-Z]+\\$?\\d+(?::\\$?[A-Z]+\\$?\\d+)?)`, "g");
        if (re.test(xml)) { xml = xml.replace(re, "#REF!"); changed = true; }
      }
      if (changed) setText(e.name, xml);
    }
  }

  // --- 2b. normalize sheet selection (prevent Excel "group" mode) ---
  // Excel treats the active sheet PLUS every sheet with tabSelected="1" as a
  // selected group, which greys out much of the Data tab (sort/filter/etc.).
  // Deleting sheets stripped activeTab above, leaving a tab-selected sheet that
  // no longer matches the (defaulted) active sheet — two selected sheets = a
  // group. Fix: make exactly one sheet (the one we wrote into) both the active
  // tab and the only tab-selected sheet.
  {
    const order = [...wbXml.matchAll(/<sheet\b[^>]*\sname="([^"]*)"[^>]*\sr:id="(rId\d+)"[^>]*\/>/g)]
      .map((m) => ({ name: m[1].replace(/&amp;/g, "&"), rid: m[2] }));
    const ridPath = {};
    for (const m of relsXml.matchAll(/<Relationship\b[^>]*Id="(rId\d+)"[^>]*Target="([^"]*)"[^>]*\/>/g)) {
      const t = m[2].replace(/^\//, "");
      ridPath[m[1]] = t.startsWith("xl/") ? t : "xl/" + t;
    }
    let activeIdx = order.findIndex((s) => s.name === sheetName);
    if (activeIdx < 0) activeIdx = 0;

    // set activeTab to the target sheet's index (replace existing or insert)
    if (/<workbookView\b[^>]*\sactiveTab="\d+"/.test(wbXml)) {
      wbXml = wbXml.replace(/(<workbookView\b[^>]*\sactiveTab=")\d+(")/, `$1${activeIdx}$2`);
    } else {
      wbXml = wbXml.replace(/<workbookView\b/, `<workbookView activeTab="${activeIdx}"`);
    }
    setText("xl/workbook.xml", wbXml);

    // strip tabSelected from every surviving sheet; add it only to the active one
    for (let i = 0; i < order.length; i++) {
      const pth = ridPath[order[i].rid];
      if (!pth) continue;
      let sx = await curText(pth);
      if (sx == null) continue;
      let changed = false;
      if (/\stabSelected="[^"]*"/.test(sx)) { sx = sx.replace(/\s*tabSelected="[^"]*"/g, ""); changed = true; }
      if (i === activeIdx && /<sheetView\s/.test(sx)) { sx = sx.replace(/<sheetView\s/, '<sheetView tabSelected="1" '); changed = true; }
      if (changed) setText(pth, sx);
    }
  }

  // --- 2c. force a full recalc on open ---
  // Patched formula cells keep the template's stale cached <v> (e.g. an AICP
  // description from before its code cell was rewritten). Excel shows that
  // cached text until something triggers a recalc, so ask for one up front.
  if (/<calcPr\b/.test(wbXml)) {
    wbXml = /fullCalcOnLoad=/.test(wbXml)
      ? wbXml.replace(/(<calcPr\b[^>]*\sfullCalcOnLoad=")[^"]*(")/, "$11$2")
      : wbXml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
  } else {
    wbXml = wbXml.replace(/<\/workbook>/, '<calcPr fullCalcOnLoad="1"/></workbook>');
  }
  setText("xl/workbook.xml", wbXml);

  // --- 3. rebuild the package: drop removed parts, swap overrides (stored), copy the rest ---
  const parts = entries.filter((e) => !dropped.has(e.name)).map((e) => {
    if (overrides[e.name]) {
      const b = overrides[e.name];
      return { name: e.name, method: 0, crc: crc32(b), compSize: b.length, uncompSize: b.length, data: b };
    }
    return { name: e.name, method: e.method, crc: e.crc, compSize: e.compSize, uncompSize: e.uncompSize, data: e.comp };
  });
  return { blob: buildZip(parts), written, sheetPath, deleted };
}
