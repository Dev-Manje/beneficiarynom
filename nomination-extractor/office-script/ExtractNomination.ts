/**
 * Office Script: read a Beneficiary Nomination Form (fillable PDF, BNF-*) and return its data as JSON.
 *
 * Called from Power Automate with "Excel Online (Business) > Run script":
 *   pdfBase64 = body('Get_file_content')?['$content']
 * Returns a JSON string (same shape as power-automate/parse-json-schema.json):
 *   { valid, errors[], warnings[], form_version, member{}, nominees[], payees[], totals{} }
 *
 * Office Scripts can't load libraries, so this includes a small PDF reader (objects, object
 * streams, incremental updates, Flate decompression) - just enough to read AcroForm fields.
 */
function main(workbook: ExcelScript.Workbook, pdfBase64: string): string {
  return JSON.stringify(extractNomination(pdfBase64));
}

// ======================================================================================
// Validation (mirrors function_app/nomination_extract.py)
// ======================================================================================

const FORM_VERSION_PREFIX = "BNF-";
const NOMINEE_ROWS = 15;
const PAYEE_ROWS = 3;
const MEMBER_FIELDS = ["fund_name", "member_name", "member_number", "branch_code", "cost_centre",
  "home_telephone", "cell_telephone", "work_telephone", "email_address", "declaration_full_names"];

interface Member {
  [key: string]: string | boolean | null;
}
interface Nominee {
  row: number; full_name: string; relationship: string; identity_number: string;
  date_of_birth: string | null; date_of_birth_text: string;
  benefit_current_pct: number | null; benefit_revised_pct: number | null;
}
interface Payee {
  row: number; full_name: string; relationship: string; identity_number: string;
  benefit_pct: number | null; monthly_maintenance: number | null;
}
interface Totals {
  nominee_count: number; payee_count: number; benefit_current_pct: number;
  benefit_revised_pct: number; effective_column: string;
}
interface Result {
  valid: boolean; errors: string[]; warnings: string[]; form_version: string | null;
  member: Member | null; nominees: Nominee[]; payees: Payee[]; totals: Totals | null;
}

function parseDate(text: string): string | null {
  const m = /^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$/.exec(text || "");
  if (!m) return null;
  const d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]);
  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (y < 1 || dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
  return `${String(y).padStart(4, "0")}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

/** '1,234.50' -> 1234.5; '' -> null; throws if not a number. */
function parseNumber(text: string): number | null {
  const t = (text || "").replace(/,/g, "").replace(/%/g, "").trim();
  if (!t) return null;
  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) throw new Error("not a number");
  return Number(t);
}

function fmt(n: number): string {
  return String(Math.round(n * 1e6) / 1e6);
}

function extractNomination(pdfBase64: string): Result {
  const errors: string[] = [];
  const warnings: string[] = [];
  const result: Result = { valid: false, errors, warnings, form_version: null,
    member: null, nominees: [], payees: [], totals: null };

  let pdf: PdfDoc;
  let values: { [name: string]: string };
  try {
    pdf = openPdf(base64ToBytes(pdfBase64 || ""));
    values = textFieldValues(pdf);
  } catch (e) {
    errors.push(`Could not read the PDF: ${(e as Error).message}`);
    return result;
  }
  const text = (name: string): string => (values[name] || "").trim();

  const version = text("form_version");
  result.form_version = version || null;
  if (!version.startsWith(FORM_VERSION_PREFIX)) {
    errors.push("This is not the fillable Beneficiary Nomination Form, or its form fields were removed " +
      "(for example by printing to PDF or scanning).");
    return result;
  }

  // --- member ---------------------------------------------------------------------
  const member: Member = {};
  for (const k of MEMBER_FIELDS) member[k] = text(k);
  const dobText = text("member_date_of_birth");
  const sigDateText = text("signature_date");
  member.date_of_birth = parseDate(dobText);
  member.date_of_birth_text = dobText;
  member.signature_date = parseDate(sigDateText);
  member.signature_date_text = sigDateText;
  member.signature_attached = text("signature_attached").toLowerCase() === "yes";
  member.signature_attached_at = text("signature_attached_at");
  member.signature_image_found = fieldHasImage(pdf, "member_signature");
  result.member = member;

  if (!member.signature_attached) {
    errors.push("The form has not been signed. Open it in Adobe Acrobat Reader and click the signature box.");
  } else if (!member.signature_image_found) {
    warnings.push("Signature is marked as attached, but no signature image was found on the form.");
  }
  if (!member.member_name) errors.push("Member name is empty.");
  if (!member.member_number) errors.push("Member number is empty.");
  if (!member.declaration_full_names) warnings.push("Full names in the declaration are empty.");
  if (dobText && !member.date_of_birth) warnings.push(`Member date of birth '${dobText}' is not a valid dd/mm/yyyy date.`);
  if (sigDateText && !member.signature_date) warnings.push(`Signature date '${sigDateText}' is not a valid dd/mm/yyyy date.`);

  // --- nominees ---------------------------------------------------------------------
  for (let i = 1; i <= NOMINEE_ROWS; i++) {
    const p = `nominee_${String(i).padStart(2, "0")}_`;
    const row = {
      full_name: text(p + "full_name"), date_of_birth: text(p + "date_of_birth"),
      relationship: text(p + "relationship"), identity_number: text(p + "identity_number"),
      benefit_current_pct: text(p + "benefit_current_pct"), benefit_revised_pct: text(p + "benefit_revised_pct"),
    };
    if (!Object.keys(row).some(k => (row as { [k: string]: string })[k])) continue;
    if (!row.full_name) {
      warnings.push(`Nominee row ${i} has details but no name; it was skipped.`);
      continue;
    }
    const n: Nominee = { row: i, full_name: row.full_name, relationship: row.relationship,
      identity_number: row.identity_number, date_of_birth: parseDate(row.date_of_birth),
      date_of_birth_text: row.date_of_birth, benefit_current_pct: null, benefit_revised_pct: null };
    for (const key of ["benefit_current_pct", "benefit_revised_pct"]) {
      const raw = (row as { [k: string]: string })[key];
      let v: number | null = null;
      try { v = parseNumber(raw); } catch (e) { errors.push(`Nominee row ${i}: '${raw}' is not a valid percentage.`); }
      if (key === "benefit_current_pct") n.benefit_current_pct = v; else n.benefit_revised_pct = v;
    }
    if (row.date_of_birth && !n.date_of_birth) {
      warnings.push(`Nominee row ${i}: date of birth '${row.date_of_birth}' is not a valid dd/mm/yyyy date.`);
    }
    result.nominees.push(n);
  }

  // --- court order payees -----------------------------------------------------------
  for (let i = 1; i <= PAYEE_ROWS; i++) {
    const p = `payee_${String(i).padStart(2, "0")}_`;
    const row = {
      full_name: text(p + "full_name"), relationship: text(p + "relationship"),
      identity_number: text(p + "identity_number"), benefit_pct: text(p + "benefit_pct"),
      monthly_maintenance: text(p + "monthly_maintenance"),
    };
    if (!Object.keys(row).some(k => (row as { [k: string]: string })[k])) continue;
    if (!row.full_name) {
      warnings.push(`Court order row ${i} has details but no name; it was skipped.`);
      continue;
    }
    const pay: Payee = { row: i, full_name: row.full_name, relationship: row.relationship,
      identity_number: row.identity_number, benefit_pct: null, monthly_maintenance: null };
    for (const key of ["benefit_pct", "monthly_maintenance"]) {
      const raw = (row as { [k: string]: string })[key];
      let v: number | null = null;
      try { v = parseNumber(raw); } catch (e) { errors.push(`Court order row ${i}: '${raw}' is not a valid number.`); }
      if (key === "benefit_pct") pay.benefit_pct = v; else pay.monthly_maintenance = v;
    }
    result.payees.push(pay);
  }

  // --- totals -------------------------------------------------------------------------
  const noms = result.nominees;
  const round2 = (x: number): number => Math.round(x * 100) / 100;
  const totCur = round2(noms.reduce((a, n) => a + (n.benefit_current_pct || 0), 0));
  const totRev = round2(noms.reduce((a, n) => a + (n.benefit_revised_pct || 0), 0));
  const usesRevised = noms.some(n => n.benefit_revised_pct !== null);
  result.totals = { nominee_count: noms.length, payee_count: result.payees.length,
    benefit_current_pct: totCur, benefit_revised_pct: totRev,
    effective_column: usesRevised ? "revised" : "current" };
  if (!noms.length) {
    errors.push("No dependants / nominees are listed.");
  } else {
    const effective = usesRevised ? totRev : totCur;
    if (Math.abs(effective - 100) > 0.01) {
      errors.push(`% Benefit (${usesRevised ? "Revised" : "Current"}) adds up to ${fmt(effective)}%, it must equal 100%.`);
    }
  }

  result.valid = errors.length === 0;
  return result;
}

// ======================================================================================
// Minimal PDF reader (plain objects and functions only: Office Scripts reject classes using `this`)
// ======================================================================================

interface PdfName { k: "name"; name: string; }
interface PdfRef { k: "ref"; num: number; gen: number; }
interface PdfStr { k: "str"; bytes: number[]; }
interface PdfArr { k: "arr"; items: PdfValue[]; }
interface PdfDict { k: "dict"; entries: { [key: string]: PdfValue }; }
type PdfValue = PdfName | PdfRef | PdfStr | PdfArr | PdfDict | number | boolean | null;

interface PdfObject { pos: number; value: PdfValue; stream: string | null; }
interface PdfDoc { text: string; objects: { [num: number]: PdfObject }; }
interface Parser { s: string; pos: number; }

function isName(v: PdfValue): v is PdfName { return v !== null && typeof v === "object" && v.k === "name"; }
function isRef(v: PdfValue): v is PdfRef { return v !== null && typeof v === "object" && v.k === "ref"; }
function isStr(v: PdfValue): v is PdfStr { return v !== null && typeof v === "object" && v.k === "str"; }
function isArr(v: PdfValue): v is PdfArr { return v !== null && typeof v === "object" && v.k === "arr"; }
function isDict(v: PdfValue): v is PdfDict { return v !== null && typeof v === "object" && v.k === "dict"; }
function dget(d: PdfDict, key: string): PdfValue { return key in d.entries ? d.entries[key] : null; }
function nameOf(v: PdfValue): string { return isName(v) ? v.name : ""; }

function openPdf(bytes: Uint8Array): PdfDoc {
  const doc: PdfDoc = { text: bytesToLatin1(bytes), objects: {} };
  if (!/^\s*%PDF-/.test(doc.text.slice(0, 1024))) throw new Error("not a PDF file");
  scanObjects(doc);
  loadObjectStreams(doc);
  if (isEncrypted(doc)) throw new Error("the PDF is password-protected or encrypted");
  return doc;
}

/** Find every "n g obj ... endobj" in file order; later definitions (incremental updates) win. */
function scanObjects(doc: PdfDoc): void {
  const text = doc.text;
  const re = /(\d+)\s+(\d+)\s+obj\b/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const num = Number(m[1]);
    try {
      const p: Parser = { s: text, pos: m.index + m[0].length };
      const value = parseValue(p);
      let stream: string | null = null;
      skipWs(p);
      if (text.startsWith("stream", p.pos)) {
        let start = p.pos + 6;
        if (text[start] === "\r") start++;
        if (text[start] === "\n") start++;
        const len = isDict(value) ? dget(value, "Length") : null;
        let end = typeof len === "number" ? start + len : -1;
        if (end < 0 || !/^\s*endstream/.test(text.slice(end, end + 20))) {
          end = text.indexOf("endstream", start);
          if (end < 0) throw new Error("unterminated stream");
          if (text[end - 1] === "\n") end--; // trim the EOL before "endstream"
          if (text[end - 1] === "\r") end--;
        }
        stream = text.slice(start, end);
        const es = text.indexOf("endstream", end);
        re.lastIndex = es >= 0 ? es + 9 : text.length;
      } else {
        re.lastIndex = Math.max(re.lastIndex, p.pos);
      }
      const prev = doc.objects[num];
      if (!prev || prev.pos <= m.index) doc.objects[num] = { pos: m.index, value: value, stream: stream };
    } catch (e) {
      re.lastIndex = m.index + 1; // false match (e.g. inside binary data): keep scanning
    }
  }
}

/** Unpack compressed object streams (/Type /ObjStm). */
function loadObjectStreams(doc: PdfDoc): void {
  const containers = Object.keys(doc.objects).map(Number).filter((n: number) => {
    const o = doc.objects[n];
    return o.stream !== null && isDict(o.value) && nameOf(dget(o.value, "Type")) === "ObjStm";
  });
  for (const cnum of containers) {
    const c = doc.objects[cnum];
    const dict = c.value as PdfDict;
    let data: string;
    try { data = decodeStream(doc, dict, c.stream as string); } catch (e) { continue; }
    const n = resolve(doc, dget(dict, "N"), 0);
    const first = resolve(doc, dget(dict, "First"), 0);
    if (typeof n !== "number" || typeof first !== "number") continue;
    const header: Parser = { s: data, pos: 0 };
    for (let i = 0; i < n; i++) {
      const num = parseValue(header), off = parseValue(header);
      if (typeof num !== "number" || typeof off !== "number") break;
      const prev = doc.objects[num];
      if (prev && prev.pos > c.pos) continue; // a later plain definition overrides
      try {
        const value = parseValue({ s: data, pos: first + off });
        doc.objects[num] = { pos: c.pos, value: value, stream: null };
      } catch (e) { /* skip unreadable object */ }
    }
  }
}

function decodeStream(doc: PdfDoc, dict: PdfDict, raw: string): string {
  const f = resolve(doc, dget(dict, "Filter"), 0);
  const filters: string[] = isArr(f) ? f.items.map((x: PdfValue) => nameOf(resolve(doc, x, 0))) : f ? [nameOf(f)] : [];
  let data = raw;
  for (const name of filters) {
    if (name === "FlateDecode") data = bytesToLatin1(inflate(latin1ToBytes(data)));
    else throw new Error(`unsupported filter ${name}`);
  }
  return data;
}

function isEncrypted(doc: PdfDoc): boolean {
  const re = /trailer\s*<</g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(doc.text)) !== null) {
    try {
      const d = parseValue({ s: doc.text, pos: m.index + 7 });
      if (isDict(d) && dget(d, "Encrypt") !== null) return true;
    } catch (e) { /* ignore */ }
  }
  for (const k of Object.keys(doc.objects)) {
    const v = doc.objects[Number(k)].value;
    if (isDict(v) && nameOf(dget(v, "Type")) === "XRef" && dget(v, "Encrypt") !== null) return true;
  }
  return false;
}

function resolve(doc: PdfDoc, v: PdfValue, depth: number): PdfValue {
  if (isRef(v) && depth < 32) {
    const o = doc.objects[v.num];
    return o ? resolve(doc, o.value, depth + 1) : null;
  }
  return v;
}

/** All terminal form fields: full name -> dictionary. */
function formFields(doc: PdfDoc): { [name: string]: PdfDict } {
  const out: { [name: string]: PdfDict } = {};
  const pos: { [name: string]: number } = {};
  for (const k of Object.keys(doc.objects)) {
    const o = doc.objects[Number(k)];
    const d = o.value;
    if (!isDict(d) || !isStr(resolve(doc, dget(d, "T"), 0))) continue;
    const name = fullName(doc, d);
    if (!(name in out) || pos[name] <= o.pos) { out[name] = d; pos[name] = o.pos; }
  }
  return out;
}

function fullName(doc: PdfDoc, d: PdfDict): string {
  const parts: string[] = [];
  let cur: PdfValue = d;
  for (let i = 0; i < 16 && isDict(cur); i++) {
    const t = resolve(doc, dget(cur, "T"), 0);
    if (isStr(t)) parts.unshift(decodePdfString(t.bytes));
    cur = resolve(doc, dget(cur, "Parent"), 0);
  }
  return parts.join(".");
}

function textFieldValues(doc: PdfDoc): { [name: string]: string } {
  const out: { [name: string]: string } = {};
  const fields = formFields(doc);
  for (const name of Object.keys(fields)) {
    const v = resolve(doc, dget(fields[name], "V"), 0);
    if (isStr(v)) out[name] = decodePdfString(v.bytes);
    else if (isName(v)) out[name] = v.name;
    else if (typeof v === "number") out[name] = String(v);
    else if (isArr(v)) {
      out[name] = v.items.map((x: PdfValue) => {
        const r = resolve(doc, x, 0);
        return isStr(r) ? decodePdfString(r.bytes) : "";
      }).join(", ");
    }
  }
  return out;
}

/** True if the field's widget appearance (or a form XObject inside it) draws an image. */
function fieldHasImage(doc: PdfDoc, fieldName: string): boolean {
  const f = formFields(doc)[fieldName];
  if (!f) return false;
  const widgets: PdfDict[] = [f];
  const kids = resolve(doc, dget(f, "Kids"), 0);
  if (isArr(kids)) {
    for (const kid of kids.items) {
      const w = resolve(doc, kid, 0);
      if (isDict(w)) widgets.push(w);
    }
  }
  for (const w of widgets) {
    const ap = resolve(doc, dget(w, "AP"), 0);
    if (!isDict(ap)) continue;
    const n = dget(ap, "N");
    const nd = resolve(doc, n, 0);
    const so = isRef(n) ? doc.objects[n.num] : undefined;
    if (isDict(nd) && !(so && so.stream !== null)) {
      // appearance sub-dictionary (states): check every state
      for (const s of Object.keys(nd.entries)) if (xobjectHasImage(doc, nd.entries[s], 0)) return true;
    } else if (xobjectHasImage(doc, n, 0)) return true;
  }
  return false;
}

function xobjectHasImage(doc: PdfDoc, ref: PdfValue, depth: number): boolean {
  if (depth > 4) return false;
  const x = resolve(doc, ref, 0);
  if (!isDict(x)) return false;
  if (nameOf(dget(x, "Subtype")) === "Image") return true;
  const res = resolve(doc, dget(x, "Resources"), 0);
  const xobjs = isDict(res) ? resolve(doc, dget(res, "XObject"), 0) : null;
  if (!isDict(xobjs)) return false;
  for (const k of Object.keys(xobjs.entries)) if (xobjectHasImage(doc, xobjs.entries[k], depth + 1)) return true;
  return false;
}

// ---- PDF object syntax parser --------------------------------------------------------

function skipWs(p: Parser): void {
  const s = p.s;
  while (p.pos < s.length) {
    const c = s.charCodeAt(p.pos);
    if (c === 0x25) { // % comment
      while (p.pos < s.length && s[p.pos] !== "\n" && s[p.pos] !== "\r") p.pos++;
    } else if (c === 0x20 || c === 0x0a || c === 0x0d || c === 0x09 || c === 0x0c || c === 0x00) {
      p.pos++;
    } else break;
  }
}

function isDelim(ch: string): boolean {
  return "()<>[]{}/%".indexOf(ch) >= 0 || /\s/.test(ch) || ch === "\0";
}

function regularToken(p: Parser): string {
  const start = p.pos;
  while (p.pos < p.s.length && !isDelim(p.s[p.pos])) p.pos++;
  return p.s.slice(start, p.pos);
}

function parseValue(p: Parser): PdfValue {
  skipWs(p);
  const s = p.s;
  if (p.pos >= s.length) throw new Error("unexpected end of data");
  const ch = s[p.pos];
  if (ch === "/") {
    p.pos++;
    const raw = regularToken(p);
    const name = raw.replace(/#([0-9A-Fa-f]{2})/g, (_m: string, h: string) => String.fromCharCode(parseInt(h, 16)));
    return { k: "name", name: name };
  }
  if (ch === "<" && s[p.pos + 1] === "<") {
    p.pos += 2;
    const d: PdfDict = { k: "dict", entries: {} };
    for (; ;) {
      skipWs(p);
      if (s.startsWith(">>", p.pos)) { p.pos += 2; return d; }
      const key = parseValue(p);
      if (!isName(key)) throw new Error("bad dictionary key");
      d.entries[key.name] = parseValue(p);
    }
  }
  if (ch === "<") {
    const end = s.indexOf(">", p.pos);
    if (end < 0) throw new Error("bad hex string");
    let hex = s.slice(p.pos + 1, end).replace(/[^0-9A-Fa-f]/g, "");
    if (hex.length % 2) hex += "0";
    const bytes: number[] = [];
    for (let i = 0; i < hex.length; i += 2) bytes.push(parseInt(hex.substr(i, 2), 16));
    p.pos = end + 1;
    return { k: "str", bytes: bytes };
  }
  if (ch === "(") return literalString(p);
  if (ch === "[") {
    p.pos++;
    const items: PdfValue[] = [];
    for (; ;) {
      skipWs(p);
      if (s[p.pos] === "]") { p.pos++; return { k: "arr", items: items }; }
      items.push(parseValue(p));
    }
  }
  const tok = regularToken(p);
  if (tok === "true") return true;
  if (tok === "false") return false;
  if (tok === "null") return null;
  if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(tok)) {
    // "num gen R" is a reference
    if (/^\d+$/.test(tok)) {
      const m = /^\s+(\d+)\s+R(?=[\s/<>\[\]()%]|$)/.exec(s.slice(p.pos, p.pos + 24));
      if (m) { p.pos += m[0].length; return { k: "ref", num: Number(tok), gen: Number(m[1]) }; }
    }
    return Number(tok);
  }
  throw new Error(`unexpected token '${tok || ch}'`);
}

const STRING_ESCAPES: { [k: string]: number } = { n: 10, r: 13, t: 9, b: 8, f: 12, "(": 40, ")": 41, "\\": 92 };

function literalString(p: Parser): PdfStr {
  const s = p.s;
  p.pos++; // (
  const out: number[] = [];
  let depth = 1;
  while (p.pos < s.length) {
    const c = s[p.pos++];
    if (c === "\\") {
      const e = s[p.pos++];
      if (e in STRING_ESCAPES) out.push(STRING_ESCAPES[e]);
      else if (e >= "0" && e <= "7") {
        let oct = e;
        while (oct.length < 3 && s[p.pos] >= "0" && s[p.pos] <= "7") oct += s[p.pos++];
        out.push(parseInt(oct, 8) & 0xff);
      } else if (e === "\r") { if (s[p.pos] === "\n") p.pos++; }
      else if (e === "\n") { /* line continuation */ }
      else out.push(e.charCodeAt(0));
    } else if (c === "(") { depth++; out.push(40); }
    else if (c === ")") { if (--depth === 0) return { k: "str", bytes: out }; out.push(41); }
    else out.push(c.charCodeAt(0));
  }
  throw new Error("unterminated string");
}

// ======================================================================================
// Text decoding
// ======================================================================================

// PDFDocEncoding differs from Latin-1 in 0x18-0x1F and 0x80-0xA0.
const PDFDOC_HIGH: { [code: number]: number } = {
  0x80: 0x2022, 0x81: 0x2020, 0x82: 0x2021, 0x83: 0x2026, 0x84: 0x2014, 0x85: 0x2013, 0x86: 0x0192,
  0x87: 0x2044, 0x88: 0x2039, 0x89: 0x203a, 0x8a: 0x2212, 0x8b: 0x2030, 0x8c: 0x201e, 0x8d: 0x201c,
  0x8e: 0x201d, 0x8f: 0x2018, 0x90: 0x2019, 0x91: 0x201a, 0x92: 0x2122, 0x93: 0xfb01, 0x94: 0xfb02,
  0x95: 0x0141, 0x96: 0x0152, 0x97: 0x0160, 0x98: 0x0178, 0x99: 0x017d, 0x9a: 0x0131, 0x9b: 0x0142,
  0x9c: 0x0153, 0x9d: 0x0161, 0x9e: 0x017e, 0xa0: 0x20ac,
};

function decodePdfString(b: number[]): string {
  if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) { // UTF-16BE
    let s = "";
    for (let i = 2; i + 1 < b.length; i += 2) s += String.fromCharCode((b[i] << 8) | b[i + 1]);
    return s;
  }
  if (b.length >= 3 && b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return utf8Decode(b.slice(3));
  let s = "";
  for (const c of b) s += String.fromCharCode(PDFDOC_HIGH[c] || c);
  return s;
}

function utf8Decode(b: number[]): string {
  let s = "";
  for (let i = 0; i < b.length;) {
    const c = b[i++];
    let cp = c;
    if (c >= 0xf0) { cp = ((c & 7) << 18) | ((b[i++] & 63) << 12) | ((b[i++] & 63) << 6) | (b[i++] & 63); }
    else if (c >= 0xe0) { cp = ((c & 15) << 12) | ((b[i++] & 63) << 6) | (b[i++] & 63); }
    else if (c >= 0xc0) { cp = ((c & 31) << 6) | (b[i++] & 63); }
    s += String.fromCodePoint(cp);
  }
  return s;
}

function filled(n: number, v: number): number[] {
  const a: number[] = [];
  for (let i = 0; i < n; i++) a.push(v);
  return a;
}

function base64ToBytes(b64: string): Uint8Array {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const lookup = filled(256, -1);
  for (let i = 0; i < 64; i++) lookup[alphabet.charCodeAt(i)] = i;
  lookup["-".charCodeAt(0)] = 62; lookup["_".charCodeAt(0)] = 63; // base64url
  const clean = b64.replace(/^data:[^,]*,/, "");
  const out = new Uint8Array(Math.floor(clean.length * 3 / 4) + 3);
  let buf = 0, bits = 0, n = 0;
  for (let i = 0; i < clean.length; i++) {
    const v = lookup[clean.charCodeAt(i)];
    if (v < 0) continue; // skips '=', whitespace, line breaks
    buf = (buf << 6) | v; bits += 6;
    if (bits >= 8) { bits -= 8; out[n++] = (buf >> bits) & 0xff; }
  }
  return out.subarray(0, n);
}

function bytesToLatin1(b: Uint8Array): string {
  const parts: string[] = [];
  for (let i = 0; i < b.length; i += 8192) {
    parts.push(String.fromCharCode(...Array.from(b.subarray(i, i + 8192))));
  }
  return parts.join("");
}

function latin1ToBytes(s: string): Uint8Array {
  const b = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i) & 0xff;
  return b;
}

// ======================================================================================
// Inflate (RFC 1950/1951), after Mark Adler's "puff"
// ======================================================================================

interface Huffman { counts: number[]; symbols: number[]; }

const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537,
  2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

function inflate(data: Uint8Array): Uint8Array {
  let pos = 0, bitBuf = 0, bitCnt = 0;
  // zlib wrapper (2-byte header) is optional
  if (data.length > 2 && (data[0] & 0x0f) === 8 && ((data[0] << 8) | data[1]) % 31 === 0) pos = 2;
  const out: number[] = [];

  const bits = (n: number): number => {
    while (bitCnt < n) {
      if (pos >= data.length) throw new Error("compressed data ended early");
      bitBuf |= data[pos++] << bitCnt;
      bitCnt += 8;
    }
    const v = bitBuf & ((1 << n) - 1);
    bitBuf >>>= n; bitCnt -= n;
    return v;
  };
  const build = (lengths: number[]): Huffman => {
    const counts = filled(16, 0);
    for (const l of lengths) counts[l]++;
    counts[0] = 0;
    const offs = filled(16, 0);
    for (let l = 1; l < 16; l++) offs[l] = offs[l - 1] + counts[l - 1];
    const symbols = filled(lengths.length, 0);
    lengths.forEach((l, sym) => { if (l) symbols[offs[l]++] = sym; });
    return { counts, symbols };
  };
  const decode = (h: Huffman): number => {
    let code = 0, first = 0, index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = h.counts[len];
      if (code - count < first) return h.symbols[index + (code - first)];
      index += count; first += count; first <<= 1; code <<= 1;
    }
    throw new Error("bad compressed data");
  };
  const codes = (lit: Huffman, dist: Huffman): void => {
    for (; ;) {
      const sym = decode(lit);
      if (sym < 256) out.push(sym);
      else if (sym === 256) return;
      else {
        const li = sym - 257;
        if (li >= 29) throw new Error("bad length code");
        const len = LEN_BASE[li] + bits(LEN_EXTRA[li]);
        const ds = decode(dist);
        if (ds >= 30) throw new Error("bad distance code");
        const d = DIST_BASE[ds] + bits(DIST_EXTRA[ds]);
        if (d > out.length) throw new Error("distance too far back");
        const start = out.length - d;
        for (let i = 0; i < len; i++) out.push(out[start + i]);
      }
    }
  };

  let fixedLit: Huffman | null = null, fixedDist: Huffman | null = null;
  let last = 0;
  do {
    last = bits(1);
    const type = bits(2);
    if (type === 0) { // stored
      bitBuf = 0; bitCnt = 0;
      if (pos + 4 > data.length) throw new Error("compressed data ended early");
      const len = data[pos] | (data[pos + 1] << 8);
      pos += 4;
      if (pos + len > data.length) throw new Error("compressed data ended early");
      for (let i = 0; i < len; i++) out.push(data[pos++]);
    } else if (type === 1) { // fixed Huffman
      if (!fixedLit) {
        const l: number[] = [];
        for (let i = 0; i < 288; i++) l.push(i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8);
        fixedLit = build(l);
        fixedDist = build(filled(30, 5));
      }
      codes(fixedLit, fixedDist as Huffman);
    } else if (type === 2) { // dynamic Huffman
      const nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
      const clLens = filled(19, 0);
      for (let i = 0; i < ncode; i++) clLens[CL_ORDER[i]] = bits(3);
      const cl = build(clLens);
      const lens: number[] = [];
      while (lens.length < nlen + ndist) {
        const sym = decode(cl);
        if (sym < 16) lens.push(sym);
        else {
          let rep = 0, val = 0;
          if (sym === 16) {
            if (!lens.length) throw new Error("bad compressed data");
            val = lens[lens.length - 1]; rep = 3 + bits(2);
          } else if (sym === 17) rep = 3 + bits(3);
          else rep = 11 + bits(7);
          for (let i = 0; i < rep; i++) lens.push(val);
        }
      }
      codes(build(lens.slice(0, nlen)), build(lens.slice(nlen, nlen + ndist)));
    } else throw new Error("bad block type");
  } while (!last);
  return Uint8Array.from(out);
}
