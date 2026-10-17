// CSV (RFC 4180): rows of fields, a field with a comma, a quote or a line
// break in quotes, a quote in it doubled. Written with \n between rows;
// read with \n or \r\n, a last line break or none, and fields split by a
// comma or another delimiter (TSV's tab, ;, |: detected, as Excel's import
// guesses), a byte-order mark left out; bytes decoded from UTF-8, UTF-16
// or Windows-1252 first.

// A field quoted when it holds the delimiter, a quote or a line break.
const quoted = (field: string, delimiter: string): string => (field.includes(delimiter) || /["\r\n]/.test(field) ? `"${field.replace(/"/g, '""')}"` : field);

export const toCsv = (rows: readonly (readonly string[])[], delimiter = ","): string => rows.map((row) => row.map((f) => quoted(f, delimiter)).join(delimiter)).join("\n");

// The delimiters a CSV file may be split by, as Excel's import guesses:
// comma, tab (TSV), semicolon (where a comma is the decimal point), bar.
export const DELIMITERS: Readonly<Record<string, string>> = { comma: ",", tab: "\t", semicolon: ";", bar: "|", space: " " };

// The delimiter a text is most likely split by: of comma, tab, semicolon
// and bar, the one giving its first lines the same number of fields, the
// most of them (outside quotes); a comma when none splits.
export function detectDelimiter(text: string): string {
  const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter((l) => l.trim() !== "").slice(0, 20);
  let best = ",", score = 0;
  for (const d of [",", "\t", ";", "|"]) {
    const counts = lines.map((l) => {
      let n = 0, q = false;
      for (const c of l) {
        if (c === '"') q = !q;
        else if (c === d && !q) n++;
      }
      return n;
    });
    if (!counts.length || counts[0] === 0) continue;
    const same = counts.filter((n) => n === counts[0]).length / counts.length;
    const s = same * 100 + counts[0];
    if (s > score) { best = d; score = s; }
  }
  return best;
}

// Bytes as text in an encoding: UTF-8 (the default; with or without its
// byte-order mark), UTF-16 little- or big-endian (a byte-order mark tells
// which, under auto), Windows-1252 and Latin-1. Auto: a byte-order mark,
// else UTF-8 if the bytes are, else Windows-1252.
export const ENCODINGS = ["auto", "utf-8", "utf-16le", "utf-16be", "windows-1252", "latin1"];
const CP1252: Record<number, number> = { 0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160, 0x8b: 0x2039, 0x8c: 0x0152, 0x8e: 0x017d,
  0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a, 0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178 };

function utf8(b: ArrayLike<number>, from: number, strict: boolean): string | null {
  let out = "";
  for (let i = from; i < b.length;) {
    const c = b[i];
    let cp: number, n: number;
    if (c < 0x80) { cp = c; n = 1; }
    else if (c >= 0xc2 && c < 0xe0) { cp = c & 0x1f; n = 2; }
    else if (c >= 0xe0 && c < 0xf0) { cp = c & 0x0f; n = 3; }
    else if (c >= 0xf0 && c < 0xf5) { cp = c & 0x07; n = 4; }
    else { if (strict) return null; out += "\ufffd"; i++; continue; }
    if (i + n > b.length) { if (strict) return null; out += "\ufffd"; break; }
    let ok = true;
    for (let k = 1; k < n; k++) {
      if ((b[i + k] & 0xc0) !== 0x80) { ok = false; break; }
      cp = (cp << 6) | (b[i + k] & 0x3f);
    }
    if (!ok) { if (strict) return null; out += "\ufffd"; i++; continue; }
    out += String.fromCodePoint(cp);
    i += n;
  }
  return out;
}

function utf16(b: ArrayLike<number>, from: number, little: boolean): string {
  let out = "";
  for (let i = from; i + 1 < b.length; i += 2) out += String.fromCharCode(little ? b[i] | (b[i + 1] << 8) : (b[i] << 8) | b[i + 1]);
  return out;
}

export function decode(b: ArrayLike<number>, encoding = "auto"): string {
  const e = encoding.trim().toLowerCase().replace(/^utf8$/, "utf-8").replace(/^cp1252$/, "windows-1252").replace(/^iso-8859-1$/, "latin1");
  if (!ENCODINGS.includes(e)) throw new RangeError(`${encoding} is not an encoding (${ENCODINGS.join(", ")})`);
  const bom8 = b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf, le = b[0] === 0xff && b[1] === 0xfe, be = b[0] === 0xfe && b[1] === 0xff;
  if (e === "utf-16le" || (e === "auto" && le)) return utf16(b, le ? 2 : 0, true);
  if (e === "utf-16be" || (e === "auto" && be)) return utf16(b, be ? 2 : 0, false);
  if (e === "utf-8") return utf8(b, bom8 ? 3 : 0, false)!;
  if (e === "latin1") return Array.from(b as ArrayLike<number>, (c) => String.fromCharCode(c)).join("");
  if (e === "windows-1252") return Array.from(b as ArrayLike<number>, (c) => String.fromCharCode(CP1252[c] ?? c)).join("");
  return utf8(b, bom8 ? 3 : 0, true) ?? decode(b, "windows-1252");
}

// Base64 as bytes (a page's binary file comes as base64 text).
export function fromBase64(s: string): Uint8Array {
  const A = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const clean = s.replace(/[^A-Za-z0-9+/]/g, "");
  const out = new Uint8Array(Math.floor((clean.length * 3) / 4));
  let o = 0;
  for (let i = 0; i < clean.length; i += 4) {
    const n = (A.indexOf(clean[i]) << 18) | (A.indexOf(clean[i + 1]) << 12) | ((A.indexOf(clean[i + 2] ?? "A") & 63) << 6) | (A.indexOf(clean[i + 3] ?? "A") & 63);
    out[o++] = (n >> 16) & 255;
    if (i + 2 < clean.length) out[o++] = (n >> 8) & 255;
    if (i + 3 < clean.length) out[o++] = n & 255;
  }
  return out.slice(0, o);
}

export function toBase64(b: ArrayLike<number>): string {
  const A = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let out = "";
  for (let i = 0; i < b.length; i += 3) {
    const n = (b[i] << 16) | ((b[i + 1] ?? 0) << 8) | (b[i + 2] ?? 0);
    out += A[(n >> 18) & 63] + A[(n >> 12) & 63] + (i + 1 < b.length ? A[(n >> 6) & 63] : "=") + (i + 2 < b.length ? A[n & 63] : "=");
  }
  return out;
}

export function fromCsv(text: string, delimiter = ","): string[][] {
  text = text.replace(/^\uFEFF/, "");
  const rows: string[][] = [];
  let row: string[] = [], field = "", i = 0, line = 1;
  const endField = () => { row.push(field); field = ""; };
  const endRow = () => { endField(); rows.push(row); row = []; line++; };
  while (i < text.length) {
    const c = text[i];
    if (c === '"' && field === "") {
      const from = line;
      for (i++; ; i++) {
        if (i >= text.length) throw new RangeError(`a quoted field that never ends, on line ${from}`);
        if (text[i] === '"') {
          if (text[i + 1] !== '"') break;
          i++;
        }
        if (text[i] === "\n") line++;
        field += text[i];
      }
      i++;
      if (i < text.length && text[i] !== delimiter && !/[\r\n]/.test(text[i])) throw new RangeError(`text after a closing quote, on line ${from}`);
      continue;
    }
    if (c === delimiter) endField();
    else if (c === "\n") endRow();
    else if (c === "\r" && text[i + 1] === "\n") { endRow(); i++; }
    else field += c;
    i++;
  }
  if (field !== "" || row.length) endRow();
  return rows;
}
