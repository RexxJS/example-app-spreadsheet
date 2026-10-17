// XLS (Excel 97 to 2003, BIFF8) read: the compound file (OLE2: its FAT,
// mini FAT and directory) to the Workbook stream, then its records: the
// sheets (BOUNDSHEET, their visibility), shared strings (SST, across
// CONTINUE records), number formats (FORMAT, XF), the 1904 date system
// (DATEMODE), and each sheet's cells: numbers (NUMBER, RK, MULRK), strings
// (LABELSST, LABEL), booleans and errors (BOOLERR), formulas (FORMULA, its
// parsed tokens written back as text, with STRING for a text result; a
// token the app cannot write keeps the value cached, said), merges
// (MERGECELLS). Written as the app's own sheet file. Pure: bytes in.
import { COLS, ROWS, colName, keyOf } from "./address.ts";
import { BUILTIN_FORMATS } from "./xlsx.ts";

const u16 = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8);
const u32 = (b: Uint8Array, i: number) => (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0;
const f64 = (b: Uint8Array, i: number) => new DataView(b.buffer, b.byteOffset + i, 8).getFloat64(0, true);

// ---- The compound file.
export function compoundStreams(b: Uint8Array): Map<string, Uint8Array> {
  const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (b.length < 512 || sig.some((x, i) => b[i] !== x)) throw new RangeError("not an XLS file (not a compound document)");
  const sectorSize = 1 << u16(b, 0x1e), miniSize = 1 << u16(b, 0x20), cutoff = u32(b, 0x38);
  const sector = (n: number) => 512 + n * sectorSize;
  // The FAT, from the DIFAT (109 in the header, then DIFAT sectors).
  const difat: number[] = [];
  for (let i = 0; i < 109; i++) { const s = u32(b, 0x4c + i * 4); if (s < 0xfffffffa) difat.push(s); }
  for (let s = u32(b, 0x44), n = u32(b, 0x48); n > 0 && s < 0xfffffffa; n--) {
    for (let i = 0; i < sectorSize / 4 - 1; i++) { const x = u32(b, sector(s) + i * 4); if (x < 0xfffffffa) difat.push(x); }
    s = u32(b, sector(s) + sectorSize - 4);
  }
  const fat: number[] = [];
  for (const s of difat) for (let i = 0; i < sectorSize / 4; i++) fat.push(u32(b, sector(s) + i * 4));
  const chain = (start: number, next: readonly number[]): number[] => {
    const out: number[] = [];
    for (let s = start, guard = 0; s < 0xfffffffa && guard < next.length + 1; s = next[s], guard++) out.push(s);
    return out;
  };
  const read = (start: number, size: number): Uint8Array => {
    const out = new Uint8Array(size);
    let o = 0;
    for (const s of chain(start, fat)) {
      const n = Math.min(sectorSize, size - o);
      if (n <= 0) break;
      out.set(b.subarray(sector(s), sector(s) + n), o);
      o += n;
    }
    return out;
  };
  const dirStart = u32(b, 0x30), dirChain = chain(dirStart, fat);
  const dir = read(dirStart, dirChain.length * sectorSize);
  const entries: { name: string; type: number; start: number; size: number }[] = [];
  for (let i = 0; i + 128 <= dir.length; i += 128) {
    const len = u16(dir, i + 64);
    let name = "";
    for (let k = 0; k + 2 < len; k += 2) name += String.fromCharCode(u16(dir, i + k));
    entries.push({ name, type: dir[i + 66], start: u32(dir, i + 0x74), size: u32(dir, i + 0x78) });
  }
  const root = entries.find((e) => e.type === 5);
  const miniStream = root ? read(root.start, root.size) : new Uint8Array(0);
  const miniFat: number[] = [];
  const mfStart = u32(b, 0x3c), mfCount = u32(b, 0x40);
  if (mfCount) {
    const mf = read(mfStart, chain(mfStart, fat).length * sectorSize);
    for (let i = 0; i + 4 <= mf.length; i += 4) miniFat.push(u32(mf, i));
  }
  const out = new Map<string, Uint8Array>();
  for (const e of entries) {
    if (e.type !== 2) continue;
    if (e.size < cutoff) {
      const data = new Uint8Array(e.size);
      let o = 0;
      for (const s of chain(e.start, miniFat)) {
        const n = Math.min(miniSize, e.size - o);
        if (n <= 0) break;
        data.set(miniStream.subarray(s * miniSize, s * miniSize + n), o);
        o += n;
      }
      out.set(e.name, data);
    } else out.set(e.name, read(e.start, e.size));
  }
  return out;
}

// ---- Records.
interface Rec { readonly type: number; readonly data: Uint8Array; readonly at: number }
function records(s: Uint8Array): Rec[] {
  const out: Rec[] = [];
  for (let i = 0; i + 4 <= s.length;) {
    const type = u16(s, i), len = u16(s, i + 2);
    out.push({ type, data: s.subarray(i + 4, i + 4 + len), at: i });
    i += 4 + len;
  }
  return out;
}

// An XLUnicodeString (a count of 8 or 16 bits, flags, then the characters),
// at `i`; its text and where it ends.
function xlString(d: Uint8Array, i: number, countBytes: 1 | 2): [string, number] {
  const n = countBytes === 1 ? d[i] : u16(d, i);
  i += countBytes;
  const flags = d[i++], wide = flags & 1;
  let s = "";
  for (let k = 0; k < n; k++) { s += String.fromCharCode(wide ? u16(d, i) : d[i]); i += wide ? 2 : 1; }
  return [s, i];
}

// The SST across its CONTINUE records: a string may break between them,
// its next part opening with a fresh flags byte.
function sharedStrings(parts: readonly Uint8Array[]): string[] {
  const out: string[] = [];
  let p = 0, i = 8;
  const total = u32(parts[0], 4);
  const byte = () => { if (i >= parts[p].length) { p++; i = 0; } return parts[p][i++]; };
  for (let k = 0; k < total && p < parts.length; k++) {
    if (i >= parts[p].length) { p++; i = 0; if (p >= parts.length) break; }
    const n = byte() | (byte() << 8);
    let flags = byte();
    const rich = flags & 8 ? byte() | (byte() << 8) : 0;
    const ext = flags & 4 ? (byte() | (byte() << 8) | (byte() << 16) | (byte() << 24)) >>> 0 : 0;
    let s = "";
    for (let c = 0; c < n; c++) {
      if (i >= parts[p].length) { p++; i = 0; flags = parts[p][i++]; }
      s += String.fromCharCode(flags & 1 ? byte() | (byte() << 8) : byte());
    }
    for (let x = 0; x < rich * 4 + ext; x++) byte();
    out.push(s);
  }
  return out;
}

const ERRORS: Record<number, string> = { 0x00: "#NULL!", 0x07: "#DIV/0!", 0x0f: "#VALUE!", 0x17: "#REF!", 0x1d: "#NAME?", 0x24: "#NUM!", 0x2a: "#N/A" };

// An RK number: a 30-bit integer or the top of a double, maybe over 100.
function rk(v: number): number {
  const n = v & 2 ? (v >> 2) : new DataView(new Uint32Array([0, v & 0xfffffffc]).buffer).getFloat64(0, true);
  return v & 1 ? n / 100 : n;
}

// ---- Formulas: the tokens ([MS-XLS] 2.5.198) back to text.
// The built-in functions by number, those with a fixed count of arguments
// with it ([MS-XLS] 2.5.198.17 Ftab).
const FIXED: Record<number, [string, number]> = {
  2: ["ISNA", 1], 3: ["ISERROR", 1], 10: ["NA", 0], 15: ["SIN", 1], 16: ["COS", 1], 17: ["TAN", 1], 18: ["ATAN", 1], 19: ["PI", 0], 20: ["SQRT", 1], 21: ["EXP", 1], 22: ["LN", 1],
  23: ["LOG10", 1], 24: ["ABS", 1], 25: ["INT", 1], 26: ["SIGN", 1], 27: ["ROUND", 2], 30: ["REPT", 2], 31: ["MID", 3], 32: ["LEN", 1], 33: ["VALUE", 1], 34: ["TRUE", 0],
  35: ["FALSE", 0], 38: ["NOT", 1], 39: ["MOD", 2], 48: ["TEXT", 2], 63: ["RAND", 0], 65: ["DATE", 3], 66: ["TIME", 3], 67: ["DAY", 1], 68: ["MONTH", 1], 69: ["YEAR", 1],
  71: ["HOUR", 1], 72: ["MINUTE", 1], 73: ["SECOND", 1], 74: ["NOW", 0], 76: ["ROWS", 1], 77: ["COLUMNS", 1], 97: ["ATAN2", 2], 98: ["ASIN", 1], 99: ["ACOS", 1], 111: ["CHAR", 1],
  112: ["LOWER", 1], 113: ["UPPER", 1], 114: ["PROPER", 1], 117: ["EXACT", 2], 118: ["TRIM", 1], 119: ["REPLACE", 4], 121: ["CODE", 1], 126: ["ISERR", 1], 127: ["ISTEXT", 1],
  128: ["ISNUMBER", 1], 129: ["ISBLANK", 1], 130: ["T", 1], 131: ["N", 1], 140: ["DATEVALUE", 1], 141: ["TIMEVALUE", 1], 162: ["CLEAN", 1], 165: ["MMULT", 2], 184: ["FACT", 1],
  198: ["ISLOGICAL", 1], 212: ["ROUNDUP", 2], 213: ["ROUNDDOWN", 2], 221: ["TODAY", 0], 252: ["FREQUENCY", 2], 276: ["COMBIN", 2], 285: ["FLOOR", 2], 288: ["CEILING", 2],
  337: ["POWER", 2], 342: ["RADIANS", 1], 343: ["DEGREES", 1], 346: ["COUNTIF", 2], 347: ["COUNTBLANK", 1],
};
const VARIABLE: Record<number, string> = {
  0: "COUNT", 1: "IF", 4: "SUM", 5: "AVERAGE", 6: "MIN", 7: "MAX", 8: "ROW", 9: "COLUMN", 11: "NPV", 12: "STDEV", 13: "DOLLAR", 14: "FIXED", 28: "LOOKUP", 29: "INDEX", 36: "AND",
  37: "OR", 46: "VAR", 56: "PV", 57: "FV", 58: "NPER", 59: "PMT", 60: "RATE", 62: "IRR", 64: "MATCH", 70: "WEEKDAY", 78: "OFFSET", 82: "SEARCH", 100: "CHOOSE", 101: "HLOOKUP",
  102: "VLOOKUP", 109: "LOG", 115: "LEFT", 116: "RIGHT", 120: "SUBSTITUTE", 124: "FIND", 148: "INDIRECT", 169: "COUNTA", 183: "PRODUCT", 193: "STDEVP", 194: "VARP", 197: "TRUNC",
  216: "RANK", 219: "ADDRESS", 220: "DAYS360", 227: "MEDIAN", 228: "SUMPRODUCT", 269: "AVEDEV", 336: "CONCATENATE", 344: "SUBTOTAL", 345: "SUMIF", 359: "HYPERLINK",
  361: "AVERAGEA", 362: "MAXA", 363: "MINA",
};
const BINARY: Record<number, string> = { 0x03: "+", 0x04: "-", 0x05: "*", 0x06: "/", 0x07: "^", 0x08: "&", 0x09: "<", 0x0a: "<=", 0x0b: "=", 0x0c: ">=", 0x0d: ">", 0x0e: "<>", 0x0f: " ", 0x10: ",", 0x11: ":" };

class Unwritable extends Error {}

// A reference's cell: a column field's low 14 bits, bit 14 the column
// relative, bit 15 the row; $ where it is not.
const refText = (row: number, colField: number): string =>
  `${colField & 0x4000 ? "" : "$"}${colName(colField & 0x3fff)}${colField & 0x8000 ? "" : "$"}${row + 1}`;

export function formulaText(d: Uint8Array, sheetName: (ixti: number) => string | null, nameOf: (i: number) => string | null): string {
  const st: string[] = [];
  const pop = () => { const x = st.pop(); if (x === undefined) throw new Unwritable("an empty stack"); return x; };
  const sheet = (ixti: number) => {
    const n = sheetName(ixti);
    if (n === null) throw new Unwritable("a reference to another workbook");
    return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? `${n}!` : `'${n.replace(/'/g, "''")}'!`;
  };
  for (let i = 0; i < d.length;) {
    const t = d[i++];
    if (BINARY[t]) { const b = pop(), a = pop(); st.push(`${a}${BINARY[t]}${b}`); continue; }
    const base = t >= 0x20 ? (t & 0x1f) | 0x20 : t;
    switch (base) {
      case 0x12: st.push(`+${pop()}`); break;
      case 0x13: st.push(`-${pop()}`); break;
      case 0x14: st.push(`${pop()}%`); break;
      case 0x15: st.push(`(${pop()})`); break;
      case 0x16: st.push(""); break;
      case 0x17: { const [s, end] = xlString(d, i, 1); st.push(`"${s.replace(/"/g, '""')}"`); i = end; break; }
      case 0x19: {
        const kind = d[i];
        if (kind & 0x04) i += 3 + (u16(d, i + 1) + 1) * 2; else i += 3;
        if (kind & 0x10) st.push(`SUM(${pop()})`);
        break;
      }
      case 0x1c: st.push(ERRORS[d[i++]] ?? "#N/A"); break;
      case 0x1d: st.push(d[i++] ? "TRUE" : "FALSE"); break;
      case 0x1e: st.push(String(u16(d, i))); i += 2; break;
      case 0x1f: st.push(String(f64(d, i))); i += 8; break;
      case 0x21: {
        const f = FIXED[u16(d, i)];
        i += 2;
        if (!f) throw new Unwritable(`function number ${u16(d, i - 2)}`);
        const args = st.splice(st.length - f[1], f[1]);
        st.push(`${f[0]}(${args.join(",")})`);
        break;
      }
      case 0x22: {
        const n = d[i] & 0x7f, id = u16(d, i + 1) & 0x7fff;
        i += 3;
        const name = VARIABLE[id] ?? FIXED[id]?.[0];
        if (!name) throw new Unwritable(id === 255 ? "an add-in or VBA function" : `function number ${id}`);
        const args = st.splice(st.length - n, n);
        st.push(`${name}(${args.join(",")})`);
        break;
      }
      case 0x23: { const n = nameOf(u32(d, i)); if (n === null) throw new Unwritable("a name"); st.push(n); i += 4; break; }
      case 0x24: st.push(refText(u16(d, i), u16(d, i + 2))); i += 4; break;
      case 0x25: st.push(`${refText(u16(d, i), u16(d, i + 4))}:${refText(u16(d, i + 2), u16(d, i + 6))}`); i += 8; break;
      case 0x2a: st.push("#REF!"); i += 4; break;
      case 0x2b: st.push("#REF!"); i += 8; break;
      case 0x3a: st.push(`${sheet(u16(d, i))}${refText(u16(d, i + 2), u16(d, i + 4))}`); i += 6; break;
      case 0x3b: st.push(`${sheet(u16(d, i))}${refText(u16(d, i + 2), u16(d, i + 6))}:${refText(u16(d, i + 4), u16(d, i + 8))}`); i += 10; break;
      case 0x26: case 0x27: i += 6; break;
      case 0x29: i += 2; break;
      default: throw new Unwritable(`token 0x${t.toString(16)}`);
    }
  }
  if (st.length !== 1) throw new Unwritable("a formula that does not come to one value");
  return st[0];
}

export function readXls(bytes: Uint8Array, fileName = "Workbook"): { file: Record<string, unknown>; report: string[] } {
  const streams = compoundStreams(bytes), report: string[] = [];
  const wbs = streams.get("Workbook") ?? streams.get("Book");
  if (!wbs) throw new RangeError("not an Excel workbook: no Workbook stream");
  if (streams.has("EncryptionInfo") || streams.has("EncryptedPackage")) throw new RangeError("an encrypted workbook");
  const recs = records(wbs);
  if (!recs.length || recs[0].type !== 0x0809 || u16(recs[0].data, 0) !== 0x0600) throw new RangeError("not a BIFF8 workbook (Excel 97 to 2003); older ones are not read");
  const sheets: { name: string; pos: number; hidden: boolean }[] = [];
  const formats = new Map<number, string>(), xfFormat: number[] = [], sstParts: Uint8Array[] = [];
  const externSheets: number[] = [], names: string[] = [];
  let date1904 = false, inSst = false, i = 0;
  for (; i < recs.length; i++) {
    const r = recs[i], d = r.data;
    if (r.type !== 0x003c) inSst = false;
    if (r.type === 0x000a) break;
    switch (r.type) {
      case 0x0085: { const [name] = xlString(d, 6, 1); sheets.push({ name, pos: u32(d, 0), hidden: d[4] !== 0 }); break; }
      case 0x00fc: sstParts.push(d); inSst = true; break;
      case 0x003c: if (inSst) sstParts.push(d); break;
      case 0x0022: date1904 = u16(d, 0) === 1; break;
      case 0x041e: { const [code] = xlString(d, 2, 2); formats.set(u16(d, 0), code); break; }
      case 0x00e0: xfFormat.push(u16(d, 2)); break;
      case 0x0017: { const n = u16(d, 0); for (let k = 0; k < n; k++) externSheets.push(u16(d, 2 + k * 6 + 2)); break; }
      case 0x0018: {
        // NAME: its length at 3, its characters at 14 (a flags byte first);
        // a built-in name (Print_Area ...) is not one a formula writes.
        const cch = d[3], wide = d[14] & 1;
        let nm = "";
        for (let k = 0; k < cch; k++) nm += String.fromCharCode(wide ? u16(d, 15 + k * 2) : d[15 + k]);
        names.push(u16(d, 0) & 0x20 ? "" : nm);
        break;
      }
    }
  }
  const strings = sstParts.length ? sharedStrings(sstParts) : [];
  const formatOf = (xf: number): string => { const id = xfFormat[xf] ?? 0; return formats.get(id) ?? BUILTIN_FORMATS[id] ?? ""; };
  const out: Record<string, unknown>[] = [];
  let off = 0, unwritten = 0;
  const why = new Set<string>();
  for (const sh of sheets) {
    const cells: Record<string, unknown> = {}, merges: string[] = [];
    const start = recs.findIndex((r) => r.at === sh.pos);
    if (start < 0 || recs[start].type !== 0x0809) { report.push(`${sh.name}: not found in the file`); continue; }
    if (u16(recs[start].data, 2) !== 0x0010) { report.push(`${sh.name}: a chart or macro sheet, left out`); continue; }
    const put = (row: number, col: number, input: string, xf: number) => {
      if (row >= ROWS || col >= COLS) { if (input !== "") off++; return; }
      const fmt = formatOf(xf), key = keyOf({ col, row });
      if (input === "") return;
      cells[key] = fmt && fmt !== "General" ? { content: input, look: { number: fmt } } : input;
    };
    const text = (s: string) => (/^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/.test(s) || s.startsWith("=") || s.startsWith("'") ? `'${s}` : s);
    for (let k = start + 1; k < recs.length && recs[k].type !== 0x000a; k++) {
      const { type, data: d } = recs[k];
      switch (type) {
        case 0x0203: put(u16(d, 0), u16(d, 2), String(f64(d, 6)), u16(d, 4)); break;
        case 0x027e: put(u16(d, 0), u16(d, 2), String(rk(u32(d, 6))), u16(d, 4)); break;
        case 0x00bd: { const row = u16(d, 0), first = u16(d, 2), n = (d.length - 6) / 6; for (let c = 0; c < n; c++) put(row, first + c, String(rk(u32(d, 4 + c * 6 + 2))), u16(d, 4 + c * 6)); break; }
        case 0x00fd: put(u16(d, 0), u16(d, 2), text(strings[u32(d, 6)] ?? ""), u16(d, 4)); break;
        case 0x0204: { const [s] = xlString(d, 6, 2); put(u16(d, 0), u16(d, 2), text(s), u16(d, 4)); break; }
        case 0x0205: put(u16(d, 0), u16(d, 2), d[7] ? ERRORS[d[6]] ?? "#N/A" : d[6] ? "TRUE" : "FALSE", u16(d, 4)); break;
        case 0x0006: {
          const row = u16(d, 0), col = u16(d, 2), xf = u16(d, 4), cce = u16(d, 20);
          let cached = "";
          if (u16(d, 12) === 0xffff) {
            const kind = d[6];
            if (kind === 0 && recs[k + 1]?.type === 0x0207) cached = text(xlString(recs[k + 1].data, 0, 2)[0]);
            else if (kind === 1) cached = d[8] ? "TRUE" : "FALSE";
            else if (kind === 2) cached = ERRORS[d[8]] ?? "#N/A";
          } else cached = String(f64(d, 6));
          try {
            const f = formulaText(d.subarray(22, 22 + cce), (ix) => { const s = externSheets[ix]; return s === undefined || s >= sheets.length ? null : sheets[s].name; }, (n) => names[n - 1] || null);
            put(row, col, `=${f}`, xf);
          } catch (e) {
            if (!(e instanceof Unwritable)) throw e;
            unwritten++;
            why.add((e as Error).message);
            put(row, col, cached, xf);
          }
          break;
        }
        case 0x00e5: {
          const n = u16(d, 0);
          for (let m = 0; m < n; m++) {
            const r1 = u16(d, 2 + m * 8), r2 = u16(d, 4 + m * 8), c1 = u16(d, 6 + m * 8), c2 = u16(d, 8 + m * 8);
            if (r2 < ROWS && c2 < COLS) merges.push(`${keyOf({ col: c1, row: r1 })}:${keyOf({ col: c2, row: r2 })}`);
          }
          break;
        }
      }
    }
    out.push({ name: sh.name, cells, merges, view: sh.hidden ? { hidden: true } : {} });
  }
  if (!out.length) throw new RangeError("an XLS workbook with no worksheets");
  if (off) report.push(`${off} cell${off === 1 ? "" : "s"} past ${colName(COLS - 1)}${ROWS}, the sheet's edge, left out`);
  if (unwritten) report.push(`${unwritten} formula${unwritten === 1 ? "" : "s"} kept as the value cached (${[...why].join(", ")})`);
  const [first, ...rest] = out;
  return { file: { formulas: "excel", sheetName: first.name, ...first, sheets: rest, name: fileName.replace(/\.xls$/i, ""), ...(date1904 ? { dateSystem: 1904 } : {}) }, report };
}
