// XLSX (Office Open XML, ECMA-376) read: the archive's workbook, its
// sheets in order (hidden ones too), each sheet's cells (numbers, shared
// and inline strings, booleans; a formula as its text, a shared formula
// moved to each cell it covers), number formats (the built-in ones by
// number, the file's own codes), each cell's look (font, fill, borders,
// alignment: its style), its comment, its validation, merged cells, column widths and row
// heights, frozen panes, hidden rows and columns; the workbook's defined
// names (a sheet's own by localSheetId) and its date system. Written as
// the app's own sheet file (file.ts reads it), with a report of what was
// left out. Pure: bytes in.
import { unzip, utf8Text } from "./zip.ts";
import { attr, all, kid, kids, parseXml, textOf } from "./xml.ts";
import type { El } from "./xml.ts";
import { COLS, ROWS, addrOf, colName, keyOf } from "./address.ts";
import { moved } from "./excel.ts";
import type { Look } from "./formats.ts";

// The built-in number formats by number (ECMA-376 18.8.30).
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = {
  1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
  18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

export interface XlsxRead {
  // The workbook as the app's sheet file (file.ts's readSheetFile reads it).
  readonly file: Record<string, unknown>;
  // What was left out, a line apiece.
  readonly report: string[];
}

const NUMERIC = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/;

// A path in the archive from a relationship's target, beside `from`.
function resolvePath(from: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = from.split("/").slice(0, -1);
  for (const p of target.split("/")) {
    if (p === "..") parts.pop();
    else if (p !== ".") parts.push(p);
  }
  return parts.join("/");
}

function rels(parts: Map<string, Uint8Array>, of: string): Map<string, string> {
  const dir = of.split("/").slice(0, -1).join("/"), name = of.split("/").pop()!;
  const r = parts.get(`${dir ? `${dir}/` : ""}_rels/${name}.rels`);
  const out = new Map<string, string>();
  if (!r) return out;
  for (const e of all(parseXml(utf8Text(r)), "Relationship")) out.set(attr(e, "Id") ?? "", resolvePath(of, attr(e, "Target") ?? ""));
  return out;
}

// A shared string: its text, or its rich runs' text together.
const stringOf = (si: El): string => {
  const t = kid(si, "t");
  return t && !kids(si, "r").length ? textOf(t) : kids(si, "r").map((r) => textOf(kid(r, "t") ?? { name: "t", attrs: {}, children: [] })).join("");
};

// A cell address as the app's ($ dropped), null when off the sheet.
const at = (ref: string): string | null => {
  const a = addrOf(ref.replace(/\$/g, ""));
  return a ? keyOf(a) : null;
};

// A defined name's range as the app writes one: Sheet1!B2:B4.
const nameRange = (text: string): string => text.replace(/\$/g, "");

const EMPTY: El = { name: "", attrs: {}, children: [] };

// A colour: rgb (ARGB, its alpha dropped), or a theme's (Excel numbers its
// first four light, dark, light, dark: themes.ts's roles), or indexed (the
// legacy palette's first eight). A tint is not kept.
const THEME_ROLES = ["light1", "dark1", "light2", "dark2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink"];
const INDEXED = ["#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff"];
export function colourOf(e: El | undefined): string | undefined {
  if (!e) return undefined;
  const rgb = attr(e, "rgb"), theme = attr(e, "theme"), indexed = attr(e, "indexed");
  if (rgb && /^[0-9a-f]{6,8}$/i.test(rgb)) return `#${rgb.slice(-6).toLowerCase()}`;
  if (theme !== undefined) return THEME_ROLES[Number(theme)];
  if (indexed !== undefined) return INDEXED[Number(indexed) % 8];
  return undefined;
}

const LINE: Record<string, string> = { thin: "thin", medium: "medium", thick: "thick", double: "double", dashed: "dashed", dotted: "dotted", hair: "hair",
  mediumDashed: "dashed", dashDot: "dashed", mediumDashDot: "dashed", dashDotDot: "dotted", mediumDashDotDot: "dotted", slantDashDot: "dashed" };

// A style (an xf) as the app's look: its font, fill, borders, alignment.
function lookOf(xf: El, fonts: El[], fills: El[], borders: El[]): Look | null {
  const look: Record<string, unknown> = {};
  const font = fonts[Number(attr(xf, "fontId") ?? 0)];
  if (font && Number(attr(xf, "fontId") ?? 0) > 0) {
    if (kid(font, "b") && attr(kid(font, "b")!, "val") !== "0") look.bold = true;
    if (kid(font, "i") && attr(kid(font, "i")!, "val") !== "0") look.italic = true;
    if (kid(font, "u") && attr(kid(font, "u")!, "val") !== "none") look.underline = true;
    if (kid(font, "strike") && attr(kid(font, "strike")!, "val") !== "0") look.strike = true;
    const sz = Number(attr(kid(font, "sz") ?? EMPTY, "val") ?? 0);
    if (sz && sz !== 11) look.size = Math.round(sz * 4 / 3);
    const c = colourOf(kid(font, "color"));
    if (c && c !== "dark1" && c !== "#000000") look.color = c;
  }
  const fill = fills[Number(attr(xf, "fillId") ?? 0)], pf = fill && kid(fill, "patternFill");
  if (pf && attr(pf, "patternType") === "solid") {
    const c = colourOf(kid(pf, "fgColor"));
    if (c) look.fill = c;
  }
  const border = borders[Number(attr(xf, "borderId") ?? 0)];
  if (border) {
    const sides: Record<string, string> = {};
    for (const side of ["top", "right", "bottom", "left"]) {
      const e = kid(border, side), style = e && attr(e, "style");
      if (!e || !style || style === "none") continue;
      const c = colourOf(kid(e, "color")) ?? "#000000", line = LINE[style] ?? "thin";
      sides[side] = line === "thin" ? c : `${line} ${c}`;
    }
    if (Object.keys(sides).length) look.borders = sides;
  }
  const al = kid(xf, "alignment");
  if (al) {
    const h = attr(al, "horizontal");
    if (h === "left") look.align = "start"; else if (h === "center" || h === "centerContinuous") look.align = "center"; else if (h === "right") look.align = "end";
    if (attr(al, "wrapText") === "1" || attr(al, "wrapText") === "true") look.wrap = true;
  }
  return Object.keys(look).length ? (look as Look) : null;
}

// A dataValidation as the app's rule, or why it is not one.
function ruleOf(dv: El): Record<string, unknown> | string {
  const type = attr(dv, "type") ?? "none", op = attr(dv, "operator") ?? "between";
  const f1 = textOf(kid(dv, "formula1") ?? EMPTY).trim(), f2 = textOf(kid(dv, "formula2") ?? EMPTY).trim();
  const mode = attr(dv, "errorStyle") === "warning" ? "warn" : attr(dv, "errorStyle") === "information" ? "info" : "stop";
  const extra = { mode, when: "always", ...(attr(dv, "prompt") ? { prompt: attr(dv, "prompt") } : {}), ...(attr(dv, "error") ? { error: attr(dv, "error") } : {}) };
  const kind = ({ whole: "whole", decimal: "decimal", textLength: "length", date: "date", time: "time", list: "list", custom: "custom" } as Record<string, string>)[type];
  if (!kind) return `of type ${type}`;
  if (kind === "list") {
    if (!/^".*"$/.test(f1)) return `a list from ${f1} (a range; the app's lists are values)`;
    return { kind, values: f1.slice(1, -1).split(",").map((x) => x.trim()).filter((x) => x), ...extra };
  }
  if (kind === "custom") return { kind, formula: f1.replace(/^=/, ""), ...extra };
  const n = (f: string) => (/^-?[\d.]+(E[-+]?\d+)?$/i.test(f) ? Number(f) : null);
  const a = n(f1), b = n(f2);
  if (a === null || (op === "between" || op === "notBetween" ? b === null : false)) return `with limits ${f1}${f2 ? `, ${f2}` : ""} (formulas)`;
  const lim = op === "between" ? { min: a, max: b } : op === "greaterThanOrEqual" ? { min: a } : op === "lessThanOrEqual" ? { max: a } : op === "equal" ? { min: a, max: a }
    : op === "greaterThan" ? { min: kind === "decimal" || kind === "date" || kind === "time" ? a : a + 1 } : op === "lessThan" ? { max: kind === "decimal" || kind === "date" || kind === "time" ? a : a - 1 } : null;
  if (!lim) return `with the operator ${op}`;
  return { kind, ...lim, ...extra };
}

export function readXlsx(bytes: Uint8Array, fileName = "Workbook"): XlsxRead {
  const parts = unzip(bytes), report: string[] = [];
  const book = parts.get("xl/workbook.xml") ?? parts.get(resolvePath("", rels(parts, "").get("rId1") ?? "xl/workbook.xml"));
  if (!book) throw new RangeError("not an XLSX workbook: no xl/workbook.xml");
  const bookPath = parts.has("xl/workbook.xml") ? "xl/workbook.xml" : resolvePath("", rels(parts, "").get("rId1") ?? "");
  const wb = parseXml(utf8Text(book)), wbRels = rels(parts, bookPath);
  const strings: string[] = [];
  const sstPath = [...wbRels.values()].find((p) => /sharedStrings\.xml$/i.test(p)) ?? "xl/sharedStrings.xml";
  if (parts.has(sstPath)) for (const si of kids(parseXml(utf8Text(parts.get(sstPath)!)), "si")) strings.push(stringOf(si));
  // Number formats and looks by style index (cellXfs).
  const formats: string[] = [], looks: (Look | null)[] = [];
  const stylesPath = [...wbRels.values()].find((p) => /styles\.xml$/i.test(p)) ?? "xl/styles.xml";
  if (parts.has(stylesPath)) {
    const st = parseXml(utf8Text(parts.get(stylesPath)!)), own = new Map<number, string>();
    for (const f of all(st, "numFmt")) own.set(Number(attr(f, "numFmtId")), attr(f, "formatCode") ?? "");
    const fonts = kids(kid(st, "fonts") ?? EMPTY, "font"), fills = kids(kid(st, "fills") ?? EMPTY, "fill"), borders = kids(kid(st, "borders") ?? EMPTY, "border");
    const xfs = kid(st, "cellXfs");
    for (const xf of xfs ? kids(xfs, "xf") : []) {
      const id = Number(attr(xf, "numFmtId") ?? 0);
      formats.push(own.get(id) ?? BUILTIN_FORMATS[id] ?? "");
      looks.push(lookOf(xf, fonts, fills, borders));
    }
  }
  const date1904 = ["1", "true"].includes((attr(kid(wb, "workbookPr") ?? { name: "", attrs: {}, children: [] }, "date1904") ?? "").toLowerCase());
  const sheetEls = kids(kid(wb, "sheets") ?? { name: "", attrs: {}, children: [] }, "sheet");
  if (!sheetEls.length) throw new RangeError("an XLSX workbook with no sheets");
  const sheets: Record<string, unknown>[] = [];
  let off = 0;
  sheetEls.forEach((se) => {
    const name = attr(se, "name") ?? `Sheet${sheets.length + 1}`, path = wbRels.get(attr(se, "id") ?? "") ?? "";
    const xml = parts.get(path);
    const cells: Record<string, unknown> = {};
    const view: Record<string, unknown> = { hiddenRows: [] as number[], hiddenCols: [] as number[], rowHeights: [] as [number, number][], colWidths: [] as [number, number][] };
    if (attr(se, "state") === "hidden" || attr(se, "state") === "veryHidden") view.hidden = true;
    const merges: string[] = [];
    if (!xml) { report.push(`${name}: not a worksheet (a chart sheet or dialog sheet), left out`); return; }
    const ws = parseXml(utf8Text(xml));
    const shared = new Map<string, { formula: string; from: string }>();
    for (const row of all(ws, "row")) {
      const r = Number(attr(row, "r") ?? 0) - 1;
      if (r >= 0 && r < ROWS) {
        if (attr(row, "hidden") === "1" || attr(row, "hidden") === "true") (view.hiddenRows as number[]).push(r);
        if (attr(row, "customHeight") === "1" && attr(row, "ht")) (view.rowHeights as [number, number][]).push([r, Math.round(Number(attr(row, "ht")) * 4 / 3)]);
      }
      for (const c of kids(row, "c")) {
        const ref = attr(c, "r") ?? "", key = at(ref);
        const t = attr(c, "t") ?? "n", v = kid(c, "v"), f = kid(c, "f");
        let input = "";
        if (f) {
          const kind = attr(f, "t"), si = attr(f, "si"), text = textOf(f);
          if (kind === "shared" && si !== undefined) {
            if (text) shared.set(si, { formula: text, from: ref });
            const master = shared.get(si);
            if (master) {
              const a = addrOf(master.from.replace(/\$/g, "")), b = addrOf(ref.replace(/\$/g, ""));
              input = a && b ? moved(`=${master.formula}`, { cols: b.col - a.col, rows: b.row - a.row }) : `=${master.formula}`;
            }
          } else if (text) input = `=${text}`;
        }
        if (!input) {
          const raw = v ? textOf(v) : "";
          if (t === "s") input = strings[Number(raw)] ?? "";
          else if (t === "inlineStr") input = textOf(kid(c, "is") ?? { name: "", attrs: {}, children: [] });
          else if (t === "b") input = raw === "1" ? "TRUE" : "FALSE";
          else if (t === "str" || t === "e") input = raw;
          else input = raw;
          // Text that the sheet would read as a number, or a formula, kept text.
          if ((t === "s" || t === "inlineStr" || t === "str") && (NUMERIC.test(input) || input.startsWith("=") || input.startsWith("'"))) input = `'${input}`;
        }
        const si = Number(attr(c, "s") ?? 0), format = formats[si] ?? "", look: Look = { ...(looks[si] ?? {}), ...(format && format !== "General" ? { number: format } : {}) };
        if (!key) { if (input !== "") off++; continue; }
        if (input === "" && !Object.keys(look).length) continue;
        cells[key] = Object.keys(look).length ? { content: input, look } : input;
      }
    }
    for (const col of all(ws, "col")) {
      const lo = Number(attr(col, "min") ?? 1) - 1, hi = Math.min(COLS - 1, Number(attr(col, "max") ?? lo + 1) - 1);
      for (let k = lo; k <= hi && k < COLS; k++) {
        if (attr(col, "hidden") === "1" || attr(col, "hidden") === "true") (view.hiddenCols as number[]).push(k);
        if (attr(col, "customWidth") === "1" && attr(col, "width")) (view.colWidths as [number, number][]).push([k, Math.round(Number(attr(col, "width")) * 7 + 5)]);
      }
    }
    const pane = all(ws, "pane")[0];
    if (pane && (attr(pane, "state") === "frozen" || attr(pane, "state") === "frozenSplit")) {
      view.frozenRows = Number(attr(pane, "ySplit") ?? 0) | 0;
      view.frozenCols = Number(attr(pane, "xSplit") ?? 0) | 0;
    }
    for (const m of all(ws, "mergeCell")) {
      const [a, b] = (attr(m, "ref") ?? "").split(":");
      if (at(a ?? "") && at(b ?? a ?? "")) merges.push(`${at(a)}:${at(b ?? a)}`);
    }
    // Its comments (a legacy comments part, by the sheet's relationships).
    const sheetRels = rels(parts, path);
    for (const target of sheetRels.values()) {
      if (!/comments\d*\.xml$/i.test(target) || !parts.has(target)) continue;
      for (const cm of all(parseXml(utf8Text(parts.get(target)!)), "comment")) {
        const k = at(attr(cm, "ref") ?? "");
        if (!k) continue;
        const t = textOf(cm).trim();
        const was = cells[k];
        cells[k] = typeof was === "object" && was ? { ...(was as object), comment: t } : { content: typeof was === "string" ? was : "", comment: t };
      }
    }
    // Its data validation, as the app's rules where they are one.
    for (const dv of all(ws, "dataValidation")) {
      const rule = ruleOf(dv);
      if (typeof rule === "string") { report.push(`${name}!${attr(dv, "sqref")}: validation ${rule}, left out`); continue; }
      // A custom formula is written for the sqref's top-left.
      const areas = (attr(dv, "sqref") ?? "").split(/\s+/).filter((x) => x).map((area) => area.split(":").map((x) => addrOf(x.replace(/\$/g, ""))));
      const corners = areas.flat().filter((x): x is NonNullable<typeof x> => !!x);
      const origin = corners.length ? keyOf({ col: Math.min(...corners.map((x) => x.col)), row: Math.min(...corners.map((x) => x.row)) }) : "A1";
      for (const [ka, kb = ka] of areas) {
        if (!ka || !kb) continue;
        for (let r = Math.min(ka.row, kb.row); r <= Math.max(ka.row, kb.row); r++) for (let cc = Math.min(ka.col, kb.col); cc <= Math.max(ka.col, kb.col); cc++) {
          const k = keyOf({ col: cc, row: r }), was = cells[k];
          const v = { ...rule, ...(rule.kind === "custom" ? { origin } : {}) };
          cells[k] = typeof was === "object" && was ? { ...(was as object), validation: v } : { content: typeof was === "string" ? was : "", validation: v };
        }
      }
    }
    sheets.push({ name, cells, view, merges, names: {} });
  });
  if (off) report.push(`${off} cell${off === 1 ? "" : "s"} past ${colName(COLS - 1)}${ROWS}, the sheet's edge, left out`);
  // The defined names: a sheet's own (localSheetId), the workbook's.
  const bookNames: Record<string, string> = {};
  for (const dn of all(wb, "definedName")) {
    const n = attr(dn, "name") ?? "", text = textOf(dn).trim();
    if (!n || n.startsWith("_xlnm.") || n.startsWith("_")) { if (n.startsWith("_xlnm.")) report.push(`${n.slice(6)} (a print setting) left out`); continue; }
    const local = attr(dn, "localSheetId");
    const range = /^(?:'[^']+'|[^'!]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i.test(text) ? nameRange(text) : null;
    if (local !== undefined && sheets[Number(local)]) {
      if (range) (sheets[Number(local)].names as Record<string, string>)[n] = range.slice(range.indexOf("!") + 1);
      else report.push(`${n}, a sheet's own formula name, left out`);
    } else bookNames[n] = range ?? `=${text}`;
  }
  const [first, ...rest] = sheets;
  const file: Record<string, unknown> = {
    name: fileName.replace(/\.xlsx$/i, ""), formulas: "excel", sheetName: first.name, ...first,
    sheets: rest, ...(Object.keys(bookNames).length ? { bookNames } : {}), ...(date1904 ? { dateSystem: 1904 } : {}),
  };
  delete file.name;
  file.name = fileName.replace(/\.xlsx$/i, "");
  return { file, report };
}
