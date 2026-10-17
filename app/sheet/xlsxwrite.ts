// XLSX write: the workbook as an Office Open XML file (xlsx.ts reads it
// back): its sheets in order (hidden ones marked), each cell's input (a
// formula as its text with the value it came to, a number, text as a
// shared string, TRUE and FALSE as booleans), number formats (the file's
// own codes, from 164), merges, column widths, row heights, frozen panes,
// hidden rows and columns, the workbook's and the sheets' own names, the
// 1904 date system. RexxJS formulas are not Excel's: their values are
// written, and said in the report.
import { zip, utf8Bytes } from "./zip.ts";
import { escapeXml } from "./xml.ts";
import { addrOf, keyOf, order } from "./address.ts";
import { CellError } from "./values.ts";
import type { Value } from "./values.ts";
import type { Page, Workbook } from "./workbook.ts";
import { BUILTIN_FORMATS } from "./xlsx.ts";

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";

// A number as XML writes it: as JavaScript does, which Excel reads.
const num = (n: number): string => String(n);

export interface XlsxWritten {
  readonly bytes: Uint8Array;
  readonly report: string[];
}

export function writeXlsx(wb: Workbook): XlsxWritten {
  const report: string[] = [];
  const strings: string[] = [], stringAt = new Map<string, number>();
  const sst = (s: string) => {
    let i = stringAt.get(s);
    if (i === undefined) { i = strings.length; strings.push(s); stringAt.set(s, i); }
    return i;
  };
  // Number formats: each its style index (cellXfs), 0 General.
  const codes: string[] = [""], styleOf = (code: string) => {
    if (!code || code === "General") return 0;
    let i = codes.indexOf(code);
    if (i < 0) { codes.push(code); i = codes.length - 1; }
    return i;
  };
  const excel = wb.sheet.language.name === "Excel";
  if (!excel) report.push("RexxJS formulas are not Excel's: their values written, not the formulas");
  const pages = wb.pages();
  const sheetXml = pages.map((page) => sheetOf(page, wb, sst, styleOf, excel));
  // The workbook.
  const names: string[] = [];
  for (const [n, r] of wb.names()) names.push(`<definedName name="${escapeXml(n)}">${escapeXml(r.startsWith("=") ? r.slice(1) : absolute(r))}</definedName>`);
  pages.forEach((p, i) => {
    for (const [n, r] of p.sheet.names()) {
      const prefix = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(p.name) ? `${p.name}!` : `'${p.name.replace(/'/g, "''")}'!`;
      names.push(`<definedName name="${escapeXml(n)}" localSheetId="${i}">${escapeXml(r.startsWith("=") ? r.slice(1) : absolute(prefix + r))}</definedName>`);
    }
  });
  const book = `${HEAD}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><workbookPr${wb.dateSystem === 1904 ? ' date1904="1"' : ""}/><sheets>${
    pages.map((p, i) => `<sheet name="${escapeXml(p.name)}" sheetId="${i + 1}"${p.hiddenSheet ? ' state="hidden"' : ""} r:id="rId${i + 1}"/>`).join("")}</sheets>${
    names.length ? `<definedNames>${names.join("")}</definedNames>` : ""}</workbook>`;
  const n = pages.length;
  const bookRels = `${HEAD}<Relationships xmlns="${PKG}">${pages.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${n + 1}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId${n + 2}" Type="${REL}/styles" Target="styles.xml"/></Relationships>`;
  const own = codes.slice(1).map((c, i) => [Number(Object.entries(BUILTIN_FORMATS).find(([, x]) => x === c)?.[0] ?? 164 + i), c] as [number, string]);
  const styles = `${HEAD}<styleSheet xmlns="${MAIN}">${own.some(([id]) => id >= 164) ? `<numFmts count="${own.filter(([id]) => id >= 164).length}">${own.filter(([id]) => id >= 164).map(([id, c]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(c)}"/>`).join("")}</numFmts>` : ""}<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${codes.length}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${
    own.map(([id]) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join("")}</cellXfs></styleSheet>`;
  const shared = `${HEAD}<sst xmlns="${MAIN}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si><t${/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ""}>${escapeXml(s)}</t></si>`).join("")}</sst>`;
  const types = `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${
    pages.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;
  const rootRels = `${HEAD}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const entries: [string, Uint8Array][] = [["[Content_Types].xml", utf8Bytes(types)], ["_rels/.rels", utf8Bytes(rootRels)], ["xl/workbook.xml", utf8Bytes(book)],
    ["xl/_rels/workbook.xml.rels", utf8Bytes(bookRels)], ["xl/styles.xml", utf8Bytes(styles)], ["xl/sharedStrings.xml", utf8Bytes(shared)],
    ...sheetXml.map((x, i): [string, Uint8Array] => [`xl/worksheets/sheet${i + 1}.xml`, utf8Bytes(x)])];
  return { bytes: zip(entries), report };
}

// Sheet1!B2:B4 as Sheet1!$B$2:$B$4, as Excel writes a name's.
const absolute = (r: string): string => r.replace(/(^|[!:])\$?([A-Z]{1,3})\$?(\d+)/g, "$1$$$2$$$3");

function sheetOf(page: Page, wb: Workbook, sst: (s: string) => number, styleOf: (code: string) => number, excel: boolean): string {
  const keys = [...new Set([...page.sheet.keys(), ...page.notes.keys()])].filter((k) => addrOf(k)).sort((a, b) => order(addrOf(a)!) - order(addrOf(b)!));
  const rows = new Map<number, string[]>();
  for (const k of keys) {
    const a = addrOf(k)!, input = page.sheet.input(k), note = page.notes.get(k);
    const code = note.format || wb.lookOf(note)?.number || "";
    const s = styleOf(/[#0?@]|general|[ymdhs]/i.test(code) ? code : "");
    if (input === "" && !s) continue;
    const v: Value = page.sheet.value(k), attrS = s ? ` s="${s}"` : "";
    let c: string;
    if (input.startsWith("=") && excel) {
      const f = `<f>${escapeXml(input.slice(1))}</f>`;
      c = v instanceof CellError ? `<c r="${k}" t="e"${attrS}>${f}<v>${escapeXml(v.code)}</v></c>`
        : typeof v === "string" ? `<c r="${k}" t="str"${attrS}>${f}<v>${escapeXml(v)}</v></c>`
        : typeof v === "boolean" ? `<c r="${k}" t="b"${attrS}>${f}<v>${v ? 1 : 0}</v></c>`
        : typeof v === "number" ? `<c r="${k}"${attrS}>${f}<v>${num(v)}</v></c>` : `<c r="${k}"${attrS}>${f}</c>`;
    } else if (v === null || v === "") c = `<c r="${k}"${attrS}/>`;
    else if (typeof v === "number") c = `<c r="${k}"${attrS}><v>${num(v)}</v></c>`;
    else if (typeof v === "boolean") c = `<c r="${k}" t="b"${attrS}><v>${v ? 1 : 0}</v></c>`;
    else if (v instanceof CellError) c = `<c r="${k}" t="e"${attrS}><v>${escapeXml(v.code)}</v></c>`;
    else c = `<c r="${k}" t="s"${attrS}><v>${sst(input.startsWith("'") ? input.slice(1) : String(v))}</v></c>`;
    const list = rows.get(a.row) ?? [];
    list.push(c);
    rows.set(a.row, list);
  }
  for (const r of [...page.hiddenRows, ...page.heights.keys()]) if (!rows.has(r)) rows.set(r, []);
  const rowXml = [...rows].sort(([x], [y]) => x - y).map(([r, cs]) => {
    const h = page.heights.get(r), hidden = page.hiddenRows.has(r);
    return `<row r="${r + 1}"${h ? ` ht="${Math.round(h * 0.75 * 100) / 100}" customHeight="1"` : ""}${hidden ? ' hidden="1"' : ""}>${cs.join("")}</row>`;
  }).join("");
  const colKeys = [...new Set([...page.widths.keys(), ...page.hiddenCols])].sort((x, y) => x - y);
  const cols = colKeys.map((c) => {
    const w = page.widths.get(c);
    return `<col min="${c + 1}" max="${c + 1}"${w ? ` width="${Math.round(((w - 5) / 7) * 100) / 100}" customWidth="1"` : ' width="8.43"'}${page.hiddenCols.has(c) ? ' hidden="1"' : ""}/>`;
  }).join("");
  const fr = page.frozen, fc = page.frozenCols;
  const pane = fr || fc ? `<pane${fc ? ` xSplit="${fc}"` : ""}${fr ? ` ySplit="${fr}"` : ""} topLeftCell="${keyOf({ col: fc, row: fr })}" activePane="${fr && fc ? "bottomRight" : fr ? "bottomLeft" : "topRight"}" state="frozen"/>` : "";
  const merges = page.merges.map(([a, b]) => `<mergeCell ref="${keyOf(a)}:${keyOf(b)}"/>`).join("");
  return `${HEAD}<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>${cols ? `<cols>${cols}</cols>` : ""}<sheetData>${rowXml}</sheetData>${
    merges ? `<mergeCells count="${page.merges.length}">${merges}</mergeCells>` : ""}</worksheet>`;
}

