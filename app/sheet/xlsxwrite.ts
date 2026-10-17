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
import type { Look } from "./formats.ts";
import type { Rule } from "./validation.ts";
import { cfXml, styleDxf, tableXml } from "./xlsxparts.ts";
import { ROLES } from "./themes.ts";

// A colour as SpreadsheetML writes one: a theme's role by its number (Excel
// numbers its first four light, dark, light, dark), else ARGB.
const THEME_INDEX = ["light1", "dark1", "light2", "dark2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink"];
function colourXml(tag: string, c: string): string {
  const k = c.trim().toLowerCase();
  if ((ROLES as readonly string[]).includes(k)) return `<${tag} theme="${THEME_INDEX.indexOf(k)}"/>`;
  return `<${tag} rgb="FF${k.replace("#", "").toUpperCase()}"/>`;
}

const CARRIED_TYPES: Record<string, string> = {
  "docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml", "docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  "docProps/custom.xml": "application/vnd.openxmlformats-officedocument.custom-properties+xml", "xl/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
};
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
  // Styles: each distinct look and number format its cellXfs index (0 the
  // plain one), its font, fill and border each theirs.
  const xfKeys: string[] = [""], xfs: { code: string; font: number; fill: number; border: number; look: Look }[] = [{ code: "", font: 0, fill: 0, border: 0, look: {} }];
  const fonts: string[] = ['<font><sz val="11"/><name val="Calibri"/></font>'], fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders: string[] = ["<border><left/><right/><top/><bottom/><diagonal/></border>"];
  const index = (list: string[], x: string) => { let i = list.indexOf(x); if (i < 0) { list.push(x); i = list.length - 1; } return i; };
  const styleOf = (code: string, look: Look) => {
    const c = !code || code === "General" ? "" : code;
    const key = JSON.stringify([c, look]);
    if (key === JSON.stringify(["", {}])) return 0;
    let i = xfKeys.indexOf(key);
    if (i >= 0) return i;
    const font = look.bold || look.italic || look.underline || look.strike || look.size || look.color
      ? index(fonts, `<font>${look.bold ? "<b/>" : ""}${look.italic ? "<i/>" : ""}${look.strike ? "<strike/>" : ""}${look.underline ? "<u/>" : ""}<sz val="${look.size ? Math.round(look.size * 0.75 * 100) / 100 : 11}"/>${look.color ? colourXml("color", look.color) : ""}<name val="Calibri"/></font>`) : 0;
    const fill = look.fill ? index(fills, `<fill><patternFill patternType="solid">${colourXml("fgColor", look.fill)}<bgColor indexed="64"/></patternFill></fill>`) : 0;
    const b = look.borders ?? {};
    const side = (n: "left" | "right" | "top" | "bottom") => {
      const kept = b[n];
      if (!kept) return `<${n}/>`;
      const [style, colour] = kept.includes(" ") ? kept.split(" ") : ["thin", kept];
      return `<${n} style="${style}">${colourXml("color", colour)}</${n}>`;
    };
    const border = Object.keys(b).length ? index(borders, `<border>${side("left")}${side("right")}${side("top")}${side("bottom")}<diagonal/></border>`) : 0;
    xfKeys.push(key);
    xfs.push({ code: c, font, fill, border, look });
    return xfs.length - 1;
  };
  const excel = wb.sheet.language.name === "Excel";
  if (!excel) report.push("RexxJS formulas are not Excel's: their values written, not the formulas");
  const charts = wb.pages().reduce((k, p) => k + p.charts.length, 0), pivots = wb.pages().reduce((k, p) => k + p.pivots.size, 0), slicers = wb.pages().reduce((k, p) => k + p.slicers.length, 0);
  if (charts) report.push(`${charts} chart${charts === 1 ? "" : "s"} not written (the app draws its own, not DrawingML)`);
  if (pivots) report.push(`${pivots} pivot table${pivots === 1 ? "" : "s"} written as their cells, not as pivots`);
  if (slicers) report.push(`${slicers} slicer${slicers === 1 ? "" : "s"} not written`);
  const pages = wb.pages();
  // Conditional formats' differential styles, each its dxf index.
  const dxfs: string[] = [];
  const dxfOf = (style: string) => { const x = styleDxf(style); let i = dxfs.indexOf(x); if (i < 0) { dxfs.push(x); i = dxfs.length - 1; } return i; };
  let tableId = 0;
  const sheetParts = pages.map((page) => sheetOf(page, wb, sst, styleOf, excel, dxfOf, () => ++tableId));
  const sheetXml = sheetParts.map((x) => x.xml);
  const sparks = sheetParts.reduce((k, x) => k + x.sparks, 0);
  if (sparks) report.push(`${sparks} sparkline${sparks === 1 ? "" : "s"} written as their text (Excel has no SPARKLINE function)`);
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
  const bookRels = `${HEAD}<Relationships xmlns="${PKG}">${pages.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${n + 1}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId${n + 2}" Type="${REL}/styles" Target="styles.xml"/>${
    wb.carried.some(([x]) => x === "xl/theme/theme1.xml") ? `<Relationship Id="rId${n + 3}" Type="${REL}/theme" Target="theme/theme1.xml"/>` : ""}</Relationships>`;
  // The number formats: Excel's built-in numbers, else the file's own from 164.
  const ownCodes: string[] = [];
  const fmtId = (c: string): number => {
    if (!c) return 0;
    const builtin = Object.entries(BUILTIN_FORMATS).find(([, x]) => x === c)?.[0];
    if (builtin) return Number(builtin);
    if (!ownCodes.includes(c)) ownCodes.push(c);
    return 164 + ownCodes.indexOf(c);
  };
  const xfXml = xfs.map((x) => {
    const id = fmtId(x.code), al = x.look.align || x.look.wrap ? `<alignment${x.look.align ? ` horizontal="${({ start: "left", center: "center", end: "right" } as Record<string, string>)[x.look.align]}"` : ""}${x.look.wrap ? ' wrapText="1"' : ""}/>` : "";
    return `<xf numFmtId="${id}" fontId="${x.font}" fillId="${x.fill}" borderId="${x.border}" xfId="0"${id ? ' applyNumberFormat="1"' : ""}${x.font ? ' applyFont="1"' : ""}${x.fill ? ' applyFill="1"' : ""}${x.border ? ' applyBorder="1"' : ""}${al ? ' applyAlignment="1">' + al + "</xf>" : "/>"}`;
  }).join("");
  const styles = `${HEAD}<styleSheet xmlns="${MAIN}">${ownCodes.length ? `<numFmts count="${ownCodes.length}">${ownCodes.map((c, i) => `<numFmt numFmtId="${164 + i}" formatCode="${escapeXml(c)}"/>`).join("")}</numFmts>` : ""}<fonts count="${fonts.length}">${fonts.join("")}</fonts><fills count="${fills.length}">${fills.join("")}</fills><borders count="${borders.length}">${borders.join("")}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfXml}</cellXfs>${dxfs.length ? `<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>` : ""}</styleSheet>`;
  const shared = `${HEAD}<sst xmlns="${MAIN}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si><t${/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ""}>${escapeXml(s)}</t></si>`).join("")}</sst>`;
  const types = `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${
    pages.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}${sheetParts.map((x, i) => (x.comments.length ? `<Override PartName="/xl/comments${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>` : "")).join("")}${
    sheetParts.flatMap((x) => x.tables).map((t) => `<Override PartName="/xl/tables/table${t.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join("")}${
    wb.carried.map(([n]) => `<Override PartName="/${n}" ContentType="${CARRIED_TYPES[n] ?? "application/xml"}"/>`).join("")}<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`;
  const carried = new Set(wb.carried.map(([n]) => n));
  const rootRels = `${HEAD}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>${
    carried.has("docProps/core.xml") ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' : ""}${
    carried.has("docProps/app.xml") ? `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>` : ""}${
    carried.has("docProps/custom.xml") ? `<Relationship Id="rId4" Type="${REL}/custom-properties" Target="docProps/custom.xml"/>` : ""}</Relationships>`;
  const entries: [string, Uint8Array][] = [["[Content_Types].xml", utf8Bytes(types)], ["_rels/.rels", utf8Bytes(rootRels)], ["xl/workbook.xml", utf8Bytes(book)],
    ["xl/_rels/workbook.xml.rels", utf8Bytes(bookRels)], ["xl/styles.xml", utf8Bytes(styles)], ["xl/sharedStrings.xml", utf8Bytes(shared)],
    ...sheetXml.map((x, i): [string, Uint8Array] => [`xl/worksheets/sheet${i + 1}.xml`, utf8Bytes(x)])];
  sheetParts.forEach(({ comments, tables, relBase, urls }, i) => {
    if (!comments.length && !tables.length && !urls.length) return;
    const relsXml = [...(comments.length ? [`<Relationship Id="rId1" Type="${REL}/comments" Target="../comments${i + 1}.xml"/>`, `<Relationship Id="rId2" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing${i + 1}.vml"/>`] : []),
      ...tables.map((t, j) => `<Relationship Id="rId${relBase + j}" Type="${REL}/table" Target="../tables/table${t.id}.xml"/>`),
      ...urls.map((u, j) => `<Relationship Id="rId${relBase + tables.length + j}" Type="${REL}/hyperlink" Target="${escapeXml(u)}" TargetMode="External"/>`)].join("");
    entries.push([`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, utf8Bytes(`${HEAD}<Relationships xmlns="${PKG}">${relsXml}</Relationships>`)]);
    if (comments.length) entries.push([`xl/comments${i + 1}.xml`, utf8Bytes(commentsXml(comments))], [`xl/drawings/vmlDrawing${i + 1}.vml`, utf8Bytes(vmlXml(comments))]);
    for (const t of tables) entries.push([`xl/tables/table${t.id}.xml`, utf8Bytes(t.xml)]);
  });
  // What the file opened carried that is written back unchanged.
  for (const [name, bytes] of wb.carried) entries.push([name, bytes]);
  return { bytes: zip(entries), report };
}

// Sheet1!B2:B4 as Sheet1!$B$2:$B$4, as Excel writes a name's.
const absolute = (r: string): string => r.replace(/(^|[!:])\$?([A-Z]{1,3})\$?(\d+)/g, "$1$$$2$$$3");

function sheetOf(page: Page, wb: Workbook, sst: (s: string) => number, styleOf: (code: string, look: Look) => number, excel: boolean, dxfOf: (style: string) => number, nextTable: () => number) {
  const keys = [...new Set([...page.sheet.keys(), ...page.notes.keys()])].filter((k) => addrOf(k)).sort((a, b) => order(addrOf(a)!) - order(addrOf(b)!));
  const rows = new Map<number, string[]>();
  let sparks = 0;
  for (const k of keys) {
    const a = addrOf(k)!, input = page.sheet.input(k), note = page.notes.get(k);
    const look = wb.lookOf(note) ?? {}, code = note.format || look.number || "";
    const { number: _n, ...drawn } = look;
    const s = styleOf(/[#0?@]|general|[ymdhs]/i.test(code) ? code : "", JSON.parse(JSON.stringify(drawn)));
    if (input === "" && !s) continue;
    const v: Value = page.sheet.value(k), attrS = s ? ` s="${s}"` : "";
    let c: string;
    // Excel has no SPARKLINE function: a sparkline's cell is its text.
    const spark = input.startsWith("=") && /\bSPARKLINE\s*\(/i.test(input);
    if (spark) sparks++;
    if (input.startsWith("=") && excel && !spark) {
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
    if (spark && typeof v === "string") c = `<c r="${k}" t="s"${attrS}><v>${sst(v)}</v></c>`;
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
  // Data validation: the cells of each rule together.
  const byRule = new Map<string, { rule: Rule; keys: string[] }>();
  for (const k of keys) {
    const rule = page.notes.get(k).validation;
    if (!rule) continue;
    const id = JSON.stringify({ ...rule, origin: undefined });
    const g = byRule.get(id) ?? { rule, keys: [] };
    g.keys.push(k);
    byRule.set(id, g);
  }
  const dvs = [...byRule.values()].map(({ rule, keys: ks }) => validationXml(rule, ks)).filter((x) => x);
  const comments = keys.filter((k) => page.notes.get(k).comment).map((k): [string, string] => [k, page.notes.get(k).comment!]);
  const cf = cfXml(page.rules, dxfOf);
  const tables = page.sheet.tables().map((t) => {
    const left = Math.min(t.range[0].col, t.range[1].col), right = Math.max(t.range[0].col, t.range[1].col), top = Math.min(t.range[0].row, t.range[1].row);
    const columns = Array.from({ length: right - left + 1 }, (_, i) => t.columns?.[i] ?? (t.headers ? String(page.sheet.value(keyOf({ col: left + i, row: top })) ?? `Column${i + 1}`) : `Column${i + 1}`));
    const id = nextTable();
    return { id, xml: tableXml(t, id, columns) };
  });
  const relBase = comments.length ? 3 : 1;
  // Its hyperlinks: a URL by an external relationship (after the tables'),
  // a place in the workbook as its location.
  const urls: string[] = [];
  const links = keys.filter((k) => page.notes.get(k).hyperlink).map((k) => {
    const h = page.notes.get(k).hyperlink!, tip = h.tip ? ` tooltip="${escapeXml(h.tip)}"` : "";
    if (/^[a-z][a-z0-9+.-]*:/i.test(h.target)) { urls.push(h.target); return `<hyperlink ref="${k}" r:id="rId${relBase + tables.length + urls.length - 1}"${tip}/>`; }
    return `<hyperlink ref="${k}" location="${escapeXml(h.target.replace(/^#/, ""))}"${tip}/>`;
  });
  const xml = `${HEAD}<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>${cols ? `<cols>${cols}</cols>` : ""}<sheetData>${rowXml}</sheetData>${
    merges ? `<mergeCells count="${page.merges.length}">${merges}</mergeCells>` : ""}${cf}${dvs.length ? `<dataValidations count="${dvs.length}">${dvs.join("")}</dataValidations>` : ""}${links.length ? `<hyperlinks>${links.join("")}</hyperlinks>` : ""}${comments.length ? '<legacyDrawing r:id="rId2"/>' : ""}${
    tables.length ? `<tableParts count="${tables.length}">${tables.map((_, i) => `<tablePart r:id="rId${relBase + i}"/>`).join("")}</tableParts>` : ""}</worksheet>`;
  return { xml, comments, tables, relBase, sparks, urls };
}

// A rule as a dataValidation over its cells.
function validationXml(rule: Rule, keys: readonly string[]): string {
  const type = ({ list: "list", whole: "whole", decimal: "decimal", length: "textLength", date: "date", time: "time", custom: "custom" } as Record<string, string>)[rule.kind];
  if (!type) return "";
  const style = rule.mode === "warn" ? ' errorStyle="warning"' : rule.mode === "info" ? ' errorStyle="information"' : "";
  const msgs = `${rule.prompt ? ` showInputMessage="1" prompt="${escapeXml(rule.prompt)}"` : ""}${rule.error ? ` showErrorMessage="1" error="${escapeXml(rule.error)}"` : ' showErrorMessage="1"'}`;
  const f = (n: number, x: string) => `<formula${n}>${escapeXml(x)}</formula${n}>`;
  let op = "", body = "";
  if (rule.kind === "list") body = f(1, `"${rule.values!.join(",")}"`);
  else if (rule.kind === "custom") {
    // Written for the rule's top-left, the first cell here.
    body = f(1, rule.formula!);
  } else if (rule.min !== undefined && rule.max !== undefined) { op = ' operator="between"'; body = f(1, String(rule.min)) + f(2, String(rule.max)); }
  else if (rule.min !== undefined) { op = ' operator="greaterThanOrEqual"'; body = f(1, String(rule.min)); }
  else if (rule.max !== undefined) { op = ' operator="lessThanOrEqual"'; body = f(1, String(rule.max)); }
  else return "";
  return `<dataValidation type="${type}"${op}${style} allowBlank="1"${msgs} sqref="${rule.kind === "custom" && rule.origin ? [rule.origin, ...keys.filter((k) => k !== rule.origin)].join(" ") : keys.join(" ")}">${body}</dataValidation>`;
}

// A sheet's comments, as the legacy comments part and the VML drawing
// Excel shows them by.
function commentsXml(comments: readonly [string, string][]): string {
  return `${HEAD}<comments xmlns="${MAIN}"><authors><author>spraedsheet</author></authors><commentList>${
    comments.map(([k, t]) => `<comment ref="${k}" authorId="0"><text><r><t xml:space="preserve">${escapeXml(t)}</t></r></text></comment>`).join("")}</commentList></comments>`;
}
function vmlXml(comments: readonly [string, string][]): string {
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${
    comments.map(([k], i) => {
      const a = addrOf(k)!;
      return `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>${a.row}</x:Row><x:Column>${a.col}</x:Column></x:ClientData></v:shape>`;
    }).join("")}</xml>`;
}

