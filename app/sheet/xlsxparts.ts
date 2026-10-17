// XLSX's conditional formatting and tables, both ways (xlsx.ts reads a
// file, xlsxwrite.ts writes one), and the parts of a file the app does
// not model: reported, and the few it can carry kept (pass-through).
// Conditional formatting: each cfRule as rules.ts's rule, its differential
// style (a dxf: font colour, bold, italic, fill) as the rule's style; a
// colour scale's, data bar's and icon set's own colours and icons. Tables:
// a table part as the sheet's table (its name, range, header row, totals
// row and each column's function, its style and stripes).
import { attr, all, kid, kids, textOf, escapeXml } from "./xml.ts";
import type { El } from "./xml.ts";
import { parseRange } from "./address.ts";
import { rangeName } from "./rules.ts";
import type { CfRule } from "./rules.ts";
import { colour } from "./formats.ts";
import type { TableDef } from "./language.ts";

const EMPTY: El = { name: "", attrs: {}, children: [] };
const THEME_ROLES = ["light1", "dark1", "light2", "dark2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink"];
const colourOf = (e: El | undefined): string | undefined => {
  if (!e) return undefined;
  const rgb = attr(e, "rgb"), theme = attr(e, "theme");
  if (rgb && /^[0-9a-f]{6,8}$/i.test(rgb)) return `#${rgb.slice(-6).toLowerCase()}`;
  if (theme !== undefined) return THEME_ROLES[Number(theme)];
  return undefined;
};
const colourXml = (tag: string, c: string): string => {
  const k = c.trim().toLowerCase();
  return THEME_ROLES.includes(k) ? `<${tag} theme="${THEME_ROLES.indexOf(k)}"/>` : `<${tag} rgb="FF${(colour(k) ?? "#000000").replace("#", "").toUpperCase()}"/>`;
};

// A dxf (a differential style) as the grid's style.
export function dxfStyle(dxf: El | undefined): string {
  if (!dxf) return "";
  const parts: string[] = [];
  const font = kid(dxf, "font"), fill = kid(dxf, "fill");
  const fc = font && colourOf(kid(font, "color"));
  const pf = fill && kid(fill, "patternFill"), bg = pf && (colourOf(kid(pf, "bgColor")) ?? colourOf(kid(pf, "fgColor")));
  if (bg) parts.push(`background: ${bg}`);
  if (fc) parts.push(`color: ${fc}`);
  if (font && kid(font, "b") && attr(kid(font, "b")!, "val") !== "0") parts.push("font-weight: bold");
  if (font && kid(font, "i") && attr(kid(font, "i")!, "val") !== "0") parts.push("font-style: italic");
  if (font && kid(font, "strike")) parts.push("text-decoration: line-through");
  else if (font && kid(font, "u")) parts.push("text-decoration: underline");
  return parts.join("; ");
}

// A rule's style as a dxf.
export function styleDxf(style: string): string {
  const d = new Map(style.split(";").map((x) => x.trim()).filter((x) => x.includes(":")).map((x) => [x.slice(0, x.indexOf(":")).trim(), x.slice(x.indexOf(":") + 1).trim()]));
  const font = `${d.get("font-weight") === "bold" ? "<b/>" : ""}${d.get("font-style") === "italic" ? "<i/>" : ""}${d.get("text-decoration")?.includes("line-through") ? "<strike/>" : ""}${d.get("text-decoration")?.includes("underline") ? "<u/>" : ""}${d.get("color") ? colourXml("color", d.get("color")!) : ""}`;
  const bg = d.get("background") ?? d.get("background-color");
  return `<dxf>${font ? `<font>${font}</font>` : ""}${bg ? `<fill><patternFill>${colourXml("bgColor", bg)}</patternFill></fill>` : ""}</dxf>`;
}

const OPS: Record<string, string> = { greaterThan: ">", lessThan: "<", equal: "=", notEqual: "<>", greaterThanOrEqual: ">=", lessThanOrEqual: "<=", between: "between", notBetween: "notBetween" };
const PERIODS: Record<string, string> = { yesterday: "yesterday", today: "today", tomorrow: "tomorrow", last7Days: "last7days", lastWeek: "lastWeek", thisWeek: "thisWeek", nextWeek: "nextWeek", lastMonth: "lastMonth", thisMonth: "thisMonth", nextMonth: "nextMonth" };
const ICON_SETS: Record<string, string> = { "3Arrows": "arrows", "3ArrowsGray": "arrows", "3Triangles": "triangles", "3TrafficLights1": "ratings", "3TrafficLights2": "ratings", "3Signs": "ratings", "3Symbols": "ratings", "3Symbols2": "ratings", "3Flags": "ratings", "3Stars": "ratings" };

// A formula's text as the value a rule compares: "text" as text.
const operand = (f: string): string => (/^".*"$/.test(f) ? f.slice(1, -1).replace(/""/g, '"') : f);

// A sheet's conditional formatting as the app's rules, by priority; what
// is not one, said in `report`.
export function cfRulesOf(ws: El, dxfs: readonly El[], sheet: string, report: string[]): Record<string, unknown>[] {
  const out: { priority: number; rule: Record<string, unknown> }[] = [];
  for (const cf of all(ws, "conditionalFormatting")) {
    const sqref = (attr(cf, "sqref") ?? "").split(/\s+/).filter((x) => x);
    for (const r of kids(cf, "cfRule")) {
      const type = attr(r, "type") ?? "", style = dxfStyle(dxfs[Number(attr(r, "dxfId") ?? -1)]) || "background: #ffc7ce; color: #9c0006";
      const fs = kids(r, "formula").map((f) => textOf(f).trim()), stop = attr(r, "stopIfTrue") === "1" ? { stop: true } : {};
      let rule: Record<string, unknown> | null = null;
      switch (type) {
        case "cellIs": {
          const op = OPS[attr(r, "operator") ?? ""];
          if (op) rule = { kind: "cellValue", op, value: operand(fs[0] ?? ""), ...(fs[1] !== undefined ? { value2: operand(fs[1]) } : {}), style };
          break;
        }
        case "containsText": rule = { kind: "cellValue", op: "contains", value: attr(r, "text") ?? "", style }; break;
        case "notContainsText": rule = { kind: "cellValue", op: "notContains", value: attr(r, "text") ?? "", style }; break;
        case "beginsWith": rule = { kind: "cellValue", op: "beginsWith", value: attr(r, "text") ?? "", style }; break;
        case "endsWith": rule = { kind: "cellValue", op: "endsWith", value: attr(r, "text") ?? "", style }; break;
        case "expression": rule = { kind: "formula", formula: fs[0] ?? "", style }; break;
        case "top10": rule = { kind: "top", n: Number(attr(r, "rank") ?? 10), bottom: attr(r, "bottom") === "1", percent: attr(r, "percent") === "1", style }; break;
        case "aboveAverage": rule = { kind: "average", below: attr(r, "aboveAverage") === "0", style }; break;
        case "duplicateValues": rule = { kind: "duplicates", unique: false, style }; break;
        case "uniqueValues": rule = { kind: "duplicates", unique: true, style }; break;
        case "containsBlanks": rule = { kind: "blanks", not: false, style }; break;
        case "notContainsBlanks": rule = { kind: "blanks", not: true, style }; break;
        case "containsErrors": rule = { kind: "errors", not: false, style }; break;
        case "notContainsErrors": rule = { kind: "errors", not: true, style }; break;
        case "timePeriod": if (PERIODS[attr(r, "timePeriod") ?? ""]) rule = { kind: "date", period: PERIODS[attr(r, "timePeriod")!], style }; break;
        case "colorScale": {
          const colors = kids(kid(r, "colorScale") ?? EMPTY, "color").map(colourOf).filter((c): c is string => !!c && c.startsWith("#"));
          if (colors.length >= 2) rule = { kind: "colorScale", colors };
          break;
        }
        case "dataBar": rule = { kind: "dataBar", color: colourOf(kid(kid(r, "dataBar") ?? EMPTY, "color")) ?? "#638ec6" }; break;
        case "iconSet": {
          const is = kid(r, "iconSet") ?? EMPTY;
          rule = { kind: "iconSet", icons: ICON_SETS[attr(is, "iconSet") ?? "3TrafficLights1"] ?? "arrows", ...(attr(is, "reverse") === "1" ? { reverse: true } : {}), ...(attr(is, "showValue") === "0" ? { iconOnly: true } : {}) };
          break;
        }
      }
      if (!rule) { report.push(`${sheet}!${sqref.join(" ")}: a ${type} rule, left out`); continue; }
      for (const area of sqref) {
        const r2 = parseRange(area.replace(/\$/g, ""));
        if (r2) out.push({ priority: Number(attr(r, "priority") ?? 0), rule: { ...rule, ...stop, range: rangeName(r2) } });
      }
    }
  }
  return out.sort((a, b) => a.priority - b.priority).map((x) => x.rule);
}

// A sheet's rules as conditionalFormatting elements; dxf(style) gives a
// style's dxf index.
export function cfXml(rules: readonly CfRule[], dxf: (style: string) => number): string {
  const op = Object.fromEntries(Object.entries(OPS).map(([k, v]) => [v, k]));
  const period = Object.fromEntries(Object.entries(PERIODS).map(([k, v]) => [v, k]));
  const text = (v: string) => (/^-?[\d.]+(E[-+]?\d+)?$/i.test(v) ? v : `"${v.replace(/"/g, '""')}"`);
  return rules.map((r, i) => {
    const ref = rangeName(r.range), p = ` priority="${i + 1}"${r.stop ? ' stopIfTrue="1"' : ""}`;
    const st = "style" in r ? ` dxfId="${dxf(r.style)}"` : "";
    let body = "";
    switch (r.kind) {
      case "cellValue":
        if (r.op === "contains" || r.op === "notContains" || r.op === "beginsWith" || r.op === "endsWith") {
          const t = { contains: "containsText", notContains: "notContainsText", beginsWith: "beginsWith", endsWith: "endsWith" }[r.op];
          body = `<cfRule type="${t}"${st}${p} operator="${t === "containsText" ? "containsText" : t === "notContainsText" ? "notContains" : t}" text="${escapeXml(r.value)}"/>`;
        } else body = `<cfRule type="cellIs"${st}${p} operator="${op[r.op]}"><formula>${escapeXml(text(r.value))}</formula>${r.value2 !== undefined ? `<formula>${escapeXml(text(r.value2))}</formula>` : ""}</cfRule>`;
        break;
      case "formula": body = `<cfRule type="expression"${st}${p}><formula>${escapeXml(r.formula)}</formula></cfRule>`; break;
      case "top": body = `<cfRule type="top10"${st}${p} rank="${r.n}"${r.bottom ? ' bottom="1"' : ""}${r.percent ? ' percent="1"' : ""}/>`; break;
      case "average": body = `<cfRule type="aboveAverage"${st}${p}${r.below ? ' aboveAverage="0"' : ""}/>`; break;
      case "duplicates": body = `<cfRule type="${r.unique ? "uniqueValues" : "duplicateValues"}"${st}${p}/>`; break;
      case "blanks": body = `<cfRule type="${r.not ? "notContainsBlanks" : "containsBlanks"}"${st}${p}/>`; break;
      case "errors": body = `<cfRule type="${r.not ? "notContainsErrors" : "containsErrors"}"${st}${p}/>`; break;
      case "date": body = `<cfRule type="timePeriod"${st}${p} timePeriod="${period[r.period]}"/>`; break;
      case "colorScale": body = `<cfRule type="colorScale"${p}><colorScale><cfvo type="min"/>${r.colors.length === 3 ? '<cfvo type="percentile" val="50"/>' : ""}<cfvo type="max"/>${r.colors.map((c) => colourXml("color", c)).join("")}</colorScale></cfRule>`; break;
      case "dataBar": body = `<cfRule type="dataBar"${p}><dataBar><cfvo type="min"/><cfvo type="max"/>${colourXml("color", r.color)}</dataBar></cfRule>`; break;
      case "iconSet": body = `<cfRule type="iconSet"${p}><iconSet iconSet="${r.icons === "arrows" ? "3Arrows" : r.icons === "triangles" ? "3Triangles" : "3TrafficLights1"}"${r.reverse ? ' reverse="1"' : ""}${r.iconOnly ? ' showValue="0"' : ""}><cfvo type="percent" val="0"/><cfvo type="percent" val="${r.thresholds?.[1] ?? 33}"/><cfvo type="percent" val="${r.thresholds?.[0] ?? 67}"/></iconSet></cfRule>`; break;
    }
    return `<conditionalFormatting sqref="${ref}">${body}</conditionalFormatting>`;
  }).join("");
}

const FUNCTIONS: Record<string, string> = { sum: "sum", average: "average", count: "countNums", counta: "count", max: "max", min: "min", stdev: "stdDev", var: "var" };

// A table part as the app's table (file.ts's tables).
export function tableOf(t: El): Record<string, unknown> | null {
  const name = attr(t, "displayName") ?? attr(t, "name"), ref = attr(t, "ref");
  const r = ref && parseRange(ref.replace(/\$/g, ""));
  if (!name || !r) return null;
  const totals = Number(attr(t, "totalsRowCount") ?? 0) > 0, headers = attr(t, "headerRowCount") !== "0";
  const cols = kids(kid(t, "tableColumns") ?? EMPTY, "tableColumn");
  const back = Object.fromEntries(Object.entries(FUNCTIONS).map(([k, v]) => [v, k]));
  const info = kid(t, "tableStyleInfo") ?? EMPTY, sn = attr(info, "name") ?? "";
  const style = /Light/.test(sn) ? "Light" : /Medium/.test(sn) ? "Medium" : /Dark/.test(sn) ? "Dark" : undefined;
  return {
    name, range: rangeName(r), headers, columns: headers ? null : cols.map((c) => attr(c, "name") ?? ""),
    ...(totals ? { totals: Object.fromEntries(cols.map((c, i) => [attr(c, "name") ?? `Column${i + 1}`, attr(c, "totalsRowLabel") ? "label" : back[attr(c, "totalsRowFunction") ?? ""] ?? "none"])) } : {}),
    ...(style ? { style } : {}), ...(attr(info, "showRowStripes") === "1" ? { banded: true } : {}), ...(attr(info, "showColumnStripes") === "1" ? { bandedColumns: true } : {}),
    ...(attr(info, "showFirstColumn") === "1" ? { firstColumn: true } : {}), ...(attr(info, "showLastColumn") === "1" ? { lastColumn: true } : {}),
  };
}

// A table as a table part (id its number in the workbook, columns its
// columns' names as the header row has them).
export function tableXml(t: TableDef, id: number, columns: readonly string[]): string {
  const ref = rangeName(t.range), data = t.totals ? ref.replace(/\d+$/, (n) => String(Number(n) - 1)) : ref;
  const style = { Light: "TableStyleLight9", Medium: "TableStyleMedium2", Dark: "TableStyleDark1", Green: "TableStyleMedium7", Orange: "TableStyleMedium3", Grey: "TableStyleMedium1" }[t.style ?? ""] ?? "TableStyleLight1";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${escapeXml(t.name)}" displayName="${escapeXml(t.name)}" ref="${ref}"${t.headers ? "" : ' headerRowCount="0"'}${t.totals ? ' totalsRowCount="1"' : ""}>${
    t.headers ? `<autoFilter ref="${data}"/>` : ""}<tableColumns count="${columns.length}">${columns.map((c, i) => {
      const f = t.totals?.[c];
      return `<tableColumn id="${i + 1}" name="${escapeXml(c)}"${f === "label" ? ' totalsRowLabel="Total"' : f && FUNCTIONS[f] ? ` totalsRowFunction="${FUNCTIONS[f]}"` : ""}/>`;
    }).join("")}</tableColumns><tableStyleInfo name="${style}" showFirstColumn="${t.firstColumn ? 1 : 0}" showLastColumn="${t.lastColumn ? 1 : 0}" showRowStripes="${t.banded ? 1 : 0}" showColumnStripes="${t.bandedColumns ? 1 : 0}"/></table>`;
}

// The parts of an archive the app does not model, said: charts, drawings
// of pictures, pivots, external links, macros, x14 sparklines and slicers.
export function unmodelled(names: readonly string[], sheetXml: readonly string[]): string[] {
  const out: string[] = [];
  const count = (re: RegExp) => names.filter((n) => re.test(n)).length;
  const said = (n: number, what: string) => { if (n) out.push(`${n} ${what}${n === 1 ? "" : "s"} left out`); };
  said(count(/^xl\/charts\/chart\d+\.xml$/), "chart");
  said(count(/^xl\/media\//), "picture");
  said(count(/^xl\/pivotTables\/pivotTable\d+\.xml$/), "pivot table");
  said(count(/^xl\/externalLinks\/externalLink\d+\.xml$/), "external link");
  said(count(/^xl\/vbaProject\.bin$/), "macro project");
  said(count(/^xl\/slicers\//), "slicer");
  said(sheetXml.filter((x) => /sparklineGroup/.test(x)).length, "sheet's sparkline group");
  return out;
}

// What a written file carries over from the one opened (pass-through):
// its document properties and its theme, unchanged.
export const PASS_THROUGH = /^(docProps\/(core|app|custom)\.xml|xl\/theme\/theme1\.xml)$/;
