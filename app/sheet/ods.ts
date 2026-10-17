// ODS (OpenDocument spreadsheet) read: content.xml's tables as sheets, each
// cell's cached value (a float, percentage, currency, date, time, boolean
// or string; repeated rows and columns expanded), its formula left as the
// value it came to, as the step asks; merged cells (spanned). Written as
// the app's own sheet file, with a report of what was left out. Pure.
import { unzip, utf8Text } from "./zip.ts";
import { attr, kid, kids, parseXml, textOf } from "./xml.ts";
import type { El } from "./xml.ts";
import { COLS, ROWS, colName, keyOf } from "./address.ts";
import { serial } from "./dates.ts";

const NUMERIC = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/;

// An ISO date or date-time as a serial number; a duration (PT1H30M) as a
// fraction of a day.
function dateSerial(iso: string): number | null {
  const m = /^(-?\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d(?:\.\d+)?))?)?/.exec(iso);
  if (!m) return null;
  const day = serial(Number(m[1]), Number(m[2]), Number(m[3]));
  return day + (m[4] ? (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6] ?? 0)) / 86400 : 0);
}
function durationDays(d: string): number | null {
  const m = /^-?P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?)?$/.exec(d);
  if (!m) return null;
  return (Number(m[1] ?? 0) * 86400 + Number(m[2] ?? 0) * 3600 + Number(m[3] ?? 0) * 60 + Number(m[4] ?? 0)) / 86400 * (d.startsWith("-") ? -1 : 1);
}

// A cell's text: its paragraphs, a line apiece, spaces (text:s) kept.
function cellText(c: El): string {
  return kids(c, "p").map((p) => p.children.map((x) => (typeof x === "string" ? x : x.name.endsWith(":s") ? " ".repeat(Number(attr(x, "c") ?? 1)) : x.name.endsWith(":line-break") ? "\n" : x.name.endsWith(":tab") ? "\t" : textOf(x))).join("")).join("\n");
}

export function readOds(bytes: Uint8Array, fileName = "Workbook"): { file: Record<string, unknown>; report: string[] } {
  const parts = unzip(bytes), report: string[] = [];
  const content = parts.get("content.xml");
  if (!content) throw new RangeError("not an ODS spreadsheet: no content.xml");
  const doc = parseXml(utf8Text(content));
  const body = kid(kid(doc, "body") ?? doc, "spreadsheet");
  if (!body) throw new RangeError("not an ODS spreadsheet: no office:spreadsheet");
  const sheets: Record<string, unknown>[] = [];
  let off = 0, formulas = 0;
  for (const t of kids(body, "table")) {
    const name = attr(t, "name") ?? `Sheet${sheets.length + 1}`, cells: Record<string, unknown> = {}, merges: string[] = [];
    let row = 0;
    const rows: El[] = [];
    const collect = (e: El) => { for (const c of kids(e)) { const n = c.name.split(":").pop(); if (n === "table-row") rows.push(c); else if (n === "table-rows" || n === "table-header-rows" || n === "table-row-group") collect(c); } };
    collect(t);
    for (const r of rows) {
      const repeat = Math.min(Number(attr(r, "number-rows-repeated") ?? 1), ROWS + 1);
      for (let rr = 0; rr < repeat; rr++, row++) {
        let col = 0;
        for (const c of kids(r).filter((x) => /table-cell$|covered-table-cell$/.test(x.name))) {
          const rep = Math.min(Number(attr(c, "number-columns-repeated") ?? 1), COLS + 1);
          const type = attr(c, "value-type"), covered = c.name.endsWith("covered-table-cell");
          let input = "";
          if (!covered) {
            if (type === "float" || type === "percentage" || type === "currency") input = attr(c, "value") ?? "";
            else if (type === "date") { const d = dateSerial(attr(c, "date-value") ?? ""); input = d === null ? cellText(c) : String(d); }
            else if (type === "time") { const d = durationDays(attr(c, "time-value") ?? ""); input = d === null ? cellText(c) : String(d); }
            else if (type === "boolean") input = attr(c, "boolean-value") === "true" ? "TRUE" : "FALSE";
            else if (type === "string") { input = attr(c, "string-value") ?? cellText(c); if (NUMERIC.test(input) || input.startsWith("=")) input = `'${input}`; }
            else input = cellText(c);
            if (attr(c, "formula")) formulas++;
          }
          for (let k = 0; k < rep; k++, col++) {
            if (input === "") continue;
            if (row >= ROWS || col >= COLS) { off++; continue; }
            const key = keyOf({ col, row });
            const fmt = type === "percentage" ? "0%" : type === "date" ? (/T/.test(attr(c, "date-value") ?? "") ? "yyyy-mm-dd hh:mm" : "yyyy-mm-dd") : type === "time" ? "[h]:mm:ss" : "";
            cells[key] = fmt ? { content: input, look: { number: fmt } } : input;
            const cs = Number(attr(c, "number-columns-spanned") ?? 1), rs = Number(attr(c, "number-rows-spanned") ?? 1);
            if ((cs > 1 || rs > 1) && k === 0) merges.push(`${key}:${keyOf({ col: Math.min(COLS - 1, col + cs - 1), row: Math.min(ROWS - 1, row + rs - 1) })}`);
          }
          if (input === "" && !covered) {
            const cs = Number(attr(c, "number-columns-spanned") ?? 1), rs = Number(attr(c, "number-rows-spanned") ?? 1);
            if ((cs > 1 || rs > 1) && row < ROWS && col - rep < COLS) merges.push(`${keyOf({ col: col - rep, row })}:${keyOf({ col: Math.min(COLS - 1, col - rep + cs - 1), row: Math.min(ROWS - 1, row + rs - 1) })}`);
          }
        }
      }
    }
    sheets.push({ name, cells, merges, view: {} });
  }
  if (!sheets.length) throw new RangeError("an ODS spreadsheet with no tables");
  if (off) report.push(`${off} cell${off === 1 ? "" : "s"} past ${colName(COLS - 1)}${ROWS}, the sheet's edge, left out`);
  if (formulas) report.push(`${formulas} formula${formulas === 1 ? "" : "s"} read as the value${formulas === 1 ? "" : "s"} cached`);
  const [first, ...rest] = sheets;
  const name = fileName.replace(/\.ods$/i, "");
  return { file: { formulas: "excel", sheetName: first.name, ...first, sheets: rest, name }, report };
}
