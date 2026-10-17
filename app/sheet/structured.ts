// Table references in Excel formulas (structured references): what is
// between a table reference's brackets, Sales[Amount], Sales[@Amount],
// Sales[[#Headers],[Qty]:[Price]], and the cells it comes to on the table
// as it is when the formula is worked out (its header row read then, so a
// column renamed is found by its new name at once).
//
//   Sales or Sales[]          its data rows
//   Sales[Amount]             one column's data
//   Sales[[Qty]:[Price]]      the columns from Qty to Price
//   Sales[#All], [#Data], [#Headers], [#Totals]   an area (#Totals is
//                             #REF! for a table with no totals row)
//   Sales[[#Headers],[#Data],[Qty]]   areas together, of a column
//   Sales[@Amount], Sales[@[Unit Price]], Sales[[#This Row],[Amount]]
//                             the formula's own row of a column
//
// A column's name is what its header cell says (or the names the table
// was given, or its letters), any case. Inside brackets ' makes the next
// character plain: Sales['#Units] is the column "#Units".
import { FormulaError } from "./lexer.ts";
import { colName } from "./address.ts";
import type { Addr } from "./address.ts";
import type { TableDef } from "./language.ts";
import { CellError } from "./values.ts";

export type Area = "all" | "data" | "headers" | "totals";

export interface TableSpec {
  // The areas, together (data alone when none was written).
  readonly areas: readonly Area[];
  // The formula's own row, not an area.
  readonly thisRow: boolean;
  // A column, or the first and last of a span of them; null for all.
  readonly columns: readonly [string, string] | null;
}

const AREAS: Readonly<Record<string, Area | "this row">> = {
  "#ALL": "all", "#DATA": "data", "#HEADERS": "headers", "#TOTALS": "totals", "#THIS ROW": "this row",
};

// A column's name as written: ' taken off what it makes plain.
const unescaped = (s: string): string => s.trim().replace(/'(.)/g, "$1");

// The items of a [..],[..] list, split at the commas between brackets.
function items(spec: string, at: number): string[] {
  const out: string[] = [];
  let depth = 0, from = 0;
  for (let j = 0; j < spec.length; j++) {
    const c = spec[j];
    if (c === "'") j++;
    else if (c === "[") depth++;
    else if (c === "]") depth--;
    else if (c === "," && depth === 0) {
      out.push(spec.slice(from, j).trim());
      from = j + 1;
    }
  }
  out.push(spec.slice(from).trim());
  if (out.some((x) => !x.startsWith("[") || !x.endsWith("]"))) throw new FormulaError(at, `a table reference's parts are each in brackets: [${spec}]`);
  return out;
}

// What is between a table reference's outer brackets; `at` is where the
// reference starts, for an error.
export function parseSpec(spec: string, at: number): TableSpec {
  const s = spec.trim();
  if (s === "") return { areas: ["data"], thisRow: false, columns: null };
  if (s.startsWith("@")) {
    const rest = s.slice(1).trim();
    const name = rest.startsWith("[") && rest.endsWith("]") ? rest.slice(1, -1) : rest;
    if (!name) throw new FormulaError(at, "@ needs a column: Sales[@Amount]");
    return { areas: [], thisRow: true, columns: [unescaped(name), unescaped(name)] };
  }
  const parts = s.startsWith("[") ? items(s, at) : [`[${s}]`];
  const areas: Area[] = [];
  let thisRow = false, columns: [string, string] | null = null;
  for (const part of parts) {
    // [A]:[B], a span of columns
    const span = /^\[((?:'.|[^\]])*)\]\s*:\s*\[((?:'.|[^\]])*)\]$/.exec(part);
    const inner = part.slice(1, -1).trim();
    const area = AREAS[inner.toUpperCase()];
    if (span || !area) {
      if (columns) throw new FormulaError(at, `a table reference names one column, or one span of them: [${spec}]`);
      columns = span ? [unescaped(span[1]), unescaped(span[2])] : [unescaped(inner), unescaped(inner)];
    } else if (area === "this row") thisRow = true;
    else if (!areas.includes(area)) areas.push(area);
  }
  if (thisRow && areas.length) throw new FormulaError(at, `#This Row goes with no other area: [${spec}]`);
  return { areas: thisRow || areas.length ? areas : ["data"], thisRow, columns };
}

// The cells a table reference comes to: its top-left and bottom-right, or
// the error it is. `header` reads a header cell's text; `row` is the
// formula's own row (null where there is none, as for a conditional
// style's expression).
export function resolve(t: TableDef, spec: TableSpec, header: (a: Addr) => string, row: number | null): [Addr, Addr] | CellError {
  const left = Math.min(t.range[0].col, t.range[1].col), right = Math.max(t.range[0].col, t.range[1].col);
  const top = Math.min(t.range[0].row, t.range[1].row), bottom = Math.max(t.range[0].row, t.range[1].row);
  const dataTop = t.headers ? top + 1 : top, totals = t.totals ? bottom : -1, dataBottom = t.totals ? bottom - 1 : bottom;
  let from: number, to: number;
  if (spec.thisRow) {
    if (row === null || row < dataTop || row > dataBottom) return new CellError("#VALUE!", `[@...] is the formula's own row of ${t.name}, and this formula is not on one of its rows`);
    from = to = row;
  } else {
    const rows = spec.areas.map((a): [number, number] | CellError => {
      switch (a) {
        case "all": return [top, bottom];
        case "data": return [dataTop, dataBottom];
        case "headers": return t.headers ? [top, top] : new CellError("#REF!", `table ${t.name} has no header row`);
        case "totals": return totals >= 0 ? [totals, totals] : new CellError("#REF!", `table ${t.name} has no totals row`);
      }
    });
    const bad = rows.find((r): r is CellError => r instanceof CellError);
    if (bad) return bad;
    from = Math.min(...(rows as [number, number][]).map((r) => r[0]));
    to = Math.max(...(rows as [number, number][]).map((r) => r[1]));
    if (from > to) return new CellError("#REF!", `table ${t.name} has no data rows`);
  }
  let first = left, last = right;
  if (spec.columns) {
    const names = Array.from({ length: right - left + 1 }, (_, i) =>
      (t.columns ? t.columns[i] ?? "" : t.headers ? header({ col: left + i, row: top }) : colName(left + i)).trim().toUpperCase());
    const at = (name: string): number | CellError => {
      const i = names.indexOf(name.toUpperCase());
      return i < 0 ? new CellError("#REF!", `table ${t.name} has no column ${name}`) : left + i;
    };
    const [a, b] = [at(spec.columns[0]), at(spec.columns[1])];
    if (a instanceof CellError) return a;
    if (b instanceof CellError) return b;
    [first, last] = [Math.min(a, b), Math.max(a, b)];
  }
  return [{ col: first, row: from }, { col: last, row: to }];
}

// A table's data rows, first and last (0-based): under its header row,
// over its totals row.
export function dataRows(t: TableDef): [number, number] {
  const top = Math.min(t.range[0].row, t.range[1].row), bottom = Math.max(t.range[0].row, t.range[1].row);
  return [t.headers ? top + 1 : top, t.totals ? bottom - 1 : bottom];
}

// Excel's SUBTOTAL numbers for a totals row's aggregates (the 100s skip
// hidden rows, as a table's totals do).
export const TOTALS: Readonly<Record<string, number>> = { average: 101, count: 102, counta: 103, max: 104, min: 105, product: 106, stdev: 107, stdevp: 108, sum: 109, var: 110, varp: 111 };
const REXX_TOTALS: Readonly<Record<string, string>> = { average: "AVERAGE_RANGE", count: "COUNT_RANGE", counta: "COUNT_RANGE", max: "MAX_RANGE", min: "MIN_RANGE", product: "PRODUCT_RANGE",
  stdev: "STDEV_RANGE", stdevp: "STDEVP_RANGE", sum: "SUM_RANGE", var: "VAR_RANGE", varp: "VARP_RANGE" };

// A totals cell's input: Excel's SUBTOTAL over the column (Sales[Amount]),
// or in RexxJS formulas its range function over the column's cells.
export function totalsInput(aggregate: string, table: string, column: string, cells: string, excel: boolean): string {
  if (aggregate === "label") return "Total";
  if (aggregate === "none") return "";
  const plain = /^[A-Za-z_][A-Za-z0-9_ ]*$/.test(column) && !column.includes(" ") ? column : `[${column.replace(/([\[\]#'])/g, "'$1")}]`;
  return excel ? `=SUBTOTAL(${TOTALS[aggregate]},${table}[${plain}])` : `=${REXX_TOTALS[aggregate]}("${cells}")`;
}

// The structured references in a formula written as the cells they come to
// (Convert to Range): $B$2:$B$4, or a row's own $B2; strings left alone.
export function unstructured(input: string, table: TableDef, header: (a: Addr) => string, row: number): string {
  if (!input.startsWith("=")) return input;
  let out = "", i = 0;
  const name = table.name.toUpperCase();
  while (i < input.length) {
    const c = input[i];
    if (c === '"') {
      const end = input.indexOf('"', i + 1);
      const j = end < 0 ? input.length : end + 1;
      out += input.slice(i, j);
      i = j;
      continue;
    }
    const m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(input.slice(i));
    if (m && input[i + m[0].length] === "[" && m[0].toUpperCase() === name && !/[A-Za-z0-9_.$!]/.test(input[i - 1] ?? "")) {
      // The brackets, balanced, ' making the next character plain.
      let j = i + m[0].length, depth = 0;
      for (; j < input.length; j++) {
        if (input[j] === "'") { j++; continue; }
        if (input[j] === "[") depth++;
        else if (input[j] === "]" && --depth === 0) break;
      }
      const inner = input.slice(i + m[0].length + 1, j);
      let written = "#REF!";
      try {
        const spec = parseSpec(inner, i);
        const r = resolve(table, spec, header, row);
        if (!(r instanceof CellError)) {
          const [a, b] = r, abs = (x: Addr) => (spec.thisRow ? `$${colName(x.col)}${x.row + 1}` : `$${colName(x.col)}$${x.row + 1}`);
          written = a.col === b.col && a.row === b.row ? abs(a) : `${abs(a)}:${abs(b)}`;
        }
      } catch (e) {
        if (!(e instanceof FormulaError)) throw e;
      }
      out += written;
      i = j + 1;
      continue;
    }
    if (m) { out += m[0]; i += m[0].length; continue; }
    out += c;
    i++;
  }
  return out;
}
