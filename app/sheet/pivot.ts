// Pivot tables, as the Tauri spreadsheet's CREATEPIVOT makes them: the
// rows of a source range (its first row their fields' names) grouped by
// some fields down the side and others across the top, one field's numbers
// aggregated where they meet. No sheet, no ui: a grid of values in, the
// pivot's cells out (commands.ts writes them, and writes them again on
// pivot.update).
import { CellError } from "./values.ts";
import type { Value } from "./values.ts";

export const AGGREGATES = ["SUM", "COUNT", "AVERAGE", "MIN", "MAX"] as const;
export type Aggregate = (typeof AGGREGATES)[number];

export interface PivotFields {
  readonly rows: readonly string[];       // grouped down the side
  readonly cols: readonly string[];       // grouped across the top
  readonly value: string;                 // aggregated
  readonly agg: Aggregate;
  readonly filters: Readonly<Record<string, string>>;   // a field's value a row must have
  readonly options?: PivotOptions;
}

// Excel's PivotTable options, past the Tauri app's (whose pivots have none
// of them, so each is off when left out): the report's layout (tabular,
// each row field its own column and every row labelled; outline, a label
// once for its group; compact, the row fields one column indented),
// subtotals for each outer row group, grand totals (a row, and a column
// when columns are grouped), the rows sorted by their labels or by their
// totals, Show Values As, a date field grouped by year, quarter or month,
// calculated fields (a formula of other fields, worked out on their
// sums, as Excel's), and outer row groups collapsed to their subtotals.
export interface PivotOptions {
  readonly layout?: "tabular" | "outline" | "compact";
  readonly subtotals?: boolean;
  readonly grandTotals?: boolean;
  readonly sort?: "asc" | "desc" | "valueAsc" | "valueDesc";
  readonly showAs?: ShowAs;
  readonly dateGroup?: { readonly field: string; readonly by: "year" | "quarter" | "month" };
  readonly calculated?: Readonly<Record<string, string>>;
  readonly collapsed?: readonly string[];
}
export const SHOW_AS = ["normal", "percentOfTotal", "percentOfRow", "percentOfColumn", "runningTotal", "rank"] as const;
export type ShowAs = (typeof SHOW_AS)[number];
export const LAYOUTS = ["tabular", "outline", "compact"];

const text = (v: Value) => (v === null ? "" : v instanceof CellError ? v.code : String(v));

// A field's column in the source, by its name in the first row; a
// RangeError naming the fields there are.
function fieldIndex(headers: readonly string[], field: string): number {
  const i = headers.indexOf(field);
  if (i < 0) throw new RangeError(`no field ${field} (${headers.join(", ")})`);
  return i;
}

function aggregate(agg: Aggregate, ns: readonly number[]): number {
  if (!ns.length) return 0;
  switch (agg) {
    case "SUM": return ns.reduce((a, b) => a + b, 0);
    case "COUNT": return ns.length;
    case "AVERAGE": return ns.reduce((a, b) => a + b, 0) / ns.length;
    case "MIN": return Math.min(...ns);
    case "MAX": return Math.max(...ns);
  }
}

// A calculated field's formula on the fields' sums: + - * / and brackets
// over numbers and field names (a name with spaces in 'quotes').
export function calculate(formula: string, sums: (field: string) => number): number {
  const toks = formula.match(/'[^']*'|\d+\.?\d*(?:[eE][-+]?\d+)?|[A-Za-z_][A-Za-z0-9_ ]*[A-Za-z0-9_]|[A-Za-z_]|[-+*/()]/g) ?? [];
  let i = 0;
  const atom = (): number => {
    const t = toks[i++];
    if (t === "(") { const v = sum(); i++; return v; }
    if (t === "-") return -atom();
    if (t === undefined) throw new RangeError(`${formula} is not a formula of fields`);
    if (/^\d/.test(t)) return Number(t);
    return sums(t.startsWith("'") ? t.slice(1, -1) : t.trim());
  };
  const product = (): number => { let v = atom(); while (toks[i] === "*" || toks[i] === "/") { const op = toks[i++], b = atom(); v = op === "*" ? v * b : b === 0 ? NaN : v / b; } return v; };
  const sum = (): number => { let v = product(); while (toks[i] === "+" || toks[i] === "-") { const op = toks[i++], b = product(); v = op === "+" ? v + b : v - b; } return v; };
  return sum();
}

const QUARTERS = ["Q1", "Q2", "Q3", "Q4"], MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
// A serial date's group: its year, quarter or month (text that is not a
// date as it is).
function dateLabel(v: Value, by: "year" | "quarter" | "month"): string {
  if (typeof v !== "number") return text(v);
  const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000), y = d.getUTCFullYear(), m = d.getUTCMonth();
  return by === "year" ? String(y) : by === "quarter" ? `${y} ${QUARTERS[Math.floor(m / 3)]}` : `${y} ${MONTH_NAMES[m]}`;
}

const num = (n: number): string => (Number.isFinite(n) ? String(Number(n.toPrecision(12))) : "#DIV/0!");

// The pivot as rows of inputs: a header row (the row fields' names, then
// each column group, or "SUM of amount" when there are no column fields),
// then a row per row group (its values, then each aggregate; 0 where no row
// met). Groups in the order they first appear, as the app's; then the
// options' layout, totals, order and Show Values As.
export function pivoted(grid: readonly (readonly Value[])[], f: PivotFields): string[][] {
  const o = f.options ?? {};
  const headers = (grid[0] ?? []).map((v, i) => text(v) || `Column${i + 1}`);
  const calc = o.calculated ?? {};
  const isCalc = (n: string) => Object.keys(calc).some((k) => k.toLowerCase() === n.toLowerCase());
  const rowAt = f.rows.map((r) => fieldIndex(headers, r)), colAt = f.cols.map((c) => fieldIndex(headers, c));
  const valueAt = isCalc(f.value) ? -1 : fieldIndex(headers, f.value);
  const filters = Object.entries(f.filters).map(([field, want]) => [fieldIndex(headers, field), want] as const);
  const dg = o.dateGroup, dgAt = dg ? fieldIndex(headers, dg.field) : -1;
  const label = (row: readonly Value[], i: number) => (i === dgAt ? dateLabel(row[i] ?? null, dg!.by) : text(row[i] ?? null));
  const rowKeys = new Map<string, string[]>(), colKeys = new Map<string, string[]>();
  // Each meeting's rows, kept so a calculated field can sum its fields.
  const met = new Map<string, (readonly Value[])[]>();
  for (const row of grid.slice(1)) {
    if (filters.some(([i, want]) => text(row[i] ?? null) !== want)) continue;
    const rv = rowAt.map((i) => label(row, i)), cv = colAt.map((i) => label(row, i));
    const rk = JSON.stringify(rv), ck = JSON.stringify(cv);
    if (!rowKeys.has(rk)) rowKeys.set(rk, rv);
    if (!colKeys.has(ck)) colKeys.set(ck, cv);
    const at = `${rk}\u0000${ck}`;
    if (!met.has(at)) met.set(at, []);
    met.get(at)!.push(row);
  }
  // A set of rows' value: the aggregate of the value field's numbers, or a
  // calculated field's formula on its fields' sums.
  const valueOf = (rows: readonly (readonly Value[])[]): number => {
    if (valueAt < 0) {
      const formula = Object.entries(calc).find(([k]) => k.toLowerCase() === f.value.toLowerCase())![1];
      return calculate(formula, (field) => rows.reduce((s, r) => { const n = parseFloat(text(r[fieldIndex(headers, field)] ?? null)); return s + (Number.isNaN(n) ? 0 : n); }, 0));
    }
    const ns = rows.map((r) => parseFloat(text(r[valueAt] ?? null))).filter((n) => !Number.isNaN(n));
    return aggregate(f.agg, ns);
  };
  const cols = [...colKeys.keys()];
  let rows = [...rowKeys.keys()];
  const cellRows = (rk: string, ck: string) => met.get(`${rk}\u0000${ck}`) ?? [];
  const rowAll = (rk: string) => cols.flatMap((ck) => cellRows(rk, ck));
  if (o.sort) {
    const by = (rk: string) => (o.sort === "asc" || o.sort === "desc" ? rowKeys.get(rk)!.join("\u0000") : valueOf(rowAll(rk)));
    rows = [...rows].sort((a, b) => { const x = by(a), y = by(b); const c = typeof x === "number" ? (x as number) - (y as number) : String(x).localeCompare(String(y), undefined, { numeric: true }); return o.sort === "desc" || o.sort === "valueDesc" ? -c : c; });
  }
  // Several row fields: each outer group's rows together, the groups in
  // the order they first come.
  if (f.rows.length > 1) {
    const order: string[] = [];
    for (const rk of rows) { const outer = rowKeys.get(rk)![0]; if (!order.includes(outer)) order.push(outer); }
    rows = order.flatMap((outer) => rows.filter((rk) => rowKeys.get(rk)![0] === outer));
  }
  // The numbers, before Show Values As.
  const raw = rows.map((rk) => cols.map((ck) => (cellRows(rk, ck).length ? valueOf(cellRows(rk, ck)) : 0)));
  const grand = valueOf(rows.flatMap(rowAll)), rowTotals = rows.map((rk) => valueOf(rowAll(rk))), colTotals = cols.map((ck) => valueOf(rows.flatMap((rk) => cellRows(rk, ck))));
  const showAs = o.showAs ?? "normal";
  const shown = (v: number, r: number, c: number): string => {
    switch (showAs) {
      case "percentOfTotal": return `${num((v / grand) * 100)}%`;
      case "percentOfRow": return `${num((v / rowTotals[r]) * 100)}%`;
      case "percentOfColumn": return `${num((v / colTotals[c]) * 100)}%`;
      case "runningTotal": return num(raw.slice(0, r + 1).reduce((s, x) => s + x[c], 0));
      case "rank": return String(1 + raw.filter((x) => x[c] > v).length);
      default: return num(v);
    }
  };
  const top = f.cols.length ? [...colKeys.values()].map((cv) => cv.join(" / ") || "(blank)") : [`${f.agg} of ${f.value}`];
  const layout = o.layout ?? "tabular", compact = layout === "compact", totalCol = o.grandTotals && f.cols.length > 0;
  const head = compact ? [f.rows.length ? "Row Labels" : ""] : [...f.rows];
  const out: string[][] = [[...head, ...top, ...(totalCol ? ["Grand Total"] : [])]];
  const line = (labels: string[], values: string[], total?: string) => [...labels, ...values, ...(totalCol ? [total ?? ""] : [])];
  const multi = f.rows.length > 1, collapsed = new Set(o.collapsed ?? []);
  let lastOuter: string | null = null;
  const subtotal = (outer: string) => {
    const ks = rows.filter((rk) => rowKeys.get(rk)![0] === outer), all = ks.flatMap(rowAll);
    const vals = cols.map((ck) => num(valueOf(ks.flatMap((rk) => cellRows(rk, ck)))));
    return line(compact ? [outer] : [`${outer} Total`, ...f.rows.slice(1).map(() => "")], vals, num(valueOf(all)));
  };
  rows.forEach((rk, r) => {
    const rv = rowKeys.get(rk)!, outer = rv[0];
    if (multi && outer !== lastOuter) {
      if (lastOuter !== null && o.subtotals && !compact) out.push(subtotal(lastOuter));
      // A compact report's outer group is its own row, its subtotal on it.
      if (compact) out.push(o.subtotals || collapsed.has(outer) ? subtotal(outer) : line([outer], cols.map(() => ""), ""));
      lastOuter = outer;
    }
    if (multi && collapsed.has(outer)) {
      // Collapsed: its detail left out, its subtotal (below, or on its row) standing for it.
      if (!compact && (r + 1 === rows.length || rowKeys.get(rows[r + 1])![0] !== outer) && !o.subtotals) out.push(subtotal(outer));
      return;
    }
    const labels = compact ? [`${"  ".repeat(rv.length - 1)}${rv[rv.length - 1]}`] : layout === "outline" ? rv.map((x, i) => (i === 0 && r > 0 && rowKeys.get(rows[r - 1])![0] === outer ? "" : x)) : rv;
    out.push(line(labels, raw[r].map((v, c) => shown(v, r, c)), num(rowTotals[r])));
  });
  if (multi && o.subtotals && !compact && lastOuter !== null) out.push(subtotal(lastOuter));
  if (o.grandTotals) out.push(line(compact ? ["Grand Total"] : ["Grand Total", ...f.rows.slice(1).map(() => "")], colTotals.map((v, c) => (showAs === "normal" ? num(v) : shown(v, rows.length - 1, c))), num(grand)));
  return out;
}
