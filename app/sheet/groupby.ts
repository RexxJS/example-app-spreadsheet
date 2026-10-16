// GROUPBY and PIVOTBY (functions.ts calls them from the evaluator, which
// holds the LAMBDA): rows of values grouped by their key fields, each
// group's values folded by an aggregate, in order, with totals. Pure: no
// sheet here.
import { CellError, compare, ordered } from "./values.ts";
import type { Value } from "./values.ts";

export type Aggregate = (values: readonly Value[]) => Value;

// Rows that share a key (every field equal, text without regard to case),
// in the order the sort asks: a column number from 1 (the output's), minus
// for descending, 0 for as found.
function grouped(keys: readonly (readonly Value[])[]): Map<number, number[]> {
  const found: number[][] = [], byRow = new Map<number, number[]>();
  for (let i = 0; i < keys.length; i++) {
    let g = found.find((rows) => keys[rows[0]].every((v, f) => alike(v, keys[i][f])));
    if (!g) found.push((g = []));
    g.push(i);
    byRow.set(i, g);
  }
  return new Map(found.map((g) => [g[0], g]));
}

const alike = (a: Value, b: Value): boolean => (a instanceof CellError || b instanceof CellError ? a === b : compare("=", a, b));

const rank = (a: Value, b: Value): number => {
  if (a instanceof CellError || b instanceof CellError) return 0;
  return ordered(a, b);
};

function inOrder(rows: Value[][], sort: number, fields: number): Value[][] {
  if (sort === 0) return rows;
  const col = Math.abs(sort) - 1, dir = Math.sign(sort);
  if (col >= rows[0]?.length) return rows;
  return [...rows].sort((x, y) => {
    const d = rank(x[col], y[col]);
    if (d !== 0 || col >= fields) return d * dir;
    // Equal on the sorted column: by the other key fields, ascending.
    for (let f = 0; f < fields; f++) { const e = rank(x[f], y[f]); if (e !== 0) return e; }
    return 0;
  });
}

export interface GroupOptions {
  // The first row of the fields and values is their headers: kept out of
  // the groups, and shown over the output when `showHeaders`.
  readonly headers: boolean;
  readonly showHeaders: boolean;
  readonly totals: boolean;
  readonly sort: number;
  // Which rows count (null: all).
  readonly keep: readonly boolean[] | null;
}

// One row a group: its key fields, then each values column folded.
// Then a Total row when asked: "Total" in the first field, the rest empty.
export function groupBy(fields: readonly (readonly Value[])[], values: readonly (readonly Value[])[], fold: Aggregate, o: GroupOptions): Value[][] {
  const from = o.headers ? 1 : 0;
  const rows = fields.map((_, i) => i).filter((i) => i >= from && (o.keep?.[i] ?? true));
  const out: Value[][] = [];
  for (const g of grouped(rows.map((i) => fields[i])).values()) {
    const members = g.map((j) => rows[j]);
    out.push([...fields[members[0]], ...values[0].map((_, c) => fold(members.map((i) => values[i][c])))]);
  }
  const body = inOrder(out, o.sort, fields[0]?.length ?? 0);
  if (o.headers && o.showHeaders) body.unshift([...fields[0], ...values[0]]);
  if (o.totals && rows.length) body.push(["Total", ...fields[0].slice(1).map(() => null), ...values[0].map((_, c) => fold(rows.map((i) => values[i][c])))]);
  return body;
}

export interface PivotOptions {
  readonly headers: boolean;
  readonly showHeaders: boolean;
  readonly rowTotals: boolean;
  readonly colTotals: boolean;
  readonly rowSort: number;
  readonly colSort: number;
  readonly keep: readonly boolean[] | null;
}

// A row a distinct row-field value, a column a distinct column-field
// value, each cell the values where both match (empty where none); a
// Total column and row when asked. The corner holds the headers when
// shown.
export function pivotBy(rowField: readonly Value[], colField: readonly Value[], values: readonly Value[], fold: Aggregate, o: PivotOptions): Value[][] {
  const from = o.headers ? 1 : 0;
  const rows = rowField.map((_, i) => i).filter((i) => i >= from && (o.keep?.[i] ?? true));
  const distinct = (field: readonly Value[], sort: number): Value[] => {
    const seen: Value[] = [];
    for (const i of rows) if (!seen.some((v) => alike(v, field[i]))) seen.push(field[i]);
    return sort === 0 ? seen : seen.sort((a, b) => rank(a, b) * Math.sign(sort));
  };
  const rowKeys = distinct(rowField, o.rowSort), colKeys = distinct(colField, o.colSort);
  const cell = (r: Value | null, c: Value | null): Value => {
    const at = rows.filter((i) => (r === null || alike(rowField[i], r)) && (c === null || alike(colField[i], c)));
    return at.length ? fold(at.map((i) => values[i])) : null;
  };
  const out: Value[][] = [];
  const corner: Value = o.headers && o.showHeaders ? rowField[0] : null;
  out.push([corner, ...colKeys, ...(o.colTotals ? ["Total"] : [])]);
  for (const r of rowKeys) out.push([r, ...colKeys.map((c) => cell(r, c)), ...(o.colTotals ? [cell(r, null)] : [])]);
  if (o.rowTotals) out.push(["Total", ...colKeys.map((c) => cell(null, c)), ...(o.colTotals ? [cell(null, null)] : [])]);
  return out;
}
