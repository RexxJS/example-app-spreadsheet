// Excel's Consolidate: several ranges (of any sheets) put together into
// one, a function over the numbers that land on each cell: by position, or
// by the labels in their top row, their left column, or both (the union of
// the labels, in the order first seen, any case). Pure: the sources' values
// in, the consolidated block out, with where each of its numbers came from
// (for a formula linking to them).
import type { Value } from "./values.ts";

export const CONSOLIDATE_FUNCTIONS: Readonly<Record<string, string>> = {
  sum: "SUM", count: "COUNT", counta: "COUNTA", average: "AVERAGE", max: "MAX", min: "MIN", product: "PRODUCT", stdev: "STDEV.S", stdevp: "STDEV.P", var: "VAR.S", varp: "VAR.P",
};

// A source: its values row by row, and how to name its cell (r, c) in a
// formula ("Jan!B2").
export interface Source {
  readonly values: readonly (readonly Value[])[];
  readonly cell: (r: number, c: number) => string;
}

export interface Consolidated {
  // The block's rows: a label, a number, or null.
  readonly rows: readonly (readonly (Value)[])[];
  // Where each number came from, by the block's (r, c): the cells, in order.
  readonly from: readonly (readonly (readonly string[])[])[];
}

const label = (v: Value): string => (v === null ? "" : String(v)).trim();

function aggregate(fn: string, xs: readonly Value[]): Value {
  const ns = xs.filter((x): x is number => typeof x === "number");
  const mean = ns.reduce((s, n) => s + n, 0) / ns.length;
  const dev = (sample: boolean) => ns.reduce((s, n) => s + (n - mean) ** 2, 0) / (ns.length - (sample ? 1 : 0));
  switch (fn) {
    case "sum": return ns.reduce((s, n) => s + n, 0);
    case "count": return ns.length;
    case "counta": return xs.filter((x) => x !== null && x !== "").length;
    case "average": return ns.length ? mean : null;
    case "max": return ns.length ? Math.max(...ns) : null;
    case "min": return ns.length ? Math.min(...ns) : null;
    case "product": return ns.length ? ns.reduce((s, n) => s * n, 1) : null;
    case "stdev": return ns.length > 1 ? Math.sqrt(dev(true)) : null;
    case "stdevp": return ns.length ? Math.sqrt(dev(false)) : null;
    case "var": return ns.length > 1 ? dev(true) : null;
    case "varp": return ns.length ? dev(false) : null;
    default: return null;
  }
}

export function consolidate(sources: readonly Source[], fn: string, top: boolean, left: boolean): Consolidated {
  const rowLabels: string[] = [], colLabels: string[] = [];
  const index = (list: string[], l: string) => {
    let i = list.findIndex((x) => x.toLowerCase() === l.toLowerCase());
    if (i < 0) { list.push(l); i = list.length - 1; }
    return i;
  };
  // Each source's numbers at their place in the block.
  const cells = new Map<string, { values: Value[]; from: string[] }>();
  let height = 0, width = 0;
  for (const s of sources) {
    const h = s.values.length, w = Math.max(0, ...s.values.map((r) => r.length));
    for (let r = top ? 1 : 0; r < h; r++) {
      for (let c = left ? 1 : 0; c < w; c++) {
        const v = s.values[r][c] ?? null;
        const R = left ? index(rowLabels, label(s.values[r][0] ?? null)) : r - (top ? 1 : 0);
        const C = top ? index(colLabels, label(s.values[0][c] ?? null)) : c - (left ? 1 : 0);
        height = Math.max(height, R + 1);
        width = Math.max(width, C + 1);
        const k = `${R},${C}`, at = cells.get(k) ?? { values: [], from: [] };
        at.values.push(v);
        if (typeof v === "number" || (v !== null && v !== "")) at.from.push(s.cell(r, c));
        cells.set(k, at);
      }
    }
  }
  const rows: Value[][] = [], from: string[][][] = [];
  if (top) {
    rows.push([...(left ? [null] : []), ...colLabels]);
    from.push(rows[0].map(() => []));
  }
  for (let R = 0; R < height; R++) {
    const row: Value[] = left ? [rowLabels[R]] : [], fromRow: string[][] = left ? [[]] : [];
    for (let C = 0; C < width; C++) {
      const at = cells.get(`${R},${C}`);
      row.push(at ? aggregate(fn, at.values) : null);
      fromRow.push(at ? at.from : []);
    }
    rows.push(row);
    from.push(fromRow);
  }
  return { rows, from };
}
