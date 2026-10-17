// An outline, as Excel's Group and Outline: rows (or columns) grouped into
// levels 1 to 8, a level's detail hidden or shown, the outline's buttons
// (show levels 1 to n: what is deeper hidden); and Subtotal's layout, the
// rows a sorted list gains at each change in a column. Pure: levels in,
// the lines to hide or show out (commands.ts keeps them on the page).
export const MOST_LEVELS = 8;

// The lines lo..hi a level deeper (or shallower), none past 8 or under 0.
export function regrouped(levels: ReadonlyMap<number, number>, lo: number, hi: number, by: 1 | -1): Map<number, number> {
  const out = new Map(levels);
  for (let i = lo; i <= hi; i++) {
    const l = Math.max(0, Math.min(MOST_LEVELS, (out.get(i) ?? 0) + by));
    if (l) out.set(i, l); else out.delete(i);
  }
  return out;
}

// The deepest level there is.
export const depth = (levels: ReadonlyMap<number, number>): number => Math.max(0, ...levels.values());

// The groups, each a run of lines at a level or deeper: [first, last, level],
// shallowest first, then in order.
export function groups(levels: ReadonlyMap<number, number>): [number, number, number][] {
  const out: [number, number, number][] = [];
  const lines = [...levels.keys()].sort((a, b) => a - b);
  for (let level = 1; level <= depth(levels); level++) {
    let start = -1, prev = -2;
    for (const i of lines) {
      if ((levels.get(i) ?? 0) < level) continue;
      if (i !== prev + 1) {
        if (start >= 0) out.push([start, prev, level]);
        start = i;
      }
      prev = i;
    }
    if (start >= 0) out.push([start, prev, level]);
  }
  return out;
}

// Show levels 1 to n (the outline's button n): the lines deeper hidden.
export const hiddenAtLevel = (levels: ReadonlyMap<number, number>, n: number): number[] =>
  [...levels].filter(([, l]) => l >= n).map(([i]) => i).sort((a, b) => a - b);

// The detail of the group a line is in, at its own level: the run of lines
// around it that deep or deeper (Hide Detail and Show Detail).
export function detailOf(levels: ReadonlyMap<number, number>, line: number): number[] {
  const level = levels.get(line) ?? 0;
  if (!level) return [];
  let lo = line, hi = line;
  while ((levels.get(lo - 1) ?? 0) >= level) lo--;
  while ((levels.get(hi + 1) ?? 0) >= level) hi++;
  return Array.from({ length: hi - lo + 1 }, (_, i) => lo + i);
}

// Excel's SUBTOTAL numbers for Subtotal's functions.
export const SUBTOTAL_FUNCTIONS: Readonly<Record<string, number>> = { average: 1, count: 2, counta: 3, max: 4, min: 5, product: 6, stdev: 7, stdevp: 8, sum: 9, var: 10, varp: 11 };

// Where each change in a list's column comes (its rows from `first` to
// `last`, `value(row)` that column's shown value): the groups, each
// [first row, last row, value], in order.
export function runs(first: number, last: number, value: (row: number) => string): [number, number, string][] {
  const out: [number, number, string][] = [];
  for (let r = first; r <= last; r++) {
    const v = value(r), top = out[out.length - 1];
    if (top && top[2].toLowerCase() === v.toLowerCase()) top[1] = r; else out.push([r, r, v]);
  }
  return out;
}
