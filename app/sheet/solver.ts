// Solver's Simplex LP: maximise or minimise c·x (or reach a value) subject
// to rows a·x (<=, >= or =) b, the variables non-negative unless freed,
// some integer or binary (by branch and bound). Pure: the coefficients in
// (commands.ts reads them off the sheet's formulas, which must be linear),
// the solution out. Two-phase simplex on a dense tableau, Bland's rule so
// it cannot cycle.
export type Relation = "<=" | ">=" | "=";
export interface LpRow { readonly a: readonly number[]; readonly rel: Relation; readonly b: number }
export interface Lp {
  readonly c: readonly number[];
  readonly maximise: boolean;
  readonly rows: readonly LpRow[];
  readonly free?: readonly boolean[];      // a variable that may go negative
  readonly integer?: readonly boolean[];
}
export type LpResult = { readonly status: "optimal"; readonly x: number[]; readonly value: number } | { readonly status: "infeasible" | "unbounded" };

const EPS = 1e-9;

// The LP in standard form, solved: max c·x, x >= 0.
function simplex(c: readonly number[], rows: readonly LpRow[]): LpResult {
  const n = c.length;
  // Each row with b >= 0, then its slack, surplus and artificial columns.
  const norm = rows.map((r) => (r.b < 0 ? { a: r.a.map((v) => -v), rel: (r.rel === "<=" ? ">=" : r.rel === ">=" ? "<=" : "=") as Relation, b: -r.b } : r));
  const m = norm.length;
  const slack = norm.map((r) => (r.rel === "=" ? 0 : 1)).reduce((s, x) => s + x, 0), art = norm.filter((r) => r.rel !== "<=").length;
  const width = n + slack + art + 1;
  const T: number[][] = [], basis: number[] = [];
  let si = n, ai = n + slack;
  norm.forEach((r) => {
    const row = new Array(width).fill(0);
    r.a.forEach((v, j) => (row[j] = v));
    row[width - 1] = r.b;
    if (r.rel === "<=") { row[si] = 1; basis.push(si++); }
    else if (r.rel === ">=") { row[si++] = -1; row[ai] = 1; basis.push(ai++); }
    else { row[ai] = 1; basis.push(ai++); }
    T.push(row);
  });
  const pivot = (r: number, col: number) => {
    const p = T[r][col];
    for (let j = 0; j < width; j++) T[r][j] /= p;
    for (let i = 0; i < T.length; i++) if (i !== r && Math.abs(T[i][col]) > EPS) { const f = T[i][col]; for (let j = 0; j < width; j++) T[i][j] -= f * T[r][j]; }
    basis[r] = col;
  };
  // Optimise the objective row z (reduced costs, maximising), over the
  // columns allowed; false when unbounded.
  const run = (z: number[], allowed: (j: number) => boolean): boolean => {
    for (let guard = 0; guard < 5000; guard++) {
      // Reduced cost of column j: z_j - Σ z_basis * T[i][j]; enter the first positive (Bland).
      let enter = -1;
      for (let j = 0; j < width - 1 && enter < 0; j++) {
        if (!allowed(j) || basis.includes(j)) continue;
        const rc = z[j] - basis.reduce((s, b, i) => s + z[b] * T[i][j], 0);
        if (rc > EPS) enter = j;
      }
      if (enter < 0) return true;
      let leave = -1, best = Infinity;
      for (let i = 0; i < m; i++) if (T[i][enter] > EPS) { const ratio = T[i][width - 1] / T[i][enter]; if (ratio < best - EPS || (Math.abs(ratio - best) <= EPS && basis[i] < basis[leave])) { best = ratio; leave = i; } }
      if (leave < 0) return false;
      pivot(leave, enter);
    }
    return true;
  };
  if (art) {
    const z = new Array(width).fill(0);
    for (let j = n + slack; j < n + slack + art; j++) z[j] = -1;
    run(z, () => true);
    const infeas = basis.reduce((s, b, i) => s + (b >= n + slack ? T[i][width - 1] : 0), 0);
    if (infeas > 1e-7) return { status: "infeasible" };
    // Artificials still basic at 0: pivot them out where a real column can.
    basis.forEach((b, i) => { if (b >= n + slack) { const j = T[i].findIndex((v, k) => k < n + slack && Math.abs(v) > EPS); if (j >= 0) pivot(i, j); } });
  }
  const z = new Array(width).fill(0);
  c.forEach((v, j) => (z[j] = v));
  if (!run(z, (j) => j < n + slack)) return { status: "unbounded" };
  const x = new Array(n).fill(0);
  basis.forEach((b, i) => { if (b < n) x[b] = T[i][width - 1]; });
  return { status: "optimal", x, value: c.reduce((s, v, j) => s + v * x[j], 0) };
}

export function solveLp(lp: Lp): LpResult {
  const n = lp.c.length, free = lp.free ?? [];
  // A free variable as the difference of two non-negative ones.
  const split = (v: readonly number[]) => [...v, ...v.filter((_, j) => free[j]).map((x) => -x)];
  const freeIdx = [...Array(n).keys()].filter((j) => free[j]);
  const join = (x: number[]) => x.slice(0, n).map((v, j) => (free[j] ? v - x[n + freeIdx.indexOf(j)] : v));
  const c = split(lp.c).map((v) => (lp.maximise ? v : -v));
  const solve = (rows: readonly LpRow[]): LpResult => {
    const r = simplex(c, rows.map((row) => ({ ...row, a: split(row.a) })));
    return r.status === "optimal" ? { status: "optimal", x: join(r.x), value: lp.maximise ? r.value : -r.value } : r;
  };
  const integer = lp.integer ?? [];
  if (!integer.some((x) => x)) return solve(lp.rows);
  // Branch and bound, depth first, the best integer solution kept.
  let best: LpResult = { status: "infeasible" };
  const better = (v: number) => best.status !== "optimal" || (lp.maximise ? v > best.value + 1e-9 : v < best.value - 1e-9);
  const branch = (rows: readonly LpRow[], depth: number) => {
    if (depth > 60) return;
    const r = solve(rows);
    if (r.status !== "optimal" || !better(r.value)) { if (r.status === "unbounded" && best.status !== "optimal") best = r; return; }
    const j = r.x.findIndex((v, k) => integer[k] && Math.abs(v - Math.round(v)) > 1e-7);
    if (j < 0) { best = { status: "optimal", x: r.x.map((v, k) => (integer[k] ? Math.round(v) : v)), value: r.value }; return; }
    const unit = (k: number) => lp.c.map((_, i) => (i === k ? 1 : 0));
    branch([...rows, { a: unit(j), rel: "<=", b: Math.floor(r.x[j]) }], depth + 1);
    branch([...rows, { a: unit(j), rel: ">=", b: Math.ceil(r.x[j]) }], depth + 1);
  };
  branch(lp.rows, 0);
  return best;
}
