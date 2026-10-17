// Trendlines, as Excel's chart trendlines fit them: linear, exponential,
// logarithmic, polynomial (order 2 to 6), power, by least squares (the
// exponential, logarithmic and power through logarithms, as Excel's,
// their R² on the transformed values), and a moving average; each its
// equation and R² as Excel's labels write them. Pure.
export const TREND_TYPES = ["linear", "exponential", "logarithmic", "polynomial", "power", "movingAverage"] as const;
export type TrendType = (typeof TREND_TYPES)[number];

export interface Fit {
  readonly at: (x: number) => number;
  readonly equation: string;
  readonly r2: number | null;
}

// Least squares: the coefficients c0..ck of c0 + c1 x + ... ck x^k.
function polyfit(xs: readonly number[], ys: readonly number[], k: number): number[] | null {
  const n = k + 1, a = Array.from({ length: n }, (_, i) => Array.from({ length: n + 1 }, (_, j) => (j < n ? xs.reduce((s, x) => s + x ** (i + j), 0) : xs.reduce((s, x, m) => s + ys[m] * x ** i, 0))));
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (Math.abs(a[p][c]) < 1e-12) return null;
    [a[c], a[p]] = [a[p], a[c]];
    for (let r = 0; r < n; r++) {
      if (r === c) continue;
      const f = a[r][c] / a[c][c];
      for (let j = c; j <= n; j++) a[r][j] -= f * a[c][j];
    }
  }
  return a.map((row, i) => row[n] / row[i]);
}

const rSquared = (ys: readonly number[], fitted: readonly number[]): number => {
  const mean = ys.reduce((s, y) => s + y, 0) / ys.length;
  const tot = ys.reduce((s, y) => s + (y - mean) ** 2, 0), res = ys.reduce((s, y, i) => s + (y - fitted[i]) ** 2, 0);
  return tot === 0 ? 1 : 1 - res / tot;
};

// A coefficient as the label writes it: four significant figures.
export const coef = (c: number): string => String(Number(c.toPrecision(4)));
const SUP: Record<string, string> = { 2: "²", 3: "³", 4: "⁴", 5: "⁵", 6: "⁶" };

// A sum's terms joined with their signs: "2x + 1", "0.5x² - 2x + 3".
function terms(cs: readonly number[]): string {
  // A coefficient lost in the arithmetic's noise is none.
  const scale = Math.max(...cs.map(Math.abs));
  const parts = cs.map((c, i) => ({ c, i })).reverse().filter(({ c }) => Math.abs(c) > scale * 1e-9);
  if (!parts.length) return "0";
  return parts.map(({ c, i }, k) => {
    const v = coef(Math.abs(c)), x = i === 0 ? "" : i === 1 ? "x" : `x${SUP[i] ?? `^${i}`}`, body = `${i > 0 && v === "1" ? "" : v}${x}`;
    return k === 0 ? `${c < 0 ? "-" : ""}${body}` : ` ${c < 0 ? "-" : "+"} ${body}`;
  }).join("");
}

export function fit(type: TrendType, xs: readonly number[], ys: readonly number[], order = 2): Fit | null {
  if (xs.length < 2) return null;
  switch (type) {
    case "linear": case "polynomial": {
      const k = type === "linear" ? 1 : Math.max(2, Math.min(6, order));
      if (xs.length <= k) return null;
      const c = polyfit(xs, ys, k);
      if (!c) return null;
      const at = (x: number) => c.reduce((s, ci, i) => s + ci * x ** i, 0);
      return { at, equation: `y = ${terms(c)}`, r2: rSquared(ys, xs.map(at)) };
    }
    case "exponential": {
      if (ys.some((y) => y <= 0)) return null;
      const l = ys.map(Math.log), c = polyfit(xs, l, 1);
      if (!c) return null;
      const a = Math.exp(c[0]), b = c[1];
      return { at: (x) => a * Math.exp(b * x), equation: `y = ${coef(a)}e^${coef(b)}x`, r2: rSquared(l, xs.map((x) => c[0] + c[1] * x)) };
    }
    case "logarithmic": {
      if (xs.some((x) => x <= 0)) return null;
      const lx = xs.map(Math.log), c = polyfit(lx, ys, 1);
      if (!c) return null;
      return { at: (x) => c[1] * Math.log(x) + c[0], equation: `y = ${coef(c[1])}ln(x) ${c[0] < 0 ? "-" : "+"} ${coef(Math.abs(c[0]))}`, r2: rSquared(ys, lx.map((x) => c[0] + c[1] * x)) };
    }
    case "power": {
      if (xs.some((x) => x <= 0) || ys.some((y) => y <= 0)) return null;
      const lx = xs.map(Math.log), ly = ys.map(Math.log), c = polyfit(lx, ly, 1);
      if (!c) return null;
      const a = Math.exp(c[0]), b = c[1];
      return { at: (x) => a * x ** b, equation: `y = ${coef(a)}x^${coef(b)}`, r2: rSquared(ly, lx.map((x) => c[0] + c[1] * x)) };
    }
    case "movingAverage": return null;
  }
}

// A moving average of `period` points: [x, mean] from the period-th on.
export function movingAverage(xs: readonly number[], ys: readonly number[], period: number): [number, number][] {
  const out: [number, number][] = [];
  for (let i = period - 1; i < ys.length; i++) out.push([xs[i], ys.slice(i - period + 1, i + 1).reduce((s, y) => s + y, 0) / period]);
  return out;
}
