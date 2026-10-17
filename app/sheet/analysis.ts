// The Analysis ToolPak's thirteen tools, as Excel's Data Analysis writes
// them: each a table of labels and values (not live formulas), from
// columns of numbers (a first row of labels when `labels`). Pure; the
// distributions are distributions.ts's.
import { fCdf, tCdf, tInverse } from "./distributions.ts";

export const TOOLS = ["descriptive", "correlation", "covariance", "histogram", "movingAverage", "exponentialSmoothing", "rankPercentile", "regression", "tTestPaired", "tTestEqual", "tTestUnequal", "fTest", "anova"] as const;
export type Tool = (typeof TOOLS)[number];
export type Cell = string | number;
export interface Column { readonly name: string; readonly values: readonly number[] }

const sum = (xs: readonly number[]) => xs.reduce((s, x) => s + x, 0);
const mean = (xs: readonly number[]) => sum(xs) / xs.length;
const svar = (xs: readonly number[]) => { const m = mean(xs); return sum(xs.map((x) => (x - m) ** 2)) / (xs.length - 1); };
const median = (xs: readonly number[]) => { const s = [...xs].sort((a, b) => a - b), n = s.length; return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; };
const mode = (xs: readonly number[]): Cell => {
  const counts = new Map<number, number>();
  for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1);
  let best: number | null = null, n = 1;
  for (const x of xs) if (counts.get(x)! > n) { best = x; n = counts.get(x)!; }
  return best === null ? "#N/A" : best;
};
const skew = (xs: readonly number[]) => { const n = xs.length, m = mean(xs), s = Math.sqrt(svar(xs)); return (n / ((n - 1) * (n - 2))) * sum(xs.map((x) => ((x - m) / s) ** 3)); };
const kurt = (xs: readonly number[]) => { const n = xs.length, m = mean(xs), s = Math.sqrt(svar(xs)); return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * sum(xs.map((x) => ((x - m) / s) ** 4)) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3)); };
const pcov = (a: readonly number[], b: readonly number[]) => { const ma = mean(a), mb = mean(b); return sum(a.map((x, i) => (x - ma) * (b[i] - mb))) / a.length; };
const correl = (a: readonly number[], b: readonly number[]) => pcov(a, b) / Math.sqrt(pcov(a, a) * pcov(b, b));
const twoTail = (t: number, df: number) => 2 * (1 - tCdf(Math.abs(t), df));
const oneTail = (t: number, df: number) => 1 - tCdf(Math.abs(t), df);

export interface Options {
  readonly alpha?: number;          // 0.05
  readonly interval?: number;       // a moving average's, 3
  readonly damping?: number;        // exponential smoothing's, 0.3
  readonly bins?: readonly number[]; // a histogram's upper bounds
  readonly difference?: number;     // the t-tests' hypothesised mean difference, 0
  readonly confidence?: number;     // descriptive statistics' level, 0.95
}

// A tool's report: rows of cells.
export function analyze(tool: Tool, cols: readonly Column[], o: Options = {}): Cell[][] {
  const alpha = o.alpha ?? 0.05;
  const need = (n: number) => { if (cols.length < n) throw new RangeError(`${tool} needs ${n} column${n === 1 ? "" : "s"} of numbers`); };
  need(1);
  switch (tool) {
    case "descriptive": {
      const out: Cell[][] = [];
      for (const c of cols) {
        const xs = c.values, n = xs.length, sd = Math.sqrt(svar(xs)), lvl = o.confidence ?? 0.95;
        out.push([c.name, ""], ["Mean", mean(xs)], ["Standard Error", sd / Math.sqrt(n)], ["Median", median(xs)], ["Mode", mode(xs)], ["Standard Deviation", sd], ["Sample Variance", svar(xs)],
          ["Kurtosis", n > 3 ? kurt(xs) : "#DIV/0!"], ["Skewness", n > 2 ? skew(xs) : "#DIV/0!"], ["Range", Math.max(...xs) - Math.min(...xs)], ["Minimum", Math.min(...xs)], ["Maximum", Math.max(...xs)],
          ["Sum", sum(xs)], ["Count", n], [`Confidence Level(${lvl * 100}%)`, tInverse(1 - (1 - lvl) / 2, n - 1) * sd / Math.sqrt(n)], ["", ""]);
      }
      return out.slice(0, -1);
    }
    case "correlation": case "covariance": {
      const f = tool === "correlation" ? correl : pcov;
      return [["", ...cols.map((c) => c.name)], ...cols.map((a, i): Cell[] => [a.name, ...cols.map((b, j) => (j <= i ? f(a.values, b.values) : ""))])];
    }
    case "histogram": {
      const xs = cols[0].values, lo = Math.min(...xs), hi = Math.max(...xs);
      // Excel's own bins when none are given: from the least, about √n of them.
      const k = Math.max(1, Math.round(Math.sqrt(xs.length))), bins = o.bins?.length ? [...o.bins].sort((a, b) => a - b) : Array.from({ length: k }, (_, i) => lo + ((hi - lo) * (i + 1)) / k);
      const counts = bins.map((b, i) => xs.filter((x) => x <= b && (i === 0 || x > bins[i - 1])).length);
      return [["Bin", "Frequency"], ...bins.map((b, i): Cell[] => [b, counts[i]]), ["More", xs.filter((x) => x > bins[bins.length - 1]).length]];
    }
    case "movingAverage": {
      const xs = cols[0].values, n = o.interval ?? 3;
      if (!(n >= 2 && n <= xs.length)) throw new RangeError(`an interval of ${n}: 2 to ${xs.length}`);
      return [["Moving Average"], ...xs.map((_, i): Cell[] => [i < n - 1 ? "#N/A" : mean(xs.slice(i - n + 1, i + 1))])];
    }
    case "exponentialSmoothing": {
      const xs = cols[0].values, d = o.damping ?? 0.3, a = 1 - d;
      if (!(d >= 0 && d <= 1)) throw new RangeError("a damping factor is 0 to 1");
      const f: Cell[] = ["#N/A"];
      let prev = xs[0];
      for (let i = 1; i < xs.length; i++) { f.push(prev); prev = a * xs[i] + d * prev; }
      return [["Forecast"], ...f.map((v): Cell[] => [v])];
    }
    case "rankPercentile": {
      const xs = cols[0].values, order = xs.map((x, i) => ({ x, i: i + 1 })).sort((p, q) => q.x - p.x), n = xs.length;
      return [["Point", cols[0].name, "Rank", "Percent"], ...order.map((e): Cell[] => {
        const rank = 1 + xs.filter((y) => y > e.x).length, below = xs.filter((y) => y < e.x).length;
        return [e.i, e.x, rank, n > 1 ? Math.floor((below / (n - 1)) * 1000) / 1000 : 1];
      })];
    }
    case "regression": {
      need(2);
      const y = cols[0].values, xs = cols.slice(1), n = y.length, k = xs.length;
      // Least squares by the normal equations, with an intercept.
      const X = y.map((_, i) => [1, ...xs.map((c) => c.values[i])]), p = k + 1;
      const xtx = Array.from({ length: p }, (_, a) => Array.from({ length: p }, (_, b) => sum(X.map((r) => r[a] * r[b]))));
      const inv = invert(xtx);
      if (!inv) throw new RangeError("the x columns are collinear");
      const xty = Array.from({ length: p }, (_, a) => sum(X.map((r, i) => r[a] * y[i])));
      const beta = inv.map((row) => sum(row.map((v, j) => v * xty[j])));
      const fitted = X.map((r) => sum(r.map((v, j) => v * beta[j]))), resid = y.map((v, i) => v - fitted[i]);
      const ym = mean(y), sst = sum(y.map((v) => (v - ym) ** 2)), sse = sum(resid.map((r) => r * r)), ssr = sst - sse;
      const dfr = k, dfe = n - k - 1, mse = sse / dfe, se = beta.map((_, j) => Math.sqrt(mse * inv[j][j]));
      const r2 = ssr / sst, F = (ssr / dfr) / mse, tc = tInverse(1 - alpha / 2, dfe);
      return [["SUMMARY OUTPUT"], [], ["Regression Statistics"], ["Multiple R", Math.sqrt(r2)], ["R Square", r2], ["Adjusted R Square", 1 - (1 - r2) * (n - 1) / dfe], ["Standard Error", Math.sqrt(mse)], ["Observations", n], [],
        ["ANOVA"], ["", "df", "SS", "MS", "F", "Significance F"], ["Regression", dfr, ssr, ssr / dfr, F, 1 - fCdf(F, dfr, dfe)], ["Residual", dfe, sse, mse], ["Total", n - 1, sst], [],
        ["", "Coefficients", "Standard Error", "t Stat", "P-value", `Lower ${100 - alpha * 100}%`, `Upper ${100 - alpha * 100}%`],
        ...beta.map((b, j): Cell[] => [j === 0 ? "Intercept" : xs[j - 1].name, b, se[j], b / se[j], twoTail(b / se[j], dfe), b - tc * se[j], b + tc * se[j]])];
    }
    case "tTestPaired": case "tTestEqual": case "tTestUnequal": case "fTest": {
      need(2);
      const [a, b] = [cols[0].values, cols[1].values], na = a.length, nb = b.length, va = svar(a), vb = svar(b), d0 = o.difference ?? 0;
      const head: Cell[][] = [["", cols[0].name, cols[1].name], ["Mean", mean(a), mean(b)], ["Variance", va, vb], ["Observations", na, nb]];
      if (tool === "fTest") {
        const F = va / vb, df1 = na - 1, df2 = nb - 1, p = F > 1 ? 1 - fCdf(F, df1, df2) : fCdf(F, df1, df2);
        const crit = F > 1 ? invertF(1 - alpha, df1, df2) : invertF(alpha, df1, df2);
        return [...head, ["df", df1, df2], ["F", F], ["P(F<=f) one-tail", p], ["F Critical one-tail", crit]];
      }
      let t: number, df: number;
      const extra: Cell[][] = [];
      if (tool === "tTestPaired") {
        if (na !== nb) throw new RangeError("a paired test's columns are the same length");
        const d = a.map((x, i) => x - b[i]);
        t = (mean(d) - d0) / Math.sqrt(svar(d) / na); df = na - 1;
        extra.push(["Pearson Correlation", correl(a, b)], ["Hypothesized Mean Difference", d0], ["df", df]);
      } else if (tool === "tTestEqual") {
        const sp = ((na - 1) * va + (nb - 1) * vb) / (na + nb - 2);
        t = (mean(a) - mean(b) - d0) / Math.sqrt(sp * (1 / na + 1 / nb)); df = na + nb - 2;
        extra.push(["Pooled Variance", sp], ["Hypothesized Mean Difference", d0], ["df", df]);
      } else {
        const sa = va / na, sb = vb / nb;
        t = (mean(a) - mean(b) - d0) / Math.sqrt(sa + sb);
        // Welch's degrees of freedom, rounded as Excel's tool reports them.
        df = Math.round((sa + sb) ** 2 / ((sa * sa) / (na - 1) + (sb * sb) / (nb - 1)));
        extra.push(["Hypothesized Mean Difference", d0], ["df", df]);
      }
      return [...head, ...extra, ["t Stat", t], ["P(T<=t) one-tail", oneTail(t, df)], ["t Critical one-tail", tInverse(1 - alpha, df)], ["P(T<=t) two-tail", twoTail(t, df)], ["t Critical two-tail", tInverse(1 - alpha / 2, df)]];
    }
    case "anova": {
      need(2);
      const all = cols.flatMap((c) => c.values), gm = mean(all), N = all.length, k = cols.length;
      const ssb = sum(cols.map((c) => c.values.length * (mean(c.values) - gm) ** 2)), ssw = sum(cols.map((c) => sum(c.values.map((x) => (x - mean(c.values)) ** 2))));
      const dfb = k - 1, dfw = N - k, F = (ssb / dfb) / (ssw / dfw);
      return [["Anova: Single Factor"], [], ["SUMMARY"], ["Groups", "Count", "Sum", "Average", "Variance"], ...cols.map((c): Cell[] => [c.name, c.values.length, sum(c.values), mean(c.values), svar(c.values)]), [],
        ["ANOVA"], ["Source of Variation", "SS", "df", "MS", "F", "P-value", "F crit"], ["Between Groups", ssb, dfb, ssb / dfb, F, 1 - fCdf(F, dfb, dfw), invertF(1 - alpha, dfb, dfw)],
        ["Within Groups", ssw, dfw, ssw / dfw], [], ["Total", ssb + ssw, N - 1]];
    }
  }
}

// F's inverse by bisection.
function invertF(p: number, d1: number, d2: number): number {
  let lo = 0, hi = 1;
  while (fCdf(hi, d1, d2) < p && hi < 1e8) hi *= 2;
  for (let i = 0; i < 200; i++) { const m = (lo + hi) / 2; if (fCdf(m, d1, d2) < p) lo = m; else hi = m; }
  return (lo + hi) / 2;
}

// A small matrix's inverse by Gauss-Jordan, or null when it is singular.
function invert(m: number[][]): number[][] | null {
  const n = m.length, a = m.map((row, i) => [...row, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (Math.abs(a[p][c]) < 1e-12) return null;
    [a[c], a[p]] = [a[p], a[c]];
    const d = a[c][c];
    for (let j = 0; j < 2 * n; j++) a[c][j] /= d;
    for (let r = 0; r < n; r++) if (r !== c) { const f = a[r][c]; for (let j = 0; j < 2 * n; j++) a[r][j] -= f * a[c][j]; }
  }
  return a.map((row) => row.slice(n));
}
