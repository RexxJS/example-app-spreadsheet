// The FORECAST.ETS family: additive Holt-Winters exponential smoothing
// (Excel's AAA: additive error, trend and season), its smoothing factors
// fitted by least squares over a grid then refined, a season found by
// autocorrelation when not given (1 for none). Excel's optimiser is its
// own; this one is documented here and tested on series whose forecast is
// known (a line goes on as the line, a season repeats). Pure.
export interface EtsModel {
  readonly alpha: number;
  readonly beta: number;
  readonly gamma: number;
  readonly season: number;
  readonly level: number;
  readonly trend: number;
  readonly seasonal: readonly number[];   // the last season's, from the series' end
  readonly sse: number;
  readonly fitted: readonly number[];     // one step ahead, from the second point
}

// The season (2 to a third of the series, or 1: none): the lag whose
// autocorrelation of the differenced series is highest and clearly so.
export function seasonality(ys: readonly number[]): number {
  const d = ys.slice(1).map((y, i) => y - ys[i]), n = d.length, m = d.reduce((s, v) => s + v, 0) / n;
  const c0 = d.reduce((s, v) => s + (v - m) ** 2, 0);
  if (n < 6 || c0 === 0) return 1;
  let best = 1, bestR = 0.5;
  for (let lag = 2; lag <= Math.floor(n / 3); lag++) {
    let c = 0;
    for (let i = lag; i < n; i++) c += (d[i] - m) * (d[i - lag] - m);
    const r = c / c0;
    if (r > bestR + 1e-9) { best = lag; bestR = r; }
  }
  return best;
}

function run(ys: readonly number[], season: number, alpha: number, beta: number, gamma: number): EtsModel {
  const L = season > 1 ? season : 1, n = ys.length;
  let level: number, trend: number;
  const seasonal: number[] = new Array(L).fill(0);
  if (L > 1 && n >= 2 * L) {
    const m1 = ys.slice(0, L).reduce((s, v) => s + v, 0) / L, m2 = ys.slice(L, 2 * L).reduce((s, v) => s + v, 0) / L;
    level = m1; trend = (m2 - m1) / L;
    for (let i = 0; i < L; i++) seasonal[i] = ys[i] - (m1 + (i - (L - 1) / 2) * trend);
  } else { level = ys[0]; trend = n > 1 ? ys[1] - ys[0] : 0; }
  let sse = 0;
  const fitted: number[] = [];
  const start = L > 1 && n >= 2 * L ? 0 : 1;
  if (start === 1) fitted.push(ys[0]);
  // From the first point (season known) the level steps forward.
  if (L > 1 && n >= 2 * L) level = level - ((L - 1) / 2) * trend - trend;
  for (let t = start; t < n; t++) {
    const s = seasonal[t % L], f = level + trend + s;
    fitted.push(f);
    const e = ys[t] - f;
    if (t >= (L > 1 ? L : 1)) sse += e * e;
    const prev = level;
    level = alpha * (ys[t] - s) + (1 - alpha) * (level + trend);
    trend = beta * (level - prev) + (1 - beta) * trend;
    seasonal[t % L] = gamma * (ys[t] - level) + (1 - gamma) * s;
  }
  // The seasonal factors from the series' end on.
  const tail = Array.from({ length: L }, (_, h) => seasonal[(n + h) % L]);
  return { alpha, beta, gamma, season: L, level, trend, seasonal: tail, sse, fitted };
}

export function fitEts(ys: readonly number[], seasonGiven?: number): EtsModel {
  if (ys.length < 3) throw new RangeError("FORECAST.ETS needs three values or more");
  const season = seasonGiven === undefined ? seasonality(ys) : Math.max(1, Math.round(seasonGiven));
  if (season > ys.length / 2) throw new RangeError(`a season of ${season} needs ${2 * season} values or more`);
  let best = run(ys, season, 0.5, 0.1, 0.1);
  const grid = [0.001, 0.1, 0.25, 0.5, 0.75, 0.999];
  for (const a of grid) for (const b of grid) for (const g of season > 1 ? grid : [0]) {
    const m = run(ys, season, a, b, g);
    if (m.sse < best.sse - 1e-12) best = m;
  }
  // Refined by coordinate steps, each halving.
  let step = 0.1;
  for (let k = 0; k < 40 && step > 1e-4; k++) {
    let moved = false;
    for (const which of ["alpha", "beta", "gamma"] as const) {
      if (which === "gamma" && season === 1) continue;
      for (const dir of [-1, 1]) {
        const p = { alpha: best.alpha, beta: best.beta, gamma: best.gamma };
        p[which] = Math.min(0.999, Math.max(0.001, p[which] + dir * step));
        const m = run(ys, season, p.alpha, p.beta, p.gamma);
        if (m.sse < best.sse - 1e-12) { best = m; moved = true; }
      }
    }
    if (!moved) step /= 2;
  }
  return best;
}

// The forecast h steps past the series' end.
export const forecastAt = (m: EtsModel, h: number): number => m.level + h * m.trend + (m.season > 1 ? m.seasonal[(h - 1) % m.season] : 0);

// A forecast's confidence interval's half width at h steps, from the
// one-step errors' deviation, widening with the steps (√h).
export function confidence(m: EtsModel, ys: readonly number[], h: number, level: number, z: (p: number) => number): number {
  const errs = ys.map((y, i) => y - m.fitted[i]).slice(m.season > 1 ? m.season : 1);
  const sd = Math.sqrt(errs.reduce((s, e) => s + e * e, 0) / Math.max(1, errs.length));
  return z(1 - (1 - level) / 2) * sd * Math.sqrt(h);
}

// FORECAST.ETS.STAT's statistics: 1 alpha, 2 beta, 3 gamma, 4 MASE, 5
// SMAPE, 6 MAE, 7 RMSE, 8 the timeline's step.
export function etsStat(m: EtsModel, ys: readonly number[], kind: number, step: number): number {
  const from = m.season > 1 ? m.season : 1, errs = ys.slice(from).map((y, i) => y - m.fitted[i + from]);
  const mae = errs.reduce((s, e) => s + Math.abs(e), 0) / errs.length;
  switch (kind) {
    case 1: return m.alpha;
    case 2: return m.beta;
    case 3: return m.gamma;
    case 4: { const naive = ys.slice(1).reduce((s, y, i) => s + Math.abs(y - ys[i]), 0) / (ys.length - 1); return naive ? mae / naive : 0; }
    case 5: return errs.reduce((s, e, i) => { const y = ys[i + from], f = m.fitted[i + from]; return s + (Math.abs(y) + Math.abs(f) ? (2 * Math.abs(e)) / (Math.abs(y) + Math.abs(f)) : 0); }, 0) / errs.length;
    case 6: return mae;
    case 7: return Math.sqrt(errs.reduce((s, e) => s + e * e, 0) / errs.length);
    case 8: return step;
    default: throw new RangeError(`statistic ${kind} is not 1 to 8`);
  }
}
