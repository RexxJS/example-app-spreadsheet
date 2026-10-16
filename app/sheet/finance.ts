// Money over time, as Excel's financial functions work it out: a rate per
// period, a number of periods, a payment each period (at its end, type 0,
// or its start, type 1), a present and a future value, money paid out
// negative. Pure arithmetic for the library.

const grow = (rate: number, n: number) => (1 + rate) ** n;

export function pmt(rate: number, n: number, pv: number, fv = 0, type = 0): number {
  if (rate === 0) return -(pv + fv) / n;
  const g = grow(rate, n);
  return -(rate * (fv + pv * g)) / ((1 + rate * type) * (g - 1));
}

export function fv(rate: number, n: number, payment: number, pv = 0, type = 0): number {
  if (rate === 0) return -(pv + payment * n);
  const g = grow(rate, n);
  return -(pv * g + (payment * (1 + rate * type) * (g - 1)) / rate);
}

export function pv(rate: number, n: number, payment: number, fv = 0, type = 0): number {
  if (rate === 0) return -(fv + payment * n);
  const g = grow(rate, n);
  return -(fv + (payment * (1 + rate * type) * (g - 1)) / rate) / g;
}

export function nper(rate: number, payment: number, pv: number, fv = 0, type = 0): number {
  if (rate === 0) return -(pv + fv) / payment;
  const p = payment * (1 + rate * type);
  return Math.log((p - fv * rate) / (p + pv * rate)) / Math.log(1 + rate);
}

// Newton's method from a guess; null if it does not settle.
function solve(f: (x: number) => number, guess: number): number | null {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x), h = 1e-7 * Math.max(1, Math.abs(x)), slope = (f(x + h) - y) / h;
    if (!Number.isFinite(slope) || slope === 0) return null;
    const next = x - y / slope;
    if (Math.abs(next - x) < 1e-12) return next;
    x = next;
  }
  return null;
}

// The rate per period that makes the payments pay pv off to fv.
export const rate = (n: number, payment: number, pv: number, fv = 0, type = 0, guess = 0.1): number | null =>
  solve((r) => (r === 0 ? pv + payment * n + fv : pv * grow(r, n) + (payment * (1 + r * type) * (grow(r, n) - 1)) / r + fv), guess);

// Cash flows at the ends of periods 1, 2 ... discounted to now.
export const npv = (rate: number, flows: readonly number[]): number => flows.reduce((s, c, i) => s + c / grow(rate, i + 1), 0);

// The rate at which flows from now (period 0) on are worth nothing now.
export const irr = (flows: readonly number[], guess = 0.1): number | null =>
  solve((r) => flows.reduce((s, c, i) => s + c / grow(r, i), 0), guess);

// The interest in period `per` of a loan paid by pmt(): at a period's end
// (type 0), on the balance after the payments before it; at its start
// (type 1), the payment then being the interest of the period before, so
// the first period's is 0.
export function ipmt(rate: number, per: number, n: number, pv: number, fv = 0, type = 0): number {
  const p = pmt(rate, n, pv, fv, type);
  if (type === 1) return per === 1 ? 0 : (fvOf(rate, per - 2, p, pv, 1) - p) * rate;
  return fvOf(rate, per - 1, p, pv, 0) * rate;
}
const fvOf = (rate: number, n: number, payment: number, pv: number, type: number): number => fv(rate, n, payment, pv, type);
export const ppmt = (rate: number, per: number, n: number, pv: number, fv = 0, type = 0): number =>
  pmt(rate, n, pv, fv, type) - ipmt(rate, per, n, pv, fv, type);

// Cash flows on dates, discounted on a 365-day year from the first date.
export const xnpv = (rate: number, flows: readonly number[], dates: readonly number[]): number =>
  flows.reduce((s, c, i) => s + c / (1 + rate) ** ((dates[i] - dates[0]) / 365), 0);

// The rate at which they are worth nothing: Newton's method with the
// derivative, from a guess, then bisection over (-1, 100) if it strays.
export function xirr(flows: readonly number[], dates: readonly number[], guess = 0.1): number | null {
  const f = (r: number) => xnpv(r, flows, dates);
  const df = (r: number) => flows.reduce((s, c, i) => {
    const t = (dates[i] - dates[0]) / 365;
    return s - (t * c) / (1 + r) ** (t + 1);
  }, 0);
  let r = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(r), d = df(r);
    if (!Number.isFinite(y) || !Number.isFinite(d) || d === 0) break;
    const next = r - y / d;
    if (next <= -1) break;
    if (Math.abs(next - r) < 1e-12) return next;
    r = next;
  }
  let lo = -0.999999, hi = 100;
  if (Math.sign(f(lo)) === Math.sign(f(hi))) return null;
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2;
    if (Math.sign(f(mid)) === Math.sign(f(lo))) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

// The modified internal rate of return: what is paid out financed at one
// rate, what comes in reinvested at another.
export function mirr(flows: readonly number[], financeRate: number, reinvestRate: number): number | null {
  const n = flows.length;
  const out = npv(financeRate, flows.map((c) => (c < 0 ? c : 0)));
  const back = npv(reinvestRate, flows.map((c) => (c > 0 ? c : 0)));
  if (out === 0 || back === 0) return null;
  return ((-back * (1 + reinvestRate) ** n) / (out * (1 + financeRate))) ** (1 / (n - 1)) - 1;
}

// Depreciation, each period's: the declining balance at `factor` times
// the straight-line rate (DDB), switching to straight line over what is
// left once that is more (VDB, unless told not to), never below salvage.
export function depreciation(cost: number, salvage: number, life: number, factor: number, toLine: boolean): number[] {
  const out: number[] = [];
  let book = cost;
  for (let i = 0; i < Math.ceil(life); i++) {
    let d = Math.min((book * factor) / life, Math.max(0, book - salvage));
    if (toLine) {
      const left = life - i, line = left > 0 ? (book - salvage) / left : 0;
      if (line > d) d = line;
    }
    d = Math.min(d, Math.max(0, book - salvage));
    out.push(d);
    book -= d;
  }
  return out;
}

// From `start` to `end` periods in (fractions of a period counted pro
// rata within it).
export function vdb(cost: number, salvage: number, life: number, start: number, end: number, factor = 2, noSwitch = false): number {
  return depreciation(cost, salvage, life, factor, !noSwitch).reduce((s, d, i) => s + Math.max(0, Math.min(end, i + 1) - Math.max(start, i)) * d, 0);
}

// The fixed-declining balance, at a rate rounded to three places, the
// first year of `month` months and a last part year to make it up.
export function db(cost: number, salvage: number, life: number, period: number, month = 12): number {
  const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
  let total = 0, d = 0;
  for (let p = 1; p <= period; p++) {
    d = p === 1 ? (cost * rate * month) / 12 : p === Math.floor(life) + 1 ? ((cost - total) * rate * (12 - month)) / 12 : (cost - total) * rate;
    total += d;
  }
  return d;
}
