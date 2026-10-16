// Bonds and bills (2.9b), as Excel's documented formulas work them out:
// coupons paid `freq` times a year (1, 2 or 4), dated back from maturity
// (a month's end kept as its end), and days counted on a basis: 0 US
// 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European
// 30/360. Pure arithmetic over date serials; a bad argument is a
// RangeError, which functions.ts makes #NUM!.
import { days360, parts, serial, yearFraction } from "./dates.ts";

export type Basis = 0 | 1 | 2 | 3 | 4;

export function checkBond(settle: number, maturity: number, freq: number, basis: number): void {
  if (!(settle < maturity)) throw new RangeError("settlement is on or after maturity");
  if (freq !== 1 && freq !== 2 && freq !== 4) throw new RangeError(`frequency is 1, 2 or 4, not ${freq}`);
  checkBasis(basis);
}

export function checkBasis(basis: number): asserts basis is Basis {
  if (!(basis >= 0 && basis <= 4 && Number.isInteger(basis))) throw new RangeError(`basis is 0 to 4, not ${basis}`);
}

const lastDay = (y: number, m: number): number => parts(serial(y, m + 1, 0)).d;

// The date `months` on from `from` (back when negative), a month's end
// kept as its end, else the day kept where the month has it.
export function addMonthsKeep(from: number, months: number, endOfMonth: boolean): number {
  const p = parts(from), total = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(total / 12), m = (total % 12) + 1;
  return serial(y, m, endOfMonth ? lastDay(y, m) : Math.min(p.d, lastDay(y, m)));
}

const isMonthEnd = (s: number): boolean => { const p = parts(s); return p.d === lastDay(p.y, p.m); };

// The coupon dates either side of settlement, and how many fall after it
// up to maturity.
export function coupons(settle: number, maturity: number, freq: number): { previous: number; next: number; count: number } {
  const step = 12 / freq, eom = isMonthEnd(maturity);
  let k = 1, d = addMonthsKeep(maturity, -step, eom);
  while (d > settle) d = addMonthsKeep(maturity, -step * ++k, eom);
  return { previous: d, next: addMonthsKeep(maturity, -step * (k - 1), eom), count: k };
}

// Days from a to b on a basis.
export const dayCount = (a: number, b: number, basis: Basis): number =>
  basis === 0 ? days360(a, b, false) : basis === 4 ? days360(a, b, true) : Math.floor(b) - Math.floor(a);

// A coupon period's length in days: actual for basis 1, else the year's
// days over the frequency.
export const periodDays = (from: number, to: number, freq: number, basis: Basis): number =>
  basis === 1 ? Math.floor(to) - Math.floor(from) : (basis === 3 ? 365 : 360) / freq;

// COUPDAYBS, COUPDAYS, COUPDAYSNC: days from the previous coupon to
// settlement, in the period, and on to the next (on the 30/360 bases, the
// period less what has gone, as Excel has it).
export function couponDays(settle: number, maturity: number, freq: number, basis: Basis): { before: number; period: number; after: number } {
  const c = coupons(settle, maturity, freq);
  const before = dayCount(c.previous, settle, basis), period = periodDays(c.previous, c.next, freq, basis);
  const after = basis === 0 || basis === 4 ? period - before : Math.floor(c.next) - Math.floor(settle);
  return { before, period, after };
}

// A year's days on a basis, between two dates: what the discount and
// money-market formulas divide by (basis 1, the actual year, averaged
// when the dates span more than one).
export const yearDays = (a: number, b: number, basis: Basis): number =>
  basis === 1 ? (Math.floor(b) - Math.floor(a)) / yearFraction(a, b, 1)! : basis === 3 ? 365 : 360;

// PRICE: a bond's price per 100 face, from its yield.
export function price(settle: number, maturity: number, rate: number, yld: number, redemption: number, freq: number, basis: Basis): number {
  const { before: a, period: e, after: dsc } = couponDays(settle, maturity, freq, basis), n = coupons(settle, maturity, freq).count;
  const coupon = (100 * rate) / freq, accrued = (coupon * a) / e;
  if (n === 1) return (redemption + coupon) / (1 + ((dsc / e) * yld) / freq) - accrued;
  const v = 1 + yld / freq;
  let total = redemption / v ** (n - 1 + dsc / e);
  for (let k = 1; k <= n; k++) total += coupon / v ** (k - 1 + dsc / e);
  return total - accrued;
}

// The rate at which the decreasing `f` gives `target`: Newton from a
// guess, then bisection over (-1, 10).
function solveFor(f: (y: number) => number, target: number, guess: number): number {
  let y = guess;
  for (let i = 0; i < 100; i++) {
    const v = f(y) - target, h = 1e-7, d = (f(y + h) - f(y - h)) / (2 * h);
    if (!Number.isFinite(v) || !Number.isFinite(d) || d === 0) break;
    const next = y - v / d;
    if (Math.abs(next - y) < 1e-13) return next;
    y = next;
  }
  let lo = -0.99, hi = 10;
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2;
    if (f(mid) > target) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

// YIELD: the yield that gives a price; one coupon left is a closed form.
export function yieldOf(settle: number, maturity: number, rate: number, pr: number, redemption: number, freq: number, basis: Basis): number {
  const { before: a, period: e, after: dsc } = couponDays(settle, maturity, freq, basis), n = coupons(settle, maturity, freq).count;
  if (n === 1) {
    const paid = pr / 100 + ((a / e) * rate) / freq;
    return ((redemption / 100 + rate / freq - paid) / paid) * ((freq * e) / dsc);
  }
  return solveFor((y) => price(settle, maturity, rate, y, redemption, freq, basis), pr, rate || 0.05);
}

// DURATION: Macaulay's, in years, per 100 face; MDURATION divides by
// (1 + yld / freq).
export function duration(settle: number, maturity: number, coupon: number, yld: number, freq: number, basis: Basis): number {
  const { period: e, after: dsc } = couponDays(settle, maturity, freq, basis), n = coupons(settle, maturity, freq).count;
  const v = 1 + yld / freq, c = (100 * coupon) / freq;
  let weighted = 0, total = 0;
  for (let k = 1; k <= n; k++) {
    const t = k - 1 + dsc / e, cash = c + (k === n ? 100 : 0), pv = cash / v ** t;
    weighted += t * pv;
    total += pv;
  }
  return weighted / total / freq;
}

// Interest accrued from `from` to `to` at `rate` on `par`, period by
// quasi-coupon period (dated from `anchor` by the frequency), each part
// of a period by its own days over the period's length.
export function accrued(from: number, to: number, anchor: number, rate: number, par: number, freq: number, basis: Basis): number {
  const step = 12 / freq, eom = isMonthEnd(anchor);
  let k = 0;
  while (addMonthsKeep(anchor, -step * k, eom) > from) k++;
  let total = 0;
  for (let start = addMonthsKeep(anchor, -step * k, eom); start < to; k--) {
    const end = addMonthsKeep(anchor, -step * (k - 1), eom);
    const lo = Math.max(start, from), hi = Math.min(end, to);
    if (hi > lo) total += dayCount(lo, hi, basis) / periodDays(start, end, freq, basis);
    start = end;
  }
  return (par * rate * total) / freq;
}

// The odd first period's coupon dates: quasi-coupon periods of the
// frequency, dated back from the first coupon, until one starts on or
// before `from`.
function quasi(from: number, firstCoupon: number, freq: number): number[] {
  const step = 12 / freq, eom = isMonthEnd(firstCoupon), dates = [firstCoupon];
  for (let k = 1; dates[dates.length - 1] > from; k++) dates.push(addMonthsKeep(firstCoupon, -step * k, eom));
  return dates.reverse();
}

// ODDFPRICE: a bond whose first period is odd, short or long (Excel's
// two documented forms; a short one is a single quasi period).
export function oddFirstPrice(settle: number, maturity: number, issue: number, firstCoupon: number, rate: number, yld: number, redemption: number, freq: number, basis: Basis): number {
  if (!(issue < settle && settle < firstCoupon && firstCoupon < maturity)) throw new RangeError("issue < settlement < first coupon < maturity, as an odd first period needs");
  const c = (100 * rate) / freq, v = 1 + yld / freq;
  const n = coupons(firstCoupon, maturity, freq).count; // coupons after the first, to maturity
  const periods = quasi(issue, firstCoupon, freq);
  // The quasi periods of the odd period: each one's days of interest
  // (from issue on), and of accrual (from issue to settlement).
  let dc = 0, acc = 0, nq = 0, dsc = 0, nl = 0;
  for (let i = 0; i + 1 < periods.length; i++) {
    const start = periods[i], end = periods[i + 1], len = periods.length === 2 ? periodDays(start, end, freq, basis) : periodDays(start, end, freq, basis);
    dc += dayCount(Math.max(start, issue), end, basis) / len;
    if (settle > start) acc += dayCount(Math.max(start, issue), Math.min(settle, end), basis) / len;
    if (settle < end) {
      if (settle >= start) {
        dsc = basis === 0 || basis === 4 ? len - dayCount(start, settle, basis) : Math.floor(end) - Math.floor(settle);
        nl = len;
      } else nq++;
    }
  }
  const lead = nq + dsc / nl;
  let total = redemption / v ** (n + lead) + (c * dc) / v ** lead;
  for (let k = 1; k <= n; k++) total += c / v ** (k + lead);
  return total - c * acc;
}

export const oddFirstYield = (settle: number, maturity: number, issue: number, firstCoupon: number, rate: number, pr: number, redemption: number, freq: number, basis: Basis): number =>
  solveFor((y) => oddFirstPrice(settle, maturity, issue, firstCoupon, rate, y, redemption, freq, basis), pr, rate || 0.05);

// The odd last period's quasi-coupon periods, dated on from the last
// coupon: each one's days of interest, of accrual (to settlement), and
// from settlement on.
function oddLast(settle: number, maturity: number, last: number, freq: number, basis: Basis): { dc: number; a: number; dsc: number } {
  if (!(last < settle && settle < maturity)) throw new RangeError("last coupon < settlement < maturity, as an odd last period needs");
  const step = 12 / freq, eom = isMonthEnd(last);
  let dc = 0, a = 0, dsc = 0;
  for (let k = 0, start = last; start < maturity; k++) {
    const end = addMonthsKeep(last, step * (k + 1), eom), len = periodDays(start, end, freq, basis), stop = Math.min(end, maturity);
    dc += dayCount(start, stop, basis) / len;
    if (settle > start) a += dayCount(start, Math.min(settle, stop), basis) / len;
    if (settle < stop) dsc += dayCount(Math.max(start, settle), stop, basis) / len;
    start = end;
  }
  return { dc, a, dsc };
}

export function oddLastPrice(settle: number, maturity: number, last: number, rate: number, yld: number, redemption: number, freq: number, basis: Basis): number {
  const { dc, a, dsc } = oddLast(settle, maturity, last, freq, basis), c = (100 * rate) / freq;
  return (redemption + dc * c) / (1 + (dsc * yld) / freq) - a * c;
}

export function oddLastYield(settle: number, maturity: number, last: number, rate: number, pr: number, redemption: number, freq: number, basis: Basis): number {
  const { dc, a, dsc } = oddLast(settle, maturity, last, freq, basis), c = (100 * rate) / freq;
  return ((redemption + dc * c - (pr + a * c)) / (pr + a * c)) * (freq / dsc);
}
