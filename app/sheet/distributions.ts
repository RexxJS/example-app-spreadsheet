// The probability distributions behind the statistics functions (2.8):
// the gamma function and its log, the regularised incomplete gamma and
// beta functions, and from them each distribution's density, its
// cumulative probability, and an inverse found by bisection. Pure maths:
// a domain error is a RangeError, which functions.ts makes #NUM!.

const ITERATIONS = 300;
const EPS = 1e-15;
const TINY = 1e-300;

const LANCZOS = [
  0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313,
  -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
];

// ln Γ(x), for x > 0 (Lanczos, g = 7).
export function lnGamma(x: number): number {
  if (!(x > 0)) throw new RangeError(`the log of gamma of ${x}: it must be positive`);
  if (x < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * x)) - lnGamma(1 - x);
  const z = x - 1;
  let a = LANCZOS[0];
  const t = z + 7.5;
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (z + i);
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
}

// Γ(x): by reflection below 1/2; a pole (0, -1, -2 ...) is a RangeError.
export function gamma(x: number): number {
  if (x <= 0 && Number.isInteger(x)) throw new RangeError(`gamma of ${x}: a pole`);
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
  if (x > 171.6) throw new RangeError(`gamma of ${x} overflows`);
  return Math.exp(lnGamma(x));
}

// P(a, x), the regularised lower incomplete gamma function: a series
// below a + 1, a continued fraction (Lentz) above.
export function lowerGamma(a: number, x: number): number {
  if (x <= 0) return 0;
  const lead = a * Math.log(x) - x - lnGamma(a);
  if (x < a + 1) {
    let term = 1 / a, total = term;
    for (let n = 1; n < ITERATIONS; n++) {
      term *= x / (a + n);
      total += term;
      if (Math.abs(term) < Math.abs(total) * EPS) break;
    }
    return Math.min(1, total * Math.exp(lead));
  }
  let b = x + 1 - a, c = 1 / TINY, d = 1 / b, h = d;
  for (let i = 1; i < ITERATIONS; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < TINY) d = TINY;
    c = b + an / c;
    if (Math.abs(c) < TINY) c = TINY;
    d = 1 / d;
    const step = d * c;
    h *= step;
    if (Math.abs(step - 1) < EPS) break;
  }
  return Math.max(0, 1 - Math.exp(lead) * h);
}

// The continued fraction of the incomplete beta function.
function betaFraction(a: number, b: number, x: number): number {
  const qab = a + b, qap = a + 1, qam = a - 1;
  let c = 1, d = 1 - (qab * x) / qap;
  if (Math.abs(d) < TINY) d = TINY;
  d = 1 / d;
  let h = d;
  for (let m = 1; m < ITERATIONS; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < TINY) d = TINY;
    c = 1 + aa / c;
    if (Math.abs(c) < TINY) c = TINY;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < TINY) d = TINY;
    c = 1 + aa / c;
    if (Math.abs(c) < TINY) c = TINY;
    d = 1 / d;
    const step = d * c;
    h *= step;
    if (Math.abs(step - 1) < EPS) break;
  }
  return h;
}

// I_x(a, b), the regularised incomplete beta function.
export function incompleteBeta(x: number, a: number, b: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const front = Math.exp(lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x));
  return x < (a + 1) / (a + b + 2) ? (front * betaFraction(a, b, x)) / a : 1 - (front * betaFraction(b, a, 1 - x)) / b;
}

// The x in [lo, hi] where the increasing `f` reaches p, by bisection.
export function inverse(f: (x: number) => number, p: number, lo: number, hi: number): number {
  for (let i = 0; i < 400 && hi - lo > 1e-15 * Math.max(1, Math.abs(lo) + Math.abs(hi)); i++) {
    const mid = (lo + hi) / 2;
    if (f(mid) < p) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

// A bound past which an increasing `f` is at least p, doubling from `start`.
export function upTo(f: (x: number) => number, p: number, start = 1): number {
  let hi = start;
  for (let i = 0; i < 2000 && f(hi) < p; i++) hi *= 2;
  return hi;
}

export const erf = (x: number): number => (x < 0 ? -lowerGamma(0.5, x * x) : lowerGamma(0.5, x * x));

// The standard normal's density and cumulative probability.
export const normalDensity = (z: number): number => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
export function normalCdf(z: number): number {
  // Far into the tails the complement keeps its precision.
  const e = lowerGamma(0.5, (z * z) / 2);
  return z < 0 ? (1 - e) / 2 : (1 + e) / 2;
}
export function normalInverse(p: number): number {
  if (!(p > 0 && p < 1)) throw new RangeError(`the normal inverse of ${p}: it must be between 0 and 1`);
  return inverse(normalCdf, p, -40, 40);
}

// Student's t: the probability that T is at most t.
export function tCdf(t: number, df: number): number {
  const tail = 0.5 * incompleteBeta(df / (df + t * t), df / 2, 0.5);
  return t >= 0 ? 1 - tail : tail;
}
export const tDensity = (t: number, df: number): number =>
  Math.exp(lnGamma((df + 1) / 2) - lnGamma(df / 2) - 0.5 * Math.log(df * Math.PI) - ((df + 1) / 2) * Math.log(1 + (t * t) / df));
export function tInverse(p: number, df: number): number {
  const f = (t: number) => tCdf(t, df);
  return p < 0.5 ? -inverse(f, 1 - p, 0, upTo(f, 1 - p)) : inverse(f, p, 0, upTo(f, p));
}

// χ² on df degrees of freedom.
export const chiCdf = (x: number, df: number): number => lowerGamma(df / 2, x / 2);
export const chiDensity = (x: number, df: number): number =>
  x <= 0 ? 0 : Math.exp((df / 2 - 1) * Math.log(x) - x / 2 - (df / 2) * Math.log(2) - lnGamma(df / 2));

// Gamma with shape α and scale β.
export const gammaCdf = (x: number, alpha: number, beta: number): number => lowerGamma(alpha, x / beta);
export const gammaDensity = (x: number, alpha: number, beta: number): number =>
  x <= 0 ? 0 : Math.exp((alpha - 1) * Math.log(x) - x / beta - lnGamma(alpha) - alpha * Math.log(beta));

// F on d1 and d2 degrees of freedom.
export const fCdf = (x: number, d1: number, d2: number): number => (x <= 0 ? 0 : incompleteBeta((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2));
export const fDensity = (x: number, d1: number, d2: number): number =>
  x <= 0 ? 0 : Math.exp(0.5 * (d1 * Math.log(d1 * x) + d2 * Math.log(d2) - (d1 + d2) * Math.log(d1 * x + d2)) - Math.log(x) - (lnGamma(d1 / 2) + lnGamma(d2 / 2) - lnGamma((d1 + d2) / 2)));

// Beta with shapes α and β, on [0, 1].
export const betaDensity = (x: number, a: number, b: number): number =>
  x <= 0 || x >= 1 ? 0 : Math.exp((a - 1) * Math.log(x) + (b - 1) * Math.log(1 - x) - (lnGamma(a) + lnGamma(b) - lnGamma(a + b)));

// ln C(n, k).
export const lnChoose = (n: number, k: number): number => lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1);

export function binomialMass(k: number, n: number, p: number): number {
  if (k < 0 || k > n) return 0;
  if (p === 0) return k === 0 ? 1 : 0;
  if (p === 1) return k === n ? 1 : 0;
  return Math.exp(lnChoose(n, k) + k * Math.log(p) + (n - k) * Math.log(1 - p));
}
export function binomialCdf(k: number, n: number, p: number): number {
  let total = 0;
  for (let i = 0; i <= k; i++) total += binomialMass(i, n, p);
  return Math.min(1, total);
}

export const poissonMass = (k: number, mean: number): number => (mean === 0 ? (k === 0 ? 1 : 0) : Math.exp(k * Math.log(mean) - mean - lnGamma(k + 1)));
export function poissonCdf(k: number, mean: number): number {
  let total = 0;
  for (let i = 0; i <= k; i++) total += poissonMass(i, mean);
  return Math.min(1, total);
}

// s successes in a sample of n, drawn from a population of N holding M.
export function hypergeometricMass(s: number, n: number, m: number, pop: number): number {
  if (s < Math.max(0, n - pop + m) || s > Math.min(n, m)) return 0;
  return Math.exp(lnChoose(m, s) + lnChoose(pop - m, n - s) - lnChoose(pop, n));
}

// f failures before the s-th success, each trial a success with p.
export const negativeBinomialMass = (f: number, s: number, p: number): number =>
  Math.exp(lnChoose(f + s - 1, s - 1) + s * Math.log(p) + f * Math.log(1 - p));
