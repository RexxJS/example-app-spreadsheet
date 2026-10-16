// The engineering functions' arithmetic (2.10): numbers in other bases
// (ten digits, two's complement for negatives, as Excel's), 48-bit
// bitwise operations without BigInt, CONVERT's units, complex numbers as
// Excel writes them ("3+4i"), and the Bessel functions by their integral
// representations. Pure; a bad argument is a RangeError (#NUM! in
// functions.ts), a value that is not one a TypeError (#VALUE!).

// --- bases ---

export interface Radix {
  readonly name: string;
  readonly base: number;
  readonly digits: RegExp;
}
export const BIN: Radix = { name: "binary", base: 2, digits: /^[01]{1,10}$/ };
export const OCT: Radix = { name: "octal", base: 8, digits: /^[0-7]{1,10}$/ };
export const HEX: Radix = { name: "hexadecimal", base: 16, digits: /^[0-9A-Fa-f]{1,10}$/ };

// Ten digits of `r` hold base^10 values; the upper half are negative.
const span = (r: Radix): number => r.base ** 10;

// Ten-digit text in a base as a number (two's complement past half).
export function fromRadix(text: string, r: Radix): number {
  const t = text.trim();
  if (t === "") return 0;
  if (!r.digits.test(t)) throw new RangeError(`'${text}' is not ${r.name}, ten digits at most`);
  const n = parseInt(t, r.base);
  return n >= span(r) / 2 ? n - span(r) : n;
}

// A whole number as text in a base: a negative as ten digits of two's
// complement (places then ignored), else padded to `places`.
export function toRadix(n: number, r: Radix, places: number | null): string {
  const x = Math.trunc(n), half = span(r) / 2;
  if (x < -half || x >= half) throw new RangeError(`${x} does not fit ten ${r.name} digits`);
  if (x < 0) return (span(r) + x).toString(r.base).toUpperCase();
  const s = x.toString(r.base).toUpperCase();
  if (places === null) return s;
  const p = Math.trunc(places);
  if (p < s.length || p > 10) throw new RangeError(`${s} does not fit ${p} places`);
  return s.padStart(p, "0");
}

// --- bits: whole numbers 0 to 2^48 - 1, as two 24-bit halves ---

const LOW = 2 ** 24, LIMIT = 2 ** 48;
export function checkBits(n: number): number {
  if (!(Number.isInteger(n) && n >= 0 && n < LIMIT)) throw new RangeError(`${n} is not a whole number from 0 to 2^48 - 1`);
  return n;
}
const halves = (n: number): [number, number] => [Math.floor(n / LOW), n % LOW];
export function bitwise(a: number, b: number, op: (x: number, y: number) => number): number {
  const [ah, al] = halves(checkBits(a)), [bh, bl] = halves(checkBits(b));
  return (op(ah, bh) >>> 0) * LOW + (op(al, bl) >>> 0);
}
export function shift(n: number, by: number): number {
  checkBits(n);
  const s = Math.trunc(by);
  if (Math.abs(s) > 53) throw new RangeError(`a shift of ${s}: at most 53`);
  const out = s >= 0 ? n * 2 ** s : Math.floor(n / 2 ** -s);
  if (out >= LIMIT) throw new RangeError(`${n} shifted ${s} passes 2^48`);
  return out;
}

// --- CONVERT ---

interface Unit {
  readonly kind: string;
  // In the kind's base unit (temperatures: their own scales, below).
  readonly factor: number;
  // Whether it takes a metric (or, for bit and byte, a binary) prefix.
  readonly prefix?: boolean;
  // The power a prefix is raised to (m2: 2, m3: 3).
  readonly power?: number;
}

const LY = 9.4607304725808e15, INCH = 0.0254, PICA_PT = INCH / 72;
const U = (kind: string, factor: number, prefix = false, power = 1): Unit => ({ kind, factor, prefix, power });
const UNITS: Readonly<Record<string, Unit>> = {
  // mass, in grams
  g: U("mass", 1, true), sg: U("mass", 14593.902937206), lbm: U("mass", 453.59237), u: U("mass", 1.66053886e-24, true), ozm: U("mass", 28.349523125),
  stone: U("mass", 6350.29318), ton: U("mass", 907184.74), grain: U("mass", 0.06479891), cwt: U("mass", 45359.237), shweight: U("mass", 45359.237),
  uk_cwt: U("mass", 50802.34544), lcwt: U("mass", 50802.34544), hweight: U("mass", 50802.34544), uk_ton: U("mass", 1016046.9088), LTON: U("mass", 1016046.9088), brton: U("mass", 1016046.9088),
  // distance, in metres
  m: U("distance", 1, true), mi: U("distance", 1609.344), Nmi: U("distance", 1852), in: U("distance", INCH), ft: U("distance", 0.3048), yd: U("distance", 0.9144),
  ang: U("distance", 1e-10, true), ell: U("distance", 1.143), ly: U("distance", LY, true), parsec: U("distance", 3.08567758128155e16, true), pc: U("distance", 3.08567758128155e16, true),
  Pica: U("distance", PICA_PT), Picapt: U("distance", PICA_PT), pica: U("distance", INCH / 6), survey_mi: U("distance", 1609.3472186944),
  // time, in seconds
  yr: U("time", 31557600), day: U("time", 86400), d: U("time", 86400), hr: U("time", 3600), mn: U("time", 60), min: U("time", 60), sec: U("time", 1, true), s: U("time", 1, true),
  // pressure, in pascals
  Pa: U("pressure", 1, true), p: U("pressure", 1, true), atm: U("pressure", 101325, true), at: U("pressure", 101325, true), mmHg: U("pressure", 133.322, true), psi: U("pressure", 6894.75729316836), Torr: U("pressure", 133.322368421053),
  // force, in newtons
  N: U("force", 1, true), dyn: U("force", 1e-5, true), dy: U("force", 1e-5, true), lbf: U("force", 4.4482216152605), pond: U("force", 0.00980665, true),
  // energy, in joules
  J: U("energy", 1, true), e: U("energy", 1e-7, true), c: U("energy", 4.184, true), cal: U("energy", 4.1868, true), eV: U("energy", 1.602176487e-19, true), ev: U("energy", 1.602176487e-19, true),
  HPh: U("energy", 2684519.53769617), hh: U("energy", 2684519.53769617), Wh: U("energy", 3600, true), wh: U("energy", 3600, true), BTU: U("energy", 1055.05585262), btu: U("energy", 1055.05585262),
  // power, in watts
  HP: U("power", 745.69987158227), h: U("power", 745.69987158227), PS: U("power", 735.49875), W: U("power", 1, true), w: U("power", 1, true),
  // magnetism, in teslas
  T: U("magnetism", 1, true), ga: U("magnetism", 1e-4, true),
  // volume, in cubic metres
  tsp: U("volume", 4.92892159375e-6), tspm: U("volume", 5e-6), tbs: U("volume", 1.478676478125e-5), oz: U("volume", 2.95735295625e-5), cup: U("volume", 2.365882365e-4),
  pt: U("volume", 4.73176473e-4), us_pt: U("volume", 4.73176473e-4), uk_pt: U("volume", 5.6826125e-4), qt: U("volume", 9.46352946e-4), uk_qt: U("volume", 1.1365225e-3),
  gal: U("volume", 3.785411784e-3), uk_gal: U("volume", 4.54609e-3), l: U("volume", 1e-3, true), L: U("volume", 1e-3, true), lt: U("volume", 1e-3, true),
  ang3: U("volume", 1e-30, true, 3), barrel: U("volume", 0.158987294928), bushel: U("volume", 0.03523907016688), ft3: U("volume", 0.3048 ** 3), in3: U("volume", INCH ** 3),
  ly3: U("volume", LY ** 3, true, 3), m3: U("volume", 1, true, 3), mi3: U("volume", 1609.344 ** 3), yd3: U("volume", 0.9144 ** 3), Nmi3: U("volume", 1852 ** 3), Pica3: U("volume", PICA_PT ** 3),
  GRT: U("volume", 2.8316846592), regton: U("volume", 2.8316846592), MTON: U("volume", 1.13267386368),
  // area, in square metres
  uk_acre: U("area", 4046.8564224), us_acre: U("area", 4046.87260987425), ang2: U("area", 1e-20, true, 2), ar: U("area", 100, true), ft2: U("area", 0.3048 ** 2), ha: U("area", 10000),
  in2: U("area", INCH ** 2), ly2: U("area", LY ** 2, true, 2), m2: U("area", 1, true, 2), Morgen: U("area", 2500), mi2: U("area", 1609.344 ** 2), Nmi2: U("area", 1852 ** 2), Pica2: U("area", PICA_PT ** 2), yd2: U("area", 0.9144 ** 2),
  // information, in bits
  bit: U("information", 1, true), byte: U("information", 8, true),
  // speed, in metres a second
  admkn: U("speed", 1853.184 / 3600), kn: U("speed", 1852 / 3600), "m/h": U("speed", 1 / 3600, true), "m/hr": U("speed", 1 / 3600, true), "m/s": U("speed", 1, true), "m/sec": U("speed", 1, true), mph: U("speed", 0.44704),
};
const PREFIXES: Readonly<Record<string, number>> = {
  Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 10, e: 10, d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24,
};
const BINARY: Readonly<Record<string, number>> = { Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10 };
// Temperatures, each to and from kelvin.
const TEMPERATURES: Readonly<Record<string, { to: (x: number) => number; from: (k: number) => number }>> = (() => {
  const c = { to: (x: number) => x + 273.15, from: (k: number) => k - 273.15 };
  const f = { to: (x: number) => ((x - 32) * 5) / 9 + 273.15, from: (k: number) => ((k - 273.15) * 9) / 5 + 32 };
  const k = { to: (x: number) => x, from: (x: number) => x };
  return { C: c, cel: c, F: f, fah: f, K: k, kel: k, Rank: { to: (x: number) => (x * 5) / 9, from: (x: number) => (x * 9) / 5 }, Reau: { to: (x: number) => x * 1.25 + 273.15, from: (x: number) => (x - 273.15) * 0.8 } };
})();

// A unit written as Excel takes it: one of the table, or a prefix and one
// that takes prefixes (km, cm2, Kibyte is not: ki is lower case).
function unitOf(name: string): Unit | null {
  if (UNITS[name]) return UNITS[name];
  for (const [p, scale] of [...Object.entries(BINARY), ...Object.entries(PREFIXES)]) {
    if (!name.startsWith(p)) continue;
    const u = UNITS[name.slice(p.length)];
    if (u?.prefix && (!(p in BINARY) || u.kind === "information")) return { ...u, factor: u.factor * scale ** (u.power ?? 1) };
  }
  return null;
}

// x in `from` as `to`; null when either is no unit, or they measure
// different things (Excel's #N/A).
export function convert(x: number, from: string, to: string): number | null {
  const tf = TEMPERATURES[from], tt = TEMPERATURES[to];
  if (tf || tt) return tf && tt ? tt.from(tf.to(x)) : null;
  const a = unitOf(from), b = unitOf(to);
  if (!a || !b || a.kind !== b.kind) return null;
  return (x * a.factor) / b.factor;
}

// --- complex numbers ---

export interface Complex {
  readonly re: number;
  readonly im: number;
  // "i" or "j", as written.
  readonly unit: "i" | "j";
}

const NUM = "[+-]?(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][+-]?\\d+)?";
const FULL = new RegExp(`^(${NUM})([+-](?:\\d+\\.?\\d*|\\.\\d+)?(?:[eE][+-]?\\d+)?)([ij])$`);
const IMAG = new RegExp(`^([+-]?(?:\\d+\\.?\\d*|\\.\\d+)?(?:[eE][+-]?\\d+)?)([ij])$`);
const REAL = new RegExp(`^${NUM}$`);

// Excel's complex text: "3+4i", "-2.5j", "i", "7", "1e3-2E-1i".
export function parseComplex(v: number | string): Complex {
  if (typeof v === "number") return { re: v, im: 0, unit: "i" };
  const t = v.trim();
  if (t === "") return { re: 0, im: 0, unit: "i" };
  const coefficient = (s: string): number => (s === "" || s === "+" ? 1 : s === "-" ? -1 : Number(s));
  let m = FULL.exec(t);
  if (m) return { re: Number(m[1]), im: coefficient(m[2]), unit: m[3] as "i" | "j" };
  m = IMAG.exec(t);
  if (m) return { re: 0, im: coefficient(m[1]), unit: m[2] as "i" | "j" };
  if (REAL.test(t)) return { re: Number(t), im: 0, unit: "i" };
  throw new RangeError(`'${v}' is not a complex number`);
}

// As Excel writes it: fifteen significant digits, a part that is 0 left
// out, a unit coefficient of 1 as the letter alone.
const part = (x: number): string => String(Number(x.toPrecision(15)));
export function formatComplex(z: Complex): string {
  const re = Number(z.re.toPrecision(15)), im = Number(z.im.toPrecision(15));
  if (!Number.isFinite(re) || !Number.isFinite(im)) throw new RangeError("the result is not a finite complex number");
  if (im === 0) return part(re);
  const coeff = im === 1 ? "" : im === -1 ? "-" : part(im);
  const imag = `${coeff}${z.unit}`;
  if (re === 0) return imag;
  return `${part(re)}${im > 0 ? "+" : ""}${imag}`;
}

const C = (re: number, im: number, unit: "i" | "j"): Complex => ({ re, im, unit });
export const cAdd = (a: Complex, b: Complex): Complex => C(a.re + b.re, a.im + b.im, a.unit);
export const cSub = (a: Complex, b: Complex): Complex => C(a.re - b.re, a.im - b.im, a.unit);
export const cMul = (a: Complex, b: Complex): Complex => C(a.re * b.re - a.im * b.im, a.re * b.im + a.im * b.re, a.unit);
export function cDiv(a: Complex, b: Complex): Complex {
  const d = b.re * b.re + b.im * b.im;
  if (d === 0) throw new RangeError("division by a complex 0");
  return C((a.re * b.re + a.im * b.im) / d, (a.im * b.re - a.re * b.im) / d, a.unit);
}
export const cAbs = (z: Complex): number => Math.hypot(z.re, z.im);
export const cArg = (z: Complex): number => Math.atan2(z.im, z.re);
export const cExp = (z: Complex): Complex => C(Math.exp(z.re) * Math.cos(z.im), Math.exp(z.re) * Math.sin(z.im), z.unit);
export function cLn(z: Complex): Complex {
  if (z.re === 0 && z.im === 0) throw new RangeError("the logarithm of 0");
  return C(Math.log(cAbs(z)), cArg(z), z.unit);
}
export function cPow(z: Complex, n: number): Complex {
  if (z.re === 0 && z.im === 0) {
    if (n <= 0) throw new RangeError("0 to a power that is not positive");
    return C(0, 0, z.unit);
  }
  const r = cAbs(z) ** n, t = cArg(z) * n;
  return C(r * Math.cos(t), r * Math.sin(t), z.unit);
}
// The principal root, by halves of |z| (no detour through polar form,
// so √-4 is 2i exactly).
export function cSqrt(z: Complex): Complex {
  const r = cAbs(z);
  return C(Math.sqrt((r + z.re) / 2), (z.im < 0 ? -1 : 1) * Math.sqrt((r - z.re) / 2), z.unit);
}
export const cSin = (z: Complex): Complex => C(Math.sin(z.re) * Math.cosh(z.im), Math.cos(z.re) * Math.sinh(z.im), z.unit);
export const cCos = (z: Complex): Complex => C(Math.cos(z.re) * Math.cosh(z.im), -Math.sin(z.re) * Math.sinh(z.im), z.unit);
export const cSinh = (z: Complex): Complex => C(Math.sinh(z.re) * Math.cos(z.im), Math.cosh(z.re) * Math.sin(z.im), z.unit);
export const cCosh = (z: Complex): Complex => C(Math.cosh(z.re) * Math.cos(z.im), Math.sinh(z.re) * Math.sin(z.im), z.unit);
const ONE = (u: "i" | "j"): Complex => C(1, 0, u);
export const cTan = (z: Complex): Complex => cDiv(cSin(z), cCos(z));
export const cCot = (z: Complex): Complex => cDiv(cCos(z), cSin(z));
export const cSec = (z: Complex): Complex => cDiv(ONE(z.unit), cCos(z));
export const cCsc = (z: Complex): Complex => cDiv(ONE(z.unit), cSin(z));
export const cSech = (z: Complex): Complex => cDiv(ONE(z.unit), cCosh(z));
export const cCsch = (z: Complex): Complex => cDiv(ONE(z.unit), cSinh(z));

// --- Bessel functions, of whole order n >= 0 ---

// ∫ f over [a, b] by Simpson's rule on `steps` (even) intervals.
function simpson(f: (t: number) => number, a: number, b: number, steps: number): number {
  const h = (b - a) / steps;
  let s = f(a) + f(b);
  for (let i = 1; i < steps; i++) s += (i % 2 ? 4 : 2) * f(a + i * h);
  return (s * h) / 3;
}
// The trapezoid rule over a whole period of a smooth periodic function
// converges faster than any power of the step.
function periodic(f: (t: number) => number, steps: number): number {
  const h = Math.PI / steps;
  let s = (f(0) + f(Math.PI)) / 2;
  for (let i = 1; i < steps; i++) s += f(i * h);
  return s * h;
}
const steps = (x: number, n: number): number => 200 + 4 * Math.ceil(Math.abs(x) + n);

// J_n(x) = (1/π) ∫_0^π cos(nτ - x sin τ) dτ.
export const besselJ = (x: number, n: number): number => periodic((t) => Math.cos(n * t - x * Math.sin(t)), steps(x, n)) / Math.PI;
// I_n(x) = (1/π) ∫_0^π e^(x cos τ) cos(nτ) dτ.
export const besselI = (x: number, n: number): number => periodic((t) => Math.exp(x * Math.cos(t)) * Math.cos(n * t), steps(x, n)) / Math.PI;
// Where e^(-x cosh t) cosh(nt) (or e^(nt - x sinh t)) has died away.
function reach(f: (t: number) => number): number {
  let t = 1;
  while (t < 60 && Math.abs(f(t)) > 1e-300) t *= 1.5;
  return t;
}
// K_n(x) = ∫_0^∞ e^(-x cosh t) cosh(nt) dt, for x > 0.
export function besselK(x: number, n: number): number {
  if (!(x > 0)) throw new RangeError(`BESSELK's x is positive, not ${x}`);
  const f = (t: number) => Math.exp(-x * Math.cosh(t)) * Math.cosh(n * t);
  return simpson(f, 0, reach(f), 20000);
}
// Y_n(x) = (1/π) ∫_0^π sin(x sin τ - nτ) dτ - (1/π) ∫_0^∞ (e^(nt) + (-1)^n e^(-nt)) e^(-x sinh t) dt, for x > 0.
export function besselY(x: number, n: number): number {
  if (!(x > 0)) throw new RangeError(`BESSELY's x is positive, not ${x}`);
  const first = simpson((t) => Math.sin(x * Math.sin(t) - n * t), 0, Math.PI, 2000);
  const tail = (t: number) => (Math.exp(n * t) + (n % 2 ? -1 : 1) * Math.exp(-n * t)) * Math.exp(-x * Math.sinh(t));
  return (first - simpson(tail, 0, reach(tail), 40000)) / Math.PI;
}
