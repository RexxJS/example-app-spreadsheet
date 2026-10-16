// The function library: each function is declared with how many arguments
// it takes, and gets them as Args, a range's cells row by row with the
// range's width (VLOOKUP and INDEX need its shape), or one value. An
// argument that is an error arrives as that error, so IFERROR and ISERROR
// can look at it; everything else rethrows it as soon as it reads it.
import { CellError, array, compare, fail, finite, isArray, scalar, toNumber, toText } from "./values.ts";
import { numberFormat } from "./numfmt.ts";
import type { ArrayValue, Comparison, Result, Value } from "./values.ts";
import { sortOrder } from "./sortfind.ts";
import { sparkline } from "./charts.ts";
import { addMonths, dateDif, days360, endOfMonth, isoWeek, now, parseDate, parseTime, parts, serial, weekOfYear, weekendOf, workdayFrom, workdaysBetween, yearFraction } from "./dates.ts";
import { fv, irr, nper, npv, pmt, pv, rate } from "./finance.ts";
import { colName, sheetPrefix } from "./address.ts";

export interface Arg {
  readonly values: readonly Value[];
  readonly cols: number;
}

export const single = (v: Value): Arg => ({ values: [v], cols: 1 });

// A function gives one value, or (SEQUENCE, FILTER ...) a dynamic array.
type Fn = (args: readonly Arg[]) => Value | ArrayValue;

interface Spec {
  readonly min: number;
  readonly max: number;
  readonly fn: Fn;
}

const library = new Map<string, Spec>();

const def = (name: string, min: number, max: number, fn: Fn): void => {
  library.set(name, { min, max, fn });
};

const plural = (n: number) => `${n} value${n === 1 ? "" : "s"}`;
const takes = (min: number, max: number): string =>
  max === min ? plural(min) : max === Infinity ? `at least ${plural(min)}` : max === min + 1 ? `${min} or ${plural(max)}` : `${min} to ${plural(max)}`;

// Call a function by name; IF is the evaluator's, since it must not work
// out the branch it does not take.
export function call(name: string, args: readonly Arg[]): Result {
  const spec = library.get(name);
  if (!spec) return fail("#NAME?", `no function ${name}`);
  if (args.length < spec.min || args.length > spec.max) return fail("#VALUE!", `${name} takes ${takes(spec.min, spec.max)}`);
  return lifted(spec.fn, args, new Set());
}

// An argument that is a range where the function takes one value (val()).
class ScalarWanted extends CellError {
  readonly index: number;
  constructor(index: number) {
    super("#VALUE!", `argument ${index + 1} is a range where one value goes`);
    this.index = index;
  }
}

// A function given a range where it takes one value is worked out once for
// each of the range's cells, as a spreadsheet's are: UPPER(A1:A3) is three
// values, LEFT(A1:A3, B1:B3) three pairs. Which arguments take one value
// is the function's own to say (val() refuses a range); the ranges given
// for them (`over`) are taken cell by cell, a one-row or one-column one
// broadcast, #N/A past an end. An error in one cell is that cell's.
function lifted(fn: Fn, args: readonly Arg[], over: ReadonlySet<number>): Result {
  if (over.size === 0) {
    try {
      return fn(args);
    } catch (e) {
      if (!(e instanceof ScalarWanted)) throw e;
      return lifted(fn, args, new Set([e.index]));
    }
  }
  const shapes = [...over].map((i) => ({ i, rows: args[i].values.length / args[i].cols, cols: args[i].cols }));
  const rows = Math.max(...shapes.map((x) => x.rows)), cols = Math.max(...shapes.map((x) => x.cols));
  const cell = (x: { i: number; rows: number; cols: number }, r: number, c: number): Value => {
    const rr = x.rows === 1 ? 0 : r, cc = x.cols === 1 ? 0 : c;
    return rr < x.rows && cc < x.cols ? args[x.i].values[rr * x.cols + cc] : new CellError("#N/A", "past the end of an array");
  };
  const values: Value[] = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const at = args.map((a, i) => (over.has(i) ? single(cell(shapes.find((x) => x.i === i)!, r, c)) : a));
      try {
        const v = fn(at);
        values.push(isArray(v) ? (v.values.length === 1 ? v.values[0] : new CellError("#CALC!", "an array in a cell of an array")) : v);
      } catch (e) {
        if (e instanceof ScalarWanted && !over.has(e.index)) return lifted(fn, args, new Set([...over, e.index]));
        if (!(e instanceof CellError)) throw e;
        values.push(e);
      }
    }
  }
  return array(rows, cols, values);
}

// The evaluator's own forms count too (IF; ROW and COLUMN read references;
// OFFSET and INDIRECT make them).
export const names = (): string[] => [...library.keys(), "IF", "ROW", "COLUMN", "OFFSET", "INDIRECT", "FORMULATEXT", "ISFORMULA", "ISREF", "SHEET", "SHEETS", "INFO", "SUBTOTAL", "AGGREGATE", "GROUPBY", "PIVOTBY"].sort();

// --- reading arguments ---

// Whether argument i was left out (missing, or written as nothing: f(a,,c)).
const absent = (args: readonly Arg[], i: number): boolean => !args[i] || args[i].values.length === 0;

// Argument i as one value (null when it was left out).
function val(args: readonly Arg[], i: number): Value {
  const a = args[i];
  if (absent(args, i)) return null;
  if (a.values.length !== 1) throw new ScalarWanted(i);
  return a.values[0];
}

const num = (args: readonly Arg[], i: number, otherwise?: number): number =>
  absent(args, i) && otherwise !== undefined ? otherwise : toNumber(val(args, i));
const str = (args: readonly Arg[], i: number): string => toText(val(args, i));
const every = (args: readonly Arg[]): Value[] => args.flatMap((a) => a.values);
const numbers = (args: readonly Arg[]): number[] =>
  every(args).map(scalar).filter((v): v is number => typeof v === "number");
const sum = (ns: readonly number[]): number => ns.reduce((s, n) => s + n, 0);
const blank = (v: Value): boolean => v === null || v === "";

// Rounding away from zero (as spreadsheets do: ROUND(-2.5, 0) is -3), up or
// down, to `digits` places; toPrecision(15) first so 2.675 * 100 is 267.5.
function rounder(how: (x: number) => number): Fn {
  return (args) => {
    const x = num(args, 0), f = 10 ** Math.trunc(num(args, 1, 0));
    return (Math.sign(x) * how(Number((Math.abs(x) * f).toPrecision(15)))) / f;
  };
}

// A COUNTIF criterion: 5, "apples", ">5", "<>apples", "" (blank).
export function criterion(c: Value): (v: Value) => boolean {
  if (typeof c !== "string") {
    const want = scalar(c);
    return (v) => typeof v === "number" && compare("=", v, want);
  }
  const m = /^(<>|<=|>=|=|<|>)?([\s\S]*)$/.exec(c)!;
  const op = (m[1] ?? "=") as Comparison, text = m[2];
  const rhs: Value = text.trim() !== "" && !Number.isNaN(Number(text)) ? Number(text) : text;
  return (v) => {
    if (v instanceof CellError) return false;
    if (rhs === "") return (op === "=") === blank(v);
    if (typeof rhs === "number" ? typeof v !== "number" : typeof v !== "string") return op === "<>";
    return compare(op, v, rhs);
  };
}

function ifs(args: readonly Arg[]): number[] {
  const test = criterion(val(args, 1));
  const from = args[0].values, of = (args[2] ?? args[0]).values;
  return from.flatMap((v, i) => {
    if (!test(v)) return [];
    const s = scalar(of[i] ?? null);
    return typeof s === "number" ? [s] : [];
  });
}

const notFound = (x: Value): never => fail("#N/A", `${toText(x)} is not there`);

// The position in `list` MATCH means by type 0 (equal), 1 (the largest at
// most x, in ascending order) or -1 (the smallest at least x, descending).
function position(x: Value, list: readonly Value[], type: number): number {
  if (type === 0) return list.findIndex((v) => !(v instanceof CellError) && v !== null && compare("=", v, x));
  const keep: Comparison = type > 0 ? "<=" : ">=";
  let at = -1;
  for (const [i, v] of list.entries()) {
    if (v === null || v instanceof CellError) continue;
    if (!compare(keep, v, x)) break;
    at = i;
  }
  return at;
}

// --- maths and statistics ---

def("SUM", 1, Infinity, (args) => sum(numbers(args)));
def("PRODUCT", 1, Infinity, (args) => {
  const ns = numbers(args);
  return ns.length ? finite(ns.reduce((p, n) => p * n, 1)) : 0;
});
def("AVERAGE", 1, Infinity, (args) => {
  const ns = numbers(args);
  return ns.length ? sum(ns) / ns.length : fail("#DIV/0!", "AVERAGE of no numbers");
});
def("MIN", 1, Infinity, (args) => { const ns = numbers(args); return ns.length ? Math.min(...ns) : 0; });
def("MAX", 1, Infinity, (args) => { const ns = numbers(args); return ns.length ? Math.max(...ns) : 0; });
def("MEDIAN", 1, Infinity, (args) => {
  const ns = numbers(args).sort((a, b) => a - b), mid = ns.length >> 1;
  if (!ns.length) return fail("#NUM!", "MEDIAN of no numbers");
  return ns.length % 2 ? ns[mid] : (ns[mid - 1] + ns[mid]) / 2;
});
def("STDEV", 1, Infinity, (args) => {
  const ns = numbers(args);
  if (ns.length < 2) return fail("#DIV/0!", "STDEV needs two numbers or more");
  const mean = sum(ns) / ns.length;
  return Math.sqrt(sum(ns.map((n) => (n - mean) ** 2)) / (ns.length - 1));
});
def("COUNT", 1, Infinity, (args) => numbers(args).length);
// The A forms count every cell that holds something: TRUE 1, FALSE and
// text 0, an error itself.
const everything = (args: readonly Arg[]): number[] => every(args).map(scalar).flatMap((v) => (v === null ? [] : typeof v === "number" ? [v] : v === true ? [1] : [0]));
def("MAXA", 1, Infinity, (args) => { const ns = everything(args); return ns.length ? Math.max(...ns) : 0; });
def("MINA", 1, Infinity, (args) => { const ns = everything(args); return ns.length ? Math.min(...ns) : 0; });
def("AVERAGEA", 1, Infinity, (args) => {
  const ns = everything(args);
  return ns.length ? sum(ns) / ns.length : fail("#DIV/0!", "AVERAGEA of nothing");
});
def("COUNTA", 1, Infinity, (args) => every(args).filter((v) => v !== null).length);
def("COUNTBLANK", 1, 1, (args) => args[0].values.filter(blank).length);
def("ABS", 1, 1, (args) => Math.abs(num(args, 0)));
def("SIGN", 1, 1, (args) => Math.sign(num(args, 0)));
def("INT", 1, 1, (args) => Math.floor(num(args, 0)));
def("ROUND", 1, 2, rounder(Math.round));
def("ROUNDUP", 1, 2, rounder(Math.ceil));
def("ROUNDDOWN", 1, 2, rounder(Math.floor));
def("MOD", 2, 2, (args) => {
  const a = num(args, 0), b = num(args, 1);
  return b === 0 ? fail("#DIV/0!", "MOD by zero") : a - b * Math.floor(a / b);
});
def("POWER", 2, 2, (args) => finite(num(args, 0) ** num(args, 1)));
def("SQRT", 1, 1, (args) => {
  const x = num(args, 0);
  return x < 0 ? fail("#NUM!", "the square root of a negative number") : Math.sqrt(x);
});
def("PI", 0, 0, () => Math.PI);

// A result that must be a real number: NaN or infinite is #NUM!.
const real = (n: number, what: string): number => (Number.isFinite(n) ? n : fail("#NUM!", what));
// x / sig as a count of multiples, settled to 15 digits so 2.675 / 0.01 is
// 267.5 and 0.3 / 0.1 is 3.
const multiples = (x: number, sig: number): number => Number((x / sig).toPrecision(15));
// Rounding to a multiple of `sig`, by `how` over the count of multiples.
const toMultiple = (x: number, sig: number, how: (n: number) => number): number => (sig === 0 ? 0 : how(multiples(x, sig)) * sig);

// CEILING(x, sig) and FLOOR(x, sig), the old ones: to a multiple of sig,
// away from zero (CEILING) or toward it (FLOOR); x and sig of different
// signs is #NUM!; FLOOR by 0 is #DIV/0!, CEILING by 0 is 0.
def("CEILING", 2, 2, (args) => {
  const x = num(args, 0), sig = num(args, 1);
  if (x * sig < 0) return fail("#NUM!", "CEILING's number and significance differ in sign");
  return toMultiple(x, sig, Math.ceil);
});
def("FLOOR", 2, 2, (args) => {
  const x = num(args, 0), sig = num(args, 1);
  if (sig === 0) return fail("#DIV/0!", "FLOOR by a significance of 0");
  if (x * sig < 0) return fail("#NUM!", "FLOOR's number and significance differ in sign");
  return toMultiple(x, sig, Math.floor);
});
// CEILING.MATH(x, [sig], [mode]) rounds up (toward +infinity); a negative
// x with a mode other than 0 rounds away from zero instead. FLOOR.MATH
// rounds down (toward -infinity); a negative x with a mode rounds toward
// zero. The significance's sign does not matter. The .PRECISE pair are the
// same with no mode.
const ceilingMath = (args: readonly Arg[], mode: boolean): number => {
  const x = num(args, 0), sig = Math.abs(num(args, 1, 1)), away = mode && num(args, 2, 0) !== 0;
  return toMultiple(x, sig, (n) => (x < 0 && away ? Math.floor(n) : Math.ceil(n)));
};
const floorMath = (args: readonly Arg[], mode: boolean): number => {
  const x = num(args, 0), sig = Math.abs(num(args, 1, 1)), toward = mode && num(args, 2, 0) !== 0;
  return toMultiple(x, sig, (n) => (x < 0 && toward ? Math.ceil(n) : Math.floor(n)));
};
def("CEILING.MATH", 1, 3, (args) => ceilingMath(args, true));
def("CEILING.PRECISE", 1, 2, (args) => ceilingMath(args, false));
def("FLOOR.MATH", 1, 3, (args) => floorMath(args, true));
def("FLOOR.PRECISE", 1, 2, (args) => floorMath(args, false));
// TRUNC(x, [digits]): the digits kept, the rest cut off (toward zero).
def("TRUNC", 1, 2, (args) => {
  const x = num(args, 0), f = 10 ** Math.trunc(num(args, 1, 0));
  return Math.trunc(Number((x * f).toPrecision(15))) / f;
});
// MROUND(x, m): to the nearest multiple of m, halves away from zero.
def("MROUND", 2, 2, (args) => {
  const x = num(args, 0), m = num(args, 1);
  if (x * m < 0) return fail("#NUM!", "MROUND's number and multiple differ in sign");
  return toMultiple(Math.abs(x), Math.abs(m), Math.round) * Math.sign(x);
});
// EVEN and ODD: to the next even or odd integer away from zero.
def("EVEN", 1, 1, (args) => {
  const x = num(args, 0), n = Math.ceil(Math.abs(x));
  return (n % 2 === 0 ? n : n + 1) * (x < 0 ? -1 : 1);
});
def("ODD", 1, 1, (args) => {
  const x = num(args, 0), n = Math.ceil(Math.abs(x));
  return (n % 2 === 1 ? n : n + 1) * (x < 0 ? -1 : 1);
});
def("EXP", 1, 1, (args) => real(Math.exp(num(args, 0)), "EXP overflows"));
def("LN", 1, 1, (args) => {
  const x = num(args, 0);
  return x > 0 ? Math.log(x) : fail("#NUM!", `LN of ${x}: not positive`);
});
def("LOG10", 1, 1, (args) => {
  const x = num(args, 0);
  return x > 0 ? Math.log10(x) : fail("#NUM!", `LOG10 of ${x}: not positive`);
});
// LOG(x, [base]): base 10 unless told; a base of 1 or less is #NUM!.
def("LOG", 1, 2, (args) => {
  const x = num(args, 0), b = num(args, 1, 10);
  if (x <= 0) return fail("#NUM!", `LOG of ${x}: not positive`);
  if (b <= 0 || b === 1) return fail("#NUM!", `LOG to base ${b}`);
  return Number((Math.log(x) / Math.log(b)).toPrecision(15));
});
def("QUOTIENT", 2, 2, (args) => {
  const n = num(args, 0), d = num(args, 1);
  return d === 0 ? fail("#DIV/0!", "QUOTIENT by zero") : Math.trunc(n / d);
});
// GCD and LCM over whole numbers (fractions cut off); a negative is #NUM!.
const wholes = (args: readonly Arg[], name: string): number[] =>
  numbers(args).map((n) => (n < 0 ? fail("#NUM!", `${name} of a negative number, ${n}`) : Math.trunc(n)));
const gcd2 = (a: number, b: number): number => (b === 0 ? a : gcd2(b, a % b));
def("GCD", 1, Infinity, (args) => wholes(args, "GCD").reduce(gcd2, 0));
def("LCM", 1, Infinity, (args) => wholes(args, "LCM").reduce((a, b) => (a === 0 || b === 0 ? 0 : (a * b) / gcd2(a, b)), 1));
const factorial = (n: number): number => {
  let f = 1;
  for (let i = 2; i <= n; i++) f *= i;
  return f;
};
def("FACT", 1, 1, (args) => {
  const n = Math.trunc(num(args, 0));
  return n < 0 ? fail("#NUM!", `FACT of ${n}: negative`) : real(factorial(n), `FACT of ${n} overflows`);
});
// COMBIN(n, k): the ways to choose k of n; PERMUT(n, k): in order.
def("COMBIN", 2, 2, (args) => {
  const n = Math.trunc(num(args, 0)), k = Math.trunc(num(args, 1));
  if (n < 0 || k < 0 || k > n) return fail("#NUM!", `COMBIN of ${k} from ${n}`);
  let c = 1;
  for (let i = 1; i <= Math.min(k, n - k); i++) c = (c * (n - i + 1)) / i;
  return real(Math.round(c), "COMBIN overflows");
});
def("PERMUT", 2, 2, (args) => {
  const n = Math.trunc(num(args, 0)), k = Math.trunc(num(args, 1));
  if (n < 0 || k < 0 || k > n) return fail("#NUM!", `PERMUT of ${k} from ${n}`);
  let p = 1;
  for (let i = 0; i < k; i++) p *= n - i;
  return real(p, "PERMUT overflows");
});
def("SUMSQ", 1, Infinity, (args) => sum(numbers(args).map((n) => n * n)));
// RAND, RANDBETWEEN and RANDARRAY draw afresh on every change (volatile:
// evaluate.ts lists them).
def("RAND", 0, 0, () => Math.random());
def("RANDBETWEEN", 2, 2, (args) => {
  const lo = Math.ceil(num(args, 0)), hi = Math.floor(num(args, 1));
  return lo > hi ? fail("#NUM!", `RANDBETWEEN ${lo} and ${hi}: nothing between`) : lo + Math.floor(Math.random() * (hi - lo + 1));
});
// RANDARRAY([rows], [cols], [min], [max], [whole]): a grid of draws, 0 to
// 1 unless told the range, whole numbers when asked.
def("RANDARRAY", 0, 5, (args) => {
  const rows = Math.trunc(num(args, 0, 1)), cols = Math.trunc(num(args, 1, 1));
  const lo = num(args, 2, 0), hi = num(args, 3, 1), whole = toNumber(val(args, 4)) !== 0;
  if (rows < 1 || cols < 1) return fail("#VALUE!", "RANDARRAY of no rows or columns");
  if (lo > hi) return fail("#VALUE!", `RANDARRAY from ${lo} to ${hi}: nothing between`);
  const draw = whole ? () => Math.ceil(lo) + Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : () => lo + Math.random() * (hi - lo);
  const values = Array.from({ length: rows * cols }, draw);
  return values.length === 1 ? values[0] : array(rows, cols, values);
});
// Trigonometry, in radians; a value outside a function's domain is #NUM!.
const trig = (name: string, f: (x: number) => number): void => {
  def(name, 1, 1, (args) => real(f(num(args, 0)), `${name} of ${num(args, 0)}: outside its domain`));
};
trig("SIN", Math.sin); trig("COS", Math.cos); trig("TAN", Math.tan);
trig("ASIN", Math.asin); trig("ACOS", Math.acos); trig("ATAN", Math.atan);
trig("SINH", Math.sinh); trig("COSH", Math.cosh); trig("TANH", Math.tanh);
trig("ASINH", Math.asinh); trig("ACOSH", Math.acosh); trig("ATANH", Math.atanh);
// ATAN2(x, y): the angle of the point (x, y), as Excel orders them.
def("ATAN2", 2, 2, (args) => {
  const x = num(args, 0), y = num(args, 1);
  return x === 0 && y === 0 ? fail("#DIV/0!", "ATAN2 of the origin") : Math.atan2(y, x);
});
def("DEGREES", 1, 1, (args) => (num(args, 0) * 180) / Math.PI);
def("RADIANS", 1, 1, (args) => (num(args, 0) * Math.PI) / 180);

// --- conditional sums and counts ---

def("SUMIF", 2, 3, (args) => sum(ifs(args)));
def("COUNTIF", 2, 2, (args) => args[0].values.filter(criterion(val(args, 1))).length);
def("AVERAGEIF", 2, 3, (args) => {
  const ns = ifs(args);
  return ns.length ? sum(ns) / ns.length : fail("#DIV/0!", "AVERAGEIF matched no numbers");
});

// --- logic and information ---

// The text "TRUE" or "FALSE" (any case) as the value, or null for other text.
const truthText = (s: string): boolean | null => {
  const t = s.trim().toUpperCase();
  return t === "TRUE" ? true : t === "FALSE" ? false : null;
};
// AND's and OR's operands as TRUE or FALSE: numbers by whether they are 0,
// the text "TRUE" and "FALSE" as those, other text and empty cells skipped.
const truths = (name: string, args: readonly Arg[]): boolean[] => {
  const ts = every(args).map(scalar).flatMap((v): boolean[] => {
    if (v === null) return [];
    if (typeof v === "string") { const t = truthText(v); return t === null ? [] : [t]; }
    return [toNumber(v) !== 0];
  });
  return ts.length ? ts : fail("#VALUE!", `${name} found no TRUE or FALSE`);
};
def("AND", 1, Infinity, (args) => truths("AND", args).every(Boolean));
def("OR", 1, Infinity, (args) => truths("OR", args).some(Boolean));
// XOR: TRUE when an odd number of its operands are.
def("XOR", 1, Infinity, (args) => truths("XOR", args).filter(Boolean).length % 2 === 1);
def("TRUE", 0, 0, () => true);
def("FALSE", 0, 0, () => false);
def("NOT", 1, 1, (args) => {
  const v = scalar(val(args, 0)), t = typeof v === "string" ? truthText(v) : null;
  return t !== null ? !t : num(args, 0) === 0;
});
def("IFERROR", 2, 2, (args) => {
  const x = val(args, 0);
  return x instanceof CellError ? val(args, 1) : x;
});
def("ISERROR", 1, 1, (args) => val(args, 0) instanceof CellError);
def("ISNUMBER", 1, 1, (args) => typeof val(args, 0) === "number");
def("ISTEXT", 1, 1, (args) => typeof val(args, 0) === "string");
def("ISBLANK", 1, 1, (args) => val(args, 0) === null);
// The rest of the IS family: an error other than #N/A (ISERR), #N/A
// (ISNA), TRUE or FALSE (ISLOGICAL), anything but text (ISNONTEXT: an
// empty cell is not text); ISEVEN and ISODD of the whole part of a number.
def("ISERR", 1, 1, (args) => { const v = val(args, 0); return v instanceof CellError && v.code !== "#N/A"; });
def("ISNA", 1, 1, (args) => { const v = val(args, 0); return v instanceof CellError && v.code === "#N/A"; });
def("ISLOGICAL", 1, 1, (args) => typeof val(args, 0) === "boolean");
def("ISNONTEXT", 1, 1, (args) => typeof val(args, 0) !== "string");
def("ISEVEN", 1, 1, (args) => Math.trunc(num(args, 0)) % 2 === 0);
def("ISODD", 1, 1, (args) => Math.trunc(num(args, 0)) % 2 !== 0);
// N(x): a number as itself, TRUE 1, FALSE and text 0, an error itself;
// T(x): text as itself, anything else "".
def("N", 1, 1, (args) => {
  const v = scalar(val(args, 0));
  return typeof v === "number" ? v : v === true ? 1 : 0;
});
def("T", 1, 1, (args) => { const v = scalar(val(args, 0)); return typeof v === "string" ? v : ""; });
// TYPE(x): 1 a number, 2 text, 4 TRUE or FALSE, 16 an error, 64 an array;
// an empty cell counts as a number.
def("TYPE", 1, 1, (args) => {
  if (args[0].values.length > 1) return 64;
  const v = val(args, 0);
  return v instanceof CellError ? 16 : typeof v === "string" ? 2 : typeof v === "boolean" ? 4 : 1;
});
// ERROR.TYPE(x): Excel's number for the error, #N/A for a value that is
// not one. #CYCLE! and #ERROR! (a formula that does not parse) are this
// app's: numbered past Excel's.
const ERROR_TYPES: Readonly<Record<string, number>> = {
  "#DIV/0!": 2, "#VALUE!": 3, "#REF!": 4, "#NAME?": 5, "#NUM!": 6, "#N/A": 7, "#SPILL!": 9, "#CALC!": 14, "#CYCLE!": 101, "#ERROR!": 102,
};
def("ERROR.TYPE", 1, 1, (args) => {
  const v = val(args, 0);
  return v instanceof CellError ? ERROR_TYPES[v.code] ?? fail("#N/A", `${v.code} has no number`) : fail("#N/A", "ERROR.TYPE of a value that is not an error");
});
def("CHOOSE", 2, Infinity, (args) => {
  const i = Math.trunc(num(args, 0));
  return i >= 1 && i < args.length ? val(args, i) : fail("#VALUE!", `CHOOSE has no choice ${i}`);
});

// --- text ---

def("CONCAT", 1, Infinity, (args) => every(args).map(toText).join(""));
def("TEXTJOIN", 3, Infinity, (args) => {
  const skip = num(args, 1) !== 0;
  return every(args.slice(2)).map(toText).filter((s) => !(skip && s === "")).join(str(args, 0));
});
// A range's numbers as a line of blocks (as Google Sheets' SPARKLINE, in text).
def("SPARKLINE", 1, 1, (args) => sparkline(numbers(args)));
def("LEN", 1, 1, (args) => str(args, 0).length);
def("UPPER", 1, 1, (args) => str(args, 0).toUpperCase());
def("LOWER", 1, 1, (args) => str(args, 0).toLowerCase());
def("PROPER", 1, 1, (args) => str(args, 0).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (_, before: string, c: string) => before + c.toUpperCase()));
def("TRIM", 1, 1, (args) => str(args, 0).trim().replace(/ +/g, " "));
def("LEFT", 1, 2, (args) => {
  const n = num(args, 1, 1);
  return n < 0 ? fail("#VALUE!", "LEFT of fewer than no characters") : str(args, 0).slice(0, n);
});
def("RIGHT", 1, 2, (args) => {
  const n = num(args, 1, 1), s = str(args, 0);
  return n < 0 ? fail("#VALUE!", "RIGHT of fewer than no characters") : n === 0 ? "" : s.slice(-n);
});
def("MID", 3, 3, (args) => {
  const start = num(args, 1), n = num(args, 2);
  return start < 1 || n < 0 ? fail("#VALUE!", "MID starts at 1 and takes no fewer than no characters") : str(args, 0).substr(start - 1, n);
});
def("FIND", 2, 3, (args) => {
  const needle = str(args, 0), hay = str(args, 1), at = hay.indexOf(needle, num(args, 2, 1) - 1);
  return at < 0 ? fail("#VALUE!", `'${needle}' is not in '${hay}'`) : at + 1;
});
def("SUBSTITUTE", 3, 3, (args) => {
  const old = str(args, 1);
  return old === "" ? str(args, 0) : str(args, 0).split(old).join(str(args, 2));
});
def("REPT", 2, 2, (args) => {
  const n = Math.trunc(num(args, 1));
  return n < 0 ? fail("#VALUE!", "REPT fewer than no times") : str(args, 0).repeat(n);
});
def("EXACT", 2, 2, (args) => str(args, 0) === str(args, 1));
def("VALUE", 1, 1, (args) => num(args, 0));

// CODE and CHAR over the first 255 codes; UNICODE and UNICHAR over all.
def("CODE", 1, 1, (args) => {
  const s = str(args, 0);
  return s === "" ? fail("#VALUE!", "CODE of no text") : s.charCodeAt(0);
});
def("CHAR", 1, 1, (args) => {
  const n = Math.trunc(num(args, 0));
  return n >= 1 && n <= 255 ? String.fromCharCode(n) : fail("#VALUE!", `CHAR(${n}): 1 to 255`);
});
def("UNICODE", 1, 1, (args) => {
  const s = str(args, 0);
  return s === "" ? fail("#VALUE!", "UNICODE of no text") : s.codePointAt(0)!;
});
def("UNICHAR", 1, 1, (args) => {
  const n = Math.trunc(num(args, 0));
  if (n < 1 || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) return fail("#VALUE!", `UNICHAR(${n}): not a character`);
  return String.fromCodePoint(n);
});
// CLEAN: the control characters (codes 0 to 31) taken out.
def("CLEAN", 1, 1, (args) => str(args, 0).replace(/[\x00-\x1f]/g, ""));
// A number to `decimals` places (a negative count rounds left of the
// point), away from zero at a half, with thousands separated unless not.
function fixed(x: number, decimals: number, commas: boolean): string {
  const d = Math.trunc(decimals), f = 10 ** Math.abs(d);
  const rounded = d >= 0 ? Math.round(Number((Math.abs(x) * f).toPrecision(15))) / f : Math.round(Math.abs(x) / f) * f;
  const [whole, frac = ""] = rounded.toFixed(Math.max(d, 0)).split(".");
  const grouped = commas ? whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",") : whole;
  return `${x < 0 && rounded !== 0 ? "-" : ""}${grouped}${d > 0 ? `.${frac}` : ""}`;
}
def("FIXED", 1, 3, (args) => fixed(num(args, 0), num(args, 1, 2), !(args.length > 2 && toNumber(val(args, 2)) !== 0)));
// DOLLAR: as FIXED with commas, a $ in front, a negative in parentheses.
def("DOLLAR", 1, 2, (args) => {
  const x = num(args, 0), s = fixed(Math.abs(x), num(args, 1, 2), true);
  return x < 0 && s !== "0" && !/^0\.0*$/.test(s) ? `($${s})` : `$${s}`;
});
// SEARCH(find, within, [start]): where the text is, counted from 1,
// without regard to case; ? stands for one character, * for any run, ~
// makes the next one plain. Not there is #VALUE!.
const searchPattern = (pattern: string): RegExp => {
  let out = "";
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === "~" && i + 1 < pattern.length) out += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    else if (c === "?") out += ".";
    else if (c === "*") out += ".*?";
    else out += c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(out, "i");
};
def("SEARCH", 2, 3, (args) => {
  const needle = str(args, 0), hay = str(args, 1), start = Math.trunc(num(args, 2, 1));
  if (start < 1 || start > hay.length + 1) return fail("#VALUE!", `SEARCH from ${start}: outside the text`);
  const at = hay.slice(start - 1).search(searchPattern(needle));
  return at < 0 ? fail("#VALUE!", `'${needle}' is not in '${hay}'`) : at + start;
});
// REPLACE(old, start, count, new): that many characters from `start`
// (counted from 1) replaced by the new text.
def("REPLACE", 4, 4, (args) => {
  const old = str(args, 0), start = Math.trunc(num(args, 1)), count = Math.trunc(num(args, 2));
  if (start < 1 || count < 0) return fail("#VALUE!", `REPLACE from ${start}, ${count} characters`);
  return old.slice(0, start - 1) + str(args, 3) + old.slice(start - 1 + count);
});
// CONCATENATE joins values (a range goes cell by cell, as CONCAT joins a
// range whole).
def("CONCATENATE", 1, Infinity, (args) => args.map((_, i) => toText(val(args, i))).join(""));
// NUMBERVALUE(text, [decimal], [group]): text as a number, with the
// separators given (the locale's, "." and ",", unless told); spaces are
// ignored, and a % at the end divides by 100 (each one).
def("NUMBERVALUE", 1, 3, (args) => {
  const text = str(args, 0), dec = (args.length > 1 ? str(args, 1) : ".")[0] ?? ".", grp = (args.length > 2 ? str(args, 2) : ",")[0] ?? ",";
  if (text.trim() === "") return 0;
  if (dec === grp) return fail("#VALUE!", "NUMBERVALUE's separators are the same");
  let s = text.replace(/\s+/g, ""), pct = 0;
  while (s.endsWith("%")) { s = s.slice(0, -1); pct++; }
  if (s.includes(grp) && s.includes(dec) && s.lastIndexOf(grp) > s.indexOf(dec)) return fail("#VALUE!", `'${text}' is not a number`);
  s = s.split(grp).join("").split(dec).join(".");
  if (!/^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/.test(s)) return fail("#VALUE!", `'${text}' is not a number`);
  return Number(s) / 100 ** pct;
});
// A value as text in the strict form: text in quotes, the rest as shown.
const strict = (v: Value): string => (typeof v === "string" ? `"${v.replace(/"/g, '""')}"` : v instanceof CellError ? v.code : toText(v));
const shown = (v: Value): string => (v instanceof CellError ? v.code : toText(v));
// VALUETOTEXT(value, [format]): the value as text, concise (0) or strict
// (1: text in quotes). ARRAYTOTEXT(array, [format]): its cells joined by
// ", " and rows by "; " (concise), or in braces as an array constant.
def("VALUETOTEXT", 1, 2, (args) => (num(args, 1, 0) === 1 ? strict : shown)(val(args, 0)));
def("ARRAYTOTEXT", 1, 2, (args) => {
  const { values, cols } = args[0], form = Math.trunc(toNumber(val(args, 1)));
  if (form !== 0 && form !== 1) return fail("#VALUE!", `ARRAYTOTEXT's format is 0 or 1, not ${form}`);
  const rows: string[] = [];
  for (let i = 0; i < values.length; i += cols) rows.push(values.slice(i, i + cols).map(form === 1 ? strict : shown).join(form === 1 ? "," : ", "));
  return form === 1 ? `{${rows.join(";")}}` : rows.join("; ");
});
// The REGEX trio on the engine's own regular expressions (ECMAScript,
// Unicode aware); a pattern that does not compile is #VALUE!. Case
// sensitivity: 0 (sensitive) unless 1.
function regex(pattern: string, flags: string): RegExp {
  try {
    return new RegExp(pattern, flags);
  } catch {
    return fail("#VALUE!", `'${pattern}' is not a regular expression`);
  }
}
const caseFlag = (args: readonly Arg[], i: number): string => (Math.trunc(num(args, i, 0)) === 1 ? "i" : "");
def("REGEXTEST", 2, 3, (args) => regex(str(args, 1), `u${caseFlag(args, 2)}`).test(str(args, 0)));
// REGEXEXTRACT(text, pattern, [mode], [case]): the first match (0), every
// match as a column (1), or the first match's groups as a row (2); none
// is #N/A.
def("REGEXEXTRACT", 2, 4, (args) => {
  const text = str(args, 0), mode = Math.trunc(num(args, 2, 0)), re = regex(str(args, 1), `gu${caseFlag(args, 3)}`);
  if (mode < 0 || mode > 2) return fail("#VALUE!", `REGEXEXTRACT's mode is 0, 1 or 2, not ${mode}`);
  const matches = [...text.matchAll(re)];
  if (!matches.length) return fail("#N/A", `'${str(args, 1)}' matches nothing in '${text}'`);
  if (mode === 1) return matches.length === 1 ? matches[0][0] : array(matches.length, 1, matches.map((m) => m[0]));
  if (mode === 0) return matches[0][0];
  const groups = matches[0].slice(1).map((g) => g ?? "");
  return groups.length === 0 ? matches[0][0] : groups.length === 1 ? groups[0] : array(1, groups.length, groups);
});
// REGEXREPLACE(text, pattern, replacement, [occurrence], [case]): every
// match replaced (0), or only the nth; $1 in the replacement is a group.
def("REGEXREPLACE", 3, 5, (args) => {
  const text = str(args, 0), rep = str(args, 2), nth = Math.trunc(num(args, 3, 0)), re = regex(str(args, 1), `gu${caseFlag(args, 4)}`);
  if (nth < 0) return fail("#VALUE!", `REGEXREPLACE's occurrence is 0 or more, not ${nth}`);
  if (nth === 0) return text.replace(re, rep);
  let seen = 0;
  return text.replace(re, (...m) => {
    seen++;
    const whole = m[0] as string;
    if (seen !== nth) return whole;
    const groups = m.slice(1, -2) as (string | undefined)[];
    return rep.replace(/\$(\d+|\$)/g, (_, k: string) => (k === "$" ? "$" : k === "0" ? whole : groups[Number(k) - 1] ?? ""));
  });
});

// TEXT: a value in an Excel number format ("0.00", "#,##0", "0%",
// "yyyy-mm-dd"); text that reads as a number is formatted as one.
def("TEXT", 2, 2, (args) => {
  const v = scalar(val(args, 0)), n = typeof v === "string" && v.trim() !== "" && !isNaN(Number(v)) ? Number(v) : v;
  return numberFormat(n, str(args, 1))?.text ?? toText(v);
});

// Where the delimiters (any of them; an array is several) are in `text`,
// left to right, none overlapping: [start, end] each.
function delimited(text: string, delims: readonly string[], anyCase: boolean): [number, number][] {
  const t = anyCase ? text.toLowerCase() : text, ds = delims.filter((d) => d !== "").map((d) => (anyCase ? d.toLowerCase() : d)).sort((x, y) => y.length - x.length);
  const out: [number, number][] = [];
  for (let i = 0; i < t.length;) {
    const d = ds.find((x) => t.startsWith(x, i));
    if (d === undefined) i++; else {
      out.push([i, i + d.length]);
      i += d.length;
    }
  }
  return out;
}
const delimiters = (a: Arg): string[] => a.values.map(toText);

// TEXTBEFORE and TEXTAFTER: the text before (after) the instance_num-th
// delimiter, counted from the end when negative; match_mode 1 ignores
// case; match_end 1 counts the text's end (start, from the end) as one;
// if_not_found, or #N/A.
const around = (before: boolean): Fn => (args) => {
  const text = str(args, 0), ds = delimiters(args[1]), n = whole(args, 2, 1);
  const anyCase = !absent(args, 3) && num(args, 3) !== 0, atEnd = !absent(args, 4) && num(args, 4) !== 0;
  if (n === 0 || Math.abs(n) > text.length + 1) return fail("#VALUE!", `${before ? "TEXTBEFORE" : "TEXTAFTER"}'s instance_num is ${n}`);
  if (ds.every((d) => d === "")) return (n > 0) === before ? "" : text;
  const found = delimited(text, ds, anyCase);
  if (atEnd) n > 0 ? found.push([text.length, text.length]) : found.unshift([0, 0]);
  const m = n > 0 ? found[n - 1] : found[found.length + n];
  if (!m) return absent(args, 5) ? fail("#N/A", `${before ? "TEXTBEFORE" : "TEXTAFTER"} found no delimiter ${n}`) : val(args, 5);
  return before ? text.slice(0, m[0]) : text.slice(m[1]);
};
def("TEXTBEFORE", 2, 6, around(true));
def("TEXTAFTER", 2, 6, around(false));

// TEXTSPLIT: text into columns at col_delimiter and rows at row_delimiter
// (either can be several, either left out); ignore_empty 1 drops empty
// pieces; match_mode 1 ignores case; short rows padded with pad_with, or
// #N/A.
def("TEXTSPLIT", 2, 6, (args) => {
  const text = str(args, 0), anyCase = !absent(args, 4) && num(args, 4) !== 0, skip = !absent(args, 3) && num(args, 3) !== 0;
  const split = (s: string, ds: readonly string[]): string[] => {
    if (!ds.some((d) => d !== "")) return [s];
    const at = delimited(s, ds, anyCase), pieces: string[] = [];
    let from = 0;
    for (const [a, b] of at) {
      pieces.push(s.slice(from, a));
      from = b;
    }
    pieces.push(s.slice(from));
    return skip ? pieces.filter((p) => p !== "") : pieces;
  };
  const cols = absent(args, 1) ? [] : delimiters(args[1]), rowDs = absent(args, 2) ? [] : delimiters(args[2]);
  if (!cols.some((d) => d !== "") && !rowDs.some((d) => d !== "")) return fail("#VALUE!", "TEXTSPLIT needs a delimiter");
  const lines = split(text, rowDs).map((line) => split(line, cols)).filter((l) => l.length);
  if (!lines.length) return fail("#CALC!", "TEXTSPLIT found nothing");
  const w = Math.max(...lines.map((l) => l.length)), pad = absent(args, 5) ? NA() : val(args, 5);
  return fromRows(lines.map((l) => [...l, ...Array.from({ length: w - l.length }, () => pad)]));
});

// --- conditionals ---

def("IFS", 2, Infinity, (args) => {
  for (let i = 0; i + 1 < args.length; i += 2) if (toNumber(val(args, i)) !== 0) return val(args, i + 1);
  return fail("#N/A", "no condition was true");
});
def("SWITCH", 3, Infinity, (args) => {
  const x = scalar(val(args, 0));
  let i = 1;
  for (; i + 1 < args.length; i += 2) if (compare("=", x, scalar(val(args, i)))) return val(args, i + 1);
  return i < args.length ? val(args, i) : fail("#N/A", `${toText(x)} matched no case`);
});
def("IFNA", 2, 2, (args) => {
  const v = val(args, 0);
  return v instanceof CellError && v.code === "#N/A" ? val(args, 1) : v;
});
def("NA", 0, 0, () => fail("#N/A", "NA()"));

// --- conditional aggregates over several criteria ---

// The values of `of` where every (range, criterion) pair holds.
function where(name: string, of: Arg, pairs: readonly Arg[]): Value[] {
  const tests: [readonly Value[], (v: Value) => boolean][] = [];
  for (let i = 0; i + 1 < pairs.length; i += 2) {
    if (pairs[i].values.length !== of.values.length) return fail("#VALUE!", `${name}'s ranges differ in size`);
    tests.push([pairs[i].values, criterion(scalar(pairs[i + 1].values[0] ?? null))]);
  }
  return of.values.filter((_, j) => tests.every(([vs, ok]) => ok(vs[j])));
}
const numbersIn = (vs: readonly Value[]): number[] => vs.map(scalar).filter((v): v is number => typeof v === "number");

def("SUMIFS", 3, Infinity, (args) => sum(numbersIn(where("SUMIFS", args[0], args.slice(1)))));
def("COUNTIFS", 2, Infinity, (args) => where("COUNTIFS", args[0], args).length);
def("AVERAGEIFS", 3, Infinity, (args) => {
  const ns = numbersIn(where("AVERAGEIFS", args[0], args.slice(1)));
  return ns.length ? sum(ns) / ns.length : fail("#DIV/0!", "AVERAGEIFS matched no numbers");
});
def("MAXIFS", 3, Infinity, (args) => {
  const ns = numbersIn(where("MAXIFS", args[0], args.slice(1)));
  return ns.length ? Math.max(...ns) : 0;
});
def("MINIFS", 3, Infinity, (args) => {
  const ns = numbersIn(where("MINIFS", args[0], args.slice(1)));
  return ns.length ? Math.min(...ns) : 0;
});
def("SUMPRODUCT", 1, Infinity, (args) => {
  const n = args[0].values.length;
  if (args.some((a) => a.values.length !== n)) return fail("#VALUE!", "SUMPRODUCT's ranges differ in size");
  let total = 0;
  for (let i = 0; i < n; i++) {
    const vs = args.map((a) => scalar(a.values[i]));
    total += vs.every((v) => typeof v === "number") ? vs.reduce((p, v) => p * (v as number), 1) : 0;
  }
  return total;
});

// --- order statistics ---

const sorted = (args: readonly Arg[], i: number): number[] => numbersIn(args[i].values).sort((x, y) => x - y);
function kth(name: string, largest: boolean): Fn {
  return (args) => {
    const ns = sorted(args, 0), k = Math.trunc(num(args, 1));
    if (k < 1 || k > ns.length) return fail("#NUM!", `${name}'s k is ${k} of ${ns.length} numbers`);
    return largest ? ns[ns.length - k] : ns[k - 1];
  };
}
def("LARGE", 2, 2, kth("LARGE", true));
def("SMALL", 2, 2, kth("SMALL", false));
def("RANK", 2, 3, (args) => {
  const x = num(args, 0), ns = sorted(args, 1), ascending = num(args, 2, 0) !== 0;
  if (!ns.includes(x)) return notFound(x);
  return 1 + ns.filter((n) => (ascending ? n < x : n > x)).length;
});
// PERCENTILE (inclusive): k from 0 to 1, between the numbers in order.
function percentile(ns: readonly number[], k: number, name: string): number {
  if (!ns.length) return fail("#NUM!", `${name} of no numbers`);
  if (k < 0 || k > 1) return fail("#NUM!", `${name}'s k is outside 0 to 1`);
  const at = k * (ns.length - 1), lo = Math.floor(at);
  return lo + 1 < ns.length ? ns[lo] + (at - lo) * (ns[lo + 1] - ns[lo]) : ns[lo];
}
def("PERCENTILE", 2, 2, (args) => percentile(sorted(args, 0), num(args, 1), "PERCENTILE"));
def("QUARTILE", 2, 2, (args) => {
  const q = Math.trunc(num(args, 1));
  return q < 0 || q > 4 ? fail("#NUM!", "QUARTILE's quart is 0 to 4") : percentile(sorted(args, 0), q / 4, "QUARTILE");
});
// The .INC pair are PERCENTILE and QUARTILE; the .EXC pair leave the ends
// out: the k-th point of n + 1, which must fall within the numbers.
const exclusive = (ns: number[], k: number): number => {
  ns.sort((a, b) => a - b);
  const rank = k * (ns.length + 1);
  if (rank < 1 || rank > ns.length) return fail("#NUM!", `the ${k} point lies outside the numbers`);
  const lo = Math.floor(rank), frac = rank - lo;
  return frac === 0 ? ns[lo - 1] : ns[lo - 1] + frac * (ns[lo] - ns[lo - 1]);
};
def("PERCENTILE.INC", 2, 2, (args) => call("PERCENTILE", args));
def("QUARTILE.INC", 2, 2, (args) => call("QUARTILE", args));
def("PERCENTILE.EXC", 2, 2, (args) => {
  const k = num(args, 1);
  if (k <= 0 || k >= 1) return fail("#NUM!", `PERCENTILE.EXC's k is between 0 and 1, not ${k}`);
  return exclusive(numbers([args[0]]), k);
});
def("QUARTILE.EXC", 2, 2, (args) => {
  const q = Math.trunc(num(args, 1));
  if (q < 1 || q > 3) return fail("#NUM!", `QUARTILE.EXC's quart is 1, 2 or 3, not ${q}`);
  return exclusive(numbers([args[0]]), q / 4);
});
def("MODE", 1, Infinity, (args) => {
  const ns = numbers(args), counts = new Map<number, number>();
  for (const n of ns) counts.set(n, (counts.get(n) ?? 0) + 1);
  let best: number | null = null;
  for (const n of ns) if (counts.get(n)! > 1 && (best === null || counts.get(n)! > counts.get(best)!)) best = n;
  return best ?? fail("#N/A", "MODE: no number repeats");
});
const squares = (ns: readonly number[]): number => {
  const m = sum(ns) / ns.length;
  return ns.reduce((s, n) => s + (n - m) ** 2, 0);
};
def("VAR", 1, Infinity, (args) => {
  const ns = numbers(args);
  return ns.length < 2 ? fail("#DIV/0!", "VAR needs two numbers or more") : squares(ns) / (ns.length - 1);
});
def("VARP", 1, Infinity, (args) => {
  const ns = numbers(args);
  return ns.length ? squares(ns) / ns.length : fail("#DIV/0!", "VARP of no numbers");
});
def("STDEVP", 1, Infinity, (args) => {
  const ns = numbers(args);
  return ns.length ? Math.sqrt(squares(ns) / ns.length) : fail("#DIV/0!", "STDEVP of no numbers");
});

// --- dates and times (dates.ts) ---

// Argument i as a serial: a number, or a date written as text.
function date(args: readonly Arg[], i: number): number {
  const v = scalar(val(args, i));
  if (typeof v === "string") return parseDate(v) ?? fail("#VALUE!", `'${v}' is not a date`);
  return toNumber(v);
}
const holidays = (args: readonly Arg[], i: number): Set<number> =>
  new Set(absent(args, i) ? [] : args[i].values.filter((v) => v !== null).map((v) => Math.floor(typeof v === "string" ? parseDate(v) ?? NaN : toNumber(v))));

def("DATE", 3, 3, (args) => {
  const y = Math.trunc(num(args, 0));
  return serial(y < 1900 ? y + 1900 : y, Math.trunc(num(args, 1)), Math.trunc(num(args, 2)));
});
def("DATEVALUE", 1, 1, (args) => date(args, 0));
def("YEAR", 1, 1, (args) => parts(date(args, 0)).y);
def("MONTH", 1, 1, (args) => parts(date(args, 0)).m);
def("DAY", 1, 1, (args) => parts(date(args, 0)).d);
def("WEEKDAY", 1, 2, (args) => {
  const w = parts(date(args, 0)).weekday, type = num(args, 1, 1);
  return type === 2 ? ((w + 6) % 7) + 1 : type === 3 ? (w + 6) % 7 : w + 1;
});
def("EDATE", 2, 2, (args) => addMonths(date(args, 0), Math.trunc(num(args, 1))));
def("EOMONTH", 2, 2, (args) => endOfMonth(date(args, 0), Math.trunc(num(args, 1))));
def("DAYS", 2, 2, (args) => Math.floor(date(args, 0)) - Math.floor(date(args, 1)));
def("DATEDIF", 3, 3, (args) => {
  const a = date(args, 0), b = date(args, 1);
  if (a > b) return fail("#NUM!", "DATEDIF's start is after its end");
  return dateDif(a, b, str(args, 2)) ?? fail("#NUM!", `${str(args, 2)} is not a unit (Y, M, D, YM, MD, YD)`);
});
def("NETWORKDAYS", 2, 3, (args) => workdaysBetween(date(args, 0), date(args, 1), holidays(args, 2)));
def("WORKDAY", 2, 3, (args) => workdayFrom(date(args, 0), Math.trunc(num(args, 1)), holidays(args, 2)));
// The .INTL pair take which days are the weekend (dates.ts weekendOf):
// a number, or seven 1s and 0s from Monday; Saturday and Sunday unless told.
const weekendArg = (args: readonly Arg[], i: number) => {
  const v = absent(args, i) ? 1 : scalar(val(args, i));
  const w = weekendOf(typeof v === "string" ? v : toNumber(v));
  return w ?? fail("#NUM!", `${toText(v)} is not a weekend: 1 to 7, 11 to 17, or seven 1s and 0s`);
};
def("NETWORKDAYS.INTL", 2, 4, (args) => workdaysBetween(date(args, 0), date(args, 1), holidays(args, 3), weekendArg(args, 2)));
def("WORKDAY.INTL", 2, 4, (args) => workdayFrom(date(args, 0), Math.trunc(num(args, 1)), holidays(args, 3), weekendArg(args, 2)));
// WEEKNUM(date, [type]): the week of the year, 1 at the week January 1 is
// in; type 1 (or 17) has weeks start on Sunday, 2 (or 11) Monday, 12
// Tuesday ... 16 Saturday; 21 is ISO 8601's week (ISOWEEKNUM).
def("WEEKNUM", 1, 2, (args) => {
  const s = date(args, 0), type = Math.trunc(num(args, 1, 1));
  if (type === 21) return isoWeek(s);
  const start = type === 1 || type === 17 ? 0 : type === 2 || type === 11 ? 1 : type >= 12 && type <= 16 ? type - 10 : -1;
  return start < 0 ? fail("#NUM!", `WEEKNUM's type is 1, 2, 11 to 17 or 21, not ${type}`) : weekOfYear(s, start);
});
def("ISOWEEKNUM", 1, 1, (args) => isoWeek(date(args, 0)));
// YEARFRAC(start, end, [basis]) and DAYS360(start, end, [european]):
// dates.ts says the conventions.
def("YEARFRAC", 2, 3, (args) => {
  const basis = Math.trunc(num(args, 2, 0));
  return yearFraction(date(args, 0), date(args, 1), basis) ?? fail("#NUM!", `YEARFRAC's basis is 0 to 4, not ${basis}`);
});
def("DAYS360", 2, 3, (args) => days360(date(args, 0), date(args, 1), !absent(args, 2) && toNumber(val(args, 2)) !== 0));
// TIMEVALUE(text): the time of day as the fraction of a day.
def("TIMEVALUE", 1, 1, (args) => {
  const s = str(args, 0);
  return parseTime(s) ?? fail("#VALUE!", `'${s}' is not a time`);
});
def("TIME", 3, 3, (args) => ((num(args, 0) * 3600 + num(args, 1) * 60 + num(args, 2)) / 86400) % 1);
const clock = (args: readonly Arg[]) => Math.round((date(args, 0) % 1) * 86400);
def("HOUR", 1, 1, (args) => Math.floor(clock(args) / 3600) % 24);
def("MINUTE", 1, 1, (args) => Math.floor(clock(args) / 60) % 60);
def("SECOND", 1, 1, (args) => clock(args) % 60);
def("TODAY", 0, 0, () => Math.floor(now()));
def("NOW", 0, 0, () => now());

// --- money (finance.ts) ---

def("PMT", 3, 5, (args) => finite(pmt(num(args, 0), num(args, 1), num(args, 2), num(args, 3, 0), num(args, 4, 0))));
def("FV", 3, 5, (args) => finite(fv(num(args, 0), num(args, 1), num(args, 2), num(args, 3, 0), num(args, 4, 0))));
def("PV", 3, 5, (args) => finite(pv(num(args, 0), num(args, 1), num(args, 2), num(args, 3, 0), num(args, 4, 0))));
def("NPER", 3, 5, (args) => finite(nper(num(args, 0), num(args, 1), num(args, 2), num(args, 3, 0), num(args, 4, 0))));
def("RATE", 3, 6, (args) =>
  rate(num(args, 0), num(args, 1), num(args, 2), num(args, 3, 0), num(args, 4, 0), num(args, 5, 0.1)) ?? fail("#NUM!", "RATE did not settle"));
def("NPV", 2, Infinity, (args) => npv(num(args, 0), numbers(args.slice(1))));
def("IRR", 1, 2, (args) => {
  const flows = numbersIn(args[0].values);
  if (!flows.some((c) => c < 0) || !flows.some((c) => c > 0)) return fail("#NUM!", "IRR needs a negative and a positive cash flow");
  return irr(flows, num(args, 1, 0.1)) ?? fail("#NUM!", "IRR did not settle");
});

// --- dynamic arrays ---

const rowsOf = (a: Arg): Value[][] => Array.from({ length: a.values.length / a.cols }, (_, i) => a.values.slice(i * a.cols, (i + 1) * a.cols));
const fromRows = (rows: readonly (readonly Value[])[]): ArrayValue => array(rows.length, rows[0]?.length ?? 0, rows.flat());
const transposed = (a: Arg): Arg => {
  const rows = a.values.length / a.cols;
  return { values: Array.from({ length: a.values.length }, (_, k) => a.values[(k % rows) * a.cols + Math.floor(k / rows)]), cols: rows };
};
const turned = (x: ArrayValue): ArrayValue => {
  const t = transposed({ values: x.values, cols: x.cols });
  return array(t.values.length / t.cols, t.cols, t.values);
};

def("SEQUENCE", 1, 4, (args) => {
  const rows = Math.trunc(num(args, 0)), cols = Math.trunc(num(args, 1, 1)), start = num(args, 2, 1), step = num(args, 3, 1);
  if (rows < 1 || cols < 1) return fail("#CALC!", `SEQUENCE of no ${rows < 1 ? "rows" : "columns"}`);
  if (rows * cols > 10_000) return fail("#NUM!", "SEQUENCE of more than 10000 values");
  return array(rows, cols, Array.from({ length: rows * cols }, (_, i) => start + step * i));
});
def("FILTER", 2, 3, (args) => {
  const a = args[0], include = args[1], rows = a.values.length / a.cols;
  const keep = include.values.map((v) => toNumber(v) !== 0);
  // A column of TRUE and FALSE keeps rows; a row of them keeps columns.
  const byRow = include.cols === 1 && keep.length === rows;
  if (!byRow && !(keep.length === a.cols && include.cols === keep.length)) return fail("#VALUE!", "FILTER's include is not the array's height or width");
  const lines = rowsOf(byRow ? a : transposed(a)).filter((_, i) => keep[i]);
  if (!lines.length) return absent(args, 2) ? fail("#CALC!", "FILTER found nothing") : val(args, 2);
  return byRow ? fromRows(lines) : turned(fromRows(lines));
});
// Rows in order of a key each (ascending, or not), equal ones as they were.
const ordered = <T>(items: readonly T[], key: (t: T) => Value, ascending: boolean): T[] =>
  items.map((t, i) => ({ t, i })).sort((p, q) => sortOrder(key(p.t), key(q.t), ascending) || p.i - q.i).map((x) => x.t);
def("SORT", 1, 4, (args) => {
  const byCol = !absent(args, 3) && num(args, 3) !== 0, a = byCol ? transposed(args[0]) : args[0];
  const index = Math.trunc(num(args, 1, 1)), ascending = num(args, 2, 1) >= 0;
  if (index < 1 || index > a.cols) return fail("#VALUE!", `SORT's index ${index} is outside the array`);
  const out = fromRows(ordered(rowsOf(a), (row) => row[index - 1], ascending));
  return byCol ? turned(out) : out;
});
def("SORTBY", 2, Infinity, (args) => {
  const rows = rowsOf(args[0]), keys: { by: readonly Value[]; ascending: boolean }[] = [];
  for (let i = 1; i < args.length; i += 2) {
    if (args[i].values.length !== rows.length) return fail("#VALUE!", "SORTBY's arrays differ in size");
    keys.push({ by: args[i].values, ascending: num(args, i + 1, 1) >= 0 });
  }
  const order = rows.map((_, i) => i).sort((p, q) => {
    for (const k of keys) {
      const c = sortOrder(k.by[p], k.by[q], k.ascending);
      if (c) return c;
    }
    return p - q;
  });
  return fromRows(order.map((i) => rows[i]));
});
def("UNIQUE", 1, 3, (args) => {
  const byCol = !absent(args, 1) && num(args, 1) !== 0, once = !absent(args, 2) && num(args, 2) !== 0;
  const lines = rowsOf(byCol ? transposed(args[0]) : args[0]);
  const keyOf = (line: readonly Value[]) => JSON.stringify(line.map((v) => (v instanceof CellError ? v.code : toText(v).toUpperCase())));
  const counts = new Map<string, number>();
  for (const line of lines) counts.set(keyOf(line), (counts.get(keyOf(line)) ?? 0) + 1);
  const seen = new Set<string>();
  const kept = lines.filter((line) => {
    const k = keyOf(line);
    if (seen.has(k) || (once && counts.get(k)! > 1)) return false;
    seen.add(k);
    return true;
  });
  if (!kept.length) return fail("#CALC!", "UNIQUE found nothing");
  return byCol ? turned(fromRows(kept)) : fromRows(kept);
});
def("TRANSPOSE", 1, 1, (args) => turned(array(args[0].values.length / args[0].cols, args[0].cols, args[0].values)));

// --- shaping arrays ---

const NA = (): CellError => new CellError("#N/A", "past the end of an array");
// A count argument: a whole number.
const whole = (args: readonly Arg[], i: number, otherwise: number): number => Math.trunc(num(args, i, otherwise));
// TAKE keeps the first rows and columns (the last, when negative); DROP
// keeps all but them.
const cut = (take: boolean): Fn => (args) => {
  const a = args[0], rows = a.values.length / a.cols;
  const part = (n: number, len: number): [number, number] =>
    take ? (n >= 0 ? [0, Math.min(n, len)] : [Math.max(0, len + n), len]) : n >= 0 ? [Math.min(n, len), len] : [0, Math.max(0, len + n)];
  const [r0, r1] = part(whole(args, 1, take ? rows : 0), rows), [c0, c1] = part(whole(args, 2, take ? a.cols : 0), a.cols);
  if (r1 <= r0 || c1 <= c0) return fail("#CALC!", `${take ? "TAKE" : "DROP"} leaves nothing`);
  return fromRows(rowsOf(a).slice(r0, r1).map((row) => row.slice(c0, c1)));
};
def("TAKE", 2, 3, cut(true));
def("DROP", 2, 3, cut(false));
// CHOOSEROWS and CHOOSECOLS: the rows (columns) by number, from the end
// when negative, in the order asked, as often as asked.
const choose = (byCol: boolean): Fn => (args) => {
  const lines = rowsOf(byCol ? transposed(args[0]) : args[0]);
  const picked = args.slice(1).flatMap((a) => a.values).map((v) => {
    const n = Math.trunc(toNumber(v)), i = n < 0 ? lines.length + n : n - 1;
    return i >= 0 && i < lines.length ? lines[i] : fail("#VALUE!", `${byCol ? "CHOOSECOLS" : "CHOOSEROWS"} has no ${byCol ? "column" : "row"} ${n}`);
  });
  return byCol ? turned(fromRows(picked)) : fromRows(picked);
};
def("CHOOSEROWS", 2, Infinity, choose(false));
def("CHOOSECOLS", 2, Infinity, choose(true));
// VSTACK: the arrays one under another, as wide as the widest, #N/A where
// one is narrower; HSTACK, side by side.
const stacked = (args: readonly Arg[]): ArrayValue => {
  const lines = args.filter((a) => a.values.length).flatMap(rowsOf), w = Math.max(...lines.map((l) => l.length));
  return fromRows(lines.map((l) => [...l, ...Array.from({ length: w - l.length }, NA)]));
};
def("VSTACK", 1, Infinity, stacked);
def("HSTACK", 1, Infinity, (args) => turned(stacked(args.filter((a) => a.values.length).map(transposed))));
// TOCOL and TOROW: an array as one column (row), by rows or (scan_by_column)
// by columns; ignore 1 skips blanks, 2 errors, 3 both.
const flattened = (args: readonly Arg[]): Value[] => {
  const ignore = whole(args, 1, 0), byCol = !absent(args, 2) && num(args, 2) !== 0;
  const a = byCol ? transposed(args[0]) : args[0];
  return a.values.filter((v) => !((ignore & 1 && blank(v)) || (ignore & 2 && v instanceof CellError)));
};
def("TOCOL", 1, 3, (args) => {
  const vs = flattened(args);
  return vs.length ? array(vs.length, 1, vs) : fail("#CALC!", "TOCOL found nothing");
});
def("TOROW", 1, 3, (args) => {
  const vs = flattened(args);
  return vs.length ? array(1, vs.length, vs) : fail("#CALC!", "TOROW found nothing");
});
// WRAPROWS and WRAPCOLS: a row or column of values in rows (columns) of
// `count`, the last padded with pad_with or #N/A.
const wrapped = (byCol: boolean): Fn => (args) => {
  const vs = args[0].values, n = whole(args, 1, 0);
  if (args[0].cols !== 1 && args[0].cols !== vs.length) return fail("#VALUE!", "only a row or a column can be wrapped");
  if (n < 1) return fail("#NUM!", "wrapping in fewer than 1");
  const pad = absent(args, 2) ? NA() : val(args, 2);
  const lines = Array.from({ length: Math.ceil(vs.length / n) }, (_, i) => Array.from({ length: n }, (_, j) => (i * n + j < vs.length ? vs[i * n + j] : pad)));
  return byCol ? turned(fromRows(lines)) : fromRows(lines);
};
def("WRAPROWS", 2, 3, wrapped(false));
def("WRAPCOLS", 2, 3, wrapped(true));
// EXPAND: an array made as big as rows by cols, padded with pad_with or #N/A.
def("EXPAND", 2, 4, (args) => {
  const a = args[0], rows = a.values.length / a.cols;
  const r = whole(args, 1, rows), c = whole(args, 2, a.cols), pad = absent(args, 3) ? NA() : val(args, 3);
  if (r < rows || c < a.cols) return fail("#VALUE!", "EXPAND cannot make an array smaller");
  return array(r, c, Array.from({ length: r * c }, (_, k) => {
    const i = Math.floor(k / c), j = k % c;
    return i < rows && j < a.cols ? a.values[i * a.cols + j] : pad;
  }));
});

// --- lookup ---

// A wildcard pattern (* any run, ? one character, ~ before either for
// itself) as a test, ignoring case.
function wildcard(pattern: string): (v: Value) => boolean {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === "~" && i + 1 < pattern.length) re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    else if (c === "*") re += ".*";
    else if (c === "?") re += ".";
    else re += c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  const r = new RegExp(`^${re}$`, "i");
  return (v) => v !== null && !(v instanceof CellError) && r.test(toText(v));
}

// XLOOKUP's and XMATCH's search: the index in `list` of x by match mode
// (0 exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcard),
// searching first to last (1) or last to first (-1); -1 when not there.
function xsearch(x: Value, list: readonly Value[], mode: number, from: number): number {
  const order = list.map((_, i) => i);
  if (from < 0) order.reverse();
  if (mode === 2) {
    const ok = wildcard(toText(x));
    return order.find((i) => ok(list[i])) ?? -1;
  }
  const real = (v: Value) => v !== null && !(v instanceof CellError);
  const exact = order.find((i) => real(list[i]) && compare("=", list[i], x));
  if (exact !== undefined || mode === 0) return exact ?? -1;
  let best = -1;
  for (const i of order) {
    const v = list[i];
    if (!real(v) || !(mode < 0 ? compare("<", v, x) : compare(">", v, x))) continue;
    if (best < 0 || (mode < 0 ? compare(">", v, list[best]) : compare("<", v, list[best]))) best = i;
  }
  return best;
}

// XLOOKUP's return: the value at the match, or, when the return range is
// wider (taller) than a column (row) of lookups, the whole row (column).
def("XLOOKUP", 3, 6, (args) => {
  const x = scalar(val(args, 0)), list = args[1].values, back = args[2];
  const lookRows = list.length / args[1].cols, backRows = back.values.length / back.cols;
  const down = args[1].cols === 1 && backRows === list.length, across = lookRows === 1 && back.cols === list.length;
  if (!down && !across) return fail("#VALUE!", "XLOOKUP's lookup and return ranges differ in size");
  const at = xsearch(x, list, num(args, 4, 0), num(args, 5, 1));
  if (at < 0) return absent(args, 3) ? notFound(x) : val(args, 3);
  if (down) return back.cols === 1 ? back.values[at] : array(1, back.cols, back.values.slice(at * back.cols, (at + 1) * back.cols));
  return backRows === 1 ? back.values[at] : array(backRows, 1, back.values.filter((_, k) => k % back.cols === at));
});
def("XMATCH", 2, 4, (args) => {
  const x = scalar(val(args, 0)), at = xsearch(x, args[1].values, num(args, 2, 0), num(args, 3, 1));
  return at < 0 ? notFound(x) : at + 1;
});
def("HLOOKUP", 3, 4, (args) => {
  const x = scalar(val(args, 0)), { values, cols } = args[1], r = Math.trunc(num(args, 2)), rows = values.length / cols;
  if (r < 1) return fail("#VALUE!", "HLOOKUP's row starts at 1");
  if (r > rows) return fail("#REF!", `HLOOKUP row ${r} of a table ${rows} high`);
  const c = position(x, values.slice(0, cols), num(args, 3, 1) !== 0 ? 1 : 0);
  return c < 0 ? notFound(x) : values[(r - 1) * cols + c];
});

def("ROWS", 1, 1, (args) => args[0].values.length / args[0].cols);
def("COLUMNS", 1, 1, (args) => args[0].cols);
// ADDRESS(row, column, [abs], [a1], [sheet]): a cell's address as text. abs
// 1 is $A$1, 2 A$1, 3 $A1, 4 A1; a1 FALSE gives the R1C1 form, its
// relative parts in brackets (R2C[3]); a sheet is put in front, quoted when
// its name needs it.
def("ADDRESS", 2, 5, (args) => {
  const r = Math.trunc(num(args, 0)), c = Math.trunc(num(args, 1)), abs = Math.trunc(num(args, 2, 1));
  const a1 = args.length > 3 && val(args, 3) !== null ? toNumber(val(args, 3)) !== 0 : true;
  if (r < 1 || c < 1 || c > 16384) return fail("#VALUE!", `ADDRESS of row ${r}, column ${c}: not on a sheet`);
  if (abs < 1 || abs > 4) return fail("#VALUE!", `ADDRESS's abs is 1 to 4, not ${abs}`);
  const rowAbs = abs === 1 || abs === 2, colAbs = abs === 1 || abs === 3;
  const cell = a1
    ? `${colAbs ? "$" : ""}${colName(c - 1)}${rowAbs ? "$" : ""}${r}`
    : `R${rowAbs ? r : `[${r}]`}C${colAbs ? c : `[${c}]`}`;
  const sheet = args.length > 4 ? str(args, 4) : "";
  return sheet === "" ? cell : `${sheetPrefix(sheet)}${cell}`;
});
// LOOKUP(value, vector, [result]): the last place in the ascending vector
// not past `value`, read from the result vector (the vector itself when
// there is none). LOOKUP(value, array) searches along the array's longer
// side (its first row when it is wider than tall, else its first column)
// and answers from the last row or column.
def("LOOKUP", 2, 3, (args) => {
  const x = scalar(val(args, 0));
  if (args.length === 3) {
    const at = position(x, args[1].values, 1);
    if (at < 0) return notFound(x);
    return at < args[2].values.length ? args[2].values[at] : fail("#N/A", `LOOKUP's result vector is shorter than its lookup vector`);
  }
  const { values, cols } = args[1], rows = values.length / cols;
  if (cols > rows) {
    const at = position(x, values.slice(0, cols), 1);
    return at < 0 ? notFound(x) : values[(rows - 1) * cols + at];
  }
  const at = position(x, Array.from({ length: rows }, (_, i) => values[i * cols]), 1);
  return at < 0 ? notFound(x) : values[at * cols + cols - 1];
});
def("INDEX", 2, 3, (args) => {
  const { values, cols } = args[0], rows = values.length / cols;
  // INDEX(list, n) counts along a single row or column.
  let r = num(args, 1), c = num(args, 2, 1);
  if (args.length === 2 && rows === 1) [r, c] = [1, r];
  return r >= 1 && r <= rows && c >= 1 && c <= cols
    ? values[(r - 1) * cols + (c - 1)]
    : fail("#REF!", `INDEX ${r}, ${c} is outside ${rows} x ${cols}`);
});
def("MATCH", 2, 3, (args) => {
  const x = scalar(val(args, 0)), at = position(x, args[1].values, Math.sign(num(args, 2, 1)));
  return at < 0 ? notFound(x) : at + 1;
});
def("VLOOKUP", 3, 4, (args) => {
  const x = scalar(val(args, 0)), { values, cols } = args[1], c = Math.trunc(num(args, 2));
  if (c < 1) return fail("#VALUE!", "VLOOKUP's column starts at 1");
  if (c > cols) return fail("#REF!", `VLOOKUP column ${c} of a table ${cols} wide`);
  const firsts = values.filter((_, i) => i % cols === 0);
  const r = position(x, firsts, num(args, 3, 1) !== 0 ? 1 : 0);
  return r < 0 ? notFound(x) : values[r * cols + c - 1];
});
