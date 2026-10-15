// How a formula's Expr is worked out against the sheet's cells. Values,
// errors and coercions are values.ts; the functions are functions.ts.
//
// A result may be a dynamic array (values.ts): a range, or an operator over
// one, worked out cell by cell (a row or a column broadcast across the
// other's, #N/A where one runs out), IF over one, or a function that gives
// one (SEQUENCE, FILTER ...). LET's and LAMBDA's names live in scopes here;
// a LAMBDA is a closure, a value only while a formula is worked out (one
// that comes to a LAMBDA is #CALC!), and MAP, REDUCE, SCAN, BYROW, BYCOL and
// MAKEARRAY call one: they are special forms here, as IF is. So are ROW
// and COLUMN, which read a reference rather than its values, and with no
// reference the cell the formula is in (`at`); and OFFSET and INDIRECT,
// which make a reference at run time (a Reference below), and tell the Env
// the cells it comes to (`read`), since the formula's static reads cannot
// name them; and FORMULATEXT, ISFORMULA, ISREF, SHEET and SHEETS, which
// ask about a reference (its cell's input, the Env's `input`) or the
// workbook (`book`).
import { COLS, ROWS, addrOf, cellAddr, keyOf, range, span, splitKey } from "./address.ts";
import type { Addr } from "./address.ts";
import { resolve } from "./structured.ts";
import type { Expr } from "./parser.ts";
import { call, single } from "./functions.ts";
import type { Arg } from "./functions.ts";
import { CellError, array, collapsed, compare, fail, finite, isArray, toNumber, toText } from "./values.ts";
import type { ArrayValue, Result, Value } from "./values.ts";
import type { Op } from "./lexer.ts";

// The forms worked out here rather than by the library (IF is listed there).
export const SPECIAL_FORMS: readonly string[] = ["LET", "LAMBDA", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY"];

export interface Env {
  get(key: string): Value;
  // What a cell spills (A1#), or null when it spills nothing.
  spill?(key: string): ArrayValue | null;
  // Told each cell a reference made at run time (OFFSET, INDIRECT) read.
  read?(key: string): void;
  // A cell's input as typed ("" for none); the workbook's sheets, in tab
  // order, and which this one is.
  input?(key: string): string;
  book?(): { readonly sheets: readonly string[]; readonly here: number };
}

// The forms that make a reference at run time, and the functions that
// draw a random number: a formula calling one is volatile (worked out
// again on every change).
export const VOLATILE_FORMS: ReadonlySet<string> = new Set(["OFFSET", "INDIRECT", "RAND", "RANDBETWEEN", "RANDARRAY"]);

// A reference as an expression comes to it: a rectangle of cells, on this
// sheet (null) or another (its name in upper case, as keys have it).
interface Reference {
  readonly sheet: string | null;
  readonly a: Addr;
  readonly b: Addr;
}

const qualify = (sheet: string | null, key: string): string => (sheet === null ? key : `${sheet}!${key}`);

// "B2", "$B$2", "Costs!B2", "'My Costs'!B2:C4": the cell or range the text
// names (A1 style), as INDIRECT reads it; or in R1C1 style, "R2C3",
// "R[1]C[-1]" (relative to `at`), "RC2". Null when it names none;
// off-sheet is #REF!.
function referenceOf(text: string, a1: boolean, at: Addr | null): Reference | null {
  const m = /^\s*(?:(?:([A-Za-z_]\w*)|'([^']+)')!)?([^!]+?)\s*$/.exec(text);
  if (!m) return null;
  const sheet = (m[1] ?? m[2])?.toUpperCase() ?? null;
  const parts = m[3].split(":");
  if (parts.length > 2) return null;
  const corner = (s: string): Addr | null => {
    if (a1) {
      const c = /^\$?([A-Za-z]{1,3})\$?([0-9]+)$/.exec(s);
      if (!c) return null;
      return addrOf(c[1] + c[2]) ?? fail("#REF!", `${s.toUpperCase()} is not on the sheet`);
    }
    const c = /^R(?:([0-9]+)|\[(-?[0-9]+)\])?C(?:([0-9]+)|\[(-?[0-9]+)\])?$/i.exec(s);
    if (!c) return null;
    const part = (abs: string | undefined, rel: string | undefined, here: number | null): number => {
      if (abs !== undefined) return Number(abs) - 1;
      if (here === null) return fail("#REF!", `${s} is relative, and the formula is in no cell`);
      return here + (rel === undefined ? 0 : Number(rel));
    };
    const row = part(c[1], c[2], at?.row ?? null), col = part(c[3], c[4], at?.col ?? null);
    return row >= 0 && row < ROWS && col >= 0 && col < COLS ? { row, col } : fail("#REF!", `${s} is not on the sheet`);
  };
  const a = corner(parts[0]), b = parts.length === 2 ? corner(parts[1]) : a;
  return a && b ? { sheet, a, b } : null;
}

interface Closure {
  readonly kind: "lambda";
  readonly params: readonly string[];
  readonly body: Expr;
  readonly scope: Scope;
}

// What a name in scope stands for.
type Bound = Result | Closure;
type Scope = ReadonlyMap<string, Bound>;

const isClosure = (v: Bound): v is Closure => typeof v === "object" && v !== null && (v as Closure).kind === "lambda";

// One operator on two values.
function operate(op: Op, a: Value, b: Value): Value {
  switch (op) {
    case "+": return finite(toNumber(a) + toNumber(b));
    case "-": return finite(toNumber(a) - toNumber(b));
    case "*": return finite(toNumber(a) * toNumber(b));
    case "/": return toNumber(b) === 0 ? fail("#DIV/0!", "divide by zero") : finite(toNumber(a) / toNumber(b));
    case "^": return finite(toNumber(a) ** toNumber(b));
    case "&": return toText(a) + toText(b);
    default: return compare(op, a, b);
  }
}

// `f` over values, or cell by cell over arrays: a one-row (one-column)
// array is broadcast down (across); past an array's end is #N/A; an error
// in a cell is that cell's.
function lift(inputs: readonly Result[], f: (vs: Value[]) => Value): Result {
  if (!inputs.some(isArray)) return f(inputs as Value[]);
  const grids = inputs.filter(isArray);
  const rows = Math.max(...grids.map((g) => g.rows)), cols = Math.max(...grids.map((g) => g.cols));
  const at = (x: Result, i: number, j: number): Value => {
    if (!isArray(x)) return x;
    const r = x.rows === 1 ? 0 : i, c = x.cols === 1 ? 0 : j;
    return r < x.rows && c < x.cols ? x.values[r * x.cols + c] : new CellError("#N/A", "past the end of an array");
  };
  const values: Value[] = [];
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      try {
        values.push(f(inputs.map((x) => at(x, i, j))));
      } catch (err) {
        if (!(err instanceof CellError)) throw err;
        values.push(err);
      }
    }
  }
  return array(rows, cols, values);
}

const truth = (v: Value): boolean => toNumber(v) !== 0;

// How deep LAMBDAs may call LAMBDAs (a named one calling itself) before
// the formula is #NUM!.
const DEPTH = 200;

class Evaluator {
  readonly #env: Env;
  readonly #at: string | null;
  #scope: Scope = new Map();
  #depth = 0;

  constructor(env: Env, at: string | null) {
    this.#env = env;
    this.#at = at;
  }

  // `work` in `scope`, the scope put back after.
  #within<T>(scope: Scope, work: () => T): T {
    const was = this.#scope;
    this.#scope = scope;
    try {
      return work();
    } finally {
      this.#scope = was;
    }
  }

  // Anything an expression comes to: a value, an array, or a closure.
  any(e: Expr): Bound {
    switch (e.kind) {
      case "local": {
        const v = this.#scope.get(e.name);
        return v === undefined ? fail("#NAME?", `no name ${e.name}`) : v;
      }
      case "let": {
        const scope = new Map(this.#scope);
        return this.#within(scope, () => {
          for (const [name, value] of e.bindings) scope.set(name, this.any(value));
          return this.any(e.body);
        });
      }
      case "lambda":
        return { kind: "lambda", params: e.params, body: e.body, scope: this.#scope };
      // A named formula is worked out in no scope but its own: what it
      // names is the sheet's, not the caller's.
      case "named":
        return this.#within(new Map(), () => this.any(e.expr()));
      case "apply": {
        const fn = this.any(e.fn);
        if (!isClosure(fn)) return fail("#VALUE!", "only a LAMBDA can be called");
        return this.#invoke(fn, e.args.map((a) => this.any(a)));
      }
      default:
        return this.result(e);
    }
  }

  #invoke(fn: Closure, args: readonly Bound[]): Bound {
    if (args.length !== fn.params.length) return fail("#VALUE!", `LAMBDA takes ${fn.params.length} values, given ${args.length}`);
    if (this.#depth >= DEPTH) return fail("#NUM!", `LAMBDAs call LAMBDAs more than ${DEPTH} deep`);
    const scope = new Map(fn.scope);
    fn.params.forEach((p, i) => scope.set(p, args[i]));
    this.#depth++;
    try {
      return this.#within(scope, () => this.any(fn.body));
    } finally {
      this.#depth--;
    }
  }

  // A value or an array: not a closure.
  result(e: Expr): Result {
    switch (e.kind) {
      case "num":
      case "str":
      case "bool":
        return e.value;
      case "ref":
        return cellAddr(e.name) ? this.#env.get(e.name) : fail("#REF!", `${e.name} is not on the sheet`);
      case "spill": {
        const spilt = this.#env.spill?.(e.name);
        if (spilt) return spilt;
        const v = this.#env.get(e.name);
        return v instanceof CellError ? v : fail("#REF!", `${e.name} spills nothing`);
      }
      case "range": {
        const s = span(e.from, e.to);
        if (!s) return fail("#REF!", `${cellAddr(e.from) ? e.to : e.from} is not on the sheet`);
        return array(s.keys.length / s.cols, s.cols, s.keys.map((k) => this.#env.get(k)));
      }
      // A table reference: its cells as the table is now (its header row
      // read for a column's name); one cell is its value.
      case "table": {
        const r = resolve(e.table, e.spec, (a) => toText(this.#env.get(keyOf(a))), e.row);
        if (r instanceof CellError) throw r;
        const keys = [...range(r[0], r[1])];
        return keys.length === 1 ? this.#env.get(keys[0]) : array(r[1].row - r[0].row + 1, r[1].col - r[0].col + 1, keys.map((k) => this.#env.get(k)));
      }
      case "missing":
        return null;
      // An error in an array constant is that cell's value.
      case "array":
        return array(e.rows, e.cols, e.values.map((v) => (typeof v === "object" ? new CellError(v.error, "written in the formula") : v)));
      case "local":
      case "let":
      case "lambda":
      case "apply":
      case "named": {
        const v = this.any(e);
        return isClosure(v) ? fail("#VALUE!", "a LAMBDA where a value goes") : v;
      }
      case "error":
        return fail(e.code, e.why ?? (e.code === "#REF!" ? "a reference moved off the sheet" : "written in the formula"));
      case "neg":
        return lift([this.result(e.arg)], ([x]) => -toNumber(x));
      case "binary":
        return lift([this.result(e.left), this.result(e.right)], ([a, b]) => operate(e.op, a, b));
      case "call":
        return this.#call(e);
      default:
        throw new TypeError(`unknown expression: ${JSON.stringify(e satisfies never)}`);
    }
  }

  // A function's argument: a range's or an array's cells and width, or one
  // value. An error is passed on as a value, for IFERROR and ISERROR to see.
  #arg(e: Expr): Arg {
    if (e.kind === "missing") return { values: [], cols: 0 };
    try {
      const r = this.result(e);
      return isArray(r) ? { values: r.values, cols: r.cols } : single(r);
    } catch (err) {
      if (err instanceof CellError && e.kind !== "range" && e.kind !== "table") return single(err);
      throw err;
    }
  }

  #grid(e: Expr): ArrayValue {
    const r = this.result(e);
    return isArray(r) ? r : array(1, 1, [r]);
  }

  #closure(e: Expr, name: string): Closure {
    const f = this.any(e);
    return isClosure(f) ? f : fail("#VALUE!", `${name} takes a LAMBDA last`);
  }

  // What a LAMBDA gave, as one value (an error in it is that value).
  #one(name: string, work: () => Bound): Value {
    try {
      const v = work();
      if (isClosure(v)) return fail("#CALC!", `${name}'s LAMBDA gives a LAMBDA`);
      const c = collapsed(v);
      return isArray(c) ? fail("#CALC!", `${name}'s LAMBDA gives an array`) : c;
    } catch (err) {
      if (!(err instanceof CellError)) throw err;
      return err;
    }
  }

  #call(e: Extract<Expr, { kind: "call" }>): Result {
    const args = e.args;
    const takes = (n: number) => (args.length !== n ? fail("#VALUE!", `${e.name} takes ${n} values`) : null);
    switch (e.name) {
      // ROW(ref) and COLUMN(ref): the reference's row or column, counted
      // from 1; over a range, each of its rows down (or columns across).
      // Without a reference, the cell the formula is in.
      case "ROW":
      case "COLUMN": {
        if (args.length > 1) return fail("#VALUE!", `${e.name} takes a reference, or nothing`);
        const which = e.name === "ROW" ? "row" : "col";
        if (args.length === 0) {
          const here = this.#at ? cellAddr(this.#at) : null;
          return here ? here[which] + 1 : fail("#REF!", `${e.name}() is in no cell`);
        }
        const r = this.#reference(args[0]);
        if (!r) return fail("#VALUE!", `${e.name} takes a reference`);
        const lo = Math.min(r.a[which], r.b[which]), hi = Math.max(r.a[which], r.b[which]);
        const ns = Array.from({ length: hi - lo + 1 }, (_, i) => lo + i + 1);
        return ns.length === 1 ? ns[0] : which === "row" ? array(ns.length, 1, ns) : array(1, ns.length, ns);
      }
      // OFFSET(reference, rows, cols, [height], [width]): the cells that
      // many rows down and columns across from the reference's top left
      // corner, as high and wide as given (else as the reference is).
      // INDIRECT(text, [a1]): the cells the text names. Both are made at
      // run time, so what they read is told the Env, and the formula is
      // volatile (excel.ts).
      case "OFFSET":
      case "INDIRECT":
        return this.#values(this.#reference(e)!, true);
      // FORMULATEXT(ref): the formula in the reference's (first) cell as
      // typed, #N/A when it holds none; ISFORMULA(ref): whether it holds
      // one; ISREF(x): whether x is a reference at all.
      case "FORMULATEXT":
      case "ISFORMULA": {
        takes(1);
        const r = this.#reference(args[0]);
        if (!r) return fail("#VALUE!", `${e.name} takes a reference`);
        const input = this.#env.input?.(qualify(r.sheet, keyOf({ row: Math.min(r.a.row, r.b.row), col: Math.min(r.a.col, r.b.col) }))) ?? "";
        const formula = input.startsWith("=");
        if (e.name === "ISFORMULA") return formula;
        return formula ? input : fail("#N/A", `${keyOf(r.a)} holds no formula`);
      }
      case "ISREF": {
        takes(1);
        try {
          return this.#reference(args[0]) !== null;
        } catch (err) {
          if (err instanceof CellError) return false;
          throw err;
        }
      }
      // SHEET([x]): the number (from 1, in tab order) of the sheet the
      // formula is on, or of the reference's sheet, or of the sheet named
      // by the text; SHEETS([ref]): how many sheets the workbook has (a
      // reference spans one).
      case "SHEET": {
        if (args.length > 1) return fail("#VALUE!", "SHEET takes a reference, a name, or nothing");
        const book = this.#book();
        if (args.length === 0) return book.here + 1;
        const r = this.#reference(args[0]);
        if (r) return r.sheet === null ? book.here + 1 : this.#sheetNumber(book.sheets, r.sheet);
        const name = this.#scalar(args[0]);
        return typeof name === "string" ? this.#sheetNumber(book.sheets, name) : fail("#N/A", `${toText(name)} is not a sheet`);
      }
      case "SHEETS": {
        if (args.length > 1) return fail("#VALUE!", "SHEETS takes a reference, or nothing");
        if (args.length === 0) return this.#book().sheets.length;
        return this.#reference(args[0]) ? 1 : fail("#VALUE!", "SHEETS takes a reference");
      }
      // INFO(type), what applies here: "numfile" is the sheets the workbook
      // has, "recalc" how it calculates (always "Automatic", until 2.15).
      case "INFO": {
        takes(1);
        const type = toText(this.#scalar(args[0])).trim().toLowerCase();
        if (type === "numfile") return this.#book().sheets.length;
        if (type === "recalc") return "Automatic";
        return fail("#VALUE!", `INFO of '${type}': only numfile and recalc apply here`);
      }
      case "IF": {
        // Only the branch taken is worked out: IF(B1=0, 0, A1/B1) is safe;
        // over an array, cell by cell.
        if (args.length < 2 || args.length > 3) return fail("#VALUE!", "IF takes 2 or 3 values");
        const cond = this.result(args[0]);
        if (isArray(cond)) return lift([cond, this.result(args[1]), args[2] ? this.result(args[2]) : false], ([c, y, n]) => (truth(c) ? y : n));
        const pick = truth(cond) ? args[1] : args[2];
        return pick ? this.result(pick) : false;
      }
      case "MAP": {
        if (args.length < 2) return fail("#VALUE!", "MAP takes an array and a LAMBDA");
        const fn = this.#closure(args[args.length - 1], "MAP"), grids = args.slice(0, -1).map((a) => this.#grid(a));
        if (grids.some((g) => g.rows !== grids[0].rows || g.cols !== grids[0].cols)) return fail("#VALUE!", "MAP's arrays differ in size");
        return array(grids[0].rows, grids[0].cols, grids[0].values.map((_, i) => this.#one("MAP", () => this.#invoke(fn, grids.map((g) => g.values[i])))));
      }
      case "REDUCE":
      case "SCAN": {
        takes(3);
        const fn = this.#closure(args[2], e.name), g = this.#grid(args[1]);
        let acc: Bound = this.result(args[0]);
        const seen: Value[] = [];
        for (const v of g.values) {
          acc = this.#invoke(fn, [acc, v]);
          seen.push(this.#one(e.name, () => acc));
        }
        return e.name === "SCAN" ? array(g.rows, g.cols, seen) : this.#one("REDUCE", () => acc);
      }
      case "BYROW":
      case "BYCOL": {
        takes(2);
        const fn = this.#closure(args[1], e.name), g = this.#grid(args[0]);
        const byRow = e.name === "BYROW";
        const lines = byRow
          ? Array.from({ length: g.rows }, (_, i) => array(1, g.cols, g.values.slice(i * g.cols, (i + 1) * g.cols)))
          : Array.from({ length: g.cols }, (_, j) => array(g.rows, 1, g.values.filter((_, k) => k % g.cols === j)));
        const out = lines.map((line) => this.#one(e.name, () => this.#invoke(fn, [line])));
        return byRow ? array(out.length, 1, out) : array(1, out.length, out);
      }
      case "MAKEARRAY": {
        takes(3);
        const rows = Math.trunc(toNumber(this.#scalar(args[0]))), cols = Math.trunc(toNumber(this.#scalar(args[1])));
        if (rows < 1 || cols < 1) return fail("#CALC!", "MAKEARRAY of no rows or columns");
        const fn = this.#closure(args[2], "MAKEARRAY"), values: Value[] = [];
        for (let i = 1; i <= rows; i++) for (let j = 1; j <= cols; j++) values.push(this.#one("MAKEARRAY", () => this.#invoke(fn, [i, j])));
        return array(rows, cols, values);
      }
      default:
        return call(e.name, args.map((a) => this.#arg(a)));
    }
  }

  #scalar(e: Expr): Value {
    const r = collapsed(this.result(e));
    return isArray(r) ? fail("#VALUE!", "an array where one value goes") : r;
  }

  // The workbook as the Env tells it: a sheet alone is a book of one.
  #book(): { readonly sheets: readonly string[]; readonly here: number } {
    return this.#env.book?.() ?? { sheets: [""], here: 0 };
  }

  // A sheet's number, from 1, by its name (any case); #N/A for none.
  #sheetNumber(sheets: readonly string[], name: string): number {
    const i = sheets.findIndex((s) => s.toUpperCase() === name.toUpperCase());
    return i < 0 ? fail("#N/A", `no sheet ${name}`) : i + 1;
  }

  // The reference an expression is: a cell, a range, what a cell spills
  // (its cell), or one OFFSET or INDIRECT makes; null when it is none
  // (a number, text, a call of anything else). One off the sheet is #REF!.
  #reference(e: Expr): Reference | null {
    switch (e.kind) {
      case "ref":
      case "spill": {
        const { sheet, cell } = splitKey(e.name);
        const a = addrOf(cell);
        return a ? { sheet, a, b: a } : fail("#REF!", `${e.name} is not on the sheet`);
      }
      case "range": {
        const a = cellAddr(e.from), b = cellAddr(e.to);
        if (!a || !b) return fail("#REF!", `${a ? e.to : e.from} is not on the sheet`);
        return { sheet: splitKey(e.from).sheet, a, b };
      }
      case "call":
        return e.name === "OFFSET" ? this.#offset(e.args) : e.name === "INDIRECT" ? this.#indirect(e.args) : null;
      default:
        return null;
    }
  }

  // A reference's cells: one value, or an array; each told the Env when
  // the reference was made at run time.
  #values(r: Reference, dynamic: boolean): Result {
    const keys = [...range(r.a, r.b)].map((k) => qualify(r.sheet, k));
    if (dynamic) for (const k of keys) this.#env.read?.(k);
    const values = keys.map((k) => this.#env.get(k));
    return keys.length === 1 ? values[0] : array(Math.abs(r.a.row - r.b.row) + 1, Math.abs(r.a.col - r.b.col) + 1, values);
  }

  #offset(args: readonly Expr[]): Reference {
    if (args.length < 3 || args.length > 5) return fail("#VALUE!", "OFFSET takes 3 to 5 values");
    const base = this.#reference(args[0]);
    if (!base) return fail("#VALUE!", "OFFSET takes a reference first");
    const given = (i: number) => i < args.length && args[i].kind !== "missing";
    const n = (i: number) => Math.trunc(toNumber(this.#scalar(args[i])));
    const top = Math.min(base.a.row, base.b.row) + n(1), left = Math.min(base.a.col, base.b.col) + n(2);
    const height = given(3) ? n(3) : Math.abs(base.a.row - base.b.row) + 1;
    const width = given(4) ? n(4) : Math.abs(base.a.col - base.b.col) + 1;
    if (height < 1 || width < 1) return fail("#REF!", `OFFSET of ${height} rows by ${width} columns is no reference`);
    const a = { row: top, col: left }, b = { row: top + height - 1, col: left + width - 1 };
    if (top < 0 || left < 0 || b.row >= ROWS || b.col >= COLS) return fail("#REF!", `OFFSET to row ${top + 1}, column ${left + 1}, ${height} by ${width}, is off the sheet`);
    return { sheet: base.sheet, a, b };
  }

  #indirect(args: readonly Expr[]): Reference {
    if (args.length < 1 || args.length > 2) return fail("#VALUE!", "INDIRECT takes 1 or 2 values");
    const text = toText(this.#scalar(args[0]));
    const a1 = args.length === 2 && args[1].kind !== "missing" ? truth(this.#scalar(args[1])) : true;
    const r = referenceOf(text, a1, this.#at ? cellAddr(this.#at) : null);
    return r ?? fail("#REF!", `INDIRECT of '${text}': not a reference`);
  }
}

// What a formula comes to: a value, or an array (of more than one value).
// `at` is the cell the formula is in, for ROW() and COLUMN().
export function evaluate(e: Expr, env: Env, at: string | null = null): Result {
  try {
    const v = new Evaluator(env, at).any(e);
    return isClosure(v) ? new CellError("#CALC!", "a formula's value is a LAMBDA: call it") : collapsed(v);
  } catch (err) {
    if (err instanceof CellError) return err;
    throw err;
  }
}
