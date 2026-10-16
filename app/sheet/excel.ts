// "Excel formulas": the sheet's own language (lexer.ts, parser.ts,
// evaluate.ts, functions.ts) as a Language. A formula reads the cells its
// references and ranges name; it is rewritten token by token, so only
// references change ($ fixes a column or row in a copy, and a reference
// moved off the sheet, or into what was deleted, is #REF!).
import { COLS, ROWS, cellAddr, colName, range, sheetPrefix, span, splitKey } from "./address.ts";
import type { Addr } from "./address.ts";
import { movedBy } from "./shift.ts";
import type { Edit, Fixed } from "./shift.ts";
import { SPECIAL_FORMS, VOLATILE_FORMS, evaluate } from "./evaluate.ts";
import { names } from "./functions.ts";
import type { Language, Offset } from "./language.ts";
import { FormulaError, tokens } from "./lexer.ts";
import type { Token } from "./lexer.ts";
import { parse } from "./parser.ts";
import type { Expr } from "./parser.ts";
import { CellError } from "./values.ts";

// The cells an expression reads, ranges spelled out; off-sheet ones are left
// to the evaluator (#REF!).
export function* refsOf(e: Expr, seen: Set<string> = new Set()): Generator<string, void, undefined> {
  switch (e.kind) {
    case "ref":
    case "spill":
      if (cellAddr(e.name)) yield e.name;
      return;
    case "range":
      yield* span(e.from, e.to)?.keys ?? [];
      return;
    // The whole table: which of its cells a reference comes to depends on
    // its header row then, so it is worked out again when any of them changes.
    case "table":
      yield* range(e.table.range[0], e.table.range[1]);
      return;
    case "neg":
    case "at":
      return yield* refsOf(e.arg, seen);
    case "binary":
      yield* refsOf(e.left, seen);
      return yield* refsOf(e.right, seen);
    case "call":
      for (const a of e.args) yield* refsOf(a, seen);
      return;
    case "let":
      for (const [, v] of e.bindings) yield* refsOf(v, seen);
      return yield* refsOf(e.body, seen);
    case "lambda": return yield* refsOf(e.body, seen);
    case "apply":
      yield* refsOf(e.fn, seen);
      for (const a of e.args) yield* refsOf(a, seen);
      return;
    // What a named formula reads, once however often it is named (or calls
    // itself).
    case "named":
      if (seen.has(e.name)) return;
      seen.add(e.name);
      return yield* refsOf(e.expr(), seen);
    default:
      return;
  }
}

// Whether an expression calls any of `forms` (a named formula's own
// expression looked into once).
function calls(e: Expr, forms: ReadonlySet<string>, seen: Set<string> = new Set()): boolean {
  switch (e.kind) {
    case "call": return forms.has(e.name) || e.args.some((a) => calls(a, forms, seen));
    // A 3-D reference's sheets are the workbook's tab order: it is worked
    // out again on every change, as OFFSET is.
    case "span3d": return forms === VOLATILE_FORMS;
    case "neg":
    case "at":
      return calls(e.arg, forms, seen);
    case "binary": return calls(e.left, forms, seen) || calls(e.right, forms, seen);
    case "let": return e.bindings.some(([, v]) => calls(v, forms, seen)) || calls(e.body, forms, seen);
    case "lambda": return calls(e.body, forms, seen);
    case "apply": return calls(e.fn, forms, seen) || e.args.some((a) => calls(a, forms, seen));
    case "named":
      if (seen.has(e.name)) return false;
      seen.add(e.name);
      return calls(e.expr(), forms, seen);
    default: return false;
  }
}

type Ref = Extract<Token, { kind: "ref" }>;

const fixedOf = (t: Ref): Fixed => ({ col: t.fixCol, row: t.fixRow });

// A reference at its new place, written as it was ($ and all), or as typed
// when it has not moved.
// (Its sheet, if it names one, as typed.)
function written(t: Ref, to: Addr): string {
  const a = cellAddr(t.name)!;
  if (to.col === a.col && to.row === a.row) return t.text;
  return `${t.text.slice(0, t.text.lastIndexOf("!") + 1)}${t.fixCol ? "$" : ""}${colName(to.col)}${t.fixRow ? "$" : ""}${to.row + 1}${t.spill ? "#" : ""}`;
}

// A cell's input with its references moved by `edit` (a constant, or a
// formula that does not lex, is unchanged). Token by token, so only the
// references change.
function rewrite(input: string, edit: Edit): string {
  if (!input.startsWith("=")) return input;
  let toks: Token[];
  try {
    toks = [...tokens(input.slice(1), 1)];
  } catch (e) {
    if (e instanceof FormulaError) return input;
    throw e;
  }
  let out = "", from = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind !== "ref") continue;
    const end = toks[i + 2];
    const other = toks[i + 1]?.kind === ":" && end?.kind === "ref" ? end : null;
    if (!edit.moves(splitKey(t.name).sheet)) {
      if (other) i += 2;
      continue;
    }
    const a = cellAddr(t.name), b = other ? cellAddr(other.name) : null;
    out += input.slice(from, t.at);
    if (other) {
      const r = a && b ? edit.range(a, b, fixedOf(t), fixedOf(other)) : null;
      out += r ? written(t, r[0]) + input.slice(t.at + t.text.length, other.at) + written(other, r[1]) : "#REF!";
    } else {
      const to = a ? edit.cell(a, fixedOf(t)) : null;
      out += to ? written(t, to) : "#REF!";
    }
    const last = other ?? t;
    from = last.at + last.text.length;
    if (other) i += 2;
  }
  return out + input.slice(from);
}

// A cell's input with its references to sheet `from` naming `to`. Token by
// token, so text in strings stays as written.
function renamed(input: string, from: string, to: string): string {
  if (!input.startsWith("=")) return input;
  let toks: Token[];
  try {
    toks = [...tokens(input.slice(1), 1)];
  } catch (e) {
    if (e instanceof FormulaError) return input;
    throw e;
  }
  let out = "", at = 0;
  for (const t of toks) {
    // A 3-D reference with the sheet at either end: that end renamed, the
    // other as typed, quoted when either needs it.
    const sheets = t.kind === "ref" ? splitKey(t.name).sheet : null;
    if (t.kind === "ref" && sheets !== null && sheets.includes(":")) {
      if (!sheets.split(":").includes(from)) continue;
      const bang = t.text.lastIndexOf("!"), typed = t.text.slice(0, bang).replace(/^'|'$/g, "").split(":");
      const ends = typed.map((n) => (n.toUpperCase() === from ? to : n));
      const both = ends.join(":");
      out += input.slice(at, t.at) + (ends.every((n) => /^[A-Za-z_]\w*$/.test(n)) ? `${both}!` : `'${both}'!`) + t.text.slice(bang + 1);
      at = t.at + t.text.length;
      continue;
    }
    // A cell there, or one of its own names (Costs!QTY).
    if ((t.kind !== "ref" && !(t.kind === "name" && t.name.includes("!"))) || splitKey(t.name).sheet !== from) continue;
    out += input.slice(at, t.at) + sheetPrefix(to) + t.text.slice(t.text.lastIndexOf("!") + 1);
    at = t.at + t.text.length;
  }
  return out + input.slice(at);
}

// A formula's references written R1C1, relative to the cell `at` it is
// in: an absolute row or column as its number (R3C3), a relative one as an
// offset in brackets (R[-1]C[2]), the cell's own row or column bare (RC).
// Token by token, so strings stay as written; a constant is unchanged.
export function toR1C1(input: string, at: string): string {
  const here = cellAddr(at);
  if (!input.startsWith("=") || !here) return input;
  let toks: Token[];
  try {
    toks = [...tokens(input.slice(1), 1)];
  } catch (e) {
    if (e instanceof FormulaError) return input;
    throw e;
  }
  let out = "", from = 0;
  for (const t of toks) {
    if (t.kind !== "ref") continue;
    const a = cellAddr(t.name);
    if (!a) continue;
    const part = (letter: string, fixed: boolean, n: number, mine: number) => (fixed ? `${letter}${n + 1}` : n === mine ? letter : `${letter}[${n - mine}]`);
    const sheet = t.text.slice(0, t.text.lastIndexOf("!") + 1);
    out += input.slice(from, t.at) + sheet + part("R", t.fixRow, a.row, here.row) + part("C", t.fixCol, a.col, here.col) + (t.spill ? "#" : "");
    from = t.at + t.text.length;
  }
  return out + input.slice(from);
}

// An R1C1 reference's text, at the start of a match: a sheet perhaps,
// then R and C, each a number, an offset in brackets, or nothing.
const R1C1 = /((?:[A-Za-z_]\w*(?::[A-Za-z_]\w*)?|'[^']+')!)?R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_([.])/iy;

// A formula typed in R1C1 written A1, from the cell `at` it is in (what
// falls off the sheet is #REF!); text in quotes stays as written.
export function fromR1C1(input: string, at: string): string {
  const here = cellAddr(at);
  if (!input.startsWith("=") || !here) return input;
  let out = "=";
  for (let i = 1; i < input.length;) {
    const c = input[i];
    if (c === '"') {
      let j = i + 1;
      while (j < input.length && !(input[j] === '"' && input[j + 1] !== '"')) j += input[j] === '"' ? 2 : 1;
      out += input.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const before = out[out.length - 1] ?? "";
    R1C1.lastIndex = i;
    const m = /[A-Za-z0-9_.\]'$]/.test(before) ? null : R1C1.exec(input);
    if (!m) {
      out += c;
      i++;
      continue;
    }
    const pick = (part: string | undefined, mine: number): [number, boolean] =>
      part === undefined ? [mine, false] : part.startsWith("[") ? [mine + Number(part.slice(1, -1)), false] : [Number(part) - 1, true];
    const [row, rowFixed] = pick(m[2], here.row), [col, colFixed] = pick(m[3], here.col);
    out += row < 0 || row >= ROWS || col < 0 || col >= COLS ? "#REF!" : `${m[1] ?? ""}${colFixed ? "$" : ""}${colName(col)}${rowFixed ? "$" : ""}${row + 1}`;
    i += m[0].length;
  }
  return out;
}

// A cell's input as it reads copied `by` cells.
export const moved = (input: string, by: Offset): string => movedBy(EXCEL, input, by);

// The references a formula's tokens name, in order, once each; tokens up
// to where it stops lexing, for a formula still being typed.
function references(formula: string): string[] {
  const toks: Token[] = [];
  try {
    for (const t of tokens(formula)) toks.push(t);
  } catch (e) {
    if (!(e instanceof FormulaError)) throw e;
  }
  const out: string[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind !== "ref") continue;
    const end = toks[i + 2];
    if (toks[i + 1]?.kind === ":" && end?.kind === "ref") {
      out.push(`${t.name}:${splitKey(end.name).cell}`);
      i += 2;
    } else out.push(t.name);
  }
  return [...new Set(out)];
}

export const EXCEL: Language = {
  name: "Excel",
  compile(formula, names, tables, formulas, at) {
    try {
      const expr = parse(formula, names ?? (() => null), formulas, tables, at ? cellAddr(at)?.row ?? null : null);
      // One calling OFFSET or INDIRECT makes references at run time: it
      // is worked out again on every change.
      return { kind: "formula", reads: [...refsOf(expr)], volatile: calls(expr, VOLATILE_FORMS), run: (env) => evaluate(expr, env, at ?? null) };
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e;
      return { kind: "broken", error: new CellError("#ERROR!", `parse error at ${e.at}: ${e.message}`) };
    }
  },
  rewrite,
  renamed,
  functions: () => [...names(), ...SPECIAL_FORMS].sort(),
  references,
};
