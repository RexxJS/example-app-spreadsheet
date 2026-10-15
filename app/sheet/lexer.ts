// A formula's tokens, one at a time, from a generator. Each knows where it
// starts (1-based, counting the formula's leading '=') so an error can say
// "at 8" of what the person typed. A reference's $ (fixing its column or
// row when the formula is copied) is in the token; its name is the cell, or
// SHEET!CELL on another sheet.
import type { ErrorCode } from "./values.ts";

export type Op = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=";

export type Token = (
  | { readonly kind: "num"; readonly value: number }
  | { readonly kind: "str"; readonly value: string }
  | { readonly kind: "ref"; readonly name: string; readonly fixCol: boolean; readonly fixRow: boolean; readonly spill: boolean }
  | { readonly kind: "error"; readonly code: ErrorCode }
  // A name (upper case); another sheet's own name is SHEET!NAME.
  | { readonly kind: "name"; readonly name: string }
  // A table reference: Sales[Amount], Sales[@Amount], Sales[[#Headers],[Qty]]:
  // the table's name (upper case) and what is between its outer brackets,
  // as written.
  | { readonly kind: "table"; readonly name: string; readonly spec: string }
  | { readonly kind: "op"; readonly op: Op }
  | { readonly kind: "(" | ")" | "," | ":" | "{" | "}" | ";" }
  | { readonly kind: "end" }
) & { readonly at: number; readonly text: string };

export class FormulaError extends Error {
  readonly at: number;

  constructor(at: number, message: string) {
    super(message);
    this.at = at;
  }
}

const OPS: readonly Op[] = ["<>", "<=", ">=", "+", "-", "*", "/", "^", "&", "=", "<", ">"];
const NUMBER = /^(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/;
// A name, or a function's with dots in it (CEILING.MATH).
const WORD = /^[A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z0-9_]+)*/;
// Another sheet's own name: Costs!QTY, 'My Costs'!QTY (a cell there,
// Costs!B2, is a reference; a call, Costs!F(, is not this).
const SHEET_NAME = /^(?:([A-Za-z_]\w*)|'([^']+)')!([A-Za-z_][A-Za-z0-9_]*)(?![\w(])/;
// A reference, perhaps on another sheet: Costs!B2, 'My Costs'!$B$2; A1#
// is the whole of what A1 spills.
const REF = /^(?:(?:([A-Za-z_]\w*)|'([^']+)')!)?(\$?)([A-Za-z]+)(\$?)([0-9]+)(?![A-Za-z0-9_])(#(?![A-Za-z/]))?/;
// The errors a formula can be written with (a copy writes #REF! for a
// reference it moved off the sheet).
const ERRORS: readonly ErrorCode[] = ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#CALC!"];

// Where the ']' that closes the '[' at `open` is (brackets nest, and ' makes
// the character after it plain, as in Excel: Sales['[Old']]), or -1.
function closing(formula: string, open: number): number {
  let depth = 0;
  for (let j = open; j < formula.length; j++) {
    const c = formula[j];
    if (c === "'") j++;
    else if (c === "[") depth++;
    else if (c === "]" && --depth === 0) return j;
  }
  return -1;
}

// `formula` is what follows the '='; `from` is where it starts in the input.
export function* tokens(formula: string, from = 2): Generator<Token, void, undefined> {
  let i = 0;
  while (i < formula.length) {
    const rest = formula.slice(i), at = from + i, c = formula[i];
    if (c === " ") { i++; continue; }
    const num = NUMBER.exec(rest)?.[0];
    if (num) {
      yield { kind: "num", value: Number(num), at, text: num };
      i += num.length;
      continue;
    }
    // A reference, unless a '(' follows: LOG10( is a call, not cell LOG10.
    const ref = REF.exec(rest);
    if (ref && rest[ref[0].length] !== "(") {
      const [text, bare, quoted, fixCol, col, fixRow, row, spill] = ref;
      const sheet = bare ?? quoted;
      const cell = col.toUpperCase() + row;
      yield { kind: "ref", name: sheet ? `${sheet.toUpperCase()}!${cell}` : cell, fixCol: fixCol === "$", fixRow: fixRow === "$", spill: !!spill, at, text };
      i += text.length;
      continue;
    }
    const qualified = SHEET_NAME.exec(rest);
    if (qualified) {
      const [text, bare, quoted, name] = qualified;
      yield { kind: "name", name: `${(bare ?? quoted).toUpperCase()}!${name.toUpperCase()}`, at, text };
      i += text.length;
      continue;
    }
    const word = WORD.exec(rest)?.[0];
    if (word && formula[i + word.length] === "[") {
      const end = closing(formula, i + word.length);
      if (end < 0) throw new FormulaError(at, `a table reference that never ends: ${rest}`);
      const text = formula.slice(i, end + 1);
      yield { kind: "table", name: word.toUpperCase(), spec: formula.slice(i + word.length + 1, end), at, text };
      i = end + 1;
      continue;
    }
    if (word) {
      yield { kind: "name", name: word.toUpperCase(), at, text: word };
      i += word.length;
      continue;
    }
    const code = ERRORS.find((e) => rest.toUpperCase().startsWith(e));
    if (code) {
      yield { kind: "error", code, at, text: rest.slice(0, code.length) };
      i += code.length;
      continue;
    }
    if (c === '"') {
      let j = i + 1, value = "";
      for (;;) {
        if (j >= formula.length) throw new FormulaError(at, "a string that never ends");
        if (formula[j] === '"') {
          if (formula[j + 1] !== '"') break;
          j++;
        }
        value += formula[j++];
      }
      yield { kind: "str", value, at, text: formula.slice(i, j + 1) };
      i = j + 1;
      continue;
    }
    if (c === "(" || c === ")" || c === "," || c === ":" || c === "{" || c === "}" || c === ";") {
      yield { kind: c, at, text: c };
      i++;
      continue;
    }
    const op = OPS.find((o) => rest.startsWith(o));
    if (op) {
      yield { kind: "op", op, at, text: op };
      i += op.length;
      continue;
    }
    throw new FormulaError(at, `unexpected '${c}'`);
  }
  yield { kind: "end", at: from + formula.length, text: "" };
}
