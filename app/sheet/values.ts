// What a cell holds once worked out, and the coercions every operator and
// function goes through. An error is a value too (a CellError): it flows
// through whatever reads it, so a cell that sums a #DIV/0! shows #DIV/0!.
export type ErrorCode = "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A" | "#CYCLE!" | "#ERROR!" | "#CALC!" | "#SPILL!";

export class CellError {
  readonly code: ErrorCode;
  readonly why: string;

  constructor(code: ErrorCode, why: string) {
    this.code = code;
    this.why = why;
  }

  toString(): string { return this.code; }
}

// null is an empty cell.
export type Value = number | string | boolean | null | CellError;

// A dynamic array: what a formula may come to (a range, SEQUENCE, FILTER
// ...), row by row; in a sheet it spills.
export interface ArrayValue {
  readonly kind: "array";
  readonly rows: number;
  readonly cols: number;
  readonly values: readonly Value[];
}

// What a formula comes to: one value, or an array.
export type Result = Value | ArrayValue;

export const isArray = (v: unknown): v is ArrayValue => typeof v === "object" && v !== null && (v as ArrayValue).kind === "array";

export const array = (rows: number, cols: number, values: readonly Value[]): ArrayValue => ({ kind: "array", rows, cols, values });

// An array of one value is that value.
export const collapsed = (r: Result): Result => (isArray(r) && r.rows === 1 && r.cols === 1 ? r.values[0] : r);

export const fail = (code: ErrorCode, why: string): never => { throw new CellError(code, why); };

// An operand that is an error is rethrown: what reads an error is one.
export function scalar(v: Value): Exclude<Value, CellError> {
  if (v instanceof CellError) throw v;
  return v;
}

export function toNumber(v: Value): number {
  const s = scalar(v);
  if (s === null) return 0;
  if (typeof s === "number") return s;
  if (typeof s === "boolean") return s ? 1 : 0;
  const n = s.trim() === "" ? NaN : Number(s);
  return Number.isNaN(n) ? fail("#VALUE!", `'${s}' is not a number`) : n;
}

export function toText(v: Value): string {
  const s = scalar(v);
  if (s === null) return "";
  if (typeof s === "boolean") return s ? "TRUE" : "FALSE";
  return typeof s === "number" ? String(Number(s.toPrecision(12))) : s;
}

export const finite = (n: number): number => (Number.isFinite(n) ? n : fail("#VALUE!", "the result is not a finite number"));

export type Comparison = "=" | "<>" | "<" | ">" | "<=" | ">=";

// How two values order, as a spreadsheet orders them: every number below
// every text, every text below FALSE, FALSE below TRUE; text without case.
// An empty cell is the other side's empty value (0, "" or FALSE), so it
// equals 0, "" and FALSE. Negative, 0 or positive.
export function ordered(a: Exclude<Value, CellError>, b: Exclude<Value, CellError>): number {
  const rank = (v: typeof a) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  const empty = (like: typeof a) => (typeof like === "number" ? 0 : typeof like === "string" ? "" : false);
  if (a === null && b === null) return 0;
  const x = a === null ? empty(b) : a, y = b === null ? empty(a) : b;
  if (rank(x) !== rank(y)) return rank(x) - rank(y);
  if (typeof x === "number") return x - (y as number);
  if (typeof x === "string") return x.toUpperCase().localeCompare((y as string).toUpperCase());
  return Number(x) - Number(y);
}

export function compare(op: Comparison, a: Value, b: Value): boolean {
  const d = ordered(scalar(a), scalar(b));
  switch (op) {
    case "=": return d === 0;
    case "<>": return d !== 0;
    case "<": return d < 0;
    case ">": return d > 0;
    case "<=": return d <= 0;
    case ">=": return d >= 0;
    default: return op satisfies never;
  }
}

// A grid with its empty (blank) rows and columns trimmed from the ends
// `rows` and `cols` say (bit 1 leading, bit 2 trailing: 3 both), or null
// when nothing is left.
export function trimmed(g: ArrayValue, rows: number, cols: number): ArrayValue | null {
  const at = (r: number, c: number) => g.values[r * g.cols + c];
  const rowEmpty = (r: number) => Array.from({ length: g.cols }, (_, c) => at(r, c)).every((v) => v === null);
  const colEmpty = (c: number, r0: number, r1: number) => Array.from({ length: r1 - r0 + 1 }, (_, i) => at(r0 + i, c)).every((v) => v === null);
  let r0 = 0, r1 = g.rows - 1, c0 = 0, c1 = g.cols - 1;
  if (rows & 1) while (r0 <= r1 && rowEmpty(r0)) r0++;
  if (rows & 2) while (r1 >= r0 && rowEmpty(r1)) r1--;
  if (r0 > r1) return null;
  if (cols & 1) while (c0 <= c1 && colEmpty(c0, r0, r1)) c0++;
  if (cols & 2) while (c1 >= c0 && colEmpty(c1, r0, r1)) c1--;
  if (c0 > c1) return null;
  const values: Value[] = [];
  for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) values.push(at(r, c));
  return array(r1 - r0 + 1, c1 - c0 + 1, values);
}
