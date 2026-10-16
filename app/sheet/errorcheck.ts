// Error Checking, as Excel's Formulas > Error Checking has its rules (the
// ones that apply here): a formula that comes to an error; a formula that
// differs from the formulas beside it, those on both sides in its row or
// its column being one formula, R1C1, and it another; a number stored as
// text (typed with an apostrophe, '42, or with thousands separators as
// text); and a formula whose range stops short of a number next to it in
// the same line (=SUM(B2:B4) with a number in B5). Each finding is a cell,
// its rule, and what it says, in sheet order; a cell ignored is left out.
import { addrOf, keyOf, parseRange, splitKey } from "./address.ts";
import type { Addr } from "./address.ts";
import { toR1C1 } from "./excel.ts";
import type { Sheet } from "./sheet.ts";
import { CellError } from "./values.ts";

export type Check = "error" | "inconsistent" | "text-number" | "omits";
export const CHECKS: readonly Check[] = ["error", "inconsistent", "text-number", "omits"];

export interface Finding {
  readonly cell: string;
  readonly check: Check;
  readonly message: string;
}

const NUMBER_TEXT = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$|^\s*[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?\s*$/;

// The findings on a sheet, those of the cells in `ignored` left out.
export function errorCheck(sheet: Sheet, ignored: ReadonlySet<string> = new Set()): Finding[] {
  const out: Finding[] = [];
  const excel = sheet.language.name === "Excel";
  const formula = (k: string) => sheet.input(k).startsWith("=");
  const r1c1 = (k: string) => (formula(k) ? toR1C1(sheet.input(k), k) : null);
  const at = (a: Addr, dc: number, dr: number) => {
    const b = { col: a.col + dc, row: a.row + dr };
    return b.col < 0 || b.row < 0 ? null : addrOf(keyOf(b)) && keyOf(b);
  };
  for (const k of sheet.keys()) {
    if (ignored.has(k)) continue;
    const input = sheet.input(k), v = sheet.value(k), a = addrOf(k)!;
    if (formula(k) && v instanceof CellError) {
      out.push({ cell: k, check: "error", message: `${v.code} ${v.why}`.trim() });
      continue;
    }
    if (!formula(k)) {
      if (typeof v === "string" && (input.startsWith("'") ? NUMBER_TEXT.test(v) : NUMBER_TEXT.test(input) && /,/.test(input)))
        out.push({ cell: k, check: "text-number", message: "a number stored as text" });
      continue;
    }
    if (!excel) continue;
    // Inconsistent: the cells on both sides one formula, this one another.
    const mine = r1c1(k);
    const differs = [[1, 0, "row"], [0, 1, "column"]].find(([dc, dr]) => {
      const p = at(a, -dc, -dr), q = at(a, dc as number, dr as number);
      if (!p || !q) return false;
      const x = r1c1(p);
      return x !== null && x === r1c1(q) && x !== mine;
    });
    if (differs) {
      out.push({ cell: k, check: "inconsistent", message: `the formula differs from the formulas beside it in its ${differs[2]}` });
      continue;
    }
    // Omits adjacent cells: a range in one row or column, a number just
    // past either end of it that the formula does not read.
    for (const ref of sheet.language.references(input.slice(1))) {
      const { sheet: other, cell } = splitKey(ref.replace(/\$/g, ""));
      const r = other === null ? parseRange(cell) : null;
      if (!r || !cell.includes(":")) continue;
      const [lo, hi] = [{ col: Math.min(r[0].col, r[1].col), row: Math.min(r[0].row, r[1].row) }, { col: Math.max(r[0].col, r[1].col), row: Math.max(r[0].row, r[1].row) }];
      const down = lo.col === hi.col, across = lo.row === hi.row;
      if (down === across) continue;
      const ends = down ? [at(lo, 0, -1), at(hi, 0, 1)] : [at(lo, -1, 0), at(hi, 1, 0)];
      const missed = ends.find((e) => e && e !== k && !formula(e) && typeof sheet.value(e) === "number");
      if (missed) {
        out.push({ cell: k, check: "omits", message: `the formula omits ${missed}, next to ${ref}` });
        break;
      }
    }
  }
  return out;
}
