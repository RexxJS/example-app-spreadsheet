// Solver, Simplex LP (6.8), on textbook problems: Hillier and Lieberman's
// Wyndor Glass (max 3x + 5y, x ≤ 4, 2y ≤ 12, 3x + 2y ≤ 18: x = 2, y = 6,
// 36); a minimum over ≥ rows (min 2x + 3y, x + y ≥ 4, x + 3y ≥ 6: 3, 1,
// 9); integers by branch and bound; infeasible and unbounded said; on the
// sheet: the coefficients read off its formulas, the solution kept (one
// step of undo), the Answer Report, a value to reach, a non-linear model
// refused; SOLVER from REXX (RexxJS has a SOLVE of its own).
import { solveLp } from "./sheet/solver.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const r = (x: number[]) => x.map((v) => Math.round(v * 1e9) / 1e9);
const wyndor = solveLp({ c: [3, 5], maximise: true, rows: [{ a: [1, 0], rel: "<=", b: 4 }, { a: [0, 2], rel: "<=", b: 12 }, { a: [3, 2], rel: "<=", b: 18 }] });
eq("Wyndor Glass", wyndor.status === "optimal" ? [r(wyndor.x), wyndor.value] : wyndor.status, [[2, 6], 36]);
const diet = solveLp({ c: [2, 3], maximise: false, rows: [{ a: [1, 1], rel: ">=", b: 4 }, { a: [1, 3], rel: ">=", b: 6 }] });
eq("a minimum over ≥ rows", diet.status === "optimal" ? [r(diet.x), diet.value] : diet.status, [[3, 1], 9]);
// max 5x + 4y, 6x + 4y ≤ 24, x + 2y ≤ 6: the LP's 3, 1.5; the integers' 4, 0 (20).
const ip = solveLp({ c: [5, 4], maximise: true, rows: [{ a: [6, 4], rel: "<=", b: 24 }, { a: [1, 2], rel: "<=", b: 6 }], integer: [true, true] });
eq("integers by branch and bound", ip.status === "optimal" ? [r(ip.x), ip.value] : ip.status, [[4, 0], 20]);
eq("infeasible, unbounded, an equality, a free variable", [solveLp({ c: [1], maximise: true, rows: [{ a: [1], rel: "<=", b: 1 }, { a: [1], rel: ">=", b: 2 }] }).status,
  solveLp({ c: [1, 1], maximise: true, rows: [{ a: [1, -1], rel: "<=", b: 1 }] }).status, (solveLp({ c: [1, 1], maximise: false, rows: [{ a: [1, 2], rel: "=", b: 4 }] }) as { x: number[] }).x,
  (solveLp({ c: [1], maximise: false, rows: [{ a: [1], rel: ">=", b: -5 }], free: [true] }) as { x: number[] }).x], ["infeasible", "unbounded", [0, 2], [-5]]);

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
for (const [k, v] of [["B1", "1"], ["B2", "1"], ["B4", "=3*B1+5*B2"], ["B5", "=B1"], ["B6", "=2*B2"], ["B7", "=3*B1+2*B2"]]) await cmds.run(`cell.set cell=${k} input="${v}"`);
const o = await cmds.run('solver.solve objective=B4 goal=max changing=B1:B2 constraints="B5<=4;B6<=12;B7<=18" report=1');
eq("on the sheet: solved, kept", [o.message, show(wb.sheet.value("B1")), show(wb.sheet.value("B2")), show(wb.sheet.value("B4"))], ["Solver found a solution: B4 = 36 at B1=2, B2=6", "2", "6", "36"]);
const rep = wb.page("Answer Report 1")!.sheet, line = (row: number) => ["A", "B", "C", "D", "E"].map((c) => show(rep.value(`${c}${row}`))).join("|");
eq("its Answer Report", [line(5), line(9), line(10), line(14), line(15), line(16)],
  ["Sheet1!B4|8|36||", "Sheet1!B1|1|2|Contin|", "Sheet1!B2|1|6|Contin|", "Sheet1!B5|2|B5 <= 4|Not Binding|2", "Sheet1!B6|12|B6 <= 12|Binding|0", "Sheet1!B7|18|B7 <= 18|Binding|0"]);
await cmds.run("edit.undo");
eq("undone as one step", [show(wb.sheet.value("B1")), show(wb.sheet.value("B2"))], ["1", "1"]);
eq("a value to reach, cells kept as they were", [(await cmds.run('solver.solve objective=B4 goal=value value=15 changing=B1:B2 constraints="B2<=0" restore=1')).message, show(wb.sheet.value("B1"))],
  ["Solver found a solution: B4 = 15 at B1=5, B2=0 (the cells as they were)", "1"]);
await cmds.run('cell.set cell=B8 input="=B1*B2"');
eq("refusals", [(await cmds.run("solver.solve objective=B8 changing=B1:B2")).message, (await cmds.run('solver.solve objective=B4 changing=B1:B2 constraints="B5>=5;B5<=4"')).message,
  (await cmds.run("solver.solve objective=B4 changing=B4")).message, show(wb.sheet.value("B1"))],
  ["the problem is not linear in B1, B2: Simplex LP needs linear formulas", "Solver could not find a feasible solution", "B4 holds a formula, not a value Solver can change", "1"]);
const said: string[] = [];
await cmds.script('SAY SOLVER("B4", "max", "B1:B2", "B5<=4;B6<=12;B7<=18;B1:B2=int")', (l) => said.push(l));
eq("SOLVER from REXX", said, ["Solver found a solution: B4 = 36 at B1=2, B2=6"]);

// expect: ok Wyndor Glass
// expect: ok a minimum over ≥ rows
// expect: ok integers by branch and bound
// expect: ok infeasible, unbounded, an equality, a free variable
// expect: ok on the sheet: solved, kept
// expect: ok its Answer Report
// expect: ok undone as one step
// expect: ok a value to reach, cells kept as they were
// expect: ok refusals
// expect: ok SOLVER from REXX
