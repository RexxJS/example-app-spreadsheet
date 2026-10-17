// Consolidation (4.14): ranges of several sheets put together by a
// function, by position or by their labels (top row, left column, the
// union of them in the order first seen, any case), as values or linked
// formulas that follow their sources; undone; CONSOLIDATE from REXX.
import { consolidate } from "./sheet/consolidate.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const src = (values: (string | number | null)[][], sheet: string) => ({ values, cell: (r: number, c: number) => `${sheet}!${"ABC"[c]}${r + 1}` });
eq("by position", consolidate([src([[1, 2], [3, 4]], "A"), src([[10, 20]], "B")], "sum", false, false).rows, [[11, 22], [3, 4]]);
const both = consolidate([src([[null, "Q1", "Q2"], ["pens", 1, 2], ["ink", 3, 4]], "Jan"), src([[null, "Q2", "Q3"], ["INK", 10, 5], ["paper", 7, 8]], "Feb")], "sum", true, true);
eq("by labels, top and left: their union, any case", [both.rows, both.from[2][2]],
  [[[null, "Q1", "Q2", "Q3"], ["pens", 1, 2, null], ["ink", 3, 14, 5], ["paper", null, 7, 8]], ["Jan!C3", "Feb!B2"]]);
eq("other functions", [consolidate([src([[2]], "A"), src([[4]], "B")], "average", false, false).rows, consolidate([src([[2]], "A"), src([[4]], "B")], "max", false, false).rows,
  consolidate([src([[2]], "A"), src([["x"]], "B")], "counta", false, false).rows], [[[3]], [[4]], [[2]]]);

const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("formulas.use language=Excel");
await cmds.run("sheet.rename name=Sheet1 to=Jan");
for (const [k, v] of [["A2", "pens"], ["B1", "Qty"], ["B2", "3"], ["A3", "ink"], ["B3", "1"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
await cmds.run("sheet.add name=Feb");
for (const [k, v] of [["A2", "ink"], ["B1", "Qty"], ["B2", "5"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
await cmds.run("sheet.add name=Year");
const col = (c: string) => [1, 2, 3].map((r) => show(wb.sheet.value(`${c}${r}`))).join(" ");
eq("two sheets by their labels, as values", [await message("data.consolidate to=A1 sources=\"Jan!A1:B3,Feb!A1:B2\" labels=both"), col("A"), col("B")],
  ["consolidated 2 ranges at A1 by sum, by the labels in the top row and left column: 6 cells set, 6 changed", " pens ink", "Qty 3 6"]);
eq("linked: a formula over the cells, following them", [await message("data.consolidate to=D1 sources=\"Jan!A1:B3,Feb!A1:B2\" labels=both links=1"), wb.sheet.input("E3"), show(wb.sheet.value("E3"))],
  ["consolidated 2 ranges at D1 by sum, by the labels in the top row and left column, linked: 6 cells set, 6 changed", "=SUM(Jan!B3,Feb!B2)", "6"]);
await cmds.run("sheet.activate name=Feb");
await cmds.run("cell.set cell=B2 input=50");
await cmds.run("sheet.activate name=Year");
eq("followed", show(wb.sheet.value("E3")), "51");
await cmds.run("edit.undo");
eq("undone, the linked total back", show(wb.page("Year")!.sheet.value("E3")), "6");
await cmds.run("sheet.activate name=Year");
eq("refusals", [await message("data.consolidate to=A1 sources=Nope!A1:B2"), await message("data.consolidate to=A1 sources=Jan!A1:B2 function=median"), await message("data.consolidate to=A1 sources=Jan!A1 labels=middle")],
  ["no sheet Nope", "median is not a consolidation function (sum, count, counta, average, max, min, product, stdev, stdevp, var, varp)", "middle is not top, left or both"]);
const said: string[] = [];
await cmds.script('SAY CONSOLIDATE("H1", "Jan!B2:B3,Feb!B2:B3", "max")', (l) => said.push(l));
eq("CONSOLIDATE from REXX", [said, show(wb.sheet.value("H1")), show(wb.sheet.value("H2"))], [["consolidated 2 ranges at H1 by max: 2 cells set, 2 changed"], "5", "1"]);

// expect: ok by position
// expect: ok by labels, top and left: their union, any case
// expect: ok other functions
// expect: ok two sheets by their labels, as values
// expect: ok linked: a formula over the cells, following them
// expect: ok followed
// expect: ok undone, the linked total back
// expect: ok refusals
// expect: ok CONSOLIDATE from REXX
