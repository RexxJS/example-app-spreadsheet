// Outlines (4.13): rows and columns grouped into levels, a group's detail
// hidden and shown, the level buttons, Clear Outline; Excel's Subtotal (a
// total under each run of a sorted list, a grand total, three levels);
// levels moving with inserted rows, undone, saved with the file; from REXX.
import { depth, detailOf, groups, hiddenAtLevel, regrouped, runs } from "./sheet/outline.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const l = regrouped(regrouped(new Map(), 1, 6, 1), 2, 3, 1);
eq("levels, groups, the buttons, a group's detail", [[...l], depth(l), groups(l), hiddenAtLevel(l, 2), detailOf(l, 2), detailOf(l, 5), detailOf(l, 0)],
  [[[1, 1], [2, 2], [3, 2], [4, 1], [5, 1], [6, 1]], 2, [[1, 6, 1], [2, 3, 2]], [2, 3], [2, 3], [1, 2, 3, 4, 5, 6], []]);
eq("runs of one value, any case", runs(0, 4, (r) => ["a", "A", "b", "b", "c"][r]), [[0, 1, "a"], [2, 3, "b"], [4, 4, "c"]]);

const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("formulas.use language=Excel");
const list = [["Region", "Sales"], ["north", "10"], ["north", "20"], ["south", "5"], ["west", "1"], ["west", "2"]];
for (const [i, r] of list.entries()) for (const [j, x] of r.entries()) await cmds.run(`cell.set cell=${"AB"[j]}${i + 1} input=${x}`);
await cmds.run("cell.set cell=D1 input=after");
eq("Subtotal: a total under each run, a grand total", [await message("data.subtotal range=A1:B6 by=A columns=B"), [...Array(10).keys()].map((i) => `${show(wb.sheet.value(`A${i + 1}`))}|${show(wb.sheet.value(`B${i + 1}`))}`), wb.sheet.input("B4"), wb.sheet.input("B10")],
  ["subtotals of A1:B6 at each change in A: 3 groups, sum of B: 8 cells set, 8 changed",
    ["Region|Sales", "north|10", "north|20", "north Total|30", "south|5", "south Total|5", "west|1", "west|2", "west Total|3", "Grand Total|38"], "=SUBTOTAL(9,B2:B3)", "=SUBTOTAL(9,B2:B9)"]);
eq("its outline: level 2 the totals, level 1 the grand total", [await message("outline.level level=2"), wb.hidden().rows, await message("outline.level level=1"), wb.hidden().rows, await message("outline.level level=3"), wb.hidden().rows],
  ["outline level 2: 5 rows hidden", [1, 2, 4, 6, 7], "outline level 1: 8 rows hidden", [1, 2, 3, 4, 5, 6, 7, 8], "outline level 3: 0 rows hidden", []]);
eq("a group's detail hidden and shown", [await message("outline.hideDetail range=5:5"), wb.hidden().rows, await message("outline.showDetail range=5:5"), wb.hidden().rows],
  ["hid the detail 5 to 5", [4], "showed the detail 5 to 5", []]);
await cmds.run("edit.undo");
await cmds.run("edit.undo");
await cmds.run("edit.undo");
await cmds.run("edit.undo");
await cmds.run("edit.undo");
await cmds.run("edit.undo");
eq("undone, the list as it was", [show(wb.sheet.value("A4")), wb.page(wb.active)!.rowLevels.size], ["south", 0]);
eq("Group: rows, then columns", [await message("outline.group range=2:5"), await message("outline.group range=3:4"), await message("outline.group range=B:B"), [...wb.page(wb.active)!.colLevels]],
  ["grouped rows 2 to 5: 1 level", "grouped rows 3 to 4: 2 levels", "grouped column B: 1 level", [[1, 1]]]);
await cmds.run("rows.insert at=1");
eq("levels move with inserted rows", [...wb.page(wb.active)!.rowLevels].sort((a, b) => a[0] - b[0]), [[2, 1], [3, 2], [4, 2], [5, 1]]);
await cmds.run("file.save name=o");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=o");
eq("saved with the file", [[...wb2.page(wb2.active)!.rowLevels].sort((a, b) => a[0] - b[0]), [...wb2.page(wb2.active)!.colLevels]], [[[2, 1], [3, 2], [4, 2], [5, 1]], [[1, 1]]]);
eq("refusals", [await message("outline.ungroup range=9:9"), await message("outline.level level=9"), await message("outline.hideDetail range=9:9"), await message("data.subtotal range=A1:B6 by=A columns=B function=median")],
  ["row 9 is not grouped", "9 is not an outline level (1 to 8)", "row 9 is in no group", "median is not a subtotal function (average, count, counta, max, min, product, stdev, stdevp, sum, var, varp)"]);
const said: string[] = [];
await cmds.script('SAY UNGROUPROWS(4, 5)\nSAY SHOWOUTLINELEVEL(1)\nSAY HIDEDETAIL("B")\nSAY CLEAROUTLINE()\nSAY GROUPCOLUMNS("C", "D")', (l) => said.push(l));
eq("from REXX", said, ["ungrouped rows 4 to 5: 1 level", "outline level 1: 4 rows hidden", "hid the detail B to B", "rows' outline cleared", "grouped columns C to D: 1 level"]);

// expect: ok levels, groups, the buttons, a group's detail
// expect: ok runs of one value, any case
// expect: ok Subtotal: a total under each run, a grand total
// expect: ok its outline: level 2 the totals, level 1 the grand total
// expect: ok a group's detail hidden and shown
// expect: ok undone, the list as it was
// expect: ok Group: rows, then columns
// expect: ok levels move with inserted rows
// expect: ok saved with the file
// expect: ok refusals
// expect: ok from REXX
