// Whole rows and columns (1.1): "B:B", "B:D", "3:3", "3:5" as Excel
// writes them, in every command that takes a range; the selection asked
// for (view.select) and read back; SELECTRANGE and GETSELECTION from REXX.
import { parseRange } from "./sheet/address.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("whole columns and rows as two corners", [parseRange("B:B"), parseRange("b:d"), parseRange("3:3"), parseRange("$2:$4"), parseRange("0:1"), parseRange("B:3")],
  [[{ col: 1, row: 0 }, { col: 1, row: 99 }], [{ col: 1, row: 0 }, { col: 3, row: 99 }], [{ col: 0, row: 2 }, { col: 25, row: 2 }], [{ col: 0, row: 1 }, { col: 25, row: 3 }], null, null]);
const wb = new Workbook();
const cmds = new Commands(wb);
for (const [k, v] of [["B1", "1"], ["B50", "2"], ["C3", "3"], ["A3", "4"]]) await cmds.execute("cell.set", { cell: k, input: v });
eq("a column cleared, one step of undo", [(await cmds.execute("range.clear", { range: "B:B" })).message, show(wb.sheet.value("B50")), show(wb.sheet.value("C3"))], ["cleared B1:B100: 100 cells set, 100 changed", "", "3"]);
await cmds.run("edit.undo");
eq("undone", show(wb.sheet.value("B50")), "2");
eq("a row's sum", (await cmds.execute("range.stats", { range: "3:3" })).message, "Average: 3.5  Count: 2  Numerical Count: 2  Min: 3  Max: 4  Sum: 7");
eq("selected, and read back", [(await cmds.execute("view.select", { range: "B:B" })).message, wb.selecting, (await cmds.execute("view.select", { range: "all" })).output, (await cmds.execute("view.selection")).output],
  ["selected B1:B100", true, "A1:Z100", "A1:Z100"]);
const said: string[] = [];
eq("SELECTRANGE and GETSELECTION from REXX", [(await cmds.script('SAY SELECTRANGE("3:5")\nSAY GETSELECTION()\nCALL CLEARCELL "C:C"\nSAY GETCELL("C3") "|"', (l) => said.push(l))).message, said], ["ran", ["A3:Z5", "A3:Z5", " |"]]);

// expect: ok whole columns and rows as two corners
// expect: ok a column cleared, one step of undo
// expect: ok undone
// expect: ok a row's sum
// expect: ok selected, and read back
// expect: ok SELECTRANGE and GETSELECTION from REXX
