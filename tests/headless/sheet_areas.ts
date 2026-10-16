// Several areas (1.2): "A1:B2,D4:D9" in the commands that work over a
// selection, as one step of undo and one line of the journal; the status
// bar's figures over them all; SELECTRANGE of several from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
for (const [k, v] of [["A1", "1"], ["B2", "2"], ["C3", "3"], ["D4", "4"], ["D9", "5"]]) await cmds.execute("cell.set", { cell: k, input: v });
const steps = () => wb.undoLabel;
eq("range.clear over two areas: one step", [(await cmds.execute("range.clear", { range: "A1:B2,D4:D9" })).message, ["A1", "B2", "C3", "D4", "D9"].map((k) => show(wb.sheet.value(k))), steps()],
  ["cleared A1:B2: 4 cells set, 4 changed; cleared D4:D9: 6 cells set, 6 changed", ["", "", "3", "", ""], "Clear contents"]);
await cmds.run("edit.undo");
eq("undone at once", ["A1", "B2", "D4", "D9"].map((k) => show(wb.sheet.value(k))), ["1", "2", "4", "5"]);
eq("one line of the journal", cmds.journal().filter((l) => l.startsWith("range.clear")), ["range.clear range=A1:B2,D4:D9"]);
await cmds.execute("format.font", { range: "A1,C3", bold: "1" });
eq("a look over areas", [wb.notes.get("A1").look, wb.notes.get("B2").look, wb.notes.get("C3").look], [{ bold: true }, undefined, { bold: true }]);
eq("an area that is none: nothing done", [(await cmds.execute("range.clear", { range: "A1,ZZ9" })).message, show(wb.sheet.value("A1"))], ["ZZ9 is not a range on the sheet", "1"]);
eq("the status bar's figures over them all, each cell once", (await cmds.execute("range.stats", { range: "A1:B2,B2:D4" })).message, "Average: 2.5  Count: 4  Numerical Count: 4  Min: 1  Max: 4  Sum: 10");
const said: string[] = [];
eq("SELECTRANGE of several from REXX", [(await cmds.script('SAY SELECTRANGE("A1:B2,D4:D9")\nSAY SELECTIONSTATS(GETSELECTION(), "sum")', (l) => said.push(l))).message, said], ["ran", ["A1:B2,D4:D9", "12"]]);

// expect: ok range.clear over two areas: one step
// expect: ok undone at once
// expect: ok one line of the journal
// expect: ok a look over areas
// expect: ok an area that is none: nothing done
// expect: ok the status bar's figures over them all, each cell once
// expect: ok SELECTRANGE of several from REXX
