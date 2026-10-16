// Sheet tabs (1.5, with the move done before): a tab's colour, hiding a
// sheet (never the last shown), and copying one with its cells and looks;
// each undone; saved with the file; from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const msg = async (id: string, p: Record<string, string>) => (await cmds.execute(id, p)).message;
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=5");
await cmds.run('cell.set cell=B1 input="=A1*2"');
await cmds.run("format.font range=A1 bold=1");
await cmds.run("cell.comment cell=A1 text=five");
await cmds.run("names.define name=QTY range=A1:A1 scope=sheet");
await cmds.run("rows.hide range=A3");
eq("a tab's colour, by name or #rrggbb, and off", [await msg("sheet.color", { name: "Sheet1", color: "green" }), wb.page("Sheet1")!.tabColor, await msg("sheet.color", { name: "Sheet1", color: "teal-ish" })],
  ["Sheet1's tab #008000", "#008000", "teal-ish is not a colour (a name such as red, or #rrggbb)"]);
eq("copied: after it, named Name copy, made active, with its cells, looks, notes, names and view", [await msg("sheet.copy", { name: "Sheet1" }), wb.sheetNames(), wb.active,
  show(wb.sheet.value("B1")), wb.notes.get("A1").look, wb.notes.get("A1").comment, wb.sheet.names(), wb.hidden().rows, wb.page("Sheet1 copy")!.tabColor],
  ["copied Sheet1 to Sheet1 copy", ["Sheet1", "Sheet1 copy"], "Sheet1 copy", "10", { bold: true }, "five", [["QTY", "A1"]], [2], "#008000"]);
await cmds.run("cell.set cell=A1 input=7");
eq("its own cells after: the first sheet unchanged", [show(wb.sheet.value("B1")), show(wb.page("Sheet1")!.sheet.value("B1"))], ["14", "10"]);
await cmds.run("edit.undo");
await cmds.run("edit.undo");
eq("the copy undone", wb.sheetNames(), ["Sheet1"]);
await msg("sheet.copy", { name: "Sheet1", to: "Plan" });
eq("hidden: the next one shown made active; the last one shown is not", [await msg("sheet.hide", { name: "Plan" }), wb.active, await msg("sheet.hide", { name: "Sheet1" })],
  ["Plan hidden", "Sheet1", "Sheet1 is the only sheet shown"]);
eq("saved with the file", ((writeSheetFile(wb) as { sheets: { name: string; view: { hidden?: boolean; tabColor?: string } }[] }).sheets[0]).view.hidden, true);
await cmds.run("edit.undo");
eq("hiding undone", wb.page("Plan")!.hiddenSheet, false);
const said: string[] = [];
eq("from REXX", [(await cmds.script('SAY HIDESHEET("Plan")\nSAY ISSHEETHIDDEN("Plan")\nSAY UNHIDESHEET("Plan")\nSAY SETSHEETCOLOR("Plan", "#ff8800")\nSAY COPYSHEET("Plan", "Plan B")', (l) => said.push(l))).message, said, wb.sheetNames()],
  ["ran", ["Plan hidden", "1", "Plan shown", "Plan's tab #ff8800", "copied Plan to Plan B"], ["Sheet1", "Plan", "Plan B"]]);

// expect: ok a tab's colour, by name or #rrggbb, and off
// expect: ok copied: after it, named Name copy, made active, with its cells, looks, notes, names and view
// expect: ok its own cells after: the first sheet unchanged
// expect: ok the copy undone
// expect: ok hidden: the next one shown made active; the last one shown is not
// expect: ok saved with the file
// expect: ok hiding undone
// expect: ok from REXX
