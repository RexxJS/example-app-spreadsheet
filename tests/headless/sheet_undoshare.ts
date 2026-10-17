// Undo's steps share what they did not change (each a snapshot of the
// whole workbook): a sheet's cells and notes saved once while they stay as
// they are, so a hundred steps on one sheet do not copy every other one a
// hundred times. Undo and redo still put each sheet back.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("cell.set cell=A1 input=1");
await cmds.run("sheet.add name=Big");
for (let i = 1; i <= 50; i++) await cmds.run(`cell.set cell=B${i} input=${i}`);
await cmds.run("cell.comment cell=B1 text=kept");
await cmds.run("sheet.activate name=Sheet1");
const big = wb.page("Big")!;
const s1 = big.sheet.saved(), n1 = big.notes.saved();
await cmds.run("cell.set cell=A2 input=2");
await cmds.run("cell.set cell=A3 input=3");
eq("an unchanged sheet's snapshot shared", [big.sheet.saved() === s1, big.notes.saved() === n1, wb.page("Sheet1")!.sheet.saved() === wb.page("Sheet1")!.sheet.saved()], [true, true, true]);
await cmds.run("sheet.activate name=Big");
await cmds.run("cell.set cell=B1 input=100");
await cmds.run("cell.comment cell=B2 text=new");
eq("a changed one's not", [big.sheet.saved() === s1, big.notes.saved() === n1, s1.inputs.get("B1"), n1.has("B2")], [false, false, "1", false]);
await cmds.run("edit.undo");
await cmds.run("edit.undo");
eq("undo puts each back", [show(wb.sheet.value("B1")), wb.notes.get("B2").comment, wb.notes.get("B1").comment], ["1", undefined, "kept"]);
await cmds.run("edit.redo");
eq("and redo", show(wb.sheet.value("B1")), "100");

// expect: ok an unchanged sheet's snapshot shared
// expect: ok a changed one's not
// expect: ok undo puts each back
// expect: ok and redo
