// Format Painter (3.2): a cell's look, number format and style onto a
// range, a range's tiled over it; not values, comments or rules; one step
// of undo; COPYFORMAT from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=1.5");
await cmds.run("format.font range=A1 bold=1");
await cmds.run("format.fill range=A2 color=yellow");
await cmds.run('cell.format cell=A1 format="0.00"');
await cmds.run("cell.comment cell=A1 text=note");
for (const k of ["B1", "B2", "B3", "B4"]) await cmds.execute("cell.set", { cell: k, input: "2" });
await cmds.run("cell.comment cell=B1 text=mine");
eq("one cell's format over a range", [(await cmds.execute("format.painter", { from: "A1", to: "B1:B2" })).message, wb.notes.get("B2").look, wb.notes.get("B2").format, wb.textOf("B2")],
  ["painted A1's format over B1:B2", { bold: true }, "0.00", "2.00"]);
eq("not the value or the comment", [wb.sheet.input("B1"), wb.notes.get("B1").comment, wb.notes.get("B2").comment], ["2", "mine", undefined]);
await cmds.run("edit.undo");
eq("one step of undo", [wb.notes.get("B2").look, wb.notes.get("B1").format], [undefined, undefined]);
await cmds.execute("format.painter", { from: "A1:A2", to: "B1:B4" });
eq("a range's tiled over the target", [1, 2, 3, 4].map((r) => JSON.stringify(wb.notes.get(`B${r}`).look ?? {})), ['{"bold":true}', '{"fill":"#ffff00"}', '{"bold":true}', '{"fill":"#ffff00"}']);
const said: string[] = [];
eq("COPYFORMAT from REXX", [(await cmds.script('SAY COPYFORMAT("A2", "C1")', (l) => said.push(l))).message, said, wb.notes.get("C1").look], ["ran", ["painted A2's format over C1"], { fill: "#ffff00" }]);

// expect: ok one cell's format over a range
// expect: ok not the value or the comment
// expect: ok one step of undo
// expect: ok a range's tiled over the target
// expect: ok COPYFORMAT from REXX
