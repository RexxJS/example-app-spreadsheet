// Borders (4.3): Excel's Borders menu over a range (its outline, every
// side, the lines inside it, across or down only, one side of it, none),
// each in a line style (thin, medium, thick, double, dashed, dotted, hair)
// and a colour; a style of none takes those lines off (the eraser);
// undone, saved with the file; SETBORDER and CLEARBORDERS from REXX.
import { borderLine } from "./sheet/formats.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
const sides = (k: string) => (["top", "right", "bottom", "left"] as const).map((e) => wb.notes.get(k).look?.borders?.[e] ?? "-").join(" ");
eq("a line's style as the grid draws it", [borderLine("#000000"), borderLine("medium #ff0000"), borderLine("thick #000000"), borderLine("dashed #333333"), borderLine("double #000000")],
  ["1px solid #000000", "2px solid #ff0000", "3px solid #000000", "1px dashed #333333", "3px double #000000"]);
eq("inside a 2 x 2 range: only the lines between its cells", [await message("format.border range=A1:B2 edge=inside"), sides("A1"), sides("B2")],
  ["A1:B2: inside borders", "- #000000 #000000 -", "#000000 - - #000000"]);
await cmds.run("format.border range=A1:B2 edge=none");
eq("across or down only", [await message("format.border range=A1:A3 edge=horizontal style=dashed"), sides("A1"), sides("A2"), await message("format.border range=C1:E1 edge=vertical"), sides("D1")],
  ["A1:A3: horizontal border, dashed", "- - dashed #000000 -", "dashed #000000 - dashed #000000 -", "C1:E1: vertical border", "- #000000 - #000000"]);
eq("one side is the range's, not every cell's", [await message("format.border range=G1:G3 edge=bottom style=thick color=red"), sides("G1"), sides("G3"), wb.styleOf("G3")],
  ["G1:G3: bottom border, thick", "- - - -", "- - thick #ff0000 -", "border-bottom: 3px solid #ff0000"]);
eq("a thick outline over a thin grid; the eraser takes one side off", [await message("format.border range=I1:J2 edge=all"), await message("format.border range=I1:J2 edge=outline style=thick"), sides("I1"),
  await message("format.border range=I1:J2 edge=top style=none"), sides("I1")],
  ["I1:J2: all borders", "I1:J2: outline borders, thick", "thick #000000 #000000 #000000 thick #000000", "I1:J2: top border, none", "- #000000 #000000 thick #000000"]);
await cmds.run("edit.undo");
eq("undone", sides("I1"), "thick #000000 #000000 #000000 thick #000000");
eq("refusals", [await message("format.border range=A1 edge=all style=wavy"), await message("format.border range=A1 edge=all color=nope")],
  ["wavy is not a line style (thin, medium, thick, double, dashed, dotted, hair, none)", "nope is not a colour (a name such as red, or #rrggbb)"]);
const saved = await cmds.run("file.save name=borders");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=borders");
eq("saved with the file", [saved.ok, (["top", "right", "bottom", "left"] as const).map((e) => wb2.notes.get("G3").look?.borders?.[e] ?? "-").join(" ")], [true, "- - thick #ff0000 -"]);
const said: string[] = [];
await cmds.script('SAY SETBORDER("L1:M2", "outline", "medium", "blue")\nSAY CLEARBORDERS("L1:L2")', (l) => said.push(l));
eq("SETBORDER and CLEARBORDERS from REXX", [said, sides("M2"), sides("L1")], [["L1:M2: outline borders, medium", "L1:L2: no borders"], "- medium #0000ff medium #0000ff -", "- - - -"]);

// expect: ok a line's style as the grid draws it
// expect: ok inside a 2 x 2 range: only the lines between its cells
// expect: ok across or down only
// expect: ok one side is the range's, not every cell's
// expect: ok a thick outline over a thin grid; the eraser takes one side off
// expect: ok undone
// expect: ok refusals
// expect: ok saved with the file
// expect: ok SETBORDER and CLEARBORDERS from REXX
