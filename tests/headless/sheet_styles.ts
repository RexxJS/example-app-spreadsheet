// Cell styles and table styles (4.6): a named look applied to cells and
// edited in one place (every cell of it drawn anew), new ones by example or
// based on another, deleted (its cells Normal again), under each cell's own
// look; table styles colouring a header row, banded rows and columns, first
// and last columns; undone, saved with the file; from REXX.
import { CELL_STYLES, under } from "./sheet/styles.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

eq("a style under a cell's own look", under({ bold: true, fill: "#c6efce", borders: { bottom: "#000000" } }, { fill: "#ffffff", borders: { top: "#ff0000" } }),
  { bold: true, fill: "#ffffff", borders: { bottom: "#000000", top: "#ff0000" } });
const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("cell.set cell=A1 input=12.5");
await cmds.run("cell.set cell=A2 input=3");
eq("the gallery: Good on two cells", [await message("style.apply range=A1:A2 name=good"), wb.styleOf("A1"), wb.styleOf("A2").includes("background: #c6efce")],
  ["A1:A2: Good", "text-align: end; background: #c6efce; color: #006100", true]);
eq("edited in one place, every cell of it anew", [await message("style.define name=Good fill=#00ff00 bold=1"), wb.styleOf("A2")],
  ["cell style Good: bold, fill #00ff00, text #006100", "text-align: end; font-weight: bold; background: #00ff00; color: #006100"]);
await cmds.run("format.fill range=A2 color=yellow");
eq("a cell's own look over its style's", wb.styleOf("A2"), "text-align: end; font-weight: bold; background: #ffff00; color: #006100");
eq("a theme's roles: Heading 1's accent line", [await message("style.apply range=B1 name=\"heading 1\""), wb.styleOf("B1")],
  ["B1:B1: Heading 1", "font-weight: bold; font-size: 15px; color: #44546a; border-bottom: 3px solid #4472c4"]);
eq("a style's number format", [await message("style.apply range=A1 name=Currency"), wb.textOf("A1")], ["A1:A1: Currency", "$12.50"]);
eq("a new style by example, and based on another", [await message("style.define name=Mine from=A2"), await message("style.define name=Loud basedOn=Bad size=20")],
  ["cell style Mine: bold, fill #ffff00, text #006100", "cell style Loud: size 20, fill #ffc7ce, text #9c0006"]);
await cmds.run("style.apply range=C1 name=Loud");
eq("deleted: its cells Normal again", [await message("style.delete name=Loud"), wb.notes.get("C1").cellStyle, await message("style.delete name=Normal"), await message("style.apply range=C1 name=Nope")],
  ["deleted cell style Loud", undefined, "Normal cannot be deleted", "no cell style Nope (" + [...wb.styles.keys()].join(", ") + ")"]);
await cmds.run("edit.undo");
eq("undone: Loud back", [wb.styles.has("Loud"), wb.notes.get("C1").cellStyle], [true, "Loud"]);
for (const [k, v] of [["E1", "Item"], ["F1", "Qty"], ["E2", "a"], ["F2", "1"], ["E3", "b"], ["F3", "2"], ["E4", "c"], ["F4", "3"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
await cmds.run("table.define name=T range=E1:F4 headers=1");
eq("a table style: the header row, banded rows, the first column", [await message("table.style name=T style=medium firstColumn=1"), wb.styleOf("E1"), wb.styleOf("E2"), wb.styleOf("F3")],
  ["T: Medium style, banded rows, first column", "background: #4472c4; color: #ffffff; font-weight: bold; border-bottom: 1px solid #4472c4", "background: #d9e1f2; font-weight: bold", "text-align: end"]);
eq("banded columns, no banded rows", [await message("table.style name=T bandedRows=0 bandedColumns=1 firstColumn=0"), wb.styleOf("E3"), wb.styleOf("F3")],
  ["T: Medium style, banded columns", "background: #d9e1f2", "text-align: end"]);
eq("refusals", await message("table.style name=T style=Plaid"), "Plaid is not a table style (Light, Medium, Dark, Green, Orange, Grey, none)");
await cmds.run("file.save name=s");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=s");
eq("saved with the file: the edited style, the new ones, the cells', the table's", [wb2.styles.get("Good"), wb2.styles.has("Mine"), wb2.notes.get("A2").cellStyle, wb2.styleOf("E3"), Object.keys(CELL_STYLES).every((n) => wb2.styles.has(n))],
  [{ fill: "#00ff00", color: "#006100", bold: true }, true, "Good", "background: #d9e1f2", true]);
const said: string[] = [];
await cmds.script('SAY DEFINESTYLE("Money", "bold=1;number=$#,##0.00")\nSAY APPLYSTYLE("G1", "Money")\nSAY SETTABLESTYLE("T", "Green", "bandedRows=1")\ns = LISTSTYLES()\nSAY s.0\nSAY DELETESTYLE("Money")', (l) => said.push(l));
eq("from REXX", said, ["cell style Money: bold, number $#,##0.00", "G1:G1: Money", "T: Green style, banded rows, banded columns", String(wb.styles.size + 1), "deleted cell style Money"]);

// expect: ok a style under a cell's own look
// expect: ok the gallery: Good on two cells
// expect: ok edited in one place, every cell of it anew
// expect: ok a cell's own look over its style's
// expect: ok a theme's roles: Heading 1's accent line
// expect: ok a style's number format
// expect: ok a new style by example, and based on another
// expect: ok deleted: its cells Normal again
// expect: ok undone: Loud back
// expect: ok a table style: the header row, banded rows, the first column
// expect: ok banded columns, no banded rows
// expect: ok refusals
// expect: ok saved with the file: the edited style, the new ones, the cells', the table's
// expect: ok from REXX
