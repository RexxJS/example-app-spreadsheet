// AutoFit (3.5): columns as wide as their widest shown text, rows as tall
// as their lines and font; an empty column, a plain row, the grid's; one
// step of undo; AUTOFIT from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.execute("cell.set", { cell, input });
await set("B1", "a long heading here");
await set("B2", "x");
await set("C1", "1");
await cmds.run("format.font range=D1 size=26");
await set("D1", "big");
const fit = async (id: string, range: string) => (await cmds.execute(id, { range })).output;
eq("each column by its widest text; an empty one the grid's", [await fit("columns.autofit", "B1:E1"), wb.colWidths()], [[157, 24, 59, 0], [[1, 157], [2, 24], [3, 59]]]);
await cmds.run("edit.undo");
eq("one step of undo", wb.colWidths(), []);
await cmds.run("format.wrap range=B5");
await set("B5", "one\ntwo\nthree");
eq("rows by their lines and font; a plain row the grid's", [await fit("rows.autofit", "A1:A6"), wb.rowHeights()], [[46, 0, 0, 0, 65, 0], [[0, 46], [4, 65]]]);
const said: string[] = [];
eq("AUTOFIT from REXX", [(await cmds.script('SAY AUTOFIT("C")\nSAY AUTOFIT(1)\nLET r = AUTOFIT("1C")', (l) => said.push(l))).message, said],
  ['AUTOFIT needs a column ("B", "B:D") or a row (3, "3:5")', ["column C fitted: 24", "row 1 fitted: 46"]]);

// expect: ok each column by its widest text; an empty one the grid's
// expect: ok one step of undo
// expect: ok rows by their lines and font; a plain row the grid's
// expect: ok AUTOFIT from REXX
