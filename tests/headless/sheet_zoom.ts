// Zoom (3.7): a percentage or Excel's steps in and out; every column and
// row, and each cell's font, scaled for the grid; a view, not a step;
// SETZOOM and GETZOOM from REXX. (The status bar's figures came with 3.0.5.)
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const z = async (p: Record<string, string>) => (await cmds.execute("view.zoom", p)).message;
await cmds.run("columns.width range=B1 width=100");
await cmds.run("cell.set cell=A1 input=hi");
eq("by a percentage, or a step in or out", [await z({ percent: "150" }), await z({ by: "in" }), await z({ by: "out" }), await z({ by: "out" }), await z({ percent: "100%" })],
  ["zoom 150%", "zoom 200%", "zoom 150%", "zoom 125%", "zoom 100%"]);
eq("refused", [await z({ percent: "5" }), await z({})], ["5 is not a zoom (10 to 400)", "view.zoom needs percent= or by=in|out"]);
eq("at 100% only the columns set", wb.shownColWidths(80), [[1, 100]]);
await z({ percent: "50" });
eq("at 50%: every column and row scaled, from its own or the grid's", [wb.shownColWidths(80).slice(0, 3), wb.shownRowHeights(24)[0], wb.colWidths()], [[[0, 40], [1, 50], [2, 40]], [0, 12], [[1, 100]]]);
eq("and the font", wb.styleOf("A1"), "font-size: 7px");
eq("a view: not a step to undo", wb.undoLabel, "Set a cell");
const said: string[] = [];
eq("SETZOOM and GETZOOM from REXX", [(await cmds.script('SAY SETZOOM(75)\nSAY SETZOOM("in") GETZOOM()', (l) => said.push(l))).message, said], ["ran", ["75", "100 100"]]);

// expect: ok by a percentage, or a step in or out
// expect: ok refused
// expect: ok at 100% only the columns set
// expect: ok at 50%: every column and row scaled, from its own or the grid's
// expect: ok and the font
// expect: ok a view: not a step to undo
// expect: ok SETZOOM and GETZOOM from REXX
