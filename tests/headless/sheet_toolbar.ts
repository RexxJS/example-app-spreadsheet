// The toolbar row (3.0.2): its new commands, Alignment, Fill Color, Font
// Color and Number Format over a range, drawn, shown, undone and saved;
// REXX's SETALIGNMENT, SETFILLCOLOR, SETTEXTCOLOR and SETNUMBERFORMAT; and
// each button runs an action the menus have too.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { MENUS } from "./sheet/menus.ts";
import type { MenuItem } from "./sheet/menus.ts";
import { TOOLS } from "./sheet/view/toolbar.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const msg = async (id: string, p: Record<string, string>) => (await cmds.execute(id, p)).message;
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=1234.5");
await cmds.run("cell.set cell=A2 input=0.25");
eq("Alignment over a range, and general again", [await msg("format.align", { range: "A1:A2", align: "center" }), wb.styleOf("A1").includes("text-align: center"), await msg("format.align", { range: "A1", align: "general" }), wb.styleOf("A1").includes("text-align: center"), await msg("format.align", { range: "A1", align: "up" })],
  ["A1:A2: aligned center", true, "A1: aligned general", false, "up is not an alignment (left, center, right, general)"]);
eq("Fill Color and Font Color, by name or #rrggbb; none takes it off", [await msg("format.fill", { range: "A1", color: "yellow" }), wb.styleOf("A1").includes("background: #ffff00"), await msg("format.color", { range: "A1", color: "#c00" }), wb.styleOf("A1").includes("color: #cc0000"),
  await msg("format.fill", { range: "A1", color: "none" }), wb.styleOf("A1").includes("background"), await msg("format.color", { range: "A1", color: "mauve-ish" })],
  ["A1: fill #ffff00", true, "A1: text #cc0000", true, "A1: fill none", false, "mauve-ish is not a colour (a name such as red, or #rrggbb)"]);
eq("Comma, Currency and Percent Style, and a pattern", [await msg("format.number", { range: "A1", format: "number" }), wb.textOf("A1"), await msg("format.number", { range: "A1", format: "currency" }), wb.textOf("A1"),
  await msg("format.number", { range: "A2", format: "percent" }), wb.textOf("A2"), await msg("format.number", { range: "A2", format: "0.0" }), wb.textOf("A2"), await msg("format.number", { range: "A2", format: "general" }), wb.textOf("A2"), await msg("format.number", { range: "A2", format: "bold" })],
  ["A1: #,##0.00", "1,234.50", "A1: $#,##0.00", "$1,234.50", "A2: 0%", "25%", "A2: 0.0", "0.3", "A2: general", "0.25", "bold is not a number format (number, currency, percent, general, or a pattern such as 0.00)"]);
await cmds.run('cell.format cell=A1 format="0.000"');
eq("the cell's own format field wins over the toolbar's", wb.textOf("A1"), "1234.500");
await cmds.run('cell.format cell=A1 format=""');
eq("saved with the file, in the look", (writeSheetFile(wb) as { cells: Record<string, { look?: unknown }> }).cells.A1.look, { align: undefined, color: "#cc0000", number: "$#,##0.00" } as unknown);
await cmds.run("edit.undo");
eq("one step of undo each", wb.textOf("A1"), "1234.500");
const said: string[] = [];
eq("from REXX", [(await cmds.script('CALL SETALIGNMENT "B1:B2", "right"\nCALL SETFILLCOLOR "B1", "green"\nCALL SETTEXTCOLOR "B1", "white"\nCALL SETNUMBERFORMAT "A2", "percent"', (l) => said.push(l))).message, wb.styleOf("B1"), wb.textOf("A2")],
  ["ran", "text-align: end; background: #008000; color: #ffffff", "25%"]);
const items = new Set(MENUS.flatMap((m) => m.items.filter((e): e is MenuItem => e !== "-").map((e) => e.label)));
eq("each toolbar button runs an action a menu item runs too", TOOLS.filter(([, , action]) => !items.has(action)).map(([l]) => l), []);
eq("each has its tooltip", TOOLS.filter(([, tip]) => !tip).length, 0);

// expect: ok Alignment over a range, and general again
// expect: ok Fill Color and Font Color, by name or #rrggbb; none takes it off
// expect: ok Comma, Currency and Percent Style, and a pattern
// expect: ok the cell's own format field wins over the toolbar's
// expect: ok saved with the file, in the look
// expect: ok one step of undo each
// expect: ok from REXX
// expect: ok each toolbar button runs an action a menu item runs too
// expect: ok each has its tooltip
