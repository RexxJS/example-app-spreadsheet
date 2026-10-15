// To do 13: another sheet's own name, written from here (Costs!QTY,
// 'My Costs'!QTY), in Excel formulas: the lexer's qualified name, the
// workbook's context answering it, other sheets compiled again when a
// sheet's own names change (and when undo puts them back), and a sheet
// renamed carrying it along.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { tokens } from "./sheet/lexer.ts";
import { show } from "./sheet/format.ts";
import { CellError } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

eq("the lexer: a qualified name, a quoted sheet's, and a cell there is still a reference",
  [...tokens("Costs!qty+'My Costs'!QTY+Costs!B2")].filter((t) => t.kind === "name" || t.kind === "ref").map((t) => `${t.kind} ${(t as { name: string }).name}`),
  ["name COSTS!QTY", "name MY COSTS!QTY", "ref COSTS!B2"]);

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.run(`cell.set cell=${cell} input="${input.replace(/"/g, '""')}"`);
const on = (sheet: string, k: string) => { const v = wb.page(sheet)!.sheet.value(k); return v instanceof CellError ? `${v.code} ${v.why}` : show(v); };

await cmds.run("sheet.add name=Costs");
for (const [r, v] of ["2", "3", "5"].entries()) await set(`B${r + 1}`, v);
await cmds.run("names.define name=QTY range=B1:B3 scope=sheet");
await cmds.run("names.define name=FIRST range=B1 scope=sheet");
await cmds.run("sheet.activate name=Sheet1");
await set("A1", "=SUM(Costs!QTY)");
await set("A2", "=Costs!first*10");
await set("A3", "=SUM(Costs!NOPE)");
await set("A4", "=QTY");
eq("Costs!QTY is Costs's own range; one cell is that cell; an unknown one #NAME?; bare QTY is not Costs's", [on("Sheet1", "A1"), on("Sheet1", "A2"), on("Sheet1", "A3"), on("Sheet1", "A4")],
  ["10", "20", "#NAME? no name Costs!NOPE", "#NAME? no name QTY"]);

await cmds.run("sheet.activate name=Costs");
await set("B2", "13");
eq("an edit there ripples here", on("Sheet1", "A1"), "20");
await cmds.run("names.define name=QTY range=B1:B2 scope=sheet");
eq("the name redefined there: read again here", on("Sheet1", "A1"), "15");
await cmds.run("names.delete name=QTY scope=sheet");
eq("deleted there: #NAME? here", on("Sheet1", "A1"), "#NAME? no name Costs!QTY");
await cmds.run("edit.undo");
eq("undone: read again", on("Sheet1", "A1"), "15");

await cmds.run('sheet.rename name=Costs to="My Costs"');
eq("Costs renamed: the formula follows it, quoted", [wb.page("Sheet1")!.sheet.input("A1"), wb.page("Sheet1")!.sheet.input("A2"), on("Sheet1", "A1")],
  ["=SUM('My Costs'!QTY)", "='My Costs'!first*10", "15"]);
await cmds.run('sheet.delete name="My Costs"');
eq("the sheet deleted: #NAME?", on("Sheet1", "A2").split(" ")[0], "#NAME?");
// expect: ok the lexer: a qualified name, a quoted sheet's, and a cell there is still a reference
// expect: ok Costs!QTY is Costs's own range; one cell is that cell; an unknown one #NAME?; bare QTY is not Costs's
// expect: ok an edit there ripples here
// expect: ok the name redefined there: read again here
// expect: ok deleted there: #NAME? here
// expect: ok undone: read again
// expect: ok Costs renamed: the formula follows it, quoted
// expect: ok the sheet deleted: #NAME?
