// app/sheet/structured.ts, parser.ts and evaluate.ts, To do 10: table
// references in Excel formulas. Sales[Amount], Sales[@Amount],
// Sales[[#Headers],[Qty]:[Price]] and the rest: what they read, as the
// table is when the formula is worked out (a header renamed, a row inserted,
// the table deleted), and what does not parse.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { parseSpec } from "./sheet/structured.ts";
import { show } from "./sheet/format.ts";
import { CellError } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const spec = (s: string) => {
  try {
    return parseSpec(s, 1);
  } catch (e) {
    return (e as Error).message;
  }
};
eq("what is between the brackets", [spec(""), spec("Amount"), spec("@Amount"), spec("@[Unit Price]"), spec("[Qty]:[Price]"), spec("#All")],
  [{ areas: ["data"], thisRow: false, columns: null }, { areas: ["data"], thisRow: false, columns: ["Amount", "Amount"] },
    { areas: [], thisRow: true, columns: ["Amount", "Amount"] }, { areas: [], thisRow: true, columns: ["Unit Price", "Unit Price"] },
    { areas: ["data"], thisRow: false, columns: ["Qty", "Price"] }, { areas: ["all"], thisRow: false, columns: null }]);
eq("areas together, #This Row, and ' making a character plain", [spec("[#Headers],[#Data],[Qty]"), spec("[#This Row],[Amount]"), spec("'#Units"), spec("[a']b]")],
  [{ areas: ["headers", "data"], thisRow: false, columns: ["Qty", "Qty"] }, { areas: [], thisRow: true, columns: ["Amount", "Amount"] },
    { areas: ["data"], thisRow: false, columns: ["#Units", "#Units"] }, { areas: ["data"], thisRow: false, columns: ["a]b", "a]b"] }]);
eq("what does not parse", [spec("[A],[B]"), spec("[#This Row],[#Data]"), spec("@"), spec("[A],B")],
  ["a table reference names one column, or one span of them: [[A],[B]]", "#This Row goes with no other area: [[#This Row],[#Data]]",
    "@ needs a column: Sales[@Amount]", "a table reference's parts are each in brackets: [[A],B]"]);

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.run(`cell.set cell=${cell} input="${input.replace(/"/g, '""')}"`);
const shown = (...keys: string[]) => keys.map((k) => show(wb.sheet.value(k)));
const why = (k: string) => { const v = wb.sheet.value(k); return v instanceof CellError ? `${v.code} ${v.why}` : show(v); };
const rows = [["Item", "Qty", "Price", "Unit Price"], ["pens", "3", "2", "1.5"], ["ink", "1", "10", "4"], ["paper", "2", "5", "0.25"]];
for (const [r, row] of rows.entries()) for (const [c, v] of row.entries()) await set(`${"ABCD"[c]}${r + 1}`, v);
await cmds.run("table.define name=Sales range=A1:D4 headers=1");

await set("F1", "=SUM(Sales[Qty])");
await set("F2", "=SUM(sales[price])");
await set("F3", "=COUNTA(Sales)");
await set("F4", "=SUM(Sales[[Qty]:[Price]])");
await set("F5", "=ROWS(Sales[#All])");
await set("F6", "=INDEX(Sales[#Headers], 1, 4)");
await set("F7", "=SUM(Sales[[#Headers],[#Data],[Qty]])");
eq("a column, any case; the table's name alone; a span of columns; #All; #Headers", shown("F1", "F2", "F3", "F4", "F5", "F6", "F7"),
  ["6", "17", "12", "23", "4", "Unit Price", "6"]);

await set("E2", "=Sales[@Qty]*Sales[@Price]");
await set("E3", "=Sales[@Qty]*Sales[@Price]");
await set("E4", "=Sales[[#This Row],[Qty]]*Sales[@[Unit Price]]");
eq("[@Column]: the formula's own row", shown("E2", "E3", "E4"), ["6", "10", "0.5"]);
await set("E6", "=Sales[@Qty]");
eq("[@...] off the table's rows is #VALUE!, saying why", why("E6"), "#VALUE! [@...] is the formula's own row of Sales, and this formula is not on one of its rows");

await set("F8", "=SUM(Sales[Total])");
await set("F9", "=SUM(Sales[#Totals])");
await set("F10", "=SUM(Costs[Qty])");
await set("F11", "=SUM(Sales[[Qty],[Price]])");
eq("no such column, no totals row, no such table: #REF!; two columns do not parse", [why("F8"), why("F9"), why("F10"), shown("F11")[0]],
  ["#REF! table Sales has no column Total", "#REF! table Sales has no totals row", "#REF! no table Costs", "#ERROR!"]);

await set("C3", "20");
eq("an edit in the table ripples", shown("F2", "E3"), ["27", "20"]);
await set("B1", "Count");
eq("a header renamed: the old name is #REF!, the new one found at once", [why("F1"), (await set("F12", "=SUM(Sales[Count])")) && shown("F12")[0]],
  ["#REF! table Sales has no column Qty", "6"]);
await set("B1", "Qty");

await cmds.run("rows.insert at=3");
eq("a row inserted inside the table: the table grows, its references with it", [(await cmds.run("table.list")).output, shown("F1")],
  [["Sales A1:D5 (Item, Qty, Price, Unit Price)"], ["6"]]);
await set("B3", "4");
eq("and the new row counts (ROWS moved down with the insert, to F6)", shown("F1", "F6"), ["10", "5"]);

await cmds.run("table.delete name=Sales");
eq("the table deleted: what read it is #REF!", why("F1"), "#REF! no table Sales");
await cmds.run("edit.undo");
eq("and undone, read again", shown("F1"), ["10"]);

await cmds.run("table.define name=Plain range=A2:B5 columns=thing,count headers=0");
await set("H1", "=SUM(Plain[count])");
eq("a table with columns named in its definition", shown("H1"), ["10"]);

await cmds.run("formulas.use language=RexxJS");
eq("RexxJS formulas keep TABLE(): no brackets there", [(await set("H2", "=Sales[Qty]")) && shown("H2")[0] !== "10"], [true]);
// expect: ok what is between the brackets
// expect: ok areas together, #This Row, and ' making a character plain
// expect: ok what does not parse
// expect: ok a column, any case; the table's name alone; a span of columns; #All; #Headers
// expect: ok [@Column]: the formula's own row
// expect: ok [@...] off the table's rows is #VALUE!, saying why
// expect: ok no such column, no totals row, no such table: #REF!; two columns do not parse
// expect: ok an edit in the table ripples
// expect: ok a header renamed: the old name is #REF!, the new one found at once
// expect: ok a row inserted inside the table: the table grows, its references with it
// expect: ok and the new row counts (ROWS moved down with the insert, to F6)
// expect: ok the table deleted: what read it is #REF!
// expect: ok and undone, read again
// expect: ok a table with columns named in its definition
// expect: ok RexxJS formulas keep TABLE(): no brackets there
