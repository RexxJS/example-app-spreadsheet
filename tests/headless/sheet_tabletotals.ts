// A table's totals row, calculated columns, banded rows and Convert to
// Range (4.9): Total Row in the empty row under the table, SUBTOTAL over
// each column (the first "Total", the last a sum, by Excel's default), the
// totals row out of the data ([#Totals] its own area); a formula down a
// column; every other data row shaded; the table gone and its structured
// references written as the cells they came to. Each undone and saved;
// REXX's SETTABLETOTALS, SETCALCULATEDCOLUMN, SETTABLEBANDING, TABLETORANGE.
import { unstructured } from "./sheet/structured.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
const v = (k: string) => show(wb.sheet.value(k));
await cmds.run("formulas.use language=Excel");
for (const [k, x] of [["A1", "Item"], ["B1", "Qty"], ["C1", "Price"], ["A2", "pens"], ["B2", "3"], ["C2", "2"], ["A3", "ink"], ["B3", "1"], ["C3", "10"]]) await cmds.run(`cell.set cell=${k} input=${x}`);
await cmds.run("table.define name=Sales range=A1:C3 headers=1");
eq("Total Row: Excel's defaults, then each column's own", [await message("table.totals name=Sales"), wb.sheet.input("A4"), wb.sheet.input("C4"), v("C4"), (await cmds.run("table.get name=Sales")).output],
  ["Sales: totals in row 4 (Item label, Qty none, Price sum): 3 cells set, 3 changed", "Total", "=SUBTOTAL(109,Sales[Price])", "12",
    { name: "Sales", range: "A1:C4", headers: true, columns: ["Item", "Qty", "Price"], types: {}, descriptions: {}, sort: null, totals: { Item: "label", Qty: "none", Price: "sum" }, banded: false }]);
eq("an aggregate per column; the totals row out of the data", [await message("table.totals name=Sales functions=Qty=average,Price=max"), v("B4"), v("C4"),
  (await cmds.run('cell.set cell=E1 input="=SUM(Sales[Qty])"')).ok, v("E1"), (await cmds.run('cell.set cell=E2 input="=Sales[[#Totals],[Qty]]"')).ok, v("E2"), wb.styleOf("B4").includes("font-weight: bold")],
  ["Sales: totals in row 4 (Item label, Qty average, Price max): 3 cells set, 3 changed", "2", "10", true, "4", true, "2", true]);
eq("refusals", [await message("table.totals name=Sales functions=Qty=median"), await message("table.totals name=Sales functions=Cost=sum"), await message("table.totals name=Nope")],
  ["median is not a total (average, count, counta, max, min, product, stdev, stdevp, sum, var, varp, none, label)", "Sales has no column Cost", "no table Nope"]);
eq("a calculated column", [await message('table.calculatedColumn name=Sales column=Item formula="=[@Qty]*[@Price]"'), wb.sheet.input("A3"), v("A2"), v("A3")],
  ["Sales[Item] = =Sales[@Qty]*Sales[@Price]: 2 cells set, 2 changed", "=Sales[@Qty]*Sales[@Price]", "6", "10"]);
eq("banded rows: the first data row and every other", [await message("table.banding name=Sales"), wb.styleOf("B2").includes("background: #d9e1f2"), wb.styleOf("B3").includes("background")],
  ["Sales: banded rows", true, false]);
await cmds.run("file.save name=t");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=t");
eq("saved with the file", [(await cmds2.run("table.get name=Sales")).output, show(wb2.sheet.value("E2"))],
  [{ name: "Sales", range: "A1:C4", headers: true, columns: ["Item", "Qty", "Price"], types: {}, descriptions: {}, sort: null, totals: { Item: "label", Qty: "average", Price: "max" }, banded: true }, "2"]);
eq("references written as cells", [unstructured('=SUM(Sales[Qty]) & "Sales[Qty]"', wb.sheet.table("Sales")!, (a) => show(wb.sheet.value(`${"ABC"[a.col]}${a.row + 1}`)), 0),
  unstructured("=Sales[@Qty]*2", wb.sheet.table("Sales")!, (a) => show(wb.sheet.value(`${"ABC"[a.col]}${a.row + 1}`)), 2)], ['=SUM($B$2:$B$3) & "Sales[Qty]"', "=$B3*2"]);
eq("Convert to Range", [await message("table.toRange name=Sales"), wb.sheet.input("E1"), wb.sheet.input("A2"), wb.sheet.input("C4"), v("C4"), wb.sheet.tables().length, wb.styleOf("B2").includes("background: #d9e1f2")],
  ["Sales is a range again: 6 cells set, 6 changed", "=SUM($B$2:$B$3)", "=$B2*$C2", "=SUBTOTAL(104,$C$2:$C$3)", "10", 0, true]);
await cmds.run("edit.undo");
eq("undone: the table back", [wb.sheet.tables().length, wb.sheet.input("E1")], [1, "=SUM(Sales[Qty])"]);
await cmds.run("table.totals name=Sales on=0");
eq("the totals row taken off", [wb.sheet.input("A4"), (await cmds.run("table.get name=Sales")).message.includes('"range":"A1:C3"')], ["", true]);
const said: string[] = [];
await cmds.script('SAY SETTABLETOTALS("Sales", "Qty=sum")\nSAY SETCALCULATEDCOLUMN("Sales", "Item", "=[@Qty]+1")\nSAY SETTABLEBANDING("Sales", 0)\nSAY TABLETORANGE("Sales")', (l) => said.push(l));
eq("from REXX", said, ["Sales: totals in row 4 (Item label, Qty sum, Price sum): 3 cells set, 4 changed", "Sales[Item] = =Sales[@Qty]+1: 2 cells set, 2 changed", "Sales: no banded rows",
  "Sales is a range again: 6 cells set, 6 changed"]);

// expect: ok Total Row: Excel's defaults, then each column's own
// expect: ok an aggregate per column; the totals row out of the data
// expect: ok refusals
// expect: ok a calculated column
// expect: ok banded rows: the first data row and every other
// expect: ok saved with the file
// expect: ok references written as cells
// expect: ok Convert to Range
// expect: ok undone: the table back
// expect: ok the totals row taken off
// expect: ok from REXX
