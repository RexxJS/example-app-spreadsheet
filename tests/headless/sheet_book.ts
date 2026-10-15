// app/sheet/workbook.ts and commands.ts, To do 8: sheets moved among the
// tabs, and names for the whole workbook (Excel's default scope): read on
// every sheet, in both languages, hidden on a sheet by its own name of the
// same spelling, following their sheet when it is renamed, #REF! when it
// goes; undone, saved and opened.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
const on = (sheet: string, k: string) => show(wb.page(sheet)!.sheet.value(k));

// Moving sheets.
await cmds.run("sheet.add name=Costs");
await cmds.run("sheet.add name=Prices");
eq("sheet.move to a place, or a way", [await message("sheet.move name=Prices to=1"), wb.sheetNames(), await message("sheet.move name=Sheet1 way=left"), wb.sheetNames()],
  ["Prices is sheet 1 of 3", ["Prices", "Sheet1", "Costs"], "Sheet1 is sheet 1 of 3", ["Sheet1", "Prices", "Costs"]]);
eq("the active sheet stays active", wb.active, "Prices");
eq("refusals: past either end, neither or both, no such sheet", [
  await message("sheet.move name=Sheet1 way=left"), await message("sheet.move name=Costs to=4"), await message("sheet.move name=Costs"),
  await message("sheet.move name=Nope to=1"), await message("sheet.move name=Costs way=up"),
], ["Sheet1 cannot move left: there are 3 sheets", "Costs cannot move to 4: there are 3 sheets", "sheet.move needs to or way", "no sheet Nope", "up is not left or right"]);
await cmds.run("edit.undo");
eq("one undo puts the order back", wb.sheetNames(), ["Prices", "Sheet1", "Costs"]);
await cmds.run("edit.redo");

// Workbook names.
await cmds.run("sheet.activate name=Sheet1");
await cmds.run("cell.set cell=B1 input=10");
await cmds.run("cell.set cell=B2 input=20");
eq("a name is the workbook's", await message("names.define name=qty range=B1:B2"), "QTY is Sheet1!B1:B2");
await cmds.run("sheet.activate name=Costs");
await cmds.run('cell.set cell=A1 input="=SUM(QTY)"');
await cmds.run('cell.set cell=A2 input="=QTY"');
eq("read on another sheet, as that sheet's range (bare, it spills)", [on("Costs", "A1"), on("Costs", "A2"), on("Costs", "A3"), wb.page("Costs")!.sheet.reads("A1")], ["30", "10", "20", ["SHEET1!B1", "SHEET1!B2"]]);
await cmds.run("sheet.activate name=Sheet1");
await cmds.run("cell.set cell=B1 input=15");
eq("and follows it", on("Costs", "A1"), "35");

await cmds.run("sheet.activate name=Costs");
await cmds.run("cell.set cell=C1 input=1");
eq("a sheet's own name of the same spelling hides it there", await message("names.define name=QTY range=C1 scope=sheet"), "Costs!QTY is C1");
eq("on that sheet only", [on("Costs", "A1"), await message('cell.set cell=B9 input="=QTY"'), wb.page("Sheet1")!.sheet.named("QTY")],
  ["1", "B9 = 1", { range: [{ col: 1, row: 0 }, { col: 1, row: 1 }], sheet: null }]);
eq("the Name Manager lists the workbook's, then each sheet's own", await message("names.list"), "QTY Sheet1!B1:B2, Costs!QTY C1");
eq("a sheet's own name is for its own range; a workbook's for any sheet's", [
  await message("names.define name=X range=Sheet1!B1 scope=sheet"), await message("names.define name=X range=Sheet1!B1"), await message("names.define name=X range=Nope!B1"),
  await message("names.define name=X range=B1 scope=book"),
], ["a sheet's own name is for a range of it: Sheet1!B1", "X is Sheet1!B1", "no sheet Nope", "book is not a scope: workbook or sheet"]);
eq("Delete name takes the sheet's own first", [await message("names.delete name=QTY"), on("Costs", "A1"), await message("names.delete name=QTY scope=sheet")],
  ["Costs!QTY deleted", "35", "no name QTY of Costs's own"]);

// RexxJS formulas read it too, and a named LAMBDA is the workbook's.
await cmds.run("formulas.use language=RexxJS");
await cmds.run('cell.set cell=D1 input="=SUM_RANGE(""QTY"") + 1"');
await cmds.run('names.define name=TWICE formula="=LAMBDA(x, x * 2)"');
await cmds.run('cell.set cell=D2 input="=TWICE(D1)"');
eq("in RexxJS formulas, a name of another sheet's range and a workbook's LAMBDA", [on("Costs", "D1"), on("Costs", "D2")], ["36", "72"]);
await cmds.run("formulas.use language=Excel");

await cmds.run("sheet.rename name=Sheet1 to=Data");
eq("renaming its sheet, it follows", [await message("names.list"), on("Costs", "A1")], ["QTY Data!B1:B2, TWICE =LAMBDA(x, x * 2), X Data!B1", "35"]);
const saved = (await cmds.run("file.export")).output as string;
eq("saved as the workbook's", JSON.parse(saved).bookNames, { QTY: "Data!B1:B2", TWICE: "=LAMBDA(x, x * 2)", X: "Data!B1" });
await cmds.run("sheet.delete name=Data");
eq("its sheet deleted, what reads it is #REF!", on("Costs", "A1"), "#REF!");
await cmds.run("edit.undo");
eq("undone, read again", on("Costs", "A1"), "35");

await cmds.run("file.new");
await cmds.run(`file.import text="${saved.replace(/"/g, '""')}"`);
await cmds.run("sheet.activate name=Costs");
await cmds.run('cell.set cell=E1 input="=TWICE(QTY)"');
eq("and opened", [await message("names.list"), on("Costs", "A1"), on("Costs", "E1"), on("Costs", "E2")], ["QTY Data!B1:B2, TWICE =LAMBDA(x, x * 2), X Data!B1", "35", "30", "40"]);
eq("sheets in the order they were in", wb.sheetNames(), ["Data", "Prices", "Costs"]);

// References made at run time reach another sheet, and follow it.
await cmds.run("sheet.activate name=Data");
await cmds.run('cell.set cell=F1 input="=INDIRECT(""Costs!E1"")"');
await cmds.run("cell.set cell=H1 input=0");
await cmds.run('cell.set cell=G1 input="=OFFSET(B1, H1, 0)"');
eq("INDIRECT of another sheet's cell reads it, and is linked to it", [on("Data", "F1"), wb.page("Data")!.sheet.reads("F1")], ["30", ["COSTS!E1"]]);
await cmds.run("sheet.activate name=Costs");
await cmds.run("cell.set cell=E1 input=99");
eq("a change at the pointed-at cell shows", on("Data", "F1"), "99");
await cmds.run("sheet.activate name=Data");
eq("OFFSET by a cell's value: linked to where it came to", [on("Data", "G1"), wb.page("Data")!.sheet.reads("G1")], ["15", ["B1", "H1"]]);
await cmds.run("cell.set cell=H1 input=1");
await cmds.run("cell.set cell=B2 input=21");
eq("moved, it follows the cell it now comes to", [on("Data", "G1"), wb.page("Data")!.sheet.reads("G1")], ["21", ["B1", "H1", "B2"]]);

// Asking about the workbook: sheets by number, a cell's formula as typed.
await cmds.run('cell.set cell=I1 input="=SHEET()"');
await cmds.run('cell.set cell=I2 input="=SHEET(Costs!A1)"');
await cmds.run('cell.set cell=I3 input="=SHEET(""prices"")"');
await cmds.run('cell.set cell=I4 input="=SHEETS()"');
await cmds.run('cell.set cell=I5 input="=FORMULATEXT(Costs!A1)"');
await cmds.run('cell.set cell=I6 input="=ISFORMULA(Costs!A1) & ISFORMULA(B1)"');
eq("SHEET, SHEETS and FORMULATEXT across the workbook", ["I1", "I2", "I3", "I4", "I5", "I6"].map((k) => on("Data", k)), ["1", "3", "2", "3", "=SUM(QTY)", "TRUEFALSE"]);
await cmds.run("sheet.move name=Costs to=1");
eq("a sheet moved, the numbers follow", ["I1", "I2", "I3"].map((k) => on("Data", k)), ["2", "1", "3"]);
await cmds.run("sheet.activate name=Costs");
await cmds.run('cell.set cell=A1 input="=SUM(QTY)*2"');
eq("the formula retyped, FORMULATEXT follows", on("Data", "I5"), "=SUM(QTY)*2");
await cmds.run("sheet.activate name=Data");
await cmds.run('cell.set cell=I7 input="=INFO(""numfile"")"');
eq("INFO(numfile) counts the workbook's sheets", on("Data", "I7"), "3");

// expect: ok sheet.move to a place, or a way
// expect: ok the active sheet stays active
// expect: ok refusals: past either end, neither or both, no such sheet
// expect: ok one undo puts the order back
// expect: ok a name is the workbook's
// expect: ok read on another sheet, as that sheet's range (bare, it spills)
// expect: ok and follows it
// expect: ok a sheet's own name of the same spelling hides it there
// expect: ok on that sheet only
// expect: ok the Name Manager lists the workbook's, then each sheet's own
// expect: ok a sheet's own name is for its own range; a workbook's for any sheet's
// expect: ok Delete name takes the sheet's own first
// expect: ok in RexxJS formulas, a name of another sheet's range and a workbook's LAMBDA
// expect: ok renaming its sheet, it follows
// expect: ok saved as the workbook's
// expect: ok its sheet deleted, what reads it is #REF!
// expect: ok undone, read again
// expect: ok and opened
// expect: ok sheets in the order they were in
// expect: ok INDIRECT of another sheet's cell reads it, and is linked to it
// expect: ok a change at the pointed-at cell shows
// expect: ok OFFSET by a cell's value: linked to where it came to
// expect: ok moved, it follows the cell it now comes to
// expect: ok SHEET, SHEETS and FORMULATEXT across the workbook
// expect: ok a sheet moved, the numbers follow
// expect: ok the formula retyped, FORMULATEXT follows
// expect: ok INFO(numfile) counts the workbook's sheets
