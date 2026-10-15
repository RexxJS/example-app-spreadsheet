// To do 12: what a spreadsheet does that the sheet did not. Comparisons and
// sorts rank every number below every text and text below FALSE and TRUE
// (values.ts's ordered, sortfind.ts's sortOrder), with empty cells last
// whichever way a sort goes; AND, OR and NOT read the text "TRUE" and
// "FALSE"; and a cycle through two sheets is found, #CYCLE! in each cell on
// it, where it used to be passed between them twenty times.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { ordered } from "./sheet/values.ts";
import { show } from "./sheet/format.ts";
import { CellError } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const sign = (n: number) => Math.sign(n);
eq("numbers below text below FALSE below TRUE; text without case", [ordered(99, "1"), ordered("a", "B"), ordered("zz", false), ordered(false, true), ordered(2, 10), ordered("A", "a")].map(sign),
  [-1, -1, -1, -1, -1, 0]);
eq("an empty cell is the other side's empty value", [ordered(null, 0), ordered(null, ""), ordered(null, false), ordered(null, -1), ordered(null, "a"), ordered(null, null)].map(sign),
  [0, 0, 0, 1, -1, 0]);

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.run(`cell.set cell=${cell} input="${input.replace(/"/g, '""')}"`);
const shown = (...keys: string[]) => keys.map((k) => show(wb.sheet.value(k)));

await set("A1", '="10">5');
await set("A2", '=TRUE>99');
await set("A3", '=TRUE=1');
await set("A4", '="abc"<"ABD"');
await set("A5", '=B9=0');
await set("A6", '=B9=""');
await set("A7", '="1"=1');
eq("in formulas: text is above every number, TRUE above both, TRUE is not 1, an empty cell is 0 and empty text",
  shown("A1", "A2", "A3", "A4", "A5", "A6", "A7"), ["TRUE", "TRUE", "FALSE", "TRUE", "TRUE", "TRUE", "FALSE"]);
await set("A8", '=COUNTIF(C1:C6, ">5")');
await set("A9", '=MAX(C1:C6)');

for (const [r, v] of ["pear", "10", "=TRUE", "", "Apple", "2"].entries()) await set(`C${r + 1}`, v);
await set("C7", "=1/0");
await cmds.run("data.sortAscending range=C1:C7 by=C");
eq("Sort A to Z: numbers, text without case, TRUE, errors, then the empty cell", shown("C1", "C2", "C3", "C4", "C5", "C6", "C7"),
  ["2", "10", "Apple", "pear", "TRUE", "#DIV/0!", ""]);
await cmds.run("data.sortDescending range=C1:C7 by=C");
eq("Sort Z to A: reversed, but the empty cell still last", shown("C1", "C2", "C3", "C4", "C5", "C6", "C7"),
  ["#DIV/0!", "TRUE", "pear", "Apple", "10", "2", ""]);
await set("E1", '=SORT({"b";3;TRUE;"A";1})');
await set("F1", '=SORT({"b";3;TRUE;"A";1}, 1, -1)');
eq("SORT the same, both ways", [shown("E1", "E2", "E3", "E4", "E5"), shown("F1", "F2", "F3", "F4", "F5")],
  [["1", "3", "A", "b", "TRUE"], ["TRUE", "b", "A", "3", "1"]]);

await set("G1", "TRUE");
await set("G2", "false");
await set("G3", "maybe");
await set("H1", '=AND("TRUE", 1)');
await set("H2", '=OR("false", G2)');
await set("H3", "=AND(G1:G3)");
await set("H4", '=NOT("false")');
await set("H5", '=AND("maybe")');
eq("AND, OR and NOT read the text TRUE and FALSE (any case); other text is skipped, and with nothing left, #VALUE!",
  shown("H1", "H2", "H3", "H4", "H5"), ["TRUE", "FALSE", "FALSE", "TRUE", "#VALUE!"]);

// A cycle through two sheets.
await cmds.run("sheet.add name=Costs");
await set("A1", "=Sheet1!J1+1");
await cmds.run("sheet.activate name=Sheet1");
await set("J1", "=Costs!A1+1");
const cyc = (v: unknown) => (v instanceof CellError ? `${v.code} ${v.why}` : show(v as never));
eq("a cycle through two sheets: #CYCLE! on both, saying the path", [cyc(wb.sheet.value("J1")), cyc(wb.page("Costs")!.sheet.value("A1"))],
  ["#CYCLE! a circular reference through more than one sheet: SHEET1!J1 → COSTS!A1 → SHEET1!J1",
    "#CYCLE! a circular reference through more than one sheet: COSTS!A1 → SHEET1!J1 → COSTS!A1"]);
await set("K1", "=J1*2");
eq("what reads it is #CYCLE! too", shown("K1"), ["#CYCLE!"]);
await set("J2", "=Costs!A2");
await cmds.run("sheet.activate name=Costs");
await set("A2", "=Sheet1!J2");
eq("even where the values would settle (each only copies the other)", shown("A2"), ["#CYCLE!"]);
await set("A1", "5");
await cmds.run("sheet.activate name=Sheet1");
eq("the cycle broken: worked out again", shown("J1", "K1"), ["6", "12"]);
await cmds.run("edit.undo");
eq("and back when undone", show(wb.page("Sheet1")!.sheet.value("J1")), "#CYCLE!");
// expect: ok numbers below text below FALSE below TRUE; text without case
// expect: ok an empty cell is the other side's empty value
// expect: ok in formulas: text is above every number, TRUE above both, TRUE is not 1, an empty cell is 0 and empty text
// expect: ok Sort A to Z: numbers, text without case, TRUE, errors, then the empty cell
// expect: ok Sort Z to A: reversed, but the empty cell still last
// expect: ok SORT the same, both ways
// expect: ok AND, OR and NOT read the text TRUE and FALSE (any case); other text is skipped, and with nothing left, #VALUE!
// expect: ok a cycle through two sheets: #CYCLE! on both, saying the path
// expect: ok what reads it is #CYCLE! too
// expect: ok even where the values would settle (each only copies the other)
// expect: ok the cycle broken: worked out again
// expect: ok and back when undone
