// R1C1 as a display and entry mode (2.14): excel.ts writes a formula's
// references R1C1 from the cell it is in, and reads them back; the
// workbook's referenceStyle shows and takes Excel formulas so, stored A1;
// formulas.style and SETREFERENCESTYLE switch it.
import { fromR1C1, toR1C1 } from "./sheet/excel.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("A1 to R1C1: relative as an offset in brackets, absolute as a number, the cell's own row or column bare", [
  toR1C1("=SUM(A1:B2)*$C$3+Costs!D4", "E5"), toR1C1("=E5+E$1+$A5", "E5"), toR1C1("=A1#", "B2"), toR1C1('="A1"&A1', "A1"), toR1C1("42", "A1"), toR1C1("=SUM(Jan:Mar!B2)", "B2"), toR1C1("='My Costs'!$A$1", "C3"),
], ["=SUM(R[-4]C[-4]:R[-3]C[-3])*R3C3+Costs!R[-1]C[-1]", "=RC+R1C+RC1", "=R[-1]C[-1]#", '="A1"&RC', "42", "=SUM(Jan:Mar!RC)", "='My Costs'!R1C1"]);
eq("R1C1 to A1: the same formulas back, in any case; names and text as written; off the sheet #REF!", [
  fromR1C1("=SUM(R[-4]C[-4]:R[-3]C[-3])*R3C3+Costs!R[-1]C[-1]", "E5"), fromR1C1("=RC+R1C+RC1", "E5"), fromR1C1('="RC"&rc[1]', "B2"), fromR1C1("=ROUND(RC[-1], 2)", "B2"),
  fromR1C1("=R[-5]C", "B2"), fromR1C1("=SUM(Jan:Mar!RC)", "B2"), fromR1C1("=ARC+RCX", "B2"), fromR1C1("7", "B2"),
], ["=SUM(A1:B2)*$C$3+Costs!D4", "=E5+E$1+$A5", '="RC"&C2', "=ROUND(A2, 2)", "=#REF!", "=SUM(Jan:Mar!B2)", "=ARC+RCX", "7"]);
eq("round trips from anywhere", ["=A1+$B$2*C$3-$D4", "=SUM(Z1:Z100)", "=AVERAGE(A1:A3)&\"R1C1\""].map((f) => fromR1C1(toR1C1(f, "M50"), "M50")), ["=A1+$B$2*C$3-$D4", "=SUM(Z1:Z100)", "=AVERAGE(A1:A3)&\"R1C1\""]);

const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("cell.set cell=A1 input=10");
await cmds.run('cell.set cell=B2 input="=A1*2"');
eq("the setting: Excel formulas shown R1C1, stored A1", [await message("formulas.style style=R1C1"), wb.shownInput("B2"), wb.sheet.input("B2"), wb.shownInput("A1")], ["R1C1 references", "=R[-1]C[-1]*2", "=A1*2", "10"]);
const typed = wb.enteredInput("C3", "=R[-1]C[-1]+R1C1");
await cmds.run(`cell.set cell=C3 input="${typed}"`);
eq("typed R1C1, stored A1, worked out", [typed, wb.sheet.input("C3"), show(wb.sheet.value("C3"))], ["=B2+$A$1", "=B2+$A$1", "30"]);
eq("and back to A1; a style that is none", [await message("formulas.style style=a1"), wb.shownInput("B2"), wb.enteredInput("C3", "=RC"), await message("formulas.style style=XY")],
  ["A1 references", "=A1*2", "=RC", "XY is not a reference style (A1, R1C1)"]);
await cmds.run("formulas.style style=R1C1");
await cmds.run("formulas.use language=RexxJS");
await cmds.run('cell.set cell=D4 input="=A1 + 1"');
eq("RexxJS formulas shown and taken as written", [wb.shownInput("D4"), wb.enteredInput("D4", "=RC + 1")], ["=A1 + 1", "=RC + 1"]);
await cmds.run("formulas.use language=Excel");
eq("sheet.state says it", ((await cmds.run("sheet.state")).output as { referenceStyle: string }).referenceStyle, "R1C1");
const said: string[] = [];
const script = async (text: string) => { said.length = 0; return (await cmds.script(text, (line) => said.push(line))).message; };
eq("SETREFERENCESTYLE and GETREFERENCESTYLE from REXX", [await script(`SAY GETREFERENCESTYLE(); SAY SETREFERENCESTYLE("A1"); SAY GETREFERENCESTYLE()`), [...said], await script(`LET r = SETREFERENCESTYLE("Z9")`)],
  ["ran", ["R1C1", "A1", "A1"], "Z9 is not a reference style (A1, R1C1)"]);

// expect: ok A1 to R1C1: relative as an offset in brackets, absolute as a number, the cell's own row or column bare
// expect: ok R1C1 to A1: the same formulas back, in any case; names and text as written; off the sheet #REF!
// expect: ok round trips from anywhere
// expect: ok the setting: Excel formulas shown R1C1, stored A1
// expect: ok typed R1C1, stored A1, worked out
// expect: ok and back to A1; a style that is none
// expect: ok RexxJS formulas shown and taken as written
// expect: ok sheet.state says it
// expect: ok SETREFERENCESTYLE and GETREFERENCESTYLE from REXX
