// Help while a formula is typed (3.3): F4 cycles a reference's $s, and
// Formula AutoComplete gives the names completing what is typed and the
// call it is in; CYCLEREFERENCE and COMPLETEFORMULA from REXX.
import { complete, completed, cycleRef } from "./sheet/complete.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const four = ["=A1+B2"];
for (let i = 0; i < 4; i++) four.push(cycleRef(four[four.length - 1]).text);
eq("F4: A1, $A$1, A$1, $A1, A1 again, the reference at the end", four, ["=A1+B2", "=A1+$B$2", "=A1+B$2", "=A1+$B2", "=A1+B2"]);
eq("at a caret; a range's end; a string, a name and a call left alone", [cycleRef("=A1+B2", 3).text, cycleRef("=SUM(A1:C3)").text, cycleRef("=SUM(A1:C3)", 10).text, cycleRef('="A1"').text, cycleRef("=LOG10(5)").text, cycleRef("=TAX2").text],
  ["=$A$1+B2", "=SUM(A1:C3)", "=SUM(A1:$C$3)", '="A1"', "=LOG10(5)", "=$TAX$2"]);
const fns = ["SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SQRT", "IF", "IFERROR", "INDEX"];
eq("the names completing what is typed", [complete("=SU", fns).names, complete("=1+if", fns).names, complete("=A1", fns).names, complete("=", fns).names, complete('="su', fns).names, complete("su", fns).names],
  [["SUM", "SUMIF", "SUMIFS", "SUMPRODUCT"], ["IF", "IFERROR"], [], [], [], []]);
eq("the call the end is in, and which argument", [complete("=SUMIF(A1:A9, ", fns), complete("=IF(A1>0, SUM(B1, (2", fns), complete("=SUM(1)", fns)].map((c) => [c.call, c.arg]),
  [["SUMIF", 2], ["SUM", 2], [null, 0]]);
eq("a completion taken", completed("=1+SU", "SU", "SUMIF"), "=1+SUMIF(");
const cmds = new Commands(new Workbook());
await cmds.run("formulas.use language=Excel");
const said: string[] = [];
eq("from REXX", [(await cmds.script('SAY CYCLEREFERENCE("=B2")\nc = COMPLETEFORMULA("=COUNTIF(A1:A3, COU")\nSAY c.0 c.1 c.CALL c.ARG', (l) => said.push(l))).message, said],
  ["ran", ["=$B$2", "8 COUNT COUNTIF 2"]]);

// expect: ok F4: A1, $A$1, A$1, $A1, A1 again, the reference at the end
// expect: ok at a caret; a range's end; a string, a name and a call left alone
// expect: ok the names completing what is typed
// expect: ok the call the end is in, and which argument
// expect: ok a completion taken
// expect: ok from REXX
