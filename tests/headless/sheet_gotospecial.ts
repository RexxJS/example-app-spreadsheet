// Go To Special (1.3): each category's cells on a sample sheet, the used
// range by default, a type for formulas and constants, refusals, and
// GOTOSPECIAL from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.execute("cell.set", { cell, input });
await cmds.run("formulas.use language=Excel");
for (const [k, v] of [["A1", "Item"], ["B1", "10"], ["A2", "=B1*2"], ["B2", "=1/0"], ["A3", "TRUE"], ["C3", '="x"']]) await set(k, v);
await cmds.run("cell.comment cell=B1 text=note");
await cmds.run("validation.set range=A3 kind=list values=TRUE,FALSE");
await cmds.run("format.highlight range=B1:B2 op=greater value=5");
const go = async (p: Record<string, string>) => (await cmds.execute("edit.goToSpecial", p)).output as string[];
eq("each category over the used range (A1:C3)", [await go({ what: "formulas" }), await go({ what: "constants" }), await go({ what: "blanks" }), await go({ what: "errors" }),
  await go({ what: "comments" }), await go({ what: "validation" }), await go({ what: "conditional" })],
  [["A2", "B2", "C3"], ["A1", "B1", "A3"], ["C1", "C2", "B3"], ["B2"], ["B1"], ["A3"], ["B1", "B2"]]);
eq("formulas and constants by type", [await go({ what: "formulas", type: "numbers" }), await go({ what: "formulas", type: "errors" }), await go({ what: "formulas", type: "text" }), await go({ what: "constants", type: "text" })],
  [["A2"], ["B2"], ["C3"], ["A1", "A3"]]);
eq("over a range given", await go({ what: "blanks", range: "C1:D2" }), ["C1", "D1", "C2", "D2"]);
eq("the message, and refusals", [(await cmds.execute("edit.goToSpecial", { what: "errors" })).message, (await cmds.execute("edit.goToSpecial", { what: "objects" })).message, (await cmds.execute("edit.goToSpecial", { what: "formulas", type: "dates" })).message],
  ["1 errors: B2", "objects is not one of formulas, constants, blanks, errors, comments, validation, conditional", "dates is not a type (numbers, text, logicals, errors)"]);
const said: string[] = [];
eq("GOTOSPECIAL from REXX: a stem of cells", [(await cmds.script('s = GOTOSPECIAL("blanks")\nSAY s.0 s.1 s.3\nt = GOTOSPECIAL("constants", "A1:B3", "numbers")\nSAY t.0 t.1', (l) => said.push(l))).message, said], ["ran", ["3 C1 B3", "1 B1"]]);

// expect: ok each category over the used range (A1:C3)
// expect: ok formulas and constants by type
// expect: ok over a range given
// expect: ok the message, and refusals
// expect: ok GOTOSPECIAL from REXX: a stem of cells
