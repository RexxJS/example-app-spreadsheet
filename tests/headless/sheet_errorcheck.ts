// Error Checking (2.16): the findings for a sample sheet, rule by rule
// (errorcheck.ts); Ignore Error passes over a cell, undoes, and is saved
// with the file; ERRORCHECK() and IGNOREERROR() from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.execute("cell.set", { cell, input });
type Found = { cell: string; check: string; message: string };
const check = async () => ((await cmds.execute("formulas.errorCheck")).output as Found[]).map((f) => `${f.cell} ${f.check}: ${f.message}`);

await cmds.run("formulas.use language=Excel");
eq("an empty sheet has nothing to check", (await cmds.execute("formulas.errorCheck")).message, "no errors found");
for (const [k, v] of [["A1", "1"], ["A2", "2"], ["A3", "3"], ["A4", "4"], ["A5", "5"], ["B1", "=A1*2"], ["B2", "=A2*2"], ["B3", "=A3+5"], ["B4", "=A4*2"],
  ["C1", "=1/0"], ["C2", "=C1+1"], ["D1", "'42"], ["D2", "1,234"], ["D3", "'abc"], ["D4", "12"], ["A6", "=SUM(A1:A4)"], ["E1", "=SUM(A1:A5)"]]) await set(k, v);
eq("an apostrophe keeps a number text, shown without it", [wb.sheet.value("D1"), wb.sheet.input("D1"), show(wb.sheet.value("D3"))], ["42", "'42", "abc"]);
eq("the findings, in sheet order", await check(), [
  "C1 error: #DIV/0! divide by zero",
  "D1 text-number: a number stored as text",
  "C2 error: #DIV/0! divide by zero",
  "D2 text-number: a number stored as text",
  "B3 inconsistent: the formula differs from the formulas beside it in its column",
  "A6 omits: the formula omits A5, next to A1:A4",
]);
eq("Ignore Error passes over a cell", [(await cmds.execute("formulas.ignoreError", { cell: "B3" })).message, (await check()).some((f) => f.startsWith("B3"))], ["B3: error ignored", false]);
eq("saved with the file", ((writeSheetFile(wb) as { cells: Record<string, { ignoreError?: boolean }> }).cells.B3).ignoreError, true);
await cmds.run("edit.undo");
eq("undone, it is checked again", (await check()).some((f) => f.startsWith("B3")), true);
await cmds.execute("formulas.ignoreError", { cell: "A6" });
await cmds.execute("formulas.ignoreError", { cell: "A6", off: "1" });
eq("and off", (await check()).some((f) => f.startsWith("A6")), true);
await set("B3", "=A3*2");
await set("A6", "=SUM(A1:A5)");
await set("C1", "=1/2");
eq("each mended: only the numbers as text are left", await check(), ["D1 text-number: a number stored as text", "D2 text-number: a number stored as text"]);
await set("B1", "=A1*3");
eq("a row's formula unlike those on both sides in its row", [await set("F2", "=B4+1"), await set("G2", "=C2*9"), await set("H2", "=D4+1")].length && (await check()).filter((f) => f.startsWith("G2")),
  ["G2 inconsistent: the formula differs from the formulas beside it in its row"]);

const said: string[] = [];
const script = async (text: string) => { said.length = 0; return (await cmds.script(text, (line) => said.push(line))).message; };
eq("ERRORCHECK() and IGNOREERROR() from REXX", [await script(['e = ERRORCHECK()', 'SAY e.0', 'SAY WORD(e.1, 1) WORD(e.1, 2)', 'CALL IGNOREERROR "D1"', 'e = ERRORCHECK()', 'SAY e.0 e.1'].join("\n")), [...said]],
  ["ran", ["3", "D1 text-number", "2 D2 text-number a number stored as text"]]);
await cmds.run("formulas.use language=RexxJS");
await set("J1", "=1/0");
eq("RexxJS formulas: errors and text numbers (no R1C1 to compare)", (await check()).map((f) => f.split(":")[0]).filter((f) => f.startsWith("J1")), ["J1 error"]);

// expect: ok an empty sheet has nothing to check
// expect: ok an apostrophe keeps a number text, shown without it
// expect: ok the findings, in sheet order
// expect: ok Ignore Error passes over a cell
// expect: ok saved with the file
// expect: ok undone, it is checked again
// expect: ok and off
// expect: ok each mended: only the numbers as text are left
// expect: ok a row's formula unlike those on both sides in its row
// expect: ok ERRORCHECK() and IGNOREERROR() from REXX
// expect: ok RexxJS formulas: errors and text numbers (no R1C1 to compare)
