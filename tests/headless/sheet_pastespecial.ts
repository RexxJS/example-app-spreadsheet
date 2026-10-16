// Paste Special (3.1): values, formulas, formats, comments, all; an
// operation with what is there (add, subtract, multiply, divide); skip
// blanks; transpose; one step of undo; and from REXX (COPY, PASTE,
// PASTESPECIAL).
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.execute("cell.set", { cell, input });
const shown = (...ks: string[]) => ks.map((k) => show(wb.sheet.value(k)));
const ps = async (p: Record<string, string>) => (await cmds.execute("edit.pasteSpecial", p)).message;
await cmds.run("formulas.use language=Excel");
await set("A1", "2"); await set("A2", ""); await set("A3", "=A1*5");
await cmds.run("cell.comment cell=A1 text=two");
await cmds.run("format.font range=A1 bold=1");
await set("C1", "10"); await set("C2", "20"); await set("C3", "=7");
await cmds.run("edit.copy range=A1:A3");
eq("add: each number added to what is there; a formula there wrapped", [await ps({ at: "C1", what: "values", operation: "add" }), shown("C1", "C2", "C3"), wb.sheet.input("C3")],
  ["pasted values, add at C1: 3 cells set, 3 changed", ["12", "20", "17"], "=(7)+10"]);
await cmds.run("edit.undo");
eq("one step of undo", shown("C1", "C2", "C3"), ["10", "20", "7"]);
eq("multiply, skipping blanks: C2 left as it is", [await ps({ at: "C1", what: "values", operation: "multiply", skipBlanks: "1" }), shown("C1", "C2", "C3")],
  ["pasted values, multiply, skipping blanks at C1: 2 cells set, 2 changed", ["20", "20", "70"]]);
await cmds.run("edit.undo");
await set("E2", "keep");
eq("skip blanks: an empty copied cell leaves its target", [await ps({ at: "E1", what: "all", skipBlanks: "1" }), wb.sheet.input("E2"), wb.sheet.input("E3")], ["pasted all, skipping blanks at E1: 2 cells set, 2 changed", "keep", "=E1*5"]);
eq("comments alone", [await ps({ at: "G1", what: "comments" }), wb.notes.get("G1").comment, wb.sheet.input("G1"), wb.notes.get("G1").look], ["pasted comments at G1: 0 cells set, 0 changed", "two", "", undefined]);
eq("formats alone: the look, not the comment", [await ps({ at: "H1", what: "formats" }), wb.notes.get("H1").look, wb.notes.get("H1").comment], ["pasted formats at H1: 0 cells set, 0 changed", { bold: true }, undefined]);
eq("transposed: a column as a row", [await ps({ at: "A10", what: "values", transpose: "1" }), shown("A10", "B10", "C10")], ["pasted values, transposed at A10: 3 cells set, 3 changed", ["2", "", "10"]]);
eq("refused", [await ps({ at: "J1", what: "pictures" }), await ps({ at: "J1", what: "values", operation: "power" }), await ps({ at: "J1", what: "formats", operation: "add" })],
  ["pictures is not all, formulas, values, formats or comments", "power is not an operation (none, add, subtract, multiply, divide)", "an operation pastes values or formulas, not formats"]);
const said: string[] = [];
eq("COPY and PASTESPECIAL from REXX", [(await cmds.script('CALL COPY "C1:C2"\nSAY PASTESPECIAL("K1", "values", "subtract")\nCALL PASTE "L1"', (l) => said.push(l))).message, said, shown("K1", "L1")],
  ["ran", ["pasted values, subtract at K1: 2 cells set, 2 changed"], ["-10", "10"]]);

// expect: ok add: each number added to what is there; a formula there wrapped
// expect: ok one step of undo
// expect: ok multiply, skipping blanks: C2 left as it is
// expect: ok skip blanks: an empty copied cell leaves its target
// expect: ok comments alone
// expect: ok formats alone: the look, not the comment
// expect: ok transposed: a column as a row
// expect: ok refused
// expect: ok COPY and PASTESPECIAL from REXX
