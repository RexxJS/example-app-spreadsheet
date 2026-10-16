// Flash Fill (3.4): the examples the reference names, split names and
// case, and initials and joins besides; refusals; one step of undo;
// FLASHFILL from REXX.
import { infer, run } from "./sheet/flashfill.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const fill = (examples: [string[], string][], rest: string[][]) => {
  const p = infer(examples);
  return p ? rest.map((s) => run(p, s)) : null;
};
eq("first names split out", fill([[["John Smith"], "John"]], [["Mary Jones"], ["Ann Lee"]]), ["Mary", "Ann"]);
eq("last names, with a middle name: counted from the end", fill([[["John Smith"], "Smith"], [["Mary Ann Jones"], "Jones"]], [["Bo Li"], ["Sam T Ray"]]), ["Li", "Ray"]);
eq("case: proper and upper", [fill([[["jOHN sMITH"], "John Smith"]], [["mary jones"]]), fill([[["ab-12"], "AB12"]], [["cd-34"]])], [["Mary Jones"], ["CD34"]]);
eq("a join with a literal, and initials", [fill([[["John Smith"], "Smith, John"]], [["Mary Jones"]]), fill([[["John Smith"], "J.S."]], [["Mary Jones"]])], [["Jones, Mary"], ["M.J."]]);
eq("two columns joined", fill([[["John", "Smith"], "john.smith@example.com"]], [["Mary", "Jones"]]), ["mary.jones@example.com"]);
eq("examples no one program gives: none", infer([[["John Smith"], "John"], [["Mary Jones"], "X"]]), null);

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.execute("cell.set", { cell, input });
for (const [k, v] of [["A1", "ada lovelace"], ["A2", "alan turing"], ["A3", "grace hopper"], ["A5", "edsger dijkstra"], ["B1", "Lovelace, A."]]) await set(k, v);
eq("the command: the empty cells filled, a row with no source left", [(await cmds.execute("data.flashFill", { range: "B1:B5" })).message, ["B2", "B3", "B4", "B5"].map((k) => show(wb.sheet.value(k)))],
  ["flash filled B1:B5 from 1 example: 3 cells set, 3 changed", ["Turing, A.", "Hopper, G.", "", "Dijkstra, E."]]);
await cmds.run("edit.undo");
eq("one step of undo", show(wb.sheet.value("B2")), "");
eq("refused", [(await cmds.execute("data.flashFill", { range: "C1:C3" })).message, (await cmds.execute("data.flashFill", { range: "B1:C3" })).message, (await set("D1", "x"), (await cmds.execute("data.flashFill", { range: "D1:D3" })).message)],
  ["type an example or two in C1:C3 first", "Flash Fill fills one column", "nothing to the left of D to fill from"]);
const said: string[] = [];
eq("FLASHFILL from REXX", [(await cmds.script('SAY FLASHFILL("B1:B3")', (l) => said.push(l))).message, said, show(wb.sheet.value("B3"))], ["ran", ["flash filled B1:B3 from 1 example: 2 cells set, 2 changed"], "Hopper, G."]);

// expect: ok first names split out
// expect: ok last names, with a middle name: counted from the end
// expect: ok case: proper and upper
// expect: ok a join with a literal, and initials
// expect: ok two columns joined
// expect: ok examples no one program gives: none
// expect: ok the command: the empty cells filled, a row with no source left
// expect: ok one step of undo
// expect: ok refused
// expect: ok FLASHFILL from REXX
