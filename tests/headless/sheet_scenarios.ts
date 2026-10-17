// Scenario Manager (6.7): named sets of values for some cells, added,
// edited, shown (one step of undo), deleted; a summary sheet of each one's
// changing cells and results beside the current values, the sheet left as
// it was; saved with the file; from REXX.
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const store = memoryStore();
const wb = new Workbook();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("formulas.use language=Excel");
for (const [k, v] of [["B1", "100"], ["B2", "0.1"], ["B3", "=B1*(1+B2)"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
eq("added: the cells' own values, or given ones", [await message("scenario.add name=Now cells=B1:B2"), await message('scenario.add name=Best cells="B1,B2" values="200,0.5"'), await message('scenario.add name=Worst cells=B1:B2 values="50,0"')],
  ["scenario Now: B1=100, B2=0.1", "scenario Best: B1=200, B2=0.5", "scenario Worst: B1=50, B2=0"]);
eq("shown, and undone", [await message("scenario.show name=best"), show(wb.sheet.value("B3")), (await cmds.run("edit.undo")).ok, show(wb.sheet.value("B3"))], ["showed scenario Best: 2 cells set, 3 changed", "300", true, "110"]);
eq("a summary sheet, the sheet as it was", [await message("scenario.summary results=B3"), wb.active, show(wb.page("Sheet1")!.sheet.value("B3"))], ["scenario summary on Scenario Summary: 23 cells set, 23 changed", "Scenario Summary", "110"]);
const row = (r: number) => ["A", "B", "C", "D", "E", "F"].map((c) => show(wb.sheet.value(`${c}${r}`))).join("|");
eq("its rows", [row(2), row(4), row(5), row(7)], ["|Cell|Current Values|Now|Best|Worst", "|Sheet1!B1|100|100|200|50", "|Sheet1!B2|0.1|0.1|0.5|0", "|Sheet1!B3|110|110|300|50"]);
await cmds.run("sheet.activate name=Sheet1");
eq("edited, deleted, refused", [await message('scenario.add name=worst cells=B1 values="1"'), (await cmds.run("scenario.list")).output, await message("scenario.delete name=Now"), await message("scenario.show name=Nope"),
  await message('scenario.add name=Bad cells=B1:B2 values="1"')],
  ["scenario worst: B1=1", ["Now: B1=100, B2=0.1", "Best: B1=200, B2=0.5", "worst: B1=1"], "deleted scenario Now", "no scenario Nope", "2 cells and 1 values"]);
await cmds.run("file.save name=s");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=s");
eq("saved with the file", (await cmds2.run("scenario.list")).output, ["Best: B1=200, B2=0.5", "worst: B1=1"]);
const said: string[] = [];
await cmds.script('SAY ADDSCENARIO("Mid", "B1", "150")\nSAY SHOWSCENARIO("Mid")\nSAY GETCELL("B3")\ns = LISTSCENARIOS()\nSAY s.0\nSAY DELETESCENARIO("Mid")', (l) => said.push(l));
eq("from REXX", said, ["scenario Mid: B1=150", "showed scenario Mid: 1 cells set, 2 changed", "165", "3", "deleted scenario Mid"]);

// expect: ok added: the cells' own values, or given ones
// expect: ok shown, and undone
// expect: ok a summary sheet, the sheet as it was
// expect: ok its rows
// expect: ok edited, deleted, refused
// expect: ok saved with the file
// expect: ok from REXX
