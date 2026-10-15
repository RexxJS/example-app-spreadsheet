// app/sheet/workbook.ts's undo: the whole workbook saved before each
// command that changes it, and put back. Looks, notes, rules, charts,
// merges, hidden rows, names, tables and sheets added, renamed or deleted
// are undone as cells are, across sheets, one step a command; what is only
// looked at or copied is no step, and a file opened is where undo stops.
// Each step says what it was ("Merge Cells"); a script is one step.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const value = (k: string) => show(wb.sheet.value(k));
const undo = async () => (await cmds.run("edit.undo")).message;
const redo = async () => (await cmds.run("edit.redo")).message;

eq("nothing to undo or redo at first", [wb.canUndo, wb.canRedo, await undo()], [false, false, "nothing to undo"]);
await cmds.run("cell.set cell=A1 input=1");
await cmds.run("cell.set cell=A2 input=2");
await cmds.run('cell.set cell=A3 input="=A1+A2"');
eq("cells, as before", [await undo(), wb.sheet.input("A3"), await redo(), value("A3")], ["undid Set a cell: 1 changed", "", "redid Set a cell: 1 changed", "3"]);

// Each of these was the sheet's own to keep, and not undone.
const look = () => [wb.notes.get("A1").look ?? null, wb.notes.get("A1").comment ?? "", wb.merges(), wb.hidden().rows, wb.charts().length,
  wb.page(wb.active)!.rules.length, wb.sheet.names().length, wb.sheet.tables().length];
const plain = look();
for (const line of ["format.font range=A1 bold=1", "cell.comment cell=A1 text=first", "cells.merge range=B1:C1", "rows.hide range=A5",
  "chart.insert kind=bar range=A1:A3", 'format.highlight range=A1:A3 op=">" value=1', "names.define name=TOP range=A1", "table.define name=T range=A1:A3"]) {
  const o = await cmds.run(line);
  if (!o.ok) throw new Error(`${line}: ${o.message}`);
}
const dressed = look();
for (let i = 0; i < 8; i++) await undo();
eq("a look, a comment, a merge, a hidden row, a chart, a rule, a name and a table: each one undo", look(), plain);
for (let i = 0; i < 8; i++) await redo();
eq("and each one redo", look(), dressed);

// Sheets: added, renamed and deleted, and what reads them.
await cmds.run("sheet.add name=Costs");
await cmds.run("cell.set cell=B1 input=40");
await cmds.run("sheet.activate name=Sheet1");
await cmds.run('cell.set cell=D1 input="=Costs!B1+1"');
eq("a formula reads another sheet", value("D1"), "41");
await cmds.run("sheet.rename name=Costs to=Prices");
eq("renamed, and what reads it follows", [wb.sheetNames(), wb.sheet.input("D1")], [["Sheet1", "Prices"], "=Prices!B1+1"]);
await undo();
eq("undone: the old name, and the formula as it was", [wb.sheetNames(), wb.sheet.input("D1"), value("D1")], [["Sheet1", "Costs"], "=Costs!B1+1", "41"]);
await cmds.run("sheet.delete name=Costs");
eq("deleted, and what read it is #REF!", [wb.sheetNames(), value("D1")], [["Sheet1"], "#REF!"]);
await undo();
eq("undone: the sheet back with its cells, and read again", [wb.sheetNames(), wb.page("Costs")!.sheet.input("B1"), value("D1")], [["Sheet1", "Costs"], "40", "41"]);

// Undo goes to the sheet the change was on.
await cmds.run("sheet.activate name=Costs");
await cmds.run("cell.set cell=B1 input=50");
await cmds.run("sheet.activate name=Sheet1");
eq("a change on another sheet reads through", value("D1"), "51");
await undo();
eq("undone from Sheet1: Costs is active, and B1 is as it was", [wb.active, wb.sheet.input("B1"), wb.page("Sheet1")!.sheet.value("D1")], ["Costs", "40", 41]);

// Not steps: a copy, going to a sheet, Circle Invalid, and refusals.
const steps = async () => {
  let n = 0;
  for (; wb.canUndo; n++) await cmds.run("edit.undo");
  for (let i = 0; i < n; i++) await cmds.run("edit.redo");
  return n;
};
const before = await steps();
await cmds.run("edit.copy range=A1:A3");
await cmds.run("sheet.activate name=Sheet1");
await cmds.run("validation.circle");
await cmds.run("cell.set cell=AA99 input=1");
eq("a copy, going to a sheet, circles and a refusal add no step", await steps(), before);

await cmds.run("cell.set cell=E1 input=7");
await undo();
await cmds.run("cell.set cell=E2 input=8");
eq("a new change forgets what could be redone", wb.canRedo, false);

for (let i = 0; i < 105; i++) await cmds.run(`cell.set cell=F1 input=${i}`);
eq("a hundred steps are kept", await steps(), 100);

// What Undo and Redo would do, by name.
await cmds.run("cells.merge range=H1:I1");
eq("each step says what it was", [wb.undoLabel, wb.redoLabel, await undo(), wb.redoLabel, await redo()],
  ["Merge Cells", null, "undid Merge Cells: 0 changed", "Merge Cells", "redid Merge Cells: 0 changed"]);

// A script is one step, however many commands it runs.
const said: string[] = [];
const depth = async () => { let n = 0; while (wb.canUndo) { await cmds.run("edit.undo"); n++; } return n; };
await cmds.run("file.new");
await cmds.run("cell.set cell=A1 input=1");
await cmds.script(`LET r = SETCELL("A2", 2); LET r = SETCELL("A3", 3); LET r = INSERTROW(1)`, (l) => said.push(l));
eq("a script that changed three things is one step, named Script", [wb.undoLabel, value("A4"), await undo(), value("A1"), wb.sheet.input("A3"), wb.undoLabel],
  ["Script", "3", "undid Script: 4 changed", "1", "", "Set a cell"]);
await redo();
eq("and redone as one", [value("A4"), wb.undoLabel], ["3", "Script"]);
await cmds.script(`LET r = SETCELL("B1", 5)`, (l) => said.push(l));
eq("one change in a script is named for its command", wb.undoLabel, "Set a cell");
await cmds.script(`SAY GETCELL("B1")`, (l) => said.push(l));
eq("a script that only reads is no step", wb.undoLabel, "Set a cell");
await cmds.script(`LET r = SETCELL("C1", 7); LET r = SETCELL("C2", 8); LET r = UNDO(); SAY GETCELL("C2") GETCELL("C1")`, (l) => said.push(l));
eq("an UNDO in a script takes back its last command; the script is still one step", [said.at(-1), wb.undoLabel, await undo(), value("C1")],
  [" 7", "Script", "undid Script: 1 changed", ""]);
await cmds.spreadsheet(`SETCELL("D1", 1) + SETCELL("D2", 2)`, (l) => said.push(l));
eq("a control-bus command is one step too", [value("D2"), wb.undoLabel, await undo(), value("D1")], ["2", "Script", "undid Script: 2 changed", ""]);
const half = await cmds.spreadsheet(`SETCELL("D1", 1); SETCELL("D2", 2); BOGUS()`, (l) => said.push(l));
eq("one that fails part way is taken back whole, and leaves nothing to redo", [half.ok, value("D1"), value("D2"), wb.undoLabel, wb.canRedo, half.message.endsWith("(nothing it did was kept)")],
  [false, "", "", "Set a cell", false, true]);
const read = await cmds.spreadsheet(`GETCELL("B1"); BOGUS()`, (l) => said.push(l));
eq("one that fails having changed nothing takes nothing back", [read.ok, wb.undoLabel, value("B1"), read.message.includes("kept")], [false, "Set a cell", "5", false]);
const stopped = await cmds.spreadsheet(`BATCH_EXECUTE('[{"command": "SETCELLS", "args": ["F1:F2", [1, 2]]}, {"command": "SORTRANGE", "args": ["F1:F2", "?"]}]', 1)`, (l) => said.push(l));
eq("BATCH_EXECUTE(json, 1): the first failure is the request's, and the request is taken back whole",
  [stopped.ok, stopped.message.startsWith("command 2 (SORTRANGE): "), stopped.message.endsWith("(nothing it did was kept)"), value("F1"), wb.undoLabel],
  [false, true, true, "", "Set a cell"]);
const carried = await cmds.spreadsheet(`BATCH_EXECUTE('[{"command": "SETCELLS", "args": ["F1:F2", [1, 2]]}, {"command": "SORTRANGE", "args": ["F1:F2", "?"]}]')`, (l) => said.push(l));
eq("without it, as the app's: each said, the rest done, one step", [carried.ok, JSON.parse(String(carried.output)).map((r: any) => r.success), value("F2"), wb.undoLabel],
  [true, [true, false], "2", "Script"]);
await undo();
// sae's bus gave up on a request (bus.serve's stopped): "deadline" while it
// still runs, so its next command fails and it is taken back as a failed
// one is; "interrupted" when it will never come back, so it is taken back
// at once, to the workbook as it was before it (a command can be stopped
// between changing a cell and recording its step). On a workbook of their
// own: here the stopped request is not killed as an interrupt kills it, so
// the command it was in finishes, and its step would join this file's count.
{
  const wb2 = new Workbook();
  const cmds2 = new Commands(wb2);
  const v2 = (k: string) => show(wb2.sheet.value(k));
  const changed = async (k: string) => {
    for (let i = 0; i < 10000 && v2(k) === ""; i++) await Promise.resolve();
    return v2(k);
  };
  await cmds2.run("cell.set cell=A1 input=kept");
  const late = cmds2.spreadsheet(`SETCELL("H1", "x"); SETCELL("H2", "y")`, () => {});
  const lateFirst = await changed("H1");
  const lateWhat = await cmds2.stopped("deadline");
  const lateDone = await late;
  eq("a request past the bus's deadline: cancelled, its next command refused, all of it taken back",
    [lateFirst, lateWhat, lateDone.ok, lateDone.message, v2("H1"), v2("H2"), v2("A1"), wb2.canRedo],
    ["x", "cancelled", false, "the request was stopped (nothing it did was kept)", "", "", "kept", false]);
  const dead = cmds2.spreadsheet(`SETCELL("H3", "x"); SETCELL("H4", "y")`, () => {});
  const deadFirst = await changed("H3");
  const deadWhat = await cmds2.stopped("interrupted");
  eq("one the page's time limit stopped: taken back there and then, to the sheet before it",
    [deadFirst, deadWhat, v2("H3"), v2("H4"), v2("A1"), wb2.canRedo], ["x", "taken back (nothing it did was kept)", "", "", "kept", false]);
  await dead;
  await cmds2.spreadsheet(`SETCELL("H5", "1"); SETCELL("H6", "2")`, () => {});
  eq("and the next request is one step again", [(await cmds2.run("edit.undo")).message, v2("H5"), v2("H6")], ["undid Script: 2 changed", "", ""]);
  eq("stopped with no request in flight does nothing", await cmds2.stopped("interrupted"), "");
}
await cmds.spreadsheet(`SETCELL("E1", 1)`, (l) => said.push(l));
await cmds.spreadsheet(`SETCELL("E2", 2)`, (l) => said.push(l));
await cmds.spreadsheet(`UNDO()`, (l) => said.push(l));
await cmds.spreadsheet(`UNDO()`, (l) => said.push(l));
eq("UNDO() sent twice goes back two steps, not back and forth", [value("E1"), value("E2"), wb.undoLabel, wb.canRedo], ["", "", "Set a cell", true]);
await cmds.spreadsheet(`REDO()`, (l) => said.push(l));
eq("and REDO() sent does the first again", [value("E1"), value("E2")], ["1", ""]);
await undo();
eq("the steps: A1's, the first script's, B1's", await depth(), 3);
eq("and nothing then", [await undo(), wb.undoLabel], ["nothing to undo", null]);

await cmds.run("file.open sample=test");
eq("a file opened is where undo stops", [wb.canUndo, wb.canRedo], [false, false]);
await cmds.run("file.new");
eq("so is a new one", wb.canUndo, false);

// expect: ok nothing to undo or redo at first
// expect: ok cells, as before
// expect: ok a look, a comment, a merge, a hidden row, a chart, a rule, a name and a table: each one undo
// expect: ok and each one redo
// expect: ok a formula reads another sheet
// expect: ok renamed, and what reads it follows
// expect: ok undone: the old name, and the formula as it was
// expect: ok deleted, and what read it is #REF!
// expect: ok undone: the sheet back with its cells, and read again
// expect: ok a change on another sheet reads through
// expect: ok undone from Sheet1: Costs is active, and B1 is as it was
// expect: ok a copy, going to a sheet, circles and a refusal add no step
// expect: ok a new change forgets what could be redone
// expect: ok a hundred steps are kept
// expect: ok each step says what it was
// expect: ok a script that changed three things is one step, named Script
// expect: ok and redone as one
// expect: ok one change in a script is named for its command
// expect: ok a script that only reads is no step
// expect: ok an UNDO in a script takes back its last command; the script is still one step
// expect: ok a control-bus command is one step too
// expect: ok one that fails part way is taken back whole, and leaves nothing to redo
// expect: ok one that fails having changed nothing takes nothing back
// expect: ok BATCH_EXECUTE(json, 1): the first failure is the request's, and the request is taken back whole
// expect: ok without it, as the app's: each said, the rest done, one step
// expect: ok a request past the bus's deadline: cancelled, its next command refused, all of it taken back
// expect: ok one the page's time limit stopped: taken back there and then, to the sheet before it
// expect: ok and the next request is one step again
// expect: ok stopped with no request in flight does nothing
// expect: ok UNDO() sent twice goes back two steps, not back and forth
// expect: ok and REDO() sent does the first again
// expect: ok the steps: A1's, the first script's, B1's
// expect: ok and nothing then
// expect: ok a file opened is where undo stops
// expect: ok so is a new one
