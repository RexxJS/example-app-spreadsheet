// Slicers (4.12): a table column's values as tiles, picking some showing
// only the table's rows holding one (one or several, toggled), cleared,
// deleted; two slicers together; undone; saved with the file, its picks
// filtering again when opened; from REXX.
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
const rows = [["Region", "Rep"], ["north", "ann"], ["south", "bob"], ["North", "cy"], ["west", "di"], ["south", "ann"]];
for (const [i, r] of rows.entries()) for (const [j, x] of r.entries()) await cmds.run(`cell.set cell=${"AB"[j]}${i + 1} input=${x}`);
await cmds.run("table.define name=Sales range=A1:B6 headers=1");
eq("a slicer: the column's values, each once, any case", [await message("slicer.add table=Sales column=region"), wb.slicerViews()],
  ["slicer Region for Sales[Region]: north, south, west", [{ name: "Region", caption: "Region (Sales[Region])", tiles: [{ value: "north", picked: false }, { value: "south", picked: false }, { value: "west", picked: false }] }]]);
eq("one tile picked: only its rows", [await message("slicer.select name=Region values=north"), wb.hidden().rows], ["Region: north (3 rows hidden)", [2, 4, 5]]);
eq("another added, then toggled off", [await message("slicer.select name=Region values=west add=1"), wb.hidden().rows, await message("slicer.select name=Region values=north toggle=1"), wb.hidden().rows],
  ["Region: north, west (2 rows hidden)", [2, 5], "Region: west (4 rows hidden)", [1, 2, 3, 5]]);
await cmds.run("slicer.add table=Sales column=Rep");
await cmds.run("slicer.select name=Region values=south");
eq("two slicers together", [await message("slicer.select name=Rep values=ann"), wb.hidden().rows, (await cmds.run("slicer.list")).output],
  ["Rep: ann (4 rows hidden)", [1, 2, 3, 4], ["Region (Sales[Region]): north [south] west", "Rep (Sales[Rep]): [ann] bob cy di"]]);
eq("cleared", [await message("slicer.clear name=Rep"), wb.hidden().rows], ["Rep: all (3 rows hidden)", [1, 3, 4]]);
await cmds.run("file.save name=sl");
await cmds.run("edit.undo");
eq("undone", wb.hidden().rows, [1, 2, 3, 4]);
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=sl");
eq("saved with the file, filtering again when opened", [wb2.hidden().rows, wb2.slicerViews().map((x) => x.name)], [[1, 3, 4], ["Region", "Rep"]]);
eq("deleted: its rows shown", [await message("slicer.delete name=Region"), wb.hidden().rows], ["deleted slicer Region (3 rows hidden)", [2, 3, 4]]);
eq("refusals", [await message("slicer.add table=Nope column=x"), await message("slicer.add table=Sales column=Cost"), await message("slicer.add table=Sales column=Rep"), await message("slicer.select name=Rep values=zed")],
  ["no table Nope", "Sales has no column Cost", "there is a slicer Rep already", "zed is not a value of Rep (ann, bob, cy, di)"]);
const said: string[] = [];
await cmds.script('SAY ADDSLICER("Sales", "Region", "Where")\nSAY SLICE("Where", "west")\nSAY SLICE("Where", "north", 1)\ns = LISTSLICERS()\nSAY s.0\nSAY CLEARSLICER("Where")\nSAY DELETESLICER("Where")', (l) => said.push(l));
eq("from REXX", said, ["slicer Where for Sales[Region]: north, south, west", "Where: west (5 rows hidden)", "Where: west, north (4 rows hidden)", "2", "Where: all (3 rows hidden)", "deleted slicer Where (3 rows hidden)"]);

// expect: ok a slicer: the column's values, each once, any case
// expect: ok one tile picked: only its rows
// expect: ok another added, then toggled off
// expect: ok two slicers together
// expect: ok cleared
// expect: ok undone
// expect: ok saved with the file, filtering again when opened
// expect: ok deleted: its rows shown
// expect: ok refusals
// expect: ok from REXX
