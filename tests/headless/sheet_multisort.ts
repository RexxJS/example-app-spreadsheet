// A multilevel sort and Advanced Filter (4.10): keys each a column and an
// order, the first deciding first, a colour on top (Sort On Cell Color or
// Font Color), a header row kept; Advanced Filter by a criteria range (a
// row's conditions all hold, any row will do), in place or copied with
// unique records only; undone; SORTRANGE's keys and ADVANCEDFILTER from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
const col = (c: string, n: number) => Array.from({ length: n }, (_, i) => show(wb.sheet.value(`${c}${i + 1}`))).join(" ");
const rows = [["Region", "Rep", "Sales"], ["north", "ann", "10"], ["south", "bob", "30"], ["north", "cy", "30"], ["south", "di", "20"], ["north", "ed", "20"]];
for (const [i, r] of rows.entries()) for (const [j, x] of r.entries()) await cmds.run(`cell.set cell=${"ABC"[j]}${i + 1} input=${x}`);
eq("by Region ascending, then Sales descending, the header kept", [await message("data.sort range=A1:C6 keys=\"A asc, C desc\" headers=1"), col("B", 6)],
  ["sorted A1:C6 by A ascending, then C descending: 18 cells set, 18 changed", "Rep cy ed ann bob di"]);
await cmds.run("format.fill range=B5 color=yellow");
eq("a cell colour on top, then the rest by name", [await message("data.sort range=A2:C6 keys=\"B fill=#ffff00, B asc\""), col("B", 6), wb.notes.get("B2").look?.fill],
  ["sorted A2:C6 by B fill #ffff00 on top, then B ascending: 15 cells set, 15 changed", "Rep bob ann cy di ed", "#ffff00"]);
await cmds.run("edit.undo");
eq("undone", col("B", 6), "Rep cy ed ann bob di");
eq("refusals", [await message("data.sort range=A1:C6 keys=\"D asc\""), await message("data.sort range=A1:C6 keys=\"A upward\""), await message("data.sort range=A1:C6 keys=\"A fill=nope\"")],
  ["D is not a column of A1:C6", "A upward is not a sort key (a column, then asc, desc, fill=colour or font=colour)", "nope is not a colour"]);
for (const [k, x] of [["E1", "Region"], ["F1", "Sales"], ["E2", "north"], ["F2", ">15"], ["F3", ">25"]]) await cmds.run(`cell.set cell=${k} input=${x}`);
eq("in place: north over 15, or anything over 25", [await message("data.advancedFilter range=A1:C6 criteria=E1:F3"), wb.hidden().rows],
  ["filtered A1:C6 by E1:F3: 3 records shown, 2 hidden", [3, 5]]);
await cmds.run("data.clearFilter");
for (const [k, x] of [["A7", "north"], ["B7", "cy"], ["C7", "30"]]) await cmds.run(`cell.set cell=${k} input=${x}`);
eq("copied, each record once", [await message("data.advancedFilter range=A1:C7 criteria=E1:F3 copyTo=H1 unique=1"), col("I", 5), col("J", 5)],
  ["copied 3 records of A1:C7 to H1: 12 cells set, 12 changed", "Rep cy ed bob ", "Sales 30 20 30 "]);
await cmds.run("cell.set cell=G1 input=Cost");
await cmds.run("cell.set cell=G2 input=>1");
eq("a criteria field that is not a field", await message("data.advancedFilter range=A1:C6 criteria=F1:G2"), "the criteria's Cost is not a field of the database");
const said: string[] = [];
await cmds.script('k.0 = 2; k.1 = "C desc"; k.2 = "B asc"\nSAY SORTRANGE("A1:C7", k, , 1)\nSAY ADVANCEDFILTER("A1:C7", "E1:F2")', (l) => said.push(l));
eq("SORTRANGE with a stem of keys, ADVANCEDFILTER from REXX", [said, col("B", 7)], [["OK", "filtered A1:C7 by E1:F2: 3 records shown, 3 hidden"], "Rep bob cy cy di ed ann"]);

// expect: ok by Region ascending, then Sales descending, the header kept
// expect: ok a cell colour on top, then the rest by name
// expect: ok undone
// expect: ok refusals
// expect: ok in place: north over 15, or anything over 25
// expect: ok copied, each record once
// expect: ok a criteria field that is not a field
// expect: ok SORTRANGE with a stem of keys, ADVANCEDFILTER from REXX
