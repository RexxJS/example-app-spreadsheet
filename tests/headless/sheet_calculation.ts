// Calculation options (2.15): manual calculation, where a change works out
// only the cells set and the rest wait, stale, for Calculate Now (every
// sheet) or Calculate Sheet (the active one); and iterative calculation,
// where a circular reference is worked round, up to the most iterations or
// until nothing moves by more than the maximum change, instead of #CYCLE!.
// Both are the workbook's settings, saved with the file, and REXX's.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.execute("cell.set", { cell, input });
const shown = (...keys: string[]) => keys.map((k) => show(wb.sheet.value(k)));
const near = (k: string, want: number, within: number) => Math.abs((wb.sheet.value(k) as number) - want) <= within;
const message = async (id: string, p: Record<string, string> = {}) => (await cmds.execute(id, p)).message;

await cmds.run("formulas.use language=Excel");
await set("B1", "=A1");
await set("A1", "=B1*0.5+10");
eq("without iteration a circular reference is #CYCLE!", shown("A1", "B1"), ["#CYCLE!", "#CYCLE!"]);
eq("iteration on, Excel's defaults", await message("formulas.iteration"), "iteration: at most 100 times, to a change of 0.001");
eq("the cycle converges on its fixed point, 20, within the maximum change", [near("A1", 20, 0.002), near("B1", 20, 0.002)], [true, true]);
await set("C1", "=A1+1");
eq("what reads the cycle reads its value", near("C1", 21, 0.002), true);

await cmds.execute("formulas.iteration", { max: "1" });
await set("E1", "=D1");
await set("D1", "=E1*0.5+10");
const once = shown("D1", "E1");
await cmds.execute("formulas.recalculate");
eq("at most one iteration: each calculation goes one further, from where the cells are", [once, shown("D1", "E1")], [["10", "10"], ["15", "15"]]);
await cmds.execute("formulas.iteration", { max: "100", change: "0" });
await set("G1", "=H1+1");
await set("H1", "=G1");
eq("one that never settles stops at the most iterations, showing where it got to", shown("G1", "H1"), ["100", "100"]);
eq("settings refused", [await message("formulas.iteration", { max: "0" }), await message("formulas.iteration", { change: "-1" }), await message("formulas.calculation", { mode: "sometimes" })],
  ["0 is not a number of iterations (1 to 32767)", "-1 is not a maximum change (0 or more)", "sometimes is not a calculation mode (automatic, manual)"]);
await cmds.execute("formulas.iteration", { off: "1" });
eq("iteration off: #CYCLE! again", shown("A1", "D1", "G1"), ["#CYCLE!", "#CYCLE!", "#CYCLE!"]);
await cmds.run("range.clear range=A1:H1");

// A cycle through two sheets, under iteration.
await cmds.execute("formulas.iteration", {});
await cmds.run("sheet.add name=Costs");
await set("A1", "=Sheet1!J1");
await cmds.run("sheet.activate name=Sheet1");
await set("J1", "=Costs!A1*0.5+10");
eq("through two sheets it converges too", [near("J1", 20, 0.002), Math.abs((wb.page("Costs")!.sheet.value("A1") as number) - 20) <= 0.002], [true, true]);
await cmds.execute("formulas.iteration", { off: "1" });
eq("and is #CYCLE! without", show(wb.sheet.value("J1")), "#CYCLE!");
await set("J1", "");

// Manual calculation.
await set("A2", "1");
await set("B2", "=A2*2");
await set("B3", "=Costs!A1+A2");
eq("manual", await message("formulas.calculation", { mode: "manual" }), "manual calculation");
await set("A2", "5");
eq("a change works out only the cell set: what reads it is stale", [shown("A2", "B2", "B3"), wb.stale()], [["5", "2", "1"], ["Sheet1!B2", "Sheet1!B3"]]);
await set("C2", "=A2+1");
eq("a formula typed is worked out", shown("C2"), ["6"]);
await set("D2", '=INFO("recalc")');
eq("INFO(\"recalc\") says Manual", shown("D2"), ["Manual"]);
await cmds.run("sheet.activate name=Costs");
await set("A1", "100");
await cmds.run("sheet.activate name=Sheet1");
eq("a change on another sheet leaves what reads it stale", [shown("B3"), wb.stale()], [["1"], ["Sheet1!B2", "Sheet1!B3"]]);
await cmds.run("sheet.activate name=Costs");
eq("Calculate Sheet works out only the active sheet", [await message("formulas.calculateSheet"), wb.page("Sheet1")!.sheet.stale()], ["Costs: 0 formulas worked out", ["B2", "B3"]]);
await cmds.run("sheet.activate name=Sheet1");
await cmds.execute("formulas.calculateSheet");
eq("then this one", [shown("B2", "B3"), wb.stale()], [["10", "105"], []]);
await set("A2", "7");
eq("Calculate Now works out every sheet", [await message("formulas.recalculate"), shown("B2", "B3"), wb.stale()], ["4 formulas worked out", ["14", "107"], []]);
await set("A2", "8");
const saved = writeSheetFile(wb) as { calculation?: unknown };
eq("saved with the file", saved.calculation, { mode: "manual" });
await cmds.execute("formulas.calculation", { mode: "automatic" });
eq("automatic again: everything worked out", [shown("B2", "B3", "D2"), wb.stale()], [["16", "108", "Automatic"], []]);

// From a file: iterating as it opens, manual after.
const file = { name: "loop", formulas: "excel", cells: { A1: "=B1*0.5+10", B1: "=A1", C1: "=A1" }, calculation: { mode: "manual", iterate: true, maxIterations: 50, maxChange: 0.0001 } };
await cmds.execute("file.import", { text: JSON.stringify(file) });
const st = (await cmds.execute("sheet.state")).output as { calculation: string; iteration: unknown };
eq("a file's settings: worked out iterating, manual after", [near("A1", 20, 0.0002), st.calculation, st.iteration], [true, "Manual", { max: 50, change: 0.0001 }]);
eq("written back as read", (writeSheetFile(wb) as { calculation?: unknown }).calculation, { mode: "manual", iterate: true, maxIterations: 50, maxChange: 0.0001 });
await cmds.execute("file.new");
eq("a new workbook calculates automatically", ((await cmds.execute("sheet.state")).output as { calculation: string }).calculation, "Automatic");

// From REXX.
const said: string[] = [];
const script = async (text: string) => { said.length = 0; return (await cmds.script(text, (line) => said.push(line))).message; };
await cmds.run("formulas.use language=Excel");
await set("A1", "2");
await set("B1", "=A1*3");
eq("SETCALCULATION, STALECELLS, CALCULATE, SETITERATION and GETITERATION from REXX", [await script([
  'SAY GETCALCULATION()', 'SAY SETCALCULATION("manual")', 'CALL SETCELL "A1", 4', 's = STALECELLS()', 'SAY s.0 s.1 GETCELL("B1")',
  'SAY CALCULATE() GETCELL("B1")', 'SAY CALCULATESHEET()', 'SAY GETITERATION()', 'SAY SETITERATION(20, 0.01)', 'SAY GETITERATION()', 'SAY SETITERATION(0)', 'SAY SETCALCULATION("Automatic")',
].join("\n")), [...said]], ["ran", ["Automatic", "Manual", "1 Sheet1!B1 6", "1 12", "1", "0", "20", "20 0.01", "0", "Automatic"]]);

// expect: ok without iteration a circular reference is #CYCLE!
// expect: ok iteration on, Excel's defaults
// expect: ok the cycle converges on its fixed point, 20, within the maximum change
// expect: ok what reads the cycle reads its value
// expect: ok at most one iteration: each calculation goes one further, from where the cells are
// expect: ok one that never settles stops at the most iterations, showing where it got to
// expect: ok settings refused
// expect: ok iteration off: #CYCLE! again
// expect: ok through two sheets it converges too
// expect: ok and is #CYCLE! without
// expect: ok manual
// expect: ok a change works out only the cell set: what reads it is stale
// expect: ok a formula typed is worked out
// expect: ok INFO("recalc") says Manual
// expect: ok a change on another sheet leaves what reads it stale
// expect: ok Calculate Sheet works out only the active sheet
// expect: ok then this one
// expect: ok Calculate Now works out every sheet
// expect: ok saved with the file
// expect: ok automatic again: everything worked out
// expect: ok a file's settings: worked out iterating, manual after
// expect: ok written back as read
// expect: ok a new workbook calculates automatically
// expect: ok SETCALCULATION, STALECELLS, CALCULATE, SETITERATION and GETITERATION from REXX
