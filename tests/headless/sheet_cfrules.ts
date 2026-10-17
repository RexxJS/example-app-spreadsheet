// Conditional rules past the first ones (4.4): a formula's rule, moved for
// each cell from the range's top-left; text that begins or ends with, or
// does not contain; blanks and no blanks, errors and no errors; A Date
// Occurring's periods; rule order (Move Up/Down) and Stop If True; saved
// with the file; ADDRULE and the rest from REXX.
import { applied, inPeriod } from "./sheet/rules.ts";
import type { CfRule } from "./sheet/rules.ts";
import { parseRange } from "./sheet/address.ts";
import { serial } from "./sheet/dates.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { CellError } from "./sheet/values.ts";
import type { Value } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const at = (s: string) => parseRange(s)!;
const RED = "background: #ffc7ce";
const cells: Record<string, Value> = { A1: "apple", A2: "pineapple", A3: "grape", A4: null, A5: new CellError("#DIV/0!", "") };
const marks = (rules: CfRule[]) => [...applied(rules, (k) => cells[k] ?? null)].map(([k, m]) => `${k} ${m.style ?? ""}`).sort();
eq("begins with, ends with, does not contain", [marks([{ kind: "cellValue", range: at("A1:A3"), op: "beginsWith", value: "AP", style: RED }]),
  marks([{ kind: "cellValue", range: at("A1:A3"), op: "endsWith", value: "apple", style: RED }]), marks([{ kind: "cellValue", range: at("A1:A4"), op: "notContains", value: "apple", style: RED }])],
  [[`A1 ${RED}`], [`A1 ${RED}`, `A2 ${RED}`], [`A3 ${RED}`, `A4 ${RED}`]]);
eq("blanks, no blanks, errors, no errors", [marks([{ kind: "blanks", range: at("A1:A5"), not: false, style: RED }]), marks([{ kind: "blanks", range: at("A3:A5"), not: true, style: RED }]),
  marks([{ kind: "errors", range: at("A1:A5"), not: false, style: RED }]), marks([{ kind: "errors", range: at("A3:A5"), not: true, style: RED }])],
  [[`A4 ${RED}`], [`A3 ${RED}`, `A5 ${RED}`], [`A5 ${RED}`], [`A3 ${RED}`, `A4 ${RED}`]]);
// Wednesday 17 January 2024.
const today = serial(2024, 1, 17);
const periods = (v: Value) => ["yesterday", "today", "tomorrow", "last7days", "lastWeek", "thisWeek", "nextWeek", "lastMonth", "thisMonth", "nextMonth"].filter((p) => inPeriod(v, p as never, today));
eq("A Date Occurring: weeks Sunday to Saturday, months by the calendar", [periods(today), periods(today - 1), periods(serial(2024, 1, 14)), periods(serial(2024, 1, 13)), periods(serial(2024, 1, 21)),
  periods(serial(2023, 12, 31)), periods(serial(2024, 2, 29) + 0.5), periods("2024-01-11"), periods("pears")],
  [["today", "last7days", "thisWeek", "thisMonth"], ["yesterday", "last7days", "thisWeek", "thisMonth"], ["last7days", "thisWeek", "thisMonth"], ["last7days", "lastWeek", "thisMonth"],
    ["nextWeek", "thisMonth"], ["lastMonth"], ["nextMonth"], ["last7days", "lastWeek", "thisMonth"], []]);
const BOLD = "font-weight: bold";
eq("the first rule wins; Stop If True leaves its cells to no later rule", [
  marks([{ kind: "blanks", range: at("A1:A5"), not: true, style: BOLD }, { kind: "cellValue", range: at("A1:A3"), op: "contains", value: "apple", style: RED }]),
  marks([{ kind: "blanks", range: at("A1:A2"), not: true, style: BOLD, stop: true }, { kind: "cellValue", range: at("A1:A3"), op: "contains", value: "apple", style: RED }])],
  [[`A1 ${RED}; ${BOLD}`, `A2 ${RED}; ${BOLD}`, `A3 ${BOLD}`, `A5 ${BOLD}`], [`A1 ${BOLD}`, `A2 ${BOLD}`]]);

const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
for (const [k, v] of [["A1", "5"], ["A2", "1"], ["A3", "7"], ["B1", "3"], ["B2", "4"], ["B3", "7"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
eq("a formula rule, written for A1 and moved to each cell", [await message('format.formulaRule range=A1:A3 formula="=A1>B1" style=bold'), wb.styleOf("A1").includes("bold"), wb.styleOf("A2").includes("bold"), wb.styleOf("A3").includes("bold")],
  ["A1:A3: formula =A1>B1", true, false, false]);
await cmds.run("cell.set cell=B2 input=0");
eq("and worked out again when what it reads changes", wb.styleOf("A2").includes("bold"), true);
eq("a broken formula is refused", await message('format.formulaRule range=A1 formula="=A1>"'), (await cmds.run('format.formulaRule range=A1 formula="=A1>"')).message);
eq("the commands", [await message("format.blanks range=C1:C3"), await message("format.errors range=C1:C3 not=1"), await message("format.datePeriod range=D1:D3 period=thisMonth"),
  await message("format.highlight range=E1:E3 op=beginsWith value=ab stop=1"), await message("format.datePeriod range=D1 period=someday")],
  ["C1:C3: blanks", "C1:C3: no errors", "D1:D3: a date this month", "E1:E3: text begins with ab (stop if true)", "someday is not a period (yesterday, today, tomorrow, last7days, lastWeek, thisWeek, nextWeek, lastMonth, thisMonth, nextMonth)"]);
eq("Move Up, and Stop If True on and off", [await message("format.moveRule index=5 to=1"), (await cmds.run("format.rules")).output, await message("format.stopIfTrue index=1 off=1"), await message("format.stopIfTrue index=2")],
  ["rule 5 is now rule 1: E1:E3: text begins with ab (stop if true)", ["1. E1:E3: text begins with ab (stop if true)", "2. A1:A3: formula =A1>B1", "3. C1:C3: blanks", "4. C1:C3: no errors", "5. D1:D3: a date this month"],
    "E1:E3: text begins with ab", "A1:A3: formula =A1>B1 (stop if true)"]);
await cmds.run("edit.undo");
eq("undone", (await cmds.run("format.rules")).output, ["1. E1:E3: text begins with ab", "2. A1:A3: formula =A1>B1", "3. C1:C3: blanks", "4. C1:C3: no errors", "5. D1:D3: a date this month"]);
await cmds.run("format.dataBar range=F1:F3");
await cmds.run("file.save name=cf");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=cf");
eq("saved with the file, data bars too", (await cmds2.run("format.rules")).output, ["1. E1:E3: text begins with ab", "2. A1:A3: formula =A1>B1", "3. C1:C3: blanks", "4. C1:C3: no errors", "5. D1:D3: a date this month", "6. F1:F3: data bars #638ec6"]);
eq("and the formula rule works there", wb2.styleOf("A2").includes("bold"), true);
const said: string[] = [];
await cmds.script('SAY ADDRULE("G1:G3", "formula", "=G1>10", "color:red")\nSAY ADDRULE("H1:H3", "between", 1, \"\", 5)\nSAY ADDRULE("H1:H3", "noBlanks")\nSAY MOVERULE(9, 1)\nSAY STOPIFTRUE(1)\nr = LISTRULES()\nSAY r.0 r.1\nSAY DELETERULE(1)\nSAY CLEARRULES("H1:H3")',
  (l) => said.push(l));
eq("ADDRULE, MOVERULE, STOPIFTRUE, LISTRULES, DELETERULE and CLEARRULES from REXX", said,
  ["G1:G3: formula =G1>10", "H1:H3: cell value between 1 and 5", "H1:H3: no blanks", "rule 9 is now rule 1: H1:H3: no blanks", "H1:H3: no blanks (stop if true)", "9 1. H1:H3: no blanks (stop if true)",
    "deleted H1:H3: no blanks (stop if true)", "cleared 1 rule"]);

// expect: ok begins with, ends with, does not contain
// expect: ok blanks, no blanks, errors, no errors
// expect: ok A Date Occurring: weeks Sunday to Saturday, months by the calendar
// expect: ok the first rule wins; Stop If True leaves its cells to no later rule
// expect: ok a formula rule, written for A1 and moved to each cell
// expect: ok and worked out again when what it reads changes
// expect: ok a broken formula is refused
// expect: ok the commands
// expect: ok Move Up, and Stop If True on and off
// expect: ok undone
// expect: ok saved with the file, data bars too
// expect: ok and the formula rule works there
// expect: ok ADDRULE, MOVERULE, STOPIFTRUE, LISTRULES, DELETERULE and CLEARRULES from REXX
