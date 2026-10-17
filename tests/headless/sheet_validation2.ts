// Validation past numbers and lists (4.11): dates and times within limits
// (as written or as serials), a custom formula (true with the input in,
// moved from the rule's top-left), a pattern; Excel's error styles (stop,
// warn, info), an input prompt (the selected cell's status line and
// tooltip) and an error message of its own; saved with the file;
// SETCELLVALIDATION's new types and options from REXX.
import { check, ruleOf } from "./sheet/validation.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const date = ruleOf({ kind: "date", min: "2024-01-01", max: "2024-12-31" }), time = ruleOf({ kind: "time", min: "09:00", max: "17:30" });
eq("dates: as written or as serials", [check(date, "2024-06-01", true), check(date, "45658", true), check(date, "2025-01-02", true), check(date, "soon", true)],
  [null, "must be a date from 2024-01-01 to 2024-12-31", "must be a date from 2024-01-01 to 2024-12-31", "must be a date from 2024-01-01 to 2024-12-31"]);
eq("times", [check(time, "12:15", true), check(time, "0.5", true), check(time, "08:59", true), check(time, "6:00 PM", true)], [null, null, "must be a time from 09:00:00 to 17:30:00", "must be a time from 09:00:00 to 17:30:00"]);
const code = ruleOf({ kind: "pattern", pattern: "[A-Z]{2}-\\d{3}", error: "a code is two letters, a dash and three digits" });
eq("a pattern, and an error message of its own", [check(code, "AB-123", true), check(code, "AB-1234", true)], [null, "a code is two letters, a dash and three digits"]);
eq("refusals", [(() => { try { return ruleOf({ kind: "date", min: "someday" }); } catch (e) { return String((e as Error).message); } })(),
  (() => { try { return ruleOf({ kind: "pattern", pattern: "(" }); } catch (e) { return String((e as Error).message); } })(),
  (() => { try { return ruleOf({ kind: "list", values: "a", mode: "loud" }); } catch (e) { return String((e as Error).message); } })()],
  ["min someday is not a date", "( is not a pattern", "loud is not stop, warn or info"]);

const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("formulas.use language=Excel");
await cmds.run("cell.set cell=A1 input=10");
await cmds.run("cell.set cell=A2 input=20");
eq("custom: each cell's own formula, worked out with the input in", [await message('validation.set range=B1:B2 kind=custom formula="=B1<A1" prompt="under its A"'),
  await message("cell.set cell=B1 input=5"), await message("cell.set cell=B2 input=25"), show(wb.sheet.value("B2")), await message("cell.set cell=B2 input=15")],
  ["B1:B2 must make =B1<A1 true", "B1 = 5", "B2 must make =B1<A1 true", "", "B2 = 15"]);
eq("the input prompt, selected", [wb.describe("B1"), wb.tooltipOf("B1")], ["B1: 5 → 5 · under its A", "under its A"]);
await cmds.run("validation.set range=C1 kind=whole min=1 max=9 mode=info");
await cmds.run("validation.set range=C2 kind=whole min=1 max=9 mode=warn");
eq("info and warn keep what is typed and say so", [await message("cell.set cell=C1 input=12"), await message("cell.set cell=C2 input=12")],
  ["C1 = 12 (note: C1 must be a whole number from 1 to 9)", "C2 = 12 (C2 must be a whole number from 1 to 9)"]);
await cmds.run("file.save name=v");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=v");
eq("saved with the file", [wb2.notes.get("B2").validation, (await cmds2.run("cell.set cell=B2 input=99")).message],
  [{ kind: "custom", formula: "B1<A1", mode: "stop", when: "always", prompt: "under its A", origin: "B1" }, "B2 must make =B1<A1 true"]);
const said: string[] = [];
await cmds.script('SAY SETCELLVALIDATION("D1", "date", "2024-01-01", "2024-12-31", "prompt=in 2024;error=not this year")\nSAY SETCELLVALIDATION("D2", "text", "", "", "[a-z]+")\nSAY SETCELLVALIDATION("D3", "custom", "=D3>A1")\nSAY VALIDATECELL("D1", "2023-05-05") VALIDATECELL("D2", "abc") VALIDATECELL("D2", "ABC")',
  (l) => said.push(l));
eq("from REXX", [said, (await cmds.run("cell.set cell=D1 input=2023-05-05")).message, (await cmds.run("cell.set cell=D3 input=3")).message], [["OK", "OK", "OK", "0 1 0"], "D1 not this year", "D3 must make =D3>A1 true"]);

// expect: ok dates: as written or as serials
// expect: ok times
// expect: ok a pattern, and an error message of its own
// expect: ok refusals
// expect: ok custom: each cell's own formula, worked out with the input in
// expect: ok the input prompt, selected
// expect: ok info and warn keep what is typed and say so
// expect: ok saved with the file
// expect: ok from REXX
