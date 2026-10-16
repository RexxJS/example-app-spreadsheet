// The status bar (3.0.5): a range's average, count, numerical count, min,
// max and sum, as Excel's status bar gives them (range.stats), and
// SELECTIONSTATS from REXX, the line or one figure.
import { Commands } from "./sheet/commands.ts";
import { Workbook, statsLine } from "./sheet/workbook.ts";
import { parseRange } from "./sheet/address.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
for (const [k, v] of [["A1", "1"], ["A2", "2"], ["A3", "x"], ["A4", "=A1+A2"], ["B1", "=1/0"]]) await cmds.execute("cell.set", { cell: k, input: v });
eq("numbers, text and an empty cell: Excel's figures", wb.stats(parseRange("A1:A5")!), { count: 4, numbers: 3, sum: 6, average: 2, min: 1, max: 3 });
eq("as the status bar says them", (await cmds.execute("range.stats", { range: "A1:A5" })).message, "Average: 2  Count: 4  Numerical Count: 3  Min: 1  Max: 3  Sum: 6");
eq("no numbers: the count alone (an error is counted, not summed)", [statsLine(wb.stats(parseRange("A3:B3")!)), statsLine(wb.stats(parseRange("B1:B2")!))], ["Count: 1", "Count: 1"]);
eq("either way round", wb.stats(parseRange("A4:A1")!).sum, 6);
const said: string[] = [];
eq("SELECTIONSTATS from REXX", [(await cmds.script('SAY SELECTIONSTATS("A1:A4")\nSAY SELECTIONSTATS("A1:A4", "sum") SELECTIONSTATS("A1:A4", "Average")', (l) => said.push(l))).message, said,
  (await cmds.script('LET r = SELECTIONSTATS("A1:A4", "median")', () => {})).message],
  ["ran", ["Average: 2  Count: 4  Numerical Count: 3  Min: 1  Max: 3  Sum: 6", "6 2"], "SELECTIONSTATS: median is not a figure (sum, average, count, numbers, min, max)"]);

// expect: ok numbers, text and an empty cell: Excel's figures
// expect: ok as the status bar says them
// expect: ok no numbers: the count alone (an error is counted, not summed)
// expect: ok either way round
// expect: ok SELECTIONSTATS from REXX
