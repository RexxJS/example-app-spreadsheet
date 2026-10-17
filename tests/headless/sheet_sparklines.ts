// Sparklines past the line (4.8): win/loss (a gain ▀, a loss ▄, a zero on
// the axis ─), column, and Excel's Show options (the high and low points
// ▲ ▼, the first and last ◆); in Excel formulas, RexxJS's SPARKLINE_RANGE,
// insert.sparkline, and INSERTSPARKLINE from REXX.
import { sparkline, sparklineOptions } from "./sheet/charts.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("win/loss, the axis drawing a zero", [sparkline([3, -2, 0, 5], "winloss"), sparkline([3, -2, 0, 5], "winloss", { axis: true })], ["▀▄ ▀", "▀▄─▀"]);
eq("the high and low points, the first and last", [sparkline([1, 8, 4, 2], "line", { high: true, low: true }), sparkline([1, 8, 4, 2], "column", { first: true, last: true })], ["▼▲▄▂", "◆█▄◆"]);
eq("options as written", [sparklineOptions("high;low"), sparklineOptions("markers"), sparklineOptions("axis=0"), sparklineOptions("bold")],
  [{ high: true, low: true }, { high: true, low: true }, { axis: false }, null]);

const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("formulas.use language=Excel");
for (const [k, v] of [["A1", "4"], ["A2", "-1"], ["A3", "0"], ["A4", "2"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
eq("in Excel formulas, and refusals", [await message("insert.sparkline range=A1:A4 at=B1 type=winloss options=axis"), wb.sheet.input("B1"),
  (await cmds.run('cell.set cell=B2 input="=SPARKLINE(A1:A4, ""column"", ""high"")"')).ok, show(wb.sheet.value("B2")), (await cmds.run('cell.set cell=B3 input="=SPARKLINE(A1:A4, ""pie"")"')).ok, show(wb.sheet.value("B3")),
  await message("insert.sparkline range=A1:A4 at=B4 type=area")],
  ["B1 = ▀▄─▀", '=SPARKLINE(A1:A4, "winloss", "axis")', true, "▲▁▂▅", true, "#VALUE!", "area is not a sparkline type (line, column, winloss)"]);
await cmds.run("formulas.use language=RexxJS");
eq("and RexxJS's", [await message("insert.sparkline range=A1:A4 at=C1 type=winloss"), wb.sheet.input("C1")], ["C1 = ▀▄ ▀", '=SPARKLINE_RANGE("A1:A4", "winloss")']);
const said: string[] = [];
await cmds.script('SAY INSERTSPARKLINE("A1:A4", "D1", "line", "high;low")', (l) => said.push(l));
eq("INSERTSPARKLINE from REXX", said, ["D1 = ▲▼▂▅"]);

// expect: ok win/loss, the axis drawing a zero
// expect: ok the high and low points, the first and last
// expect: ok options as written
// expect: ok in Excel formulas, and refusals
// expect: ok and RexxJS's
// expect: ok INSERTSPARKLINE from REXX
