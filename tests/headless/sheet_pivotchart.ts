// PivotCharts (6.6): a chart of a pivot, its data what the pivot last
// wrote (grand totals left out), following it when the pivot is written
// again from a changed source; saved with the file; INSERTPIVOTCHART.
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const store = memoryStore();
const wb = new Workbook();
const cmds = new Commands(wb, store);
const rows = [["region", "amount"], ["North", "10"], ["South", "4"], ["North", "6"]];
for (const [r, row] of rows.entries()) for (const [c, v] of row.entries()) await cmds.run(`cell.set cell=${"AB"[c]}${r + 1} input=${v}`);
await cmds.run("pivot.create id=p source=A1:B4 value=amount agg=sum output=D1 rows=region");
await cmds.run("pivot.options id=p grandTotals=1");
eq("a PivotChart of the pivot, its grand total left out", [(await cmds.run("chart.pivot id=p kind=pie")).message, wb.charts()[0].range, wb.charts()[0].data],
  ["inserted a pie PivotChart of pivot p", "pivot p", { labels: ["North", "South"], series: [{ name: "SUM of amount", values: [16, 4] }] }]);
await cmds.run("cell.set cell=A5 input=West");
await cmds.run("cell.set cell=B5 input=9");
await cmds.run("pivot.create id=p source=A1:B5 value=amount agg=sum output=D1 rows=region");
eq("it follows the pivot written again", wb.charts()[0].data.labels, ["North", "South", "West"]);
await cmds.run("file.save name=pc");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=pc");
eq("saved with the file", [wb2.charts()[0].range, wb2.charts()[0].data.labels.length], ["pivot p", 3]);
const said: string[] = [];
await cmds.script('SAY INSERTPIVOTCHART("p", "bar", "By region")\nSAY INSERTPIVOTCHART("q")', (l) => said.push(l));
eq("from REXX", said[0], "inserted a bar PivotChart of pivot p");

// expect: ok a PivotChart of the pivot, its grand total left out
// expect: ok it follows the pivot written again
// expect: ok saved with the file
// expect: ok from REXX
