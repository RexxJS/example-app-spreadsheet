// Chart options (6.2): a title, axis titles, the legend's place (or none),
// data labels, gridlines off, Switch Row/Column, Select Data and Change
// Chart Type; undone, saved with the file; from REXX.
import { chartData, layout } from "./sheet/charts.ts";
import type { Shape } from "./sheet/charts.ts";
import { parseRange } from "./sheet/address.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import type { Value } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const cells: Record<string, Value> = { B1: "North", C1: "South", A2: "Q1", A3: "Q2", B2: 10, B3: 20, C2: 30, C3: 5 };
const data = chartData(parseRange("A1:C3")!, (k) => cells[k] ?? null);
eq("Switch Row/Column: the rows the series", chartData(parseRange("A1:C3")!, (k) => cells[k] ?? null, true), { labels: ["North", "South"], series: [{ name: "Q1", values: [10, 30] }, { name: "Q2", values: [20, 5] }] });
const texts = (sh: Shape[]) => sh.filter((s): s is Extract<Shape, { kind: "text" }> => s.kind === "text").map((t) => t.text);
const lines = (sh: Shape[]) => sh.filter((s) => s.kind === "line").length;
const plain = texts(layout("bar", data, "T", 400, 240)), dressed = texts(layout("bar", data, "T", 400, 240, { xTitle: "Quarter", yTitle: "Units", labels: true }));
eq("axis titles, and a data label a value", [dressed.includes("Quarter"), dressed.includes("Units"), plain.includes("Quarter"), dressed.length - plain.length, dressed.filter((t) => t === "5").length],
  [true, true, false, 6, 1]);
eq("gridlines off: the zero line alone", [lines(layout("bar", data, "", 400, 240)) > 2, lines(layout("bar", data, "", 400, 240, { gridlines: false }))], [true, 1]);
const swatches = (sh: Shape[]) => sh.filter((s): s is Extract<Shape, { kind: "rect" }> => s.kind === "rect" && s.w === 10 && s.h === 10);
eq("the legend: bottom, right (one under another), none", [swatches(layout("bar", data, "", 400, 240)).map((r) => r.y > 200), swatches(layout("bar", data, "", 400, 240, { legend: "right" })).map((r) => r.x > 300),
  swatches(layout("bar", data, "", 400, 240, { legend: "none" })).length], [[true, true], [true, true], 0]);
const store = memoryStore();
const wb = new Workbook();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
for (const [k, v] of Object.entries(cells)) await cmds.run(`cell.set cell=${k} input=${v}`);
await cmds.run("chart.insert kind=bar range=A1:C3");
eq("the commands", [await message("chart.option index=1 title=Sales xTitle=Quarter legend=right labels=1 gridlines=0"), await message("chart.option index=1 byRows=1"), wb.charts()[0].data.labels,
  await message("chart.data index=1 range=A1:B3 kind=area"), wb.charts()[0].kind, await message("chart.option index=1 legend=middle")],
  ["chart 1: title Sales, xTitle Quarter, legend right, labels true, gridlines false", "chart 1: xTitle Quarter, legend right, labels true, gridlines false, byRows true", ["North", "South"],
    "chart 1: a area chart of A1:B3", "area", "middle is not a legend's place (bottom, top, right, left, none)"]);
await cmds.run("edit.undo");
eq("undone", [wb.charts()[0].kind, wb.charts()[0].range], ["bar", "A1:C3"]);
await cmds.run("file.save name=o");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=o");
eq("saved with the file", [wb2.page(wb2.active)!.charts[0].options, wb2.charts()[0].title], [{ xTitle: "Quarter", legend: "right", labels: true, gridlines: false, byRows: true }, "Sales"]);
const said: string[] = [];
await cmds.script('SAY SETCHARTOPTION(1, "yTitle", "Units")\nSAY SETCHARTTYPE(1, "doughnut")\nSAY SETCHARTDATA(1, "A1:B3")', (l) => said.push(l));
eq("from REXX", said, ["chart 1: xTitle Quarter, legend right, labels true, gridlines false, byRows true, yTitle Units", "chart 1: a doughnut chart of A1:C3", "chart 1: a doughnut chart of A1:B3"]);

// expect: ok Switch Row/Column: the rows the series
// expect: ok axis titles, and a data label a value
// expect: ok gridlines off: the zero line alone
// expect: ok the legend: bottom, right (one under another), none
// expect: ok the commands
// expect: ok undone
// expect: ok saved with the file
// expect: ok from REXX
