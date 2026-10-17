// Stock and combo charts (6.4c): high-low-close (a line, a close tick) and
// open-high-low-close (a box, white when it rose, dark when it fell); a
// combo's first series columns, the rest lines, those asked for on a
// secondary axis at the right with its own scale.
import { layout } from "./sheet/charts.ts";
import type { Shape } from "./sheet/charts.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const of = <K extends Shape["kind"]>(shapes: Shape[], kind: K) => shapes.filter((s): s is Extract<Shape, { kind: K }> => s.kind === kind);
const hlc = layout("stock", { labels: ["Mon", "Tue"], series: [{ name: "High", values: [12, 14] }, { name: "Low", values: [8, 9] }, { name: "Close", values: [10, 13] }] }, "", 400, 240);
eq("high-low-close: a line a day, its close a tick", of(hlc, "line").filter((l) => l.stroke === "#404040").length, 4);
const ohlc = of(layout("stock", { labels: ["Mon", "Tue"], series: [{ name: "Open", values: [9, 13] }, { name: "High", values: [12, 14] }, { name: "Low", values: [8, 9] }, { name: "Close", values: [11, 10] }] }, "", 400, 240), "rect");
eq("open-high-low-close: a rise white, a fall dark", ohlc.map((r) => r.fill), ["#ffffff", "#404040"]);
const data = { labels: ["Q1", "Q2"], series: [{ name: "Units", values: [10, 20] }, { name: "Margin", values: [0.1, 0.4] }] };
const one = layout("combo", data, "", 400, 240), two = layout("combo", data, "", 400, 240, { secondary: [1] });
const texts = (sh: Shape[]) => of(sh, "text").map((t) => t.text);
eq("combo: columns then a line", [of(one, "rect").filter((r) => r.w > 10).length, of(one, "polyline").length], [2, 1]);
eq("the line's values on their own axis at the right", [texts(two).includes("0.4") || texts(two).includes("0.5"), (() => { const p = of(two, "polyline")[0].points; return p[3] < p[1] && p[1] - p[3] > 100; })()], [true, true]);
const wb = new Workbook();
const cmds = new Commands(wb);
for (const [k, v] of [["A2", "Q1"], ["A3", "Q2"], ["B1", "U"], ["C1", "M"], ["B2", "10"], ["B3", "20"], ["C2", "0.1"], ["C3", "0.4"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
await cmds.run("chart.insert kind=combo range=A1:C3");
eq("the secondary axis by command", [(await cmds.run("chart.option index=1 secondary=2")).message, (await cmds.run("chart.option index=1 secondary=x")).message],
  ["chart 1: secondary 1", "x is not series numbers (1, 2 ...)"]);

// expect: ok high-low-close: a line a day, its close a tick
// expect: ok open-high-low-close: a rise white, a fall dark
// expect: ok combo: columns then a line
// expect: ok the line's values on their own axis at the right
// expect: ok the secondary axis by command
