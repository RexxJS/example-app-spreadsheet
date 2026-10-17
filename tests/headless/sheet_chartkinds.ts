// Chart kinds (6.1): stacked and 100% stacked columns, horizontal bars the
// same three, a stacked line, areas, a doughnut, scatter (with lines) and
// bubble, each as plain shapes in proportion; saved with the file;
// INSERTCHART, LISTCHARTS, DELETECHART from REXX.
import { layout } from "./sheet/charts.ts";
import type { ChartData, Shape } from "./sheet/charts.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const data: ChartData = { labels: ["Q1", "Q2"], series: [{ name: "North", values: [10, 30] }, { name: "South", values: [30, 10] }] };
const of = <K extends Shape["kind"]>(shapes: Shape[], kind: K) => shapes.filter((s): s is Extract<Shape, { kind: K }> => s.kind === kind);
const r = (n: number) => Math.round(n * 100) / 100;
const bars = (kind: Parameters<typeof layout>[0]) => of(layout(kind, data, "", 400, 240), "rect").filter((x) => x.w > 10 || x.h > 10).filter((x) => !(x.w === 10 && x.h === 10));
const st = bars("barStacked");
eq("stacked columns: South on North, one column a category, heights as values", [st.length, r(st[1].h / st[0].h), r(st[0].y), r(st[1].y + st[1].h), st[0].x === st[1].x], [4, 3, r(st[1].y + st[1].h), r(st[0].y), true]);
const pc = bars("bar100");
eq("100% stacked: every column the same height, shared", [r(pc[0].h + pc[1].h), r(pc[2].h + pc[3].h), r(pc[0].h / (pc[0].h + pc[1].h))], [r(pc[2].h + pc[3].h), r(pc[0].h + pc[1].h), 0.25]);
const hb = bars("hbar");
eq("horizontal bars: widths as values, one row a category", [r(hb[1].w / hb[0].w), hb[0].h === hb[1].h, hb[0].y < hb[2].y], [3, true, true]);
const hs = bars("hbarStacked");
eq("stacked horizontal: end to end", r(hs[0].x + hs[0].w), r(hs[1].x));
const ar = of(layout("area", data, "", 400, 240), "path"), as = of(layout("areaStacked", data, "", 400, 240), "path");
eq("areas: a filled path a series, closed to the axis; stacked on the one under", [ar.length, ar[0].d.endsWith("Z"), as[1].d.split(" ").length > as[0].d.split(" ").length - 1], [2, true, true]);
const ls = of(layout("lineStacked", data, "", 400, 240), "polyline");
eq("a stacked line: the second the sum's", ls[1].points[1] < ls[0].points[1], true);
const dn = of(layout("doughnut", { labels: ["A", "B"], series: [{ name: "S", values: [75, 25] }] }, "", 300, 240), "path");
eq("a doughnut: slices with a hole", [dn.length, dn[0].d.split("A").length - 1], [2, 2]);
const sc = layout("scatter", { labels: ["1", "2", "3"], series: [{ name: "x", values: [1, 2, 4] }, { name: "y", values: [10, 20, 40] }] }, "", 400, 240);
const pts = of(sc, "circle");
eq("scatter: the first column x, a point a row, in proportion", [pts.length, r((pts[2].cx - pts[0].cx) / (pts[1].cx - pts[0].cx)), r((pts[0].cy - pts[2].cy) / (pts[0].cy - pts[1].cy))], [3, 3, 3]);
eq("scatter with lines: a polyline through them", of(layout("scatterLines", { labels: ["1", "2"], series: [{ name: "x", values: [1, 2] }, { name: "y", values: [1, 2] }] }, "", 400, 240), "polyline").length, 1);
const bb = of(layout("bubble", { labels: ["1", "2"], series: [{ name: "x", values: [1, 2] }, { name: "y", values: [5, 5] }, { name: "size", values: [1, 4] }] }, "", 400, 240), "circle");
eq("bubbles: their areas as their sizes", r(((bb[1].r - 3) / (bb[0].r - 3)) ** 2), 4);

const store = memoryStore();
const wb = new Workbook();
const cmds = new Commands(wb, store);
for (const [k, v] of [["A2", "Q1"], ["A3", "Q2"], ["B1", "N"], ["B2", "1"], ["B3", "2"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
eq("the commands: any kind, any case", [(await cmds.run("chart.insert kind=AREASTACKED range=A1:B3")).message, wb.charts().map((c) => c.kind)], ["inserted a areaStacked chart of A1:B3", ["areaStacked"]]);
await cmds.run("file.save name=c");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=c");
eq("saved with the file", wb2.charts().map((c) => c.kind), ["areaStacked"]);
const said: string[] = [];
await cmds.script('SAY INSERTCHART("A1:B3", "doughnut", "Share")\nc = LISTCHARTS()\nSAY c.0 c.2\nSAY DELETECHART(1)', (l) => said.push(l));
eq("from REXX", said, ["inserted a doughnut chart of A1:B3", "2 2. Share: doughnut chart of A1:B3", "deleted Chart 1"]);

// expect: ok stacked columns: South on North, one column a category, heights as values
// expect: ok 100% stacked: every column the same height, shared
// expect: ok horizontal bars: widths as values, one row a category
// expect: ok stacked horizontal: end to end
// expect: ok areas: a filled path a series, closed to the axis; stacked on the one under
// expect: ok a stacked line: the second the sum's
// expect: ok a doughnut: slices with a hole
// expect: ok scatter: the first column x, a point a row, in proportion
// expect: ok scatter with lines: a polyline through them
// expect: ok bubbles: their areas as their sizes
// expect: ok the commands: any kind, any case
// expect: ok saved with the file
// expect: ok from REXX
