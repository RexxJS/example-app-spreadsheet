// Trendlines (6.3): each fit against data made from its own kind of curve
// (so the coefficients come back and R² is 1), a noisy line's R² against
// one worked by hand, a moving average, the equations as Excel's labels
// write them; drawn on a chart; ADDTRENDLINE from REXX.
import { fit, movingAverage } from "./sheet/trend.ts";
import { layout } from "./sheet/charts.ts";
import type { Shape } from "./sheet/charts.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const xs = [1, 2, 3, 4, 5];
const e = (t: Parameters<typeof fit>[0], ys: number[], order?: number) => { const f = fit(t, xs, ys, order)!; return [f.equation, Math.round(f.r2! * 1e6) / 1e6]; };
eq("linear", e("linear", xs.map((x) => 2 * x + 1)), ["y = 2x + 1", 1]);
eq("exponential", e("exponential", xs.map((x) => 3 * Math.exp(0.5 * x))), ["y = 3e^0.5x", 1]);
eq("logarithmic", e("logarithmic", xs.map((x) => 2 * Math.log(x) - 1)), ["y = 2ln(x) - 1", 1]);
eq("power", e("power", xs.map((x) => 2 * x ** 1.5)), ["y = 2x^1.5", 1]);
eq("polynomial, order 2 and 3", [e("polynomial", xs.map((x) => 0.5 * x * x - 2 * x + 3), 2), e("polynomial", xs.map((x) => x ** 3 - x), 3)], [["y = 0.5x² - 2x + 3", 1], ["y = x³ - x", 1]]);
// y = 1, 3, 2, 5, 4: slope 0.8, intercept 0.6; SSres 3.6, SStot 10: R² 0.64.
eq("a noisy line's R²", e("linear", [1, 3, 2, 5, 4]), ["y = 0.8x + 0.6", 0.64]);
eq("a moving average", movingAverage(xs, [2, 4, 6, 8, 10], 3), [[3, 4], [4, 6], [5, 8]]);
eq("what cannot be fitted", [fit("exponential", xs, [1, -1, 2, 3, 4]), fit("polynomial", [1, 2], [1, 2], 3)], [null, null]);
const data = { labels: ["a", "b", "c", "d"], series: [{ name: "S", values: [2, 4, 6, 8] }] };
const shapes = layout("line", data, "", 400, 240, { trendlines: [{ series: 0, type: "linear", equation: true, r2: true }] });
const texts = shapes.filter((s): s is Extract<Shape, { kind: "text" }> => s.kind === "text").map((t) => t.text);
eq("drawn on a chart: the line, its label", [shapes.filter((s) => s.kind === "polyline").length, texts.includes("y = 2x   R² = 1")], [2, true]);
const wb = new Workbook();
const cmds = new Commands(wb);
for (const [k, v] of [["A2", "a"], ["A3", "b"], ["A4", "c"], ["B1", "N"], ["B2", "1"], ["B3", "3"], ["B4", "5"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
await cmds.run("chart.insert kind=line range=A1:B4");
eq("the command", [(await cmds.run("chart.trendline index=1 series=N type=linear equation=1 r2=1")).message, (await cmds.run("chart.trendline index=1 series=1 type=polynomial order=9")).message,
  (await cmds.run("chart.trendline index=1 series=2 type=linear")).message, (await cmds.run("chart.trendline index=1 series=1 type=none")).message],
  ["chart 1, N: a linear trendline: y = 2x - 1   R² = 1", "9 is not an order (2 to 6)", "2 is not a series of chart 1 (N)", "chart 1, N: no trendline"]);
const said: string[] = [];
await cmds.script('SAY ADDTRENDLINE(1, 1, "movingAverage", "period=2")', (l) => said.push(l));
eq("ADDTRENDLINE from REXX", said, ["chart 1, N: a 2-point moving average trendline"]);

// expect: ok linear
// expect: ok exponential
// expect: ok logarithmic
// expect: ok power
// expect: ok polynomial, order 2 and 3
// expect: ok a noisy line's R²
// expect: ok a moving average
// expect: ok what cannot be fitted
// expect: ok drawn on a chart: the line, its label
// expect: ok the command
// expect: ok ADDTRENDLINE from REXX
