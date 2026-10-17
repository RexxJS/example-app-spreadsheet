// The Analysis ToolPak's tools (6.9), each against values worked in
// Python from the formulas (statistics.stdev, sums by hand); p-values on
// distributions.ts's t and F (2.8's, held to Excel's examples); the
// command writing a report's values; ANALYZE from REXX.
import { analyze } from "./sheet/analysis.ts";
import type { Cell } from "./sheet/analysis.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const A = { name: "A", values: [2, 4, 4, 4, 5, 5, 7, 9] }, B = { name: "B", values: [1, 3, 2, 5, 4, 6, 7, 8] };
const r = (v: Cell) => (typeof v === "number" ? Math.round(v * 1e9) / 1e9 : v);
const get = (t: Cell[][], label: string, col = 1) => r(t.find((row) => row[0] === label)![col]);
const d = analyze("descriptive", [A]);
eq("descriptive statistics", ["Mean", "Median", "Mode", "Standard Deviation", "Sample Variance", "Kurtosis", "Skewness", "Range", "Sum", "Count"].map((l) => get(d, l)),
  [5, 4.5, 4, 2.138089935, 4.571428571, 0.940625, 0.818487553, 7, 40, 8]);
eq("correlation and covariance (Excel's tool's: the population's)", [r(analyze("correlation", [A, B])[2][1]), r(analyze("covariance", [A, B])[2][1]), analyze("correlation", [A, B])[1]], [0.900148797, 4.125, ["A", 1, ""]]);
eq("a histogram's bins, and More", analyze("histogram", [A], { bins: [3, 5] }), [["Bin", "Frequency"], [3, 1], [5, 5], ["More", 2]]);
eq("a moving average", analyze("movingAverage", [A], { interval: 3 }).slice(1, 5).map((x) => r(x[0])), ["#N/A", "#N/A", 3.333333333, 4]);
eq("exponential smoothing, damping 0.3", analyze("exponentialSmoothing", [A], { damping: 0.3 }).slice(1, 5).map((x) => r(x[0])), ["#N/A", 2, 3.4, 3.82]);
eq("rank and percentile", analyze("rankPercentile", [A]).slice(1, 4), [[8, 9, 1, 1], [7, 7, 2, 0.857], [5, 5, 3, 0.571]]);
const reg = analyze("regression", [B, A]);
eq("regression: the line, R², its ANOVA", [r(reg.find((x) => x[0] === "Intercept")![1]), r(reg.find((x) => x[0] === "A")![1]), get(reg, "R Square"), get(reg, "Observations"), reg.find((x) => x[0] === "Residual")![1]],
  [-0.65625, 1.03125, 0.810267857, 8, 6]);
eq("t-tests: paired, equal variances; the F-test", [get(analyze("tTestPaired", [A, B]), "t Stat"), get(analyze("tTestEqual", [A, B]), "t Stat"), get(analyze("fTest", [A, B]), "F"),
  get(analyze("tTestEqual", [A, B]), "df"), get(analyze("tTestUnequal", [A, B]), "df")], [1.322875656, 0.434958836, 0.761904762, 14, 14]);
const p = get(analyze("tTestPaired", [A, B]), "P(T<=t) two-tail") as number, p1 = get(analyze("tTestPaired", [A, B]), "P(T<=t) one-tail") as number;
eq("its p-values, two tails twice one", [p > 0.2 && p < 0.25, Math.abs(p - 2 * p1) < 1e-9], [true, true]);
// With 2 degrees of freedom between, F's tail is (1 + 2F/d2)^(-d2/2): (1 + 9)^-3, 0.001.
const an = analyze("anova", [{ name: "x", values: [1, 2, 3] }, { name: "y", values: [4, 5, 6] }, { name: "z", values: [7, 8, 9] }]);
eq("single-factor ANOVA", [an.find((x) => x[0] === "Between Groups")!.slice(1, 5).map(r), r(an.find((x) => x[0] === "Within Groups")![1]), Math.abs((an.find((x) => x[0] === "Between Groups")![5] as number) - 0.001) < 1e-9],
  [[54, 2, 27, 27], 6, true]);
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("cell.set cell=A1 input=Score");
for (const [i, v] of A.values.entries()) await cmds.run(`cell.set cell=A${i + 2} input=${v}`);
eq("the command: values at the output", [(await cmds.run("data.analyze tool=descriptive input=A1:A9 output=C1 labels=1")).message, show(wb.sheet.value("C1")), show(wb.sheet.value("D2")), show(wb.sheet.value("D14"))],
  ["descriptive of A1:A9 at C1: 30 cells set, 30 changed", "Score", "5", "8"]);
eq("refusals", [(await cmds.run("data.analyze tool=fourier input=A1:A9 output=C1")).message, (await cmds.run("data.analyze tool=regression input=A1:A9 output=C1 labels=1")).message],
  ["fourier is not a tool (descriptive, correlation, covariance, histogram, movingAverage, exponentialSmoothing, rankPercentile, regression, tTestPaired, tTestEqual, tTestUnequal, fTest, anova)", "regression needs 2 columns of numbers"]);
const said: string[] = [];
await cmds.script('SAY ANALYZE("movingAverage", "A2:A9", "F1", "interval=2")', (l) => said.push(l));
eq("ANALYZE from REXX", [said, show(wb.sheet.value("F3"))], [["movingAverage of A2:A9 at F1: 9 cells set, 9 changed"], "3"]);

// expect: ok descriptive statistics
// expect: ok correlation and covariance (Excel's tool's: the population's)
// expect: ok a histogram's bins, and More
// expect: ok a moving average
// expect: ok exponential smoothing, damping 0.3
// expect: ok rank and percentile
// expect: ok regression: the line, R², its ANOVA
// expect: ok t-tests: paired, equal variances; the F-test
// expect: ok its p-values, two tails twice one
// expect: ok single-factor ANOVA
// expect: ok the command: values at the output
// expect: ok refusals
// expect: ok ANALYZE from REXX
