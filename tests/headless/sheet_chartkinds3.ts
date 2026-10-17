// Treemap, sunburst, box and whisker (6.4b): a two-level range (a group
// column, a label column) read as such; squarified rectangles, their
// areas as values; rings of groups and labels, arcs as shares; Excel's
// exclusive-median quartiles, whiskers within 1.5 IQR, outliers.
import { boxStats, chartData, layout, squarify } from "./sheet/charts.ts";
import type { Shape } from "./sheet/charts.ts";
import { parseRange } from "./sheet/address.ts";
import type { Value } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const cells: Record<string, Value> = { A1: "Group", B1: "Item", C1: "Sales", A2: "Fruit", B2: "Apples", C2: 6, A3: null, B3: "Pears", C3: 2, A4: "Veg", B4: "Kale", C4: 4 };
const data = chartData(parseRange("A1:C4")!, (k) => cells[k] ?? null);
eq("two levels: the group carried down", data, { labels: ["Apples", "Pears", "Kale"], series: [{ name: "Sales", values: [6, 2, 4] }], groups: ["Fruit", "Fruit", "Veg"] });
const r = (n: number) => Math.round(n * 1000) / 1000;
const sq = squarify([6, 6, 4, 3, 2, 2, 1], 0, 0, 6, 4);
eq("squarified: areas as values, filling the box", [sq.map((b) => r(b.w * b.h)), r(sq.reduce((s, b) => s + b.w * b.h, 0))], [[6, 6, 4, 3, 2, 2, 1], 24]);
eq("and nearest square: no box thinner than a third of its length", sq.every((b) => Math.min(b.w, b.h) / Math.max(b.w, b.h) > 0.3), true);
const of = <K extends Shape["kind"]>(shapes: Shape[], kind: K) => shapes.filter((s): s is Extract<Shape, { kind: K }> => s.kind === kind);
const tm = layout("treemap", data, "", 400, 240);
eq("a treemap: a box a group, a box an item within it, labelled", [of(tm, "rect").length, of(tm, "text").map((t) => t.text)], [5, ["Fruit", "Apples", "Pears", "Veg", "Kale"]]);
eq("a sunburst: the groups' ring, the items' ring", of(layout("sunburst", data, "", 300, 240), "path").length, 5);
eq("Excel's quartiles, exclusive median", [boxStats([1, 2, 3, 4, 5, 6, 7]), boxStats([1, 2, 3, 4, 100])],
  [{ q1: 2, median: 4, q3: 6, mean: 4, low: 1, high: 7, outliers: [] }, { q1: 1.5, median: 3, q3: 52, mean: 22, low: 1, high: 100, outliers: [] }]);
eq("an outlier past 1.5 IQR", boxStats([1, 2, 3, 4, 5, 6, 30])!.outliers, [30]);
const bx = layout("box", { labels: ["1"], series: [{ name: "A", values: [1, 2, 3, 4, 5, 6, 30] }] }, "", 400, 240);
eq("a box: its box, median, whiskers, mean, outlier", [of(bx, "rect").length, of(bx, "circle").length, of(bx, "text").some((t) => t.text === "×")], [1, 1, true]);

// expect: ok two levels: the group carried down
// expect: ok squarified: areas as values, filling the box
// expect: ok and nearest square: no box thinner than a third of its length
// expect: ok a treemap: a box a group, a box an item within it, labelled
// expect: ok a sunburst: the groups' ring, the items' ring
// expect: ok Excel's quartiles, exclusive median
// expect: ok an outlier past 1.5 IQR
// expect: ok a box: its box, median, whiskers, mean, outlier
