// More chart kinds (6.4a): radar (a spoke a category, a polygon a series,
// distance as value), histogram (Excel's automatic bins, Scott's width; the
// bars touching), waterfall (each bar from the running total, a Total its
// sum, rises and falls in their colours, connectors), funnel (bars centred,
// widths as values).
import { bins, layout, PALETTE } from "./sheet/charts.ts";
import type { Shape } from "./sheet/charts.ts";
import { eq } from "./lib/check.ts";

const of = <K extends Shape["kind"]>(shapes: Shape[], kind: K) => shapes.filter((s): s is Extract<Shape, { kind: K }> => s.kind === kind);
const r = (n: number) => Math.round(n * 100) / 100;
const rad = layout("radar", { labels: ["a", "b", "c", "d"], series: [{ name: "S", values: [10, 5, 10, 5] }] }, "", 400, 240);
const pts = of(rad, "circle");
const cx = r((pts[0].cx + pts[2].cx) / 2), cy = r((pts[0].cy + pts[2].cy) / 2), d = (p: { cx: number; cy: number }) => Math.hypot(p.cx - cx, p.cy - cy);
eq("radar: a point a spoke, its distance its value, the polygon closed", [pts.length, r(d(pts[0]) / d(pts[1])), of(rad, "polyline").filter((p) => p.stroke === PALETTE[0])[0].points.length], [4, 2, 10]);
// n 9, s 1.2247: width 3.5 × 1.2247 / 9^⅓ = 2.0608, so two bins: [1, 3.06], (3.06, 5.12].
eq("histogram bins: Scott's width from the least", bins([1, 2, 2, 3, 3, 3, 4, 4, 5]).map((b) => [b.count, Math.round(b.to * 100) / 100]), [[6, 3.06], [3, 5.12]]);
const hist = of(layout("histogram", { labels: [], series: [{ name: "S", values: [1, 2, 2, 3, 3, 3, 4, 4, 5] }] }, "", 400, 240), "rect");
eq("its bars touch", r(hist[0].x + hist[0].w), r(hist[1].x));
const wf = of(layout("waterfall", { labels: ["Start", "Up", "Down", "Total"], series: [{ name: "S", values: [10, 5, -3, 0] }] }, "", 400, 240), "rect");
eq("waterfall: from the running total, a Total its sum", [wf.map((b) => b.fill), r(wf[1].y + wf[1].h), r(wf[0].y), r(wf[3].h / wf[0].h)], [[PALETTE[0], PALETTE[0], PALETTE[1], PALETTE[2]], r(wf[0].y), r(wf[1].y + wf[1].h), 1.2]);
eq("and its connectors", of(layout("waterfall", { labels: ["a", "b"], series: [{ name: "S", values: [1, 2] }] }, "", 400, 240), "line").filter((l) => l.stroke === "#808080" && l.y1 === l.y2 && l.x2 - l.x1 < 100).length >= 1, true);
const fn = of(layout("funnel", { labels: ["Leads", "Calls", "Sales"], series: [{ name: "S", values: [100, 50, 25] }] }, "", 400, 240), "rect");
eq("funnel: centred, widths as values", [r(fn[0].x + fn[0].w / 2) === r(fn[2].x + fn[2].w / 2), r(fn[0].w / fn[1].w), r(fn[1].w / fn[2].w)], [true, 2, 2]);

// expect: ok radar: a point a spoke, its distance its value, the polygon closed
// expect: ok histogram bins: Scott's width from the least
// expect: ok its bars touch
// expect: ok waterfall: from the running total, a Total its sum
// expect: ok and its connectors
// expect: ok funnel: centred, widths as values
