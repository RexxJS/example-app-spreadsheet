// Charts laid out with no toolkit, as GridCraft's chart crate lays them out:
// a range read as categories and series (chartData), a scale with round
// ticks (niceScale), and any of fifteen of Excel's kinds as plain shapes (layout)
// that the page draws in vg and a spec can read back. And sparklines, as
// text a cell can hold: eight heights of block.
import { range } from "./address.ts";
import type { Addr } from "./address.ts";
import { show } from "./format.ts";
import type { Value } from "./values.ts";
import { coef, fit, movingAverage } from "./trend.ts";
import type { TrendType } from "./trend.ts";

// Excel's kinds: columns (bar here, as the app first named it), stacked
// and 100% stacked; horizontal bars (hbar) the same three; a line with
// markers, a stacked one; areas, stacked; pie and doughnut; scatter, with
// lines, bubble.
export type ChartKind = "bar" | "barStacked" | "bar100" | "hbar" | "hbarStacked" | "hbar100" | "line" | "lineStacked" | "area" | "areaStacked" | "pie" | "doughnut" | "scatter" | "scatterLines" | "bubble";
export const CHART_KINDS: readonly ChartKind[] = ["bar", "barStacked", "bar100", "hbar", "hbarStacked", "hbar100", "line", "lineStacked", "area", "areaStacked", "pie", "doughnut", "scatter", "scatterLines", "bubble"];

// Office's palette, series by series.
export const PALETTE = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"];

// Chart Elements and Select Data: axis titles, the legend's place (or
// none), data labels, gridlines (on unless off), the series by rows
// (Switch Row/Column).
export interface Trendline {
  readonly series: number;          // from 0
  readonly type: TrendType;
  readonly order?: number;          // a polynomial's
  readonly period?: number;         // a moving average's
  readonly equation?: boolean;
  readonly r2?: boolean;
}
export interface ChartOptions {
  readonly trendlines?: readonly Trendline[];
  readonly xTitle?: string;
  readonly yTitle?: string;
  readonly legend?: "bottom" | "top" | "right" | "left" | "none";
  readonly labels?: boolean;
  readonly gridlines?: boolean;
  readonly byRows?: boolean;
}
export const LEGENDS = ["bottom", "top", "right", "left", "none"];

export interface Series {
  readonly name: string;
  readonly values: readonly number[];
}

export interface ChartData {
  readonly labels: readonly string[];
  readonly series: readonly Series[];
}

export type Shape =
  | { readonly kind: "rect"; readonly x: number; readonly y: number; readonly w: number; readonly h: number; readonly fill: string }
  | { readonly kind: "line"; readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number; readonly stroke: string; readonly width: number }
  | { readonly kind: "polyline"; readonly points: readonly number[]; readonly stroke: string; readonly width: number }
  | { readonly kind: "circle"; readonly cx: number; readonly cy: number; readonly r: number; readonly fill: string }
  | { readonly kind: "path"; readonly d: string; readonly fill: string; readonly stroke: string }
  | { readonly kind: "text"; readonly x: number; readonly y: number; readonly size: number; readonly anchor: "start" | "middle" | "end"; readonly text: string; readonly fill: string };

const NUMBER = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/;
const isNumber = (v: Value) => typeof v === "number" || (typeof v === "string" && NUMBER.test(v));
const numberOf = (v: Value): number => (isNumber(v) ? Number(v) : 0);
const isText = (v: Value) => v !== null && !isNumber(v);

// A range as Excel reads one for a chart: when its first column is text
// below its first row, those are the categories; when its first row is text
// (or empty) beside them, those name the series; each other column is a
// series, anything not a number in it 0.
export function chartData([a, b]: [Addr, Addr], value: (key: string) => Value, byRows = false): ChartData {
  const keys = [...range(a, b)], cols = Math.abs(a.col - b.col) + 1;
  const byCol = Array.from({ length: keys.length / cols }, (_, r) => keys.slice(r * cols, (r + 1) * cols).map(value));
  // Switch Row/Column: the series are the rows, the categories the columns.
  const grid = byRows ? byCol[0].map((_, c) => byCol.map((row) => row[c])) : byCol;
  const labelled = cols > 1 && grid.slice(1).some((row) => isText(row[0]));
  const headed = grid.length > 1 && grid[0].slice(labelled ? 1 : 0).some(isText);
  const rows = headed ? grid.slice(1) : grid;
  const first = labelled ? 1 : 0;
  return {
    labels: rows.map((row, i) => (labelled ? show(row[0]) : String(i + 1))),
    series: grid[0].slice(first).map((h, j) => ({
      name: headed ? show(h) : `Series ${j + 1}`,
      values: rows.map((row) => numberOf(row[first + j])),
    })),
  };
}

const tidy = (n: number) => Number(n.toPrecision(12));

// A scale from lo to hi (0 included) in round steps: 1, 2 or 5 times a
// power of ten, about five of them.
export function niceScale(lo: number, hi: number): { lo: number; hi: number; step: number } {
  let a = Math.min(0, lo), b = Math.max(0, hi);
  if (a === b) b = a + 1;
  const rough = (b - a) / 5, mag = 10 ** Math.floor(Math.log10(rough));
  const step = tidy(([1, 2, 5, 10].find((m) => m * mag >= rough - 1e-12) ?? 10) * mag);
  return { lo: tidy(Math.floor(a / step + 1e-9) * step), hi: tidy(Math.ceil(b / step - 1e-9) * step), step };
}

const text = (x: number, y: number, s: string, anchor: "start" | "middle" | "end" = "middle", size = 10, fill = "#404040"): Shape =>
  ({ kind: "text", x, y, size, anchor, text: s, fill });

// The plot's box once the title, the legend and the axis titles have
// their room; and the legend where it is asked for.
interface Frame { readonly x0: number; readonly x1: number; readonly y0: number; readonly y1: number }
function frame(w: number, h: number, title: string, o: ChartOptions, series: number, left = 40): Frame {
  const place = o.legend ?? "bottom", shown = series > 1 && place !== "none";
  return {
    x0: left + (o.yTitle ? 14 : 0) + (shown && place === "left" ? 80 : 0),
    x1: w - 10 - (shown && place === "right" ? 80 : 0),
    y0: (title ? 28 : 12) + (shown && place === "top" ? 16 : 0),
    y1: h - 24 - (shown && place === "bottom" ? 16 : 0) - (o.xTitle ? 14 : 0),
  };
}
function placedLegend(names: readonly string[], o: ChartOptions, f: Frame, w: number, h: number, title: string): Shape[] {
  const place = o.legend ?? "bottom";
  if (names.length < 2 || place === "none") return [];
  if (place === "top") return legend(names, f.x0, (title ? 28 : 12) + 8);
  if (place === "bottom") return legend(names, f.x0, h - 8);
  // Right or left: one under another.
  const x = place === "right" ? w - 84 : 4, out: Shape[] = [];
  names.forEach((name, i) => {
    out.push({ kind: "rect", x, y: f.y0 + i * 14, w: 10, h: 10, fill: PALETTE[i % PALETTE.length] });
    out.push(text(x + 14, f.y0 + i * 14 + 9, name, "start"));
  });
  return out;
}
const axisTitles = (o: ChartOptions, f: Frame, h: number): Shape[] => [
  ...(o.xTitle ? [text((f.x0 + f.x1) / 2, f.y1 + 28, o.xTitle, "middle", 10, "#595959")] : []),
  ...(o.yTitle ? [text(f.x0 - 30 > 2 ? 2 : 2, f.y0 - 4, o.yTitle, "start", 10, "#595959")] : []),
];

// A row of swatches and names under the plot.
function legend(names: readonly string[], x: number, y: number): Shape[] {
  const out: Shape[] = [];
  for (const [i, name] of names.entries()) {
    out.push({ kind: "rect", x, y: y - 9, w: 10, h: 10, fill: PALETTE[i % PALETTE.length] });
    out.push(text(x + 14, y, name, "start"));
    x += 24 + name.length * 6;
  }
  return out;
}

// Columns, bars, lines and areas: categories along one axis, values the
// other; stacked kinds sum a category's series, 100% kinds share it.
function axes(kind: ChartKind, data: ChartData, w: number, h: number, title: string, o: ChartOptions = {}): Shape[] {
  const stacked = /Stacked$|100$/.test(kind), share = kind.endsWith("100"), horizontal = kind.startsWith("hbar");
  const columns = kind.startsWith("bar") || horizontal, area = kind.startsWith("area");
  const n = data.labels.length;
  // Each series' values as drawn: from a base (the ones under it stacked).
  const totals = data.labels.map((_, i) => data.series.reduce((t, s) => t + Math.abs(s.values[i]), 0));
  const drawn = data.series.map((s) => s.values.map((v, i) => (share ? (totals[i] ? (v / totals[i]) * 100 : 0) : v)));
  const bases = drawn.map((vs, j) => vs.map((_, i) => (stacked || area && kind === "areaStacked" ? drawn.slice(0, j).reduce((t, x) => t + x[i], 0) : 0)));
  const tops = drawn.map((vs, j) => vs.map((v, i) => bases[j][i] + v));
  const all = stacked ? tops.flat() : drawn.flat();
  const scale = share ? { lo: 0, hi: 100, step: 20 } : niceScale(Math.min(...all, 0), Math.max(...all, 0));
  const f = frame(w, h, title, o, data.series.length, horizontal ? 60 : 40), { x0, x1, y0, y1 } = f;
  // The value axis: up the left, or (horizontal) along the bottom.
  const vOf = (v: number) => (horizontal ? x0 + ((v - scale.lo) / (scale.hi - scale.lo)) * (x1 - x0) : y1 - ((v - scale.lo) / (scale.hi - scale.lo)) * (y1 - y0));
  const band = ((horizontal ? y1 - y0 : x1 - x0)) / Math.max(1, n);
  const cOf = (i: number) => (horizontal ? y0 + (i + 0.5) * band : x0 + (i + 0.5) * band);
  const out: Shape[] = title ? [text(w / 2, 16, title, "middle", 13)] : [];
  const ticks: Shape[] = [];
  for (let t = scale.lo; t <= scale.hi + scale.step / 2; t = tidy(t + scale.step)) {
    const at = vOf(t), grid = t === 0 ? "#808080" : "#d9d9d9";
    if (o.gridlines !== false || t === 0) out.push(horizontal ? { kind: "line", x1: at, y1: y0, x2: at, y2: y1, stroke: grid, width: 1 } : { kind: "line", x1: x0, y1: at, x2: x1, y2: at, stroke: grid, width: 1 });
    const label = `${show(t)}${share ? "%" : ""}`;
    ticks.push(horizontal ? text(at, y1 + 14, label) : text(x0 - 4, at + 3, label, "end"));
  }
  out.push(...ticks);
  if (columns) {
    const width = stacked ? band * 0.6 : (band * 0.8) / Math.max(1, data.series.length);
    data.labels.forEach((_, i) => data.series.forEach((_s, j) => {
      const from = stacked ? bases[j][i] : 0, to = stacked ? tops[j][i] : drawn[j][i];
      const a = vOf(Math.min(from, to)), b = vOf(Math.max(from, to)), offset = stacked ? band * 0.2 : band * 0.1 + j * width;
      const fill = PALETTE[j % PALETTE.length];
      out.push(horizontal ? { kind: "rect", x: a, y: y0 + i * band + offset, w: b - a, h: width, fill } : { kind: "rect", x: x0 + i * band + offset, y: b, w: width, h: a - b, fill });
      if (o.labels) out.push(horizontal ? text(b + 3, y0 + i * band + offset + width / 2 + 3, show(data.series[j].values[i]), "start", 9) : text(x0 + i * band + offset + width / 2, b - 3, show(data.series[j].values[i]), "middle", 9));
    }));
  } else {
    data.series.forEach((_s, j) => {
      const colour = PALETTE[j % PALETTE.length];
      const line = (kind === "lineStacked" || kind === "areaStacked" ? tops[j] : drawn[j]).flatMap((v, i) => [cOf(i), vOf(v)]);
      if (area) {
        const under = kind === "areaStacked" ? bases[j].flatMap((v, i) => [cOf(i), vOf(v)]) : data.labels.flatMap((_, i) => [cOf(i), vOf(Math.max(scale.lo, 0))]);
        const pts: string[] = [];
        for (let k = 0; k < line.length; k += 2) pts.push(`${k ? "L" : "M"} ${line[k]} ${line[k + 1]}`);
        for (let k = under.length - 2; k >= 0; k -= 2) pts.push(`L ${under[k]} ${under[k + 1]}`);
        out.push({ kind: "path", d: `${pts.join(" ")} Z`, fill: colour, stroke: colour });
      } else {
        out.push({ kind: "polyline", points: line, stroke: colour, width: 2 });
        for (let k = 0; k < line.length; k += 2) out.push({ kind: "circle", cx: line[k], cy: line[k + 1], r: 2.5, fill: colour });
      }
      if (o.labels) for (let k = 0; k < line.length; k += 2) out.push(text(line[k], line[k + 1] - 5, show(data.series[j].values[k / 2]), "middle", 9));
    });
  }
  data.labels.forEach((l, i) => out.push(horizontal ? text(x0 - 4, cOf(i) + 3, l, "end") : text(cOf(i), y1 + 14, l)));
  // Trendlines on a category chart: x the category's number, from 1.
  if (!horizontal && !stacked) out.push(...trendShapes(o, data.series, data.labels.map((_, i) => i + 1), (x) => cOf(x - 1), vOf, f));
  out.push(...axisTitles(o, f, h), ...placedLegend(data.series.map((s) => s.name), o, f, w, h, title));
  return out;
}

// Each trendline asked for: the fit drawn across the plot, its equation
// and R² at the top right; a moving average through its means.
function trendShapes(o: ChartOptions, series: readonly Series[], xs: readonly number[], X: (x: number) => number, Y: (y: number) => number, f: Frame): Shape[] {
  const out: Shape[] = [];
  let line = 0;
  for (const t of o.trendlines ?? []) {
    const s = series[t.series];
    if (!s) continue;
    const colour = PALETTE[t.series % PALETTE.length], lo = Math.min(...xs), hi = Math.max(...xs);
    if (t.type === "movingAverage") {
      const pts = movingAverage(xs, s.values, Math.max(2, t.period ?? 2)).flatMap(([x, y]) => [X(x), Y(y)]);
      if (pts.length >= 4) out.push({ kind: "polyline", points: pts, stroke: colour, width: 1.5 });
      continue;
    }
    const fitted = fit(t.type, xs, s.values, t.order ?? 2);
    if (!fitted) continue;
    const pts: number[] = [];
    for (let k = 0; k <= 40; k++) {
      const x = lo + ((hi - lo) * k) / 40, y = fitted.at(x);
      if (isFinite(y)) pts.push(X(x), Math.max(f.y0, Math.min(f.y1, Y(y))));
    }
    out.push({ kind: "polyline", points: pts, stroke: colour, width: 1.5 });
    const said = [t.equation ? fitted.equation : "", t.r2 && fitted.r2 !== null ? `R² = ${coef(fitted.r2)}` : ""].filter((x) => x).join("   ");
    if (said) out.push(text(f.x1 - 4, f.y0 + 12 + line++ * 12, said, "end", 9, colour));
  }
  return out;
}

// Scatter and bubble charts: numbers on both axes, the first column the
// x of every point, each other column a series' y (bubble: x, y and a
// size, each three columns a series).
function xy(kind: ChartKind, data: ChartData, w: number, h: number, title: string, o: ChartOptions = {}): Shape[] {
  // A range of numbers alone: its first column the x of every point.
  const counted = data.labels.every((l, i) => l === String(i + 1)) && data.series.length > 1;
  const xs = counted ? [...data.series[0].values] : data.labels.map((l) => Number(l)).map((x, i) => (isFinite(x) ? x : i + 1));
  const ys0 = counted ? data.series.slice(1) : data.series;
  const series = kind === "bubble" ? ys0.filter((_, j) => j % 2 === 0).map((s, k) => ({ s, size: ys0[k * 2 + 1] })) : ys0.map((s) => ({ s, size: undefined }));
  const ys = series.flatMap(({ s }) => s.values);
  const sx = niceScale(Math.min(...xs, 0), Math.max(...xs, 0)), sy = niceScale(Math.min(...ys, 0), Math.max(...ys, 0));
  const f = frame(w, h, title, o, series.length), { x0, x1, y0, y1 } = f;
  const X = (v: number) => x0 + ((v - sx.lo) / (sx.hi - sx.lo)) * (x1 - x0), Y = (v: number) => y1 - ((v - sy.lo) / (sy.hi - sy.lo)) * (y1 - y0);
  const out: Shape[] = title ? [text(w / 2, 16, title, "middle", 13)] : [];
  for (let t = sy.lo; t <= sy.hi + sy.step / 2; t = tidy(t + sy.step)) { if (o.gridlines !== false) out.push({ kind: "line", x1: x0, y1: Y(t), x2: x1, y2: Y(t), stroke: "#d9d9d9", width: 1 }); out.push(text(x0 - 4, Y(t) + 3, show(t), "end")); }
  for (let t = sx.lo; t <= sx.hi + sx.step / 2; t = tidy(t + sx.step)) out.push(text(X(t), y1 + 14, show(t)));
  const biggest = Math.max(1e-9, ...series.flatMap(({ size }) => size?.values.map(Math.abs) ?? []));
  series.forEach(({ s, size }, j) => {
    const colour = PALETTE[j % PALETTE.length], points = s.values.flatMap((v, i) => [X(xs[i]), Y(v)]);
    if (kind === "scatterLines") out.push({ kind: "polyline", points, stroke: colour, width: 2 });
    for (let k = 0; k < points.length; k += 2) out.push({ kind: "circle", cx: points[k], cy: points[k + 1], r: size ? 3 + 17 * Math.sqrt(Math.abs(size.values[k / 2]) / biggest) : 3, fill: colour });
    if (o.labels) for (let k = 0; k < points.length; k += 2) out.push(text(points[k], points[k + 1] - 6, show(s.values[k / 2]), "middle", 9));
  });
  if (kind !== "bubble") out.push(...trendShapes(o, series.map(({ s }) => s), xs, X, Y, f));
  out.push(...axisTitles(o, f, h), ...placedLegend(series.map(({ s }) => s.name), o, f, w, h, title));
  return out;
}

function pie(data: ChartData, w: number, h: number, title: string, hole = 0, o: ChartOptions = {}): Shape[] {
  const values = (data.series[0]?.values ?? []).map((v) => Math.max(0, v));
  const total = values.reduce((s, v) => s + v, 0);
  const top = title ? 28 : 12, bottom = h - 24;
  const r = Math.min(w - 20, bottom - top) / 2, cx = w / 2, cy = (top + bottom) / 2;
  const out: Shape[] = title ? [text(w / 2, 16, title, "middle", 13)] : [];
  const at = (angle: number, radius: number) => [cx + radius * Math.sin(angle), cy - radius * Math.cos(angle)];
  const labels: Shape[] = [];
  let angle = 0;
  values.forEach((v, i) => {
    if (!total || !v) return;
    const sweep = (v / total) * 2 * Math.PI, [ax, ay] = at(angle, r), [bx, by] = at(angle + sweep, r);
    const ri = r * hole, [cxa, cya] = at(angle, ri), [cxb, cyb] = at(angle + sweep, ri);
    const d = sweep >= 2 * Math.PI - 1e-9
      ? `M ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy} Z${hole ? ` M ${cx - ri} ${cy} A ${ri} ${ri} 0 1 0 ${cx + ri} ${cy} A ${ri} ${ri} 0 1 0 ${cx - ri} ${cy} Z` : ""}`
      : hole ? `M ${cxa} ${cya} L ${ax} ${ay} A ${r} ${r} 0 ${sweep > Math.PI ? 1 : 0} 1 ${bx} ${by} L ${cxb} ${cyb} A ${ri} ${ri} 0 ${sweep > Math.PI ? 1 : 0} 0 ${cxa} ${cya} Z`
      : `M ${cx} ${cy} L ${ax} ${ay} A ${r} ${r} 0 ${sweep > Math.PI ? 1 : 0} 1 ${bx} ${by} Z`;
    out.push({ kind: "path", d, fill: PALETTE[i % PALETTE.length], stroke: "#ffffff" });
    const [lx, ly] = at(angle + sweep / 2, r * (hole ? (1 + hole) / 2 : 0.65));
    labels.push(text(lx, ly + 3, `${Math.round((v / total) * 100)}%`, "middle", 10, "#ffffff"));
    angle += sweep;
  });
  return [...out, ...(o.labels === false ? [] : labels), ...((o.legend ?? "bottom") === "none" ? [] : legend(data.labels, 10, h - 8))];
}

// A chart as shapes, in a w x h box.
export function layout(kind: ChartKind, data: ChartData, title: string, w: number, h: number, o: ChartOptions = {}): Shape[] {
  return kind === "pie" ? pie(data, w, h, title, 0, o) : kind === "doughnut" ? pie(data, w, h, title, 0.5, o) : kind.startsWith("scatter") || kind === "bubble" ? xy(kind, data, w, h, title, o) : axes(kind, data, w, h, title, o);
}

const BLOCKS = "▁▂▃▄▅▆▇█";

export const SPARKLINE_TYPES = ["line", "column", "winloss"] as const;
export type SparklineType = (typeof SPARKLINE_TYPES)[number];
// Excel's Show group and its axis: the high and low points marked, the
// first and last, and (win/loss) a zero drawn on the axis.
export interface SparklineOptions {
  readonly high?: boolean;
  readonly low?: boolean;
  readonly first?: boolean;
  readonly last?: boolean;
  readonly axis?: boolean;
}

// The options as written: "high;low;axis" (or "high=1;axis=0").
export function sparklineOptions(text: string): SparklineOptions | null {
  const out: Record<string, boolean> = {};
  for (const part of text.split(/[;,]/).map((x) => x.trim().toLowerCase()).filter((x) => x)) {
    const [k, v = "1"] = part.split("=").map((x) => x.trim());
    if (!["high", "low", "first", "last", "axis", "markers"].includes(k)) return null;
    out[k === "markers" ? "high" : k] = /^(1|true|yes|on)$/.test(v);
    if (k === "markers") out.low = out.high;
  }
  return out;
}

// Numbers as a line of blocks, lowest ▁ to highest █ (all alike, ▄): a
// line or a column sparkline in text, the marked points ▲ (high), ▼ (low),
// ◆ (first, last). Win/loss: ▀ a gain, ▄ a loss, a zero blank or, on the
// axis, ─.
export function sparkline(ns: readonly number[], type: SparklineType = "line", o: SparklineOptions = {}): string {
  if (type === "winloss") return ns.map((n) => (n > 0 ? "▀" : n < 0 ? "▄" : o.axis ? "─" : " ")).join("");
  const lo = Math.min(...ns), hi = Math.max(...ns);
  const hiAt = ns.indexOf(hi), loAt = ns.indexOf(lo);
  return ns.map((n, i) => (o.high && i === hiAt ? "▲" : o.low && i === loAt ? "▼" : (o.first && i === 0) || (o.last && i === ns.length - 1) ? "◆"
    : BLOCKS[hi === lo ? 3 : Math.round(((n - lo) / (hi - lo)) * 7)])).join("");
}
