// Conditional-format rules, as Excel's Conditional Formatting menu has them
// (the Tauri app's conditional styles, an expression per cell, are notes.ts's
// styleExpression). A rule covers a range: Highlight Cells (a comparison,
// between, text that contains), Top/Bottom (n cells, or n%), Above/Below
// Average, Duplicate or Unique values, a two- or three-colour scale, an
// icon set. applied() works out what each cell of a sheet gets: a style, or
// an icon before its text, or a data bar (a share of the largest number,
// from the cell's end when negative, as the grid's --fill draws it).
// Excel's other rules too: a formula's (true for the cell, written for the
// range's top-left and moved for each cell as a copy would be), text that
// begins or ends with, or does not contain; blanks and no blanks, errors
// and no errors; a date in a period (yesterday, today, tomorrow, the last
// seven days, last, this or next week or month). A rule may stop the ones
// after it (Stop If True) for the cells it marks.
import { keyOf, range } from "./address.ts";
import type { Addr } from "./address.ts";
import { show } from "./format.ts";
import { now, parseDate, parts, serial } from "./dates.ts";
import { CellError } from "./values.ts";
import type { Value } from "./values.ts";

export type Comparison = ">" | "<" | ">=" | "<=" | "=" | "<>" | "between" | "notBetween" | "contains" | "notContains" | "beginsWith" | "endsWith";

export const PERIODS = ["yesterday", "today", "tomorrow", "last7days", "lastWeek", "thisWeek", "nextWeek", "lastMonth", "thisMonth", "nextMonth"] as const;
export type Period = (typeof PERIODS)[number];
const PERIOD_SAID: Record<Period, string> = { yesterday: "yesterday", today: "today", tomorrow: "tomorrow", last7days: "in the last 7 days", lastWeek: "last week",
  thisWeek: "this week", nextWeek: "next week", lastMonth: "last month", thisMonth: "this month", nextMonth: "next month" };

export type CfRule = (
  | { readonly kind: "cellValue"; readonly op: Comparison; readonly value: string; readonly value2?: string; readonly style: string }
  | { readonly kind: "top"; readonly n: number; readonly bottom: boolean; readonly percent: boolean; readonly style: string }
  | { readonly kind: "average"; readonly below: boolean; readonly style: string }
  | { readonly kind: "duplicates"; readonly unique: boolean; readonly style: string }
  | { readonly kind: "colorScale"; readonly colors: readonly string[] }
  | { readonly kind: "iconSet"; readonly icons: IconSet; readonly thresholds?: readonly [number, number]; readonly by?: Threshold; readonly reverse?: boolean; readonly iconOnly?: boolean }
  | { readonly kind: "dataBar"; readonly color: string; readonly solid?: boolean; readonly barOnly?: boolean; readonly min?: number; readonly max?: number }
  | { readonly kind: "formula"; readonly formula: string; readonly style: string }
  | { readonly kind: "blanks"; readonly not: boolean; readonly style: string }
  | { readonly kind: "errors"; readonly not: boolean; readonly style: string }
  | { readonly kind: "date"; readonly period: Period; readonly style: string }
) & { readonly range: [Addr, Addr]; readonly stop?: boolean };

export const RULE_KINDS = ["cellValue", "top", "average", "duplicates", "colorScale", "iconSet", "dataBar", "formula", "blanks", "errors", "date"];

export const ICONS = { arrows: ["↑", "→", "↓"], triangles: ["▲", "▬", "▼"], ratings: ["●●●", "●●○", "●○○"] } as const;
export type IconSet = keyof typeof ICONS;
// How an icon set's two thresholds read (Excel's Type): a percent of the
// way from the lowest to the highest, a percentile, or the numbers.
export type Threshold = "percent" | "percentile" | "number";
export const THRESHOLDS: readonly Threshold[] = ["percent", "percentile", "number"];

export interface Mark {
  readonly style?: string;
  readonly icon?: string;
  // Show Bar Only, Show Icon Only: the cell's text not shown.
  readonly hide?: boolean;
}

const NUMBER = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/;
const numberOf = (v: Value): number | null => (typeof v === "number" ? v : typeof v === "string" && NUMBER.test(v) ? Number(v) : null);

const hex = (c: string): number[] => [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16));
const mix = (a: string, b: string, t: number): string =>
  `#${hex(a).map((x, i) => Math.round(x + (hex(b)[i] - x) * t).toString(16).padStart(2, "0")).join("")}`;

const median = (ns: readonly number[]): number => {
  const s = [...ns].sort((x, y) => x - y), m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};

function compared(op: Comparison, v: Value, a: string, b?: string): boolean {
  const n = numberOf(v), x = Number(a), y = Number(b);
  const t = v === null || v instanceof CellError ? null : show(v).toLowerCase(), want = a.toLowerCase();
  if (op === "contains") return t !== null && t.includes(want);
  if (op === "notContains") return t === null ? v === null : !t.includes(want);
  if (op === "beginsWith") return t !== null && t.startsWith(want);
  if (op === "endsWith") return t !== null && t.endsWith(want);
  if (op === "between") return n !== null && n >= Math.min(x, y) && n <= Math.max(x, y);
  if (op === "notBetween") return n !== null && !(n >= Math.min(x, y) && n <= Math.max(x, y));
  if (n === null || isNaN(x)) {
    // Text compares as text, any case.
    if (v === null || v instanceof CellError) return false;
    const s = show(v).toLowerCase(), t = a.toLowerCase();
    return op === "=" ? s === t : op === "<>" ? s !== t : false;
  }
  switch (op) {
    case ">": return n > x;
    case "<": return n < x;
    case ">=": return n >= x;
    case "<=": return n <= x;
    case "=": return n === x;
    default: return n !== x;
  }
}

// Whether a value is a date in a period, today being `today` (a serial):
// weeks run Sunday to Saturday, as Excel's do.
export function inPeriod(v: Value, period: Period, today: number): boolean {
  const d = typeof v === "number" ? Math.floor(v) : typeof v === "string" ? parseDate(v) : null;
  if (d === null) return false;
  const day = Math.floor(d), sunday = today - parts(today).weekday;
  const month = (by: number) => {
    const p = parts(today), m0 = p.m - 1 + by, y = p.y + Math.floor(m0 / 12), m = ((m0 % 12) + 12) % 12 + 1;
    const first = serial(y, m, 1), next = serial(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1);
    return day >= first && day < next;
  };
  switch (period) {
    case "yesterday": return day === today - 1;
    case "today": return day === today;
    case "tomorrow": return day === today + 1;
    case "last7days": return day > today - 7 && day <= today;
    case "lastWeek": return day >= sunday - 7 && day < sunday;
    case "thisWeek": return day >= sunday && day < sunday + 7;
    case "nextWeek": return day >= sunday + 7 && day < sunday + 14;
    case "lastMonth": return month(-1);
    case "thisMonth": return month(0);
    case "nextMonth": return month(1);
  }
}

// What one rule gives the cells of its range; a formula rule's cells are
// the ones its formula came true for (`hits`, worked out by the sheet).
function marked(rule: CfRule, value: (key: string) => Value, hits: ReadonlySet<string>, today: number): [string, Mark][] {
  const keys = [...range(rule.range[0], rule.range[1])];
  const numbers = keys.map((k) => numberOf(value(k)));
  const ns = numbers.filter((n): n is number => n !== null);
  const styled = (pick: (k: string, i: number) => boolean, style: string): [string, Mark][] =>
    keys.filter(pick).map((k) => [k, { style }]);
  switch (rule.kind) {
    case "cellValue": return styled((k) => compared(rule.op, value(k), rule.value, rule.value2), rule.style);
    case "formula": return styled((k) => hits.has(k), rule.style);
    case "blanks": return styled((k) => (value(k) === null || value(k) === "") !== rule.not, rule.style);
    case "errors": return styled((k) => value(k) instanceof CellError !== rule.not, rule.style);
    case "date": return styled((k) => inPeriod(value(k), rule.period, today), rule.style);
    case "top": {
      const count = Math.max(1, rule.percent ? Math.floor((ns.length * rule.n) / 100) : rule.n);
      const sorted = [...ns].sort((x, y) => (rule.bottom ? x - y : y - x));
      const edge = sorted[Math.min(count, sorted.length) - 1];
      return styled((_, i) => numbers[i] !== null && (rule.bottom ? numbers[i]! <= edge : numbers[i]! >= edge), rule.style);
    }
    case "average": {
      const mean = ns.reduce((s, n) => s + n, 0) / ns.length;
      return styled((_, i) => numbers[i] !== null && (rule.below ? numbers[i]! < mean : numbers[i]! > mean), rule.style);
    }
    case "duplicates": {
      const text = (k: string) => show(value(k)).toLowerCase();
      const counts = new Map<string, number>();
      for (const k of keys) if (text(k) !== "") counts.set(text(k), (counts.get(text(k)) ?? 0) + 1);
      return styled((k) => text(k) !== "" && (counts.get(text(k))! > 1) !== rule.unique, rule.style);
    }
    case "colorScale": {
      if (!ns.length) return [];
      const lo = Math.min(...ns), hi = Math.max(...ns), mid = median(ns);
      const [a, b, c] = rule.colors;
      const at = (n: number): string => {
        if (!c) return mix(a, b, hi === lo ? 0 : (n - lo) / (hi - lo));
        return n <= mid ? mix(a, b, mid === lo ? 0 : (n - lo) / (mid - lo)) : mix(b, c, hi === mid ? 1 : (n - mid) / (hi - mid));
      };
      return keys.flatMap((k, i): [string, Mark][] => (numbers[i] === null ? [] : [[k, { style: `background: ${at(numbers[i]!)}` }]]));
    }
    case "dataBar": {
      // A share of the largest size (or of min to max when they are set),
      // from the cell's end when negative.
      const most = Math.max(...ns.map(Math.abs));
      const fixed = rule.min !== undefined || rule.max !== undefined;
      const lo = rule.min ?? (fixed ? Math.min(0, ...ns) : 0), hi = rule.max ?? (fixed ? Math.max(...ns) : most);
      if (!fixed && !most) return [];
      const share = (n: number) => (fixed ? (hi === lo ? 1 : Math.min(1, Math.max(0, (n - lo) / (hi - lo)))) : n / most);
      return keys.flatMap((k, i): [string, Mark][] => (numbers[i] === null ? []
        : [[k, { style: `--fill: ${Number(share(numbers[i]!).toFixed(4))}; --fill-color: ${rule.color}`, ...(rule.barOnly ? { hide: true } : {}) }]]));
    }
    case "iconSet": {
      if (!ns.length) return [];
      const lo = Math.min(...ns), hi = Math.max(...ns), [up, level, down] = rule.reverse ? [...ICONS[rule.icons]].reverse() : ICONS[rule.icons];
      const [high, mid] = rule.thresholds ?? [67, 33], by = rule.by ?? "percent";
      const sorted = [...ns].sort((x, y) => x - y);
      // A percentile as PERCENTILE.INC works it out.
      const percentile = (p: number) => {
        const at = (sorted.length - 1) * Math.min(1, Math.max(0, p / 100)), j = Math.floor(at);
        return sorted[j] + (sorted[Math.min(j + 1, sorted.length - 1)] - sorted[j]) * (at - j);
      };
      const cut = (t: number) => (by === "number" ? t : by === "percentile" ? percentile(t) : lo + ((hi - lo) * t) / 100);
      const [a, b] = [cut(high), cut(mid)];
      return keys.flatMap((k, i): [string, Mark][] => {
        const n = numbers[i];
        if (n === null) return [];
        const icon = hi === lo && by === "percent" ? up : n >= a ? up : n >= b ? level : down;
        return [[k, { icon, ...(rule.iconOnly ? { hide: true } : {}) }]];
      });
    }
  }
}

// Every rule's marks, by cell; where two rules mark a cell, the first rule
// listed wins (its style comes last, and its icon is the one shown), and a
// rule that stops if true leaves the cells it marks to no rule after it.
// `hits` holds each formula rule's cells that came true.
export function applied(rules: readonly CfRule[], value: (key: string) => Value, hits: ReadonlyMap<CfRule, ReadonlySet<string>> = new Map(), today = Math.floor(now())): Map<string, Mark> {
  const out = new Map<string, Mark>(), stopped = new Set<string>();
  for (const rule of rules) {
    const marks = marked(rule, value, hits.get(rule) ?? new Set(), today).filter(([k]) => !stopped.has(k));
    for (const [k, m] of marks) {
      const was = out.get(k) ?? {};
      out.set(k, { ...(m.style || was.style ? { style: [m.style, was.style].filter((x) => x).join("; ") } : {}), ...(was.icon || m.icon ? { icon: was.icon ?? m.icon } : {}),
        ...(was.hide || m.hide ? { hide: true } : {}) });
    }
    if (rule.stop) for (const [k] of marks) stopped.add(k);
  }
  return out;
}

export const rangeName = ([a, b]: readonly [Addr, Addr]): string => (keyOf(a) === keyOf(b) ? keyOf(a) : `${keyOf(a)}:${keyOf(b)}`);

// What a rule does, as Manage Rules lists it.
export function describeRule(rule: CfRule): string {
  const what = (() => {
    switch (rule.kind) {
      case "cellValue":
        return rule.op === "between" || rule.op === "notBetween" ? `cell value ${rule.op === "between" ? "between" : "not between"} ${rule.value} and ${rule.value2}`
          : rule.op === "contains" ? `text contains ${rule.value}` : rule.op === "notContains" ? `text does not contain ${rule.value}`
          : rule.op === "beginsWith" ? `text begins with ${rule.value}` : rule.op === "endsWith" ? `text ends with ${rule.value}` : `cell value ${rule.op} ${rule.value}`;
      case "formula": return `formula =${rule.formula.replace(/^=/, "")}`;
      case "blanks": return rule.not ? "no blanks" : "blanks";
      case "errors": return rule.not ? "no errors" : "errors";
      case "date": return `a date ${PERIOD_SAID[rule.period]}`;
      case "top": return `${rule.bottom ? "bottom" : "top"} ${rule.n}${rule.percent ? "%" : ""}`;
      case "average": return `${rule.below ? "below" : "above"} average`;
      case "duplicates": return `${rule.unique ? "unique" : "duplicate"} values`;
      case "colorScale": return `colour scale ${rule.colors.join(", ")}`;
      case "iconSet": return `icon set ${rule.icons}${rule.reverse ? " reversed" : ""}${rule.thresholds ? ` at ${rule.thresholds.join(", ")}${(rule.by ?? "percent") === "percent" ? "%" : ` (${rule.by})`}` : ""}${rule.iconOnly ? ", icons only" : ""}`;
      case "dataBar": return `data bars ${rule.color}${rule.solid ? " solid" : ""}${rule.min !== undefined || rule.max !== undefined ? ` from ${rule.min ?? "automatic"} to ${rule.max ?? "automatic"}` : ""}${rule.barOnly ? ", bar only" : ""}`;
    }
  })();
  return `${rangeName(rule.range)}: ${what}${rule.stop ? " (stop if true)" : ""}`;
}
