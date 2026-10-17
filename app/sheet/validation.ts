// Data validation: a cell's rule for what may be typed into it, as Excel's
// Data Validation and the Tauri spreadsheet's rules have it. A list of
// values (any case), whole numbers or decimals within limits, or a text
// length; "stop" refuses what breaks it, "warn" keeps it and says so; and,
// from the Tauri app, "create" checks only what is typed into an empty cell,
// "update" only what replaces something. An empty input, and a formula, are
// not checked (Excel checks what is typed, not what a formula comes to).
// Excel's other kinds: a date or a time within limits (a date as a serial
// or as written, a time as a fraction of a day or as written), a custom
// formula (true for what is typed, worked out by the sheet once it is in:
// commands.ts), and the Tauri app's pattern (a regular expression the
// whole text matches). Excel's error styles: stop refuses, warn and info
// keep it and say so (info quieter); an input prompt shown when the cell
// is selected, and an error message of its own.
import { parseDate, parseTime } from "./dates.ts";
import { numberFormat } from "./numfmt.ts";

export interface Rule {
  readonly kind: "list" | "whole" | "decimal" | "length" | "date" | "time" | "custom" | "pattern";
  readonly values?: readonly string[];
  readonly min?: number;
  readonly max?: number;
  readonly formula?: string;
  // The rule's range's top-left, which a custom formula is written for.
  readonly origin?: string;
  readonly pattern?: string;
  readonly mode: "stop" | "warn" | "info";
  readonly when: "always" | "create" | "update";
  readonly prompt?: string;
  readonly error?: string;
}

export const RULE_KINDS = ["list", "whole", "decimal", "length", "date", "time", "custom", "pattern"];

// A date or a time as a number, as it is typed or as a serial.
const dateOf = (s: string): number | null => (NUMBERLIKE.test(s) ? Number(s) : parseDate(s));
const timeOf = (s: string): number | null => (NUMBERLIKE.test(s) ? Number(s) : parseTime(s) ?? (parseDate(s) !== null ? parseDate(s)! % 1 : null));
const NUMBERLIKE = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)\s*$/;

const number = (s: string | undefined, what: string): number | undefined => {
  if (s === undefined || s.trim() === "") return undefined;
  const n = Number(s);
  if (isNaN(n)) throw new RangeError(`${what} ${s} is not a number`);
  return n;
};

// A rule from a command's parameters (kind, values, min, max, formula,
// pattern, mode, when, prompt, error).
export function ruleOf(p: Readonly<Record<string, string | undefined>>): Rule {
  const kind = String(p.kind ?? "").toLowerCase();
  if (!RULE_KINDS.includes(kind)) throw new RangeError(`${p.kind} is not a rule (${RULE_KINDS.join(", ")})`);
  const mode = String(p.mode ?? "stop").toLowerCase(), when = String(p.when ?? "always").toLowerCase();
  if (mode !== "stop" && mode !== "warn" && mode !== "info") throw new RangeError(`${p.mode} is not stop, warn or info`);
  if (!["always", "create", "update"].includes(when)) throw new RangeError(`${p.when} is not always, create or update`);
  const extra = { ...(p.prompt?.trim() ? { prompt: p.prompt.trim() } : {}), ...(p.error?.trim() ? { error: p.error.trim() } : {}) };
  if (kind === "list") {
    const values = String(p.values ?? "").split(",").map((v) => v.trim()).filter((v) => v);
    if (!values.length) throw new RangeError("a list rule needs values");
    return { kind, values, mode, when, ...extra } as Rule;
  }
  if (kind === "custom") {
    const formula = String(p.formula ?? "").trim().replace(/^=/, "");
    if (!formula) throw new RangeError("a custom rule needs a formula");
    return { kind, formula, mode, when, ...extra } as Rule;
  }
  if (kind === "pattern") {
    const pattern = String(p.pattern ?? p.values ?? "");
    if (!pattern) throw new RangeError("a pattern rule needs a pattern");
    try { new RegExp(pattern); } catch { throw new RangeError(`${pattern} is not a pattern`); }
    return { kind, pattern, mode, when, ...extra } as Rule;
  }
  if (kind === "date" || kind === "time") {
    const of = kind === "date" ? dateOf : timeOf;
    const end = (v: string | undefined, what: string) => {
      if (v === undefined || v.trim() === "") return undefined;
      const n = of(v.trim());
      if (n === null) throw new RangeError(`${what} ${v} is not a ${kind}`);
      return n;
    };
    return { kind, min: end(p.min, "min"), max: end(p.max, "max"), mode, when, ...extra } as Rule;
  }
  return { kind, min: number(p.min, "min"), max: number(p.max, "max"), mode, when, ...extra } as Rule;
}

const limits = (r: Rule): string =>
  r.min !== undefined && r.max !== undefined ? `from ${r.min} to ${r.max}` : r.min !== undefined ? `of at least ${r.min}` : r.max !== undefined ? `of at most ${r.max}` : "";

// What the rule asks, as "must be ...".
export function describeRule(r: Rule): string {
  switch (r.kind) {
    case "list": return `must be one of ${r.values!.join(", ")}`;
    case "whole": return `must be a whole number ${limits(r)}`.trim();
    case "decimal": return `must be a number ${limits(r)}`.trim();
    case "length":
      return r.min !== undefined && r.max !== undefined ? `must be ${r.min} to ${r.max} characters long`
        : r.max !== undefined ? `must be at most ${r.max} characters long` : `must be at least ${r.min ?? 0} characters long`;
    case "date": return `must be a date ${dateLimits(r, (n) => numberFormat(n, "yyyy-mm-dd")?.text ?? String(n))}`.trim();
    case "time": return `must be a time ${dateLimits(r, (n) => numberFormat(n, "hh:mm:ss")?.text ?? String(n))}`.trim();
    case "custom": return `must make =${r.formula} true`;
    case "pattern": return `must match ${r.pattern}`;
  }
}

const dateLimits = (r: Rule, shown: (n: number) => string): string =>
  r.min !== undefined && r.max !== undefined ? `from ${shown(r.min)} to ${shown(r.max)}` : r.min !== undefined ? `on or after ${shown(r.min)}` : r.max !== undefined ? `on or before ${shown(r.max)}` : "";

// What a stop says: the rule's own error message, or what it asks.
export const errorOf = (r: Rule): string => r.error ?? describeRule(r);

const within = (n: number, r: Rule) => (r.min === undefined || n >= r.min) && (r.max === undefined || n <= r.max);

// Null when `input` keeps the rule (or the rule does not apply: an empty
// input, a formula, or not when it checks), else why not. `wasEmpty`: the
// cell held nothing before.
export function check(r: Rule, input: string, wasEmpty: boolean): string | null {
  const t = input.trim();
  if (t === "" || t.startsWith("=")) return null;
  if ((r.when === "create" && !wasEmpty) || (r.when === "update" && wasEmpty)) return null;
  let ok: boolean;
  switch (r.kind) {
    case "list": ok = r.values!.some((v) => v.toLowerCase() === t.toLowerCase()); break;
    case "whole": ok = /^[-+]?\d+$/.test(t) && within(Number(t), r); break;
    case "decimal": ok = !isNaN(Number(t)) && within(Number(t), r); break;
    case "length": ok = within(t.length, r); break;
    case "date": { const d = dateOf(t); ok = d !== null && within(d, r); break; }
    case "time": { const d = timeOf(t); ok = d !== null && within(d, r); break; }
    case "pattern": ok = new RegExp(`^(?:${r.pattern})$`).test(t); break;
    // Worked out by the sheet with the input in (commands.ts).
    case "custom": ok = true; break;
  }
  return ok ? null : errorOf(r);
}

// Limits as a person types them: "1..10", "1 to 10", "1-10", ">=0",
// "<= 5", or one number (at least it).
export function limitsOf(text: string): { min?: string; max?: string } {
  const t = text.trim(), n = "(-?\\d+(?:\\.\\d+)?)";
  let m = new RegExp(`^${n}\\s*(?:\\.\\.|to|-)\\s*${n}$`, "i").exec(t);
  if (m) return { min: m[1], max: m[2] };
  m = new RegExp(`^(>=?|<=?)\\s*${n}$`).exec(t);
  if (m) return m[1].startsWith(">") ? { min: m[2] } : { max: m[2] };
  return new RegExp(`^${n}$`).test(t) ? { min: t } : {};
}
