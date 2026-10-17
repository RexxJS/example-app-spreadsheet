// The Tauri spreadsheet's control functions (its src/spreadsheet-control-
// functions.js, at RexxJS/example-app-spreadsheet 45b0993) by name, each as
// demo 18's commands: SETFORMAT("A1", "bold") runs cell.format cell=A1
// format=bold, and gives back what the app's gives. A script (script.ts)
// has them all, so the app's REXX scripts run against demo 18 as they are.
// The functions see the sheet only through `exec`, a
// command and its parameters in and its outcome out, as a panel does.
import { COLS, ROWS, addrOf, colName, keyOf, parseRange, range } from "./address.ts";
import { check } from "./validation.ts";
import type { Rule } from "./validation.ts";
import type { Addr } from "./address.ts";
import { DESCRIBED, signature } from "./described.ts";

export interface Done {
  readonly ok: boolean;
  readonly message: string;
  readonly output?: unknown;
}
export type Exec = (id: string, params: Record<string, string>) => Promise<Done>;
type Control = (exec: Exec, ...args: any[]) => unknown;

// What cell.about and sheet.state give (commands.ts).
interface About {
  readonly value: unknown;
  readonly input: string;
  readonly format: string;
  readonly comment: string;
  readonly wrap: number;
  readonly merged: string;
  readonly validation: Rule | null;
  readonly editor: { readonly type: string; readonly config: object } | null;
}
interface State {
  readonly name: string;
  readonly sheets: readonly string[];
  readonly hiddenRows: readonly number[];
  readonly hiddenCols: readonly string[];
  readonly frozenRows: number;
  readonly frozenCols: number;
  readonly rowHeights: readonly (readonly [number, number])[];
  readonly colWidths: readonly (readonly [string, number])[];
  readonly filter: { readonly range: string; readonly column: string; readonly criterion: string } | null;
  readonly canUndo: number;
  readonly canRedo: number;
  readonly autoId: { readonly column: string; readonly next: string } | null;
  readonly language: string;
  readonly referenceStyle: string;
  readonly panel: string;
  readonly split: number;
  readonly zoom: number;
  readonly hiddenSheets: readonly string[];
  readonly tabColors: Readonly<Record<string, string>>;
  readonly dateSystem: number;
  readonly calculation: string;
  readonly iteration: { readonly max: number; readonly change: number } | null;
  readonly stale: readonly string[];
}

// A command that must be done: a refusal is the function's error, as the
// app's functions throw theirs.
async function must(exec: Exec, id: string, params: Record<string, string> = {}): Promise<Done> {
  const o = await exec(id, params);
  if (!o.ok) throw new Error(o.message);
  return o;
}
const about = async (exec: Exec, cell: unknown) => (await must(exec, "cell.about", { cell: text(cell) })).output as About;
const state = async (exec: Exec) => (await must(exec, "sheet.state")).output as State;

// The app's arguments: text, 1 for yes, rows counted from 1, a column by
// its letter or its number.
const text = (a: unknown) => (a === undefined || a === null ? "" : String(a));
const yes = (a: unknown) => a === 1 || a === "1" || a === true || a === "true";
function row(a: unknown, fn: string): number {
  const n = parseInt(text(a), 10);
  if (!(n >= 1)) throw new Error(`${fn} needs a row number (e.g. 2)`);
  return n;
}
function col(a: unknown, fn: string): string {
  const t = text(a).trim();
  const n = /^\d+$/.test(t) ? parseInt(t, 10) : NaN;
  if (n >= 1) return colName(n - 1);
  if (/^[A-Za-z]{1,3}$/.test(t)) return t.toUpperCase();
  throw new Error(`${fn} needs a column, a letter or a number (e.g. "B" or 2)`);
}
function cell(a: unknown, fn: string) {
  const at = addrOf(text(a).trim());
  if (!at) throw new Error(`${fn} needs a cell reference (e.g. "A1")`);
  return at;
}
// A REXX stem: {0: count, 1: first, ...}.
// "a=1;b=two" as { a: "1", b: "two" }.
const pairs = (a: unknown): Record<string, string> => Object.fromEntries(text(a).split(";").map((x) => x.trim()).filter((x) => x.includes("=")).map((x) => [x.slice(0, x.indexOf("=")).trim(), x.slice(x.indexOf("=") + 1).trim()]));
const stem = (xs: readonly unknown[]): Record<string, unknown> => Object.fromEntries([[0, xs.length], ...xs.map((x, i) => [i + 1, x])]);
const unstem = (v: unknown): unknown[] => {
  if (Array.isArray(v)) return v;
  if (typeof v === "string") return JSON.parse(v) as unknown[];
  const s = v as Record<string, unknown>;
  return Array.from({ length: Number(s?.[0] ?? 0) }, (_, i) => s[i + 1]);
};
// The range a source and a target span, as the app's FILLDOWN takes them
// apart ("A1" and "A2:A5" is A1:A5), and how many lines of it are the
// source, repeated into the rest.
function spanned(a: string, b: string, fn: string, across: boolean): { range: string; source: string } {
  const r = [parseRange(a), parseRange(b)];
  if (!r[0] || !r[1]) throw new Error(`${fn} needs a source and a target range (e.g. "A1", "A2:A10")`);
  const all = [...r[0], ...r[1]], cols = all.map((x) => x.col), rows = all.map((x) => x.row);
  const lines = across ? Math.abs(r[0][0].col - r[0][1].col) + 1 : Math.abs(r[0][0].row - r[0][1].row) + 1;
  return { range: `${keyOf({ col: Math.min(...cols), row: Math.min(...rows) })}:${keyOf({ col: Math.max(...cols), row: Math.max(...rows) })}`, source: String(lines) };
}
const rowsOf = (n: number) => `A${n}`;
const VERSION = "sae demo 18";

export const CONTROL: Readonly<Record<string, Control>> = {
  SPREADSHEET_VERSION: () => VERSION,
  // Not the app's (its sheets are RexxJS formulas only): the formula
  // language, "Excel" or "RexxJS", as the Formula language setting.
  SETREFERENCESTYLE: async (exec, s) => (await must(exec, "formulas.style", { style: text(s) }), (await state(exec)).referenceStyle),
  GETREFERENCESTYLE: async (exec) => (await state(exec)).referenceStyle,
  SETFORMULALANGUAGE: async (exec, l) => (await must(exec, "formulas.use", { language: text(l) }), (await state(exec)).language),
  GETFORMULALANGUAGE: async (exec) => (await state(exec)).language,
  EXPORT: async (exec) => (await must(exec, "file.export")).output,
  IMPORT: async (exec, json) => (await must(exec, "file.import", { text: text(json) }), "OK"),
  EVALUATE: async (exec, formula) => (await must(exec, "formulas.calculate", { formula: text(formula) })).output,
  RECALCULATE: async (exec) => (await must(exec, "formulas.recalculate")).output,
  // Calculation Options: "Automatic" or "Manual"; Calculate Now and
  // Calculate Sheet; the cells out of date; iterative calculation (0 off).
  SETCALCULATION: async (exec, m) => (await must(exec, "formulas.calculation", { mode: text(m) }), (await state(exec)).calculation),
  GETCALCULATION: async (exec) => (await state(exec)).calculation,
  // The clipboard: a range copied or cut, pasted at a cell (formulas moved).
  COPY: async (exec, r) => (await must(exec, "edit.copy", { range: text(r) })).message,
  CUT: async (exec, r) => (await must(exec, "edit.cut", { range: text(r) })).message,
  PASTE: async (exec, at) => (await must(exec, "edit.paste", { at: text(at) })).message,
  // Paste Special at a cell: what (all, formulas, values, formats,
  // comments), an operation (add, subtract, multiply, divide), skip blanks,
  // transpose; what was copied (COPY, or the page's Copy).
  PASTESPECIAL: async (exec, at, what, op, skip, turn) => (await must(exec, "edit.pasteSpecial", {
    at: text(at), what: text(what) || "all", ...(text(op).trim() ? { operation: text(op).trim() } : {}), ...(yes(skip) ? { skipBlanks: "1" } : {}), ...(yes(turn) ? { transpose: "1" } : {}),
  })).message,
  // The status bar's figures for a range: its line, or one figure (sum,
  // average, count, numbers, min, max).
  SELECTIONSTATS: async (exec, r, which) => {
    const o = await must(exec, "range.stats", { range: text(r) });
    const w = text(which).trim().toLowerCase();
    if (!w) return o.message;
    const st = o.output as Record<string, unknown>;
    if (!(w in st)) throw new Error(`SELECTIONSTATS: ${which} is not a figure (sum, average, count, numbers, min, max)`);
    return st[w] ?? "";
  },
  // Split: SPLITWINDOW(1) two panes, SPLITWINDOW(0) one; ISSPLIT().
  SPLITWINDOW: async (exec, on) => (await must(exec, "view.split", text(on).trim() === "0" ? { off: "1" } : {})).output,
  ISSPLIT: async (exec) => (await state(exec)).split,
  // The panel open under the grid: SHOWPANEL("rules"), "none" to close.
  SHOWPANEL: async (exec, n) => (await must(exec, "view.panel", { name: text(n) })).output,
  GETPANEL: async (exec) => (await state(exec)).panel,
  // Help while typing: F4's next form of the reference at the end, and the
  // function names completing a formula (a stem), with the call it is in.
  CYCLEREFERENCE: async (exec, t) => (await must(exec, "formulas.cycleRef", { text: text(t) })).output,
  COMPLETEFORMULA: async (exec, t) => {
    const c = (await must(exec, "formulas.complete", { text: text(t) })).output as { names: string[]; call: string | null; arg: number };
    return { ...stem(c.names), CALL: c.call ?? "", ARG: c.arg };
  },
  // The menu bar's items: a stem of "File > Save | file.save | Ctrl+S".
  LISTMENUS: async (exec) => stem((await must(exec, "menus.list")).output as string[]),
  // Error Checking: a stem of "cell check message" lines (check: error,
  // inconsistent, text-number, omits); Ignore Error, and checked again.
  ERRORCHECK: async (exec) => stem(((await must(exec, "formulas.errorCheck")).output as { cell: string; check: string; message: string }[]).map((f) => `${f.cell} ${f.check} ${f.message}`)),
  IGNOREERROR: async (exec, c, off) => (await must(exec, "formulas.ignoreError", { cell: text(c), ...(yes(off) ? { off: "1" } : {}) }), ""),
  CALCULATE: async (exec) => (await must(exec, "formulas.recalculate")).output,
  CALCULATESHEET: async (exec) => (await must(exec, "formulas.calculateSheet")).output,
  STALECELLS: async (exec) => stem((await state(exec)).stale),
  SETITERATION: async (exec, max, change) => {
    const n = text(max).trim();
    await must(exec, "formulas.iteration", n === "0" ? { off: "1" } : { ...(n ? { max: n } : {}), ...(text(change).trim() ? { change: text(change).trim() } : {}) });
    return (await state(exec)).iteration?.max ?? 0;
  },
  GETITERATION: async (exec) => {
    const it = (await state(exec)).iteration;
    return it ? `${it.max} ${it.change}` : "0";
  },
  GETSETUPSCRIPT: async (exec) => (await must(exec, "setup.get")).output,
  SETSETUPSCRIPT: async (exec, script) => (await must(exec, "setup.set", { text: text(script) }), "OK"),
  EXECUTESETUPSCRIPT: async (exec) => (await must(exec, "setup.run"), "OK"),
  // [{address, value}, ...]: each set that can be; a summary, as the app's.
  BATCH_SET_CELLS: async (exec, updates) => {
    const errors: { address: string; error: string }[] = [];
    const list = unstem(updates) as { address?: unknown; value?: unknown }[];
    for (const u of list) {
      const o = await exec("cell.set", { cell: text(u.address), input: text(u.value) });
      if (!o.ok) errors.push({ address: text(u.address), error: o.message });
    }
    return JSON.stringify({ total: list.length, success: list.length - errors.length, errors: errors.length, ...(errors.length ? { errorDetails: errors } : {}) });
  },
  GETEXPRESSION: async (exec, c) => {
    const input = (await about(exec, c)).input;
    return input.startsWith("=") ? input.slice(1) : "";
  },
  CLEARCELL: async (exec, c) => (await must(exec, "range.clear", { range: text(c) }), ""),
  SETFORMAT: async (exec, c, f) => (await must(exec, "cell.format", { cell: text(c), format: text(f) }), text(f)),
  GETFORMAT: async (exec, c) => (await about(exec, c)).format,
  SETCOMMENT: async (exec, c, t) => (await must(exec, "cell.comment", { cell: text(c), text: text(t) }), text(t)),
  GETCOMMENT: async (exec, c) => (await about(exec, c)).comment,
  GETROW: (_, c) => cell(c, "GETROW").row + 1,
  GETCOL: (_, c) => cell(c, "GETCOL").col + 1,
  GETCOLNAME: (_, c) => colName(cell(c, "GETCOLNAME").col),
  MAKECELLREF: (_, c, r) => `${col(c, "MAKECELLREF")}${row(r, "MAKECELLREF")}`,
  GETCELLS: async (exec, r) => {
    const span = parseRange(text(r));
    if (!span) throw new Error('GETCELLS needs a range (e.g. "A1:B5")');
    const cells: unknown[] = [];
    for (const k of range(span[0], span[1])) {
      const a = await about(exec, k);
      cells.push({ ref: k, value: a.value, expression: a.input.startsWith("=") ? a.input.slice(1) : "", comment: a.comment, format: a.format });
    }
    return stem(cells);
  },
  SETCELLS: async (exec, r, values) => {
    const span = parseRange(text(r));
    if (!span) throw new Error('SETCELLS needs a range (e.g. "A1:A3")');
    const vs = unstem(values);
    let i = 0;
    for (const k of range(span[0], span[1])) if (i < vs.length) await must(exec, "cell.set", { cell: k, input: text(vs[i++]) });
    return "OK";
  },
  CLEAR: async (exec) => (await must(exec, "sheet.clear"), "OK"),
  GETSHEETNAME: async (exec) => (await state(exec)).name,
  INSERTROW: async (exec, n) => (await must(exec, "rows.insert", { at: String(row(n, "INSERTROW")) }), "OK"),
  DELETEROW: async (exec, n) => (await must(exec, "rows.delete", { at: String(row(n, "DELETEROW")) }), "OK"),
  INSERTCOLUMN: async (exec, c) => (await must(exec, "columns.insert", { at: col(c, "INSERTCOLUMN") }), "OK"),
  DELETECOLUMN: async (exec, c) => (await must(exec, "columns.delete", { at: col(c, "DELETECOLUMN") }), "OK"),
  FILLDOWN: async (exec, s, t) => (await must(exec, "edit.fillDown", spanned(text(s), text(t), "FILLDOWN", false)), "OK"),
  FILLRIGHT: async (exec, s, t) => (await must(exec, "edit.fillRight", spanned(text(s), text(t), "FILLRIGHT", true)), "OK"),
  FILLUP: async (exec, s, t) => (await must(exec, "edit.fillUp", spanned(text(s), text(t), "FILLUP", false)), "OK"),
  FILLLEFT: async (exec, s, t) => (await must(exec, "edit.fillLeft", spanned(text(s), text(t), "FILLLEFT", true)), "OK"),
  // The way the target lies from the source, as the app's: below it, down.
  AUTOFILL: async (exec, s, t) => {
    const [src, to] = [parseRange(text(s)), parseRange(text(t))];
    if (!src || !to) throw new Error('AUTOFILL needs a source and a target range (e.g. "A1:A3", "A4:A10")');
    const lo = (r: [Addr, Addr], k: "row" | "col") => Math.min(r[0][k], r[1][k]), hi = (r: [Addr, Addr], k: "row" | "col") => Math.max(r[0][k], r[1][k]);
    const way = lo(to, "row") > hi(src, "row") ? "Down" : hi(to, "row") < lo(src, "row") ? "Up" : lo(to, "col") > hi(src, "col") ? "Right" : hi(to, "col") < lo(src, "col") ? "Left" : null;
    if (!way) throw new Error("AUTOFILL: the target must lie beside the source");
    return (await must(exec, `edit.fill${way}`, spanned(text(s), text(t), "AUTOFILL", way === "Right" || way === "Left")), "");
  },
  // SORTRANGE("A1:C9", "B") by one column; ("A1:C9", "B desc, C asc")
  // or a stem of such keys (k.1 = "B desc"), several levels; a key may be
  // "C fill=#ffff00", that cell colour on top; headers 1 keeps the first row.
  SORTRANGE: async (exec, r, by, ascending, headers) => {
    const keys = typeof by === "object" && by !== null ? unstem(by).map(text).join(", ") : text(by);
    if (/[,\s=]/.test(keys.trim())) return (await must(exec, "data.sort", { range: text(r), keys, ...(yes(headers) ? { headers: "1" } : {}) }), "OK");
    const id = ascending === undefined || text(ascending).trim() === "" || yes(ascending) ? "data.sortAscending" : "data.sortDescending";
    return (await must(exec, id, { range: text(r), by: col(by, "SORTRANGE") }), "OK");
  },
  // AutoFit: AUTOFIT("B") a column (or "B:D"), AUTOFIT(3) a row (or "3:5").
  AUTOFIT: async (exec, what) => {
    const t = text(what).trim();
    const rows = /^\d+(:\d+)?$/.test(t), parts = t.split(":");
    if (!rows && !/^[A-Za-z]{1,3}(:[A-Za-z]{1,3})?$/.test(t)) throw new Error(`AUTOFIT needs a column ("B", "B:D") or a row (3, "3:5")`);
    const r = rows ? `A${parts[0]}:A${parts[1] ?? parts[0]}` : `${parts[0]}1:${parts[1] ?? parts[0]}1`;
    return (await must(exec, rows ? "rows.autofit" : "columns.autofit", { range: r })).message;
  },
  // A sheet's tab colour, hiding, and a copy of it (with its cells and looks).
  SETSHEETCOLOR: async (exec, n, c) => (await must(exec, "sheet.color", { name: text(n), color: text(c) })).message,
  HIDESHEET: async (exec, n, off) => (await must(exec, "sheet.hide", { name: text(n), ...(yes(off) ? { off: "1" } : {}) })).message,
  UNHIDESHEET: async (exec, n) => (await must(exec, "sheet.hide", { name: text(n), off: "1" })).message,
  ISSHEETHIDDEN: async (exec, n) => ((await state(exec)).hiddenSheets.some((x) => x.toUpperCase() === text(n).toUpperCase()) ? 1 : 0),
  COPYSHEET: async (exec, n, to) => (await must(exec, "sheet.copy", { name: text(n), ...(text(to).trim() ? { to: text(to) } : {}) })).message,
  // The date system: SETDATESYSTEM(1904), GETDATESYSTEM().
  SETDATESYSTEM: async (exec, sys) => (await must(exec, "workbook.dateSystem", { system: text(sys).trim() }), (await state(exec)).dateSystem),
  GETDATESYSTEM: async (exec) => (await state(exec)).dateSystem,
  // Zoom: SETZOOM(150), SETZOOM("in"), SETZOOM("out"); GETZOOM().
  SETZOOM: async (exec, z) => (await must(exec, "view.zoom", /^(in|out)$/i.test(text(z).trim()) ? { by: text(z).trim().toLowerCase() } : { percent: text(z).trim() })).output,
  GETZOOM: async (exec) => (await state(exec)).zoom,
  // The selection: SELECTRANGE("B:B"), ("3:5"), ("all"), GETSELECTION().
  SELECTRANGE: async (exec, r) => (await must(exec, "view.select", { range: text(r) })).output,
  GETSELECTION: async (exec) => (await must(exec, "view.selection")).output,
  // Go To Special: a stem of the cells of a kind (GOTOSPECIAL("blanks"),
  // GOTOSPECIAL("constants", "B2:D9", "numbers")).
  GOTOSPECIAL: async (exec, what, r, type) => stem((await must(exec, "edit.goToSpecial", { what: text(what), ...(text(r).trim() ? { range: text(r) } : {}), ...(text(type).trim() ? { type: text(type) } : {}) })).output as string[]),
  // Ctrl+End's cell, and the keyboard map ("Ctrl+C | copy | edit.copy").
  GETLASTCELL: async (exec) => (await must(exec, "sheet.lastCell")).output,
  LISTKEYS: async (exec) => stem((await must(exec, "keys.list")).output as string[]),
  // Flash Fill: FLASHFILL("B1:B9"), its first cells the examples.
  FLASHFILL: async (exec, r) => (await must(exec, "data.flashFill", { range: text(r) })).message,
  // Format Painter: COPYFORMAT("A1", "B1:B9").
  COPYFORMAT: async (exec, from, to) => (await must(exec, "format.painter", { from: text(from), to: text(to) })).message,
  // The toolbar's looks over a range.
  SETALIGNMENT: async (exec, r, how) => (await must(exec, "format.align", { range: text(r), align: text(how) }), "OK"),
  SETFILLCOLOR: async (exec, r, c) => (await must(exec, "format.fill", { range: text(r), color: text(c) }), "OK"),
  SETTEXTCOLOR: async (exec, r, c) => (await must(exec, "format.color", { range: text(r), color: text(c) }), "OK"),
  // Conditional formatting: ADDRULE("A1:A9", "formula", "=A1>B1",
  // "bold;color:red"), ("B1:B9", "greater", 10), ("C1:C9", "between", 1,
  // style, 5), ("D1:D9", "blanks"), ("E1:E9", "date", "thisWeek"),
  // ("F1:F9", "top", 3) ...; LISTRULES(), DELETERULE(n), MOVERULE(n, to),
  // STOPIFTRUE(n, [0]), CLEARRULES(range).
  ADDRULE: async (exec, r, kind, arg, style, arg2) => {
    const k = text(kind).trim().toLowerCase(), a = text(arg), range = text(r);
    const st = text(style).trim() ? { style: text(style) } : {};
    const run = (id: string, ps: Record<string, string>) => must(exec, id, { range, ...ps, ...st });
    if (k === "databar" || k === "iconset") return (await must(exec, k === "databar" ? "format.dataBar" : "format.iconSet", { range, ...(a.trim() ? { [k === "databar" ? "color" : "icons"]: a } : {}), ...pairs(style) })).message;
    const done = await (
      k === "formula" ? run("format.formulaRule", { formula: a })
      : k === "blanks" || k === "noblanks" ? run("format.blanks", k === "noblanks" ? { not: "1" } : {})
      : k === "errors" || k === "noerrors" ? run("format.errors", k === "noerrors" ? { not: "1" } : {})
      : k === "date" ? run("format.datePeriod", { period: a })
      : /^(top|bottom)(percent)?$/.test(k) ? run("format.topBottom", { n: a || "10", ...(k.startsWith("bottom") ? { bottom: "1" } : {}), ...(k.endsWith("percent") ? { percent: "1" } : {}) })
      : k === "above" || k === "below" ? run("format.aboveAverage", k === "below" ? { below: "1" } : {})
      : k === "duplicates" || k === "unique" ? run("format.duplicates", k === "unique" ? { unique: "1" } : {})
      : run("format.highlight", { op: k, value: a, ...(text(arg2).trim() ? { value2: text(arg2) } : {}) }));
    return done.message;
  },
  // The outline: GROUPROWS(2, 9), UNGROUPROWS(2, 9), GROUPCOLUMNS("B",
  // "D"), UNGROUPCOLUMNS; SHOWOUTLINELEVEL(1), (2, "columns"); HIDEDETAIL(3),
  // SHOWDETAIL(3) (a row, or a column's letter); CLEAROUTLINE(); and
  // SUBTOTALS("A1:C20", "A", "C", "sum").
  GROUPROWS: async (exec, a, b) => (await must(exec, "outline.group", { range: `${row(a, "GROUPROWS")}:${b === undefined || text(b).trim() === "" ? row(a, "GROUPROWS") : row(b, "GROUPROWS")}`, axis: "rows" })).message,
  UNGROUPROWS: async (exec, a, b) => (await must(exec, "outline.ungroup", { range: `${row(a, "UNGROUPROWS")}:${b === undefined || text(b).trim() === "" ? row(a, "UNGROUPROWS") : row(b, "UNGROUPROWS")}`, axis: "rows" })).message,
  GROUPCOLUMNS: async (exec, a, b) => (await must(exec, "outline.group", { range: `${col(a, "GROUPCOLUMNS")}:${b === undefined || text(b).trim() === "" ? col(a, "GROUPCOLUMNS") : col(b, "GROUPCOLUMNS")}`, axis: "columns" })).message,
  UNGROUPCOLUMNS: async (exec, a, b) => (await must(exec, "outline.ungroup", { range: `${col(a, "UNGROUPCOLUMNS")}:${b === undefined || text(b).trim() === "" ? col(a, "UNGROUPCOLUMNS") : col(b, "UNGROUPCOLUMNS")}`, axis: "columns" })).message,
  SHOWOUTLINELEVEL: async (exec, n, axis) => (await must(exec, "outline.level", { level: text(n), ...(text(axis).trim() ? { axis: text(axis) } : {}) })).message,
  HIDEDETAIL: async (exec, at) => (await must(exec, "outline.hideDetail", /^\d+$/.test(text(at).trim()) ? { range: `${text(at).trim()}:${text(at).trim()}`, axis: "rows" } : { range: `${col(at, "HIDEDETAIL")}:${col(at, "HIDEDETAIL")}`, axis: "columns" })).message,
  SHOWDETAIL: async (exec, at) => (await must(exec, "outline.showDetail", /^\d+$/.test(text(at).trim()) ? { range: `${text(at).trim()}:${text(at).trim()}`, axis: "rows" } : { range: `${col(at, "SHOWDETAIL")}:${col(at, "SHOWDETAIL")}`, axis: "columns" })).message,
  CLEAROUTLINE: async (exec, axis) => (await must(exec, "outline.clear", text(axis).trim() ? { axis: text(axis) } : {})).message,
  SUBTOTALS: async (exec, r, by, cols, fn) => (await must(exec, "data.subtotal", { range: text(r), by: text(by), columns: text(cols), ...(text(fn).trim() ? { function: text(fn) } : {}) })).message,
  // CSV and TSV: IMPORTCSV("A1", text) (its delimiter detected), ("A1",
  // text, "tab"), ("A1", base64, "auto", "windows-1252") bytes in an
  // encoding; EXPORTCSV("A1:C9", "tab").
  IMPORTCSV: async (exec, at, t, d, enc) => (await must(exec, "file.importCsv", { at: text(at), ...(text(enc).trim() ? { base64: text(t), encoding: text(enc) } : { text: text(t) }), ...(text(d).trim() ? { delimiter: text(d) } : {}) })).message,
  EXPORTCSV: async (exec, r, d) => (await must(exec, "file.exportCsv", { ...(text(r).trim() ? { range: text(r) } : {}), ...(text(d).trim() ? { delimiter: text(d) } : {}) })).output,
  // XLSX: OPENXLSX(base64, "Book.xlsx") opens one; FILEREPORT() a stem of
  // what it held that the app left out.
  OPENXLSX: async (exec, b64, n, pw) => (await must(exec, "file.importXlsx", { base64: text(b64), ...(text(n).trim() ? { name: text(n) } : {}), ...(text(pw) ? { password: text(pw) } : {}) })).message,
  // OPENXLS(base64, "Book.xls"): an Excel 97-2003 workbook.
  OPENXLS: async (exec, b64, n) => (await must(exec, "file.importXls", { base64: text(b64), ...(text(n).trim() ? { name: text(n) } : {}) })).message,
  // OPENODS(base64, "Book.ods"): an ODS file's tables, their cached values.
  OPENODS: async (exec, b64, n) => (await must(exec, "file.importOds", { base64: text(b64), ...(text(n).trim() ? { name: text(n) } : {}) })).message,
  // EXPORTPDF("A1:F40", "paper=letter;landscape=1;gridlines=1;headings=1"),
  // EXPORTPDF("", "sheets=all"): the PDF's bytes in base64.
  EXPORTPDF: async (exec, r, opts) => (await must(exec, "file.exportPdf", { ...(text(r).trim() ? { range: text(r) } : {}), ...pairs(opts) })).output,
  // SAVEXLSX(): the workbook as an XLSX file, its bytes in base64.
  SAVEXLSX: async (exec) => (await must(exec, "file.exportXlsx")).output,
  FILEREPORT: async (exec) => stem((await must(exec, "file.report")).output as string[]),
  // JSON and HTML: IMPORTJSON("A1", text), IMPORTHTML("A1", html, 2) its
  // second table, EXPORTHTML("A1:C9", 1) the first row a header.
  IMPORTJSON: async (exec, at, t) => (await must(exec, "file.importJson", { at: text(at), text: text(t) })).message,
  IMPORTHTML: async (exec, at, t, n) => (await must(exec, "file.importHtml", { at: text(at), text: text(t), ...(text(n).trim() ? { table: text(n) } : {}) })).message,
  EXPORTHTML: async (exec, r, header) => (await must(exec, "file.exportHtml", { ...(text(r).trim() ? { range: text(r) } : {}), ...(yes(header) ? { header: "1" } : {}) })).output,
  // Slicers: ADDSLICER("Sales", "region"), SLICE("region", "north,south"),
  // SLICE("region", "west", 1) to add it; CLEARSLICER(slicer),
  // DELETESLICER(slicer), LISTSLICERS().
  ADDSLICER: async (exec, t, c, n) => (await must(exec, "slicer.add", { table: text(t), column: text(c), ...(text(n).trim() ? { name: text(n) } : {}) })).message,
  SLICE: async (exec, n, values, add) => (await must(exec, "slicer.select", { name: text(n), values: text(values), ...(yes(add) ? { add: "1" } : {}) })).message,
  CLEARSLICER: async (exec, n) => (await must(exec, "slicer.clear", { name: text(n) })).message,
  DELETESLICER: async (exec, n) => (await must(exec, "slicer.delete", { name: text(n) })).message,
  LISTSLICERS: async (exec) => stem((await must(exec, "slicer.list")).output as string[]),
  // Themes: SETTHEME("Harbour"), GETTHEME(), SETTHEMECOLOR("accent1",
  // "#123456"), SETTHEMEFONTS("Georgia", "Verdana").
  SETTHEME: async (exec, n) => (await must(exec, "theme.set", { name: text(n) })).message,
  GETTHEME: async (exec) => ((await must(exec, "theme.get")).output as { name: string }).name,
  SETTHEMECOLOR: async (exec, role, c) => (await must(exec, "theme.color", { role: text(role), color: text(c) })).message,
  SETTHEMEFONTS: async (exec, major, minor) => (await must(exec, "theme.fonts", { ...(text(major).trim() ? { major: text(major) } : {}), ...(text(minor).trim() ? { minor: text(minor) } : {}) })).message,
  // Cell and table styles: DEFINESTYLE("Money", "bold=1;fill=#e2efda;number=$#,##0.00"),
  // DEFINESTYLE("Mine", "from=B2"); APPLYSTYLE("B2:B9", "Good"); LISTSTYLES();
  // DELETESTYLE("Mine"); SETTABLESTYLE("Sales", "Medium", "firstColumn=1;bandedColumns=1").
  DEFINESTYLE: async (exec, n, look) => (await must(exec, "style.define", { name: text(n), ...pairs(look) })).message,
  APPLYSTYLE: async (exec, r, n) => (await must(exec, "style.apply", { range: text(r), name: text(n) })).message,
  DELETESTYLE: async (exec, n) => (await must(exec, "style.delete", { name: text(n) })).message,
  LISTSTYLES: async (exec) => stem((await must(exec, "style.list")).output as string[]),
  SETTABLESTYLE: async (exec, n, st, opts) => (await must(exec, "table.style", { name: text(n), ...(text(st).trim() ? { style: text(st) } : {}), ...pairs(opts) })).message,
  // Consolidate: CONSOLIDATE("F1", "Jan!A1:C4,Feb!A1:C4", "sum", "both", 1).
  CONSOLIDATE: async (exec, to, sources, fn, labels, links) => (await must(exec, "data.consolidate", { to: text(to), sources: text(sources), ...(text(fn).trim() ? { function: text(fn) } : {}), ...(text(labels).trim() ? { labels: text(labels) } : {}), ...(yes(links) ? { links: "1" } : {}) })).message,
  // Advanced Filter: ADVANCEDFILTER("A1:D20", "F1:G3") in place, or
  // ("A1:D20", "F1:G3", "I1", 1) copied to I1, each record once.
  ADVANCEDFILTER: async (exec, r, c, to, unique) => (await must(exec, "data.advancedFilter", { range: text(r), criteria: text(c), ...(text(to).trim() ? { copyTo: text(to) } : {}), ...(yes(unique) ? { unique: "1" } : {}) })).message,
  // A table's totals row, calculated columns, banding and Convert to
  // Range: SETTABLETOTALS("Sales", "Qty=sum,Price=average"), ("Sales", 0)
  // off; SETCALCULATEDCOLUMN("Sales", "Total", "=[@Qty]*[@Price]");
  // SETTABLEBANDING("Sales", 1); TABLETORANGE("Sales").
  SETTABLETOTALS: async (exec, n, fns) => (await must(exec, "table.totals", { name: text(n), ...(text(fns).trim() === "0" ? { on: "0" } : text(fns).trim() ? { functions: text(fns) } : {}) })).message,
  SETCALCULATEDCOLUMN: async (exec, n, c, f) => (await must(exec, "table.calculatedColumn", { name: text(n), column: text(c), formula: text(f) })).message,
  SETTABLEBANDING: async (exec, n, on) => (await must(exec, "table.banding", { name: text(n), rows: text(on).trim() === "" || yes(on) ? "1" : "0" })).message,
  TABLETORANGE: async (exec, n) => (await must(exec, "table.toRange", { name: text(n) })).message,
  // Charts: INSERTCHART("A1:C9", "barStacked", "Sales") (bar, barStacked,
  // bar100, hbar, hbarStacked, hbar100, line, lineStacked, area,
  // areaStacked, pie, doughnut, scatter, scatterLines, bubble);
  // LISTCHARTS(), DELETECHART(n).
  INSERTCHART: async (exec, r, kind, title) => (await must(exec, "chart.insert", { range: text(r), kind: text(kind) || "bar", ...(text(title).trim() ? { title: text(title) } : {}) })).message,
  // SETCHARTOPTION(1, "legend", "right"), (1, "xTitle", "Quarter"),
  // (1, "labels", 1), (1, "gridlines", 0), (1, "byRows", 1), (1, "title", "Sales");
  // SETCHARTDATA(1, "A1:C9"); SETCHARTTYPE(1, "area").
  SETCHARTOPTION: async (exec, n, name, v) => {
    const k = ["title", "xTitle", "yTitle", "legend", "labels", "gridlines", "byRows"].find((x) => x.toLowerCase() === text(name).trim().toLowerCase());
    if (!k) throw new Error(`SETCHARTOPTION: ${text(name)} is not title, xTitle, yTitle, legend, labels, gridlines or byRows`);
    return (await must(exec, "chart.option", { index: text(n), [k]: text(v) })).message;
  },
  // ADDTRENDLINE(1, "North", "linear", "equation=1;r2=1"), (1, 2,
  // "polynomial", "order=3"), (1, 1, "movingAverage", "period=3"),
  // (1, 1, "none") to take it off.
  ADDTRENDLINE: async (exec, n, series, type, opts) => (await must(exec, "chart.trendline", { index: text(n), series: text(series), type: text(type), ...pairs(opts) })).message,
  SETCHARTDATA: async (exec, n, r) => (await must(exec, "chart.data", { index: text(n), range: text(r) })).message,
  SETCHARTTYPE: async (exec, n, kind) => (await must(exec, "chart.data", { index: text(n), kind: text(kind) })).message,
  LISTCHARTS: async (exec) => stem((await must(exec, "chart.list")).output as string[]),
  DELETECHART: async (exec, n) => (await must(exec, "chart.delete", { index: text(n) })).message,
  // A sparkline: INSERTSPARKLINE("A1:A9", "B1", "winloss", "axis").
  INSERTSPARKLINE: async (exec, r, at, type, options) => (await must(exec, "insert.sparkline", { range: text(r), at: text(at), ...(text(type).trim() ? { type: text(type) } : {}), ...(text(options).trim() ? { options: text(options) } : {}) })).message,
  // Data bars and icon sets with Excel's options, as name=value pairs:
  // ADDDATABAR("B1:B9", "#638ec6", "fill=solid;showValue=0;min=0;max=100"),
  // ADDICONSET("C1:C9", "arrows", "thresholds=80,50;by=percentile;reverse=1;showValue=0").
  ADDDATABAR: async (exec, r, c, opts) => (await must(exec, "format.dataBar", { range: text(r), ...(text(c).trim() ? { color: text(c) } : {}), ...pairs(opts) })).message,
  ADDICONSET: async (exec, r, icons, opts) => (await must(exec, "format.iconSet", { range: text(r), ...(text(icons).trim() ? { icons: text(icons) } : {}), ...pairs(opts) })).message,
  LISTRULES: async (exec) => stem((await must(exec, "format.rules")).output as string[]),
  DELETERULE: async (exec, n) => (await must(exec, "format.deleteRule", { index: text(n) })).message,
  MOVERULE: async (exec, n, to) => (await must(exec, "format.moveRule", { index: text(n), to: text(to) })).message,
  STOPIFTRUE: async (exec, n, on) => (await must(exec, "format.stopIfTrue", { index: text(n), ...(text(on).trim() !== "" && !yes(on) ? { off: "1" } : {}) })).message,
  CLEARRULES: async (exec, r) => (await must(exec, "format.clearRules", { range: text(r) })).message,
  // Borders: SETBORDER("A1:C3", "outline", "thick", "red"), the edge
  // outline, all, inside, horizontal, vertical, top, right, bottom, left or
  // none; CLEARBORDERS("A1:C3").
  SETBORDER: async (exec, r, edge, style, c) => (await must(exec, "format.border", { range: text(r), edge: text(edge) || "outline", ...(text(style).trim() ? { style: text(style) } : {}), ...(text(c).trim() ? { color: text(c) } : {}) })).message,
  CLEARBORDERS: async (exec, r) => (await must(exec, "format.border", { range: text(r), edge: "none" })).message,
  // A value in a number format: FORMATVALUE(1234.5, "#,##0.00"), FORMATVALUE(0.5, "fraction").
  FORMATVALUE: async (exec, v, f) => (await must(exec, "format.preview", { value: text(v), format: text(f) })).output,
  SETNUMBERFORMAT: async (exec, r, f) => (await must(exec, "format.number", { range: text(r), format: text(f) }), "OK"),
  SETWRAPTEXT: async (exec, c, on) => (await must(exec, "format.wrap", yes(on) ? { range: text(c) } : { range: text(c), off: "1" }), "OK"),
  GETWRAPTEXT: async (exec, c) => (await about(exec, c)).wrap,
  FIND: async (exec, t, matchCase, wholeCell, inFormulas) =>
    stem((await must(exec, "edit.find", { text: text(t), ...options(matchCase, wholeCell, inFormulas) })).output as string[]),
  REPLACE: async (exec, t, by, matchCase, wholeCell, inFormulas) =>
    (await must(exec, "edit.replace", { text: text(t), with: text(by), ...options(matchCase, wholeCell, inFormulas) })).output,
  MOVECOLUMNLEFT: async (exec, c) => (await must(exec, "columns.move", { at: col(c, "MOVECOLUMNLEFT"), way: "left" }), ""),
  MOVECOLUMNRIGHT: async (exec, c) => (await must(exec, "columns.move", { at: col(c, "MOVECOLUMNRIGHT"), way: "right" }), ""),
  RENAMESHEET: async (exec, from, to) => (await must(exec, "sheet.rename", { name: text(from), to: text(to) }), text(to)),
  // The app's SETSHEETNAME names its document, which it never shows; here
  // it renames the active sheet, as GETSHEETNAME reads it.
  SETSHEETNAME: async (exec, to) => (await must(exec, "sheet.rename", { name: (await state(exec)).name, to: text(to) }), text(to)),
  // A cell's editor (editors.ts): type checkbox, dropdown or date, and the
  // app's config ({"options": [...]}, {"format": ...}) as JSON.
  SETCELLEDITOR: async (exec, c, type, config) => {
    const conf = (typeof config === "string" && config ? JSON.parse(config) : config ?? {}) as { options?: unknown; format?: unknown };
    await must(exec, "cell.editor", {
      cell: text(c), kind: text(type), ...(Array.isArray(conf.options) ? { options: conf.options.map(text).join(",") } : {}), ...(conf.format ? { format: text(conf.format) } : {}),
    });
    return text(c);
  },
  GETCELLEDITOR: async (exec, c) => {
    const e = (await about(exec, c)).editor;
    return e ? JSON.stringify(e) : "";
  },
  REMOVECELLEDITOR: async (exec, c) => ((await about(exec, c)).editor ? (await must(exec, "cell.editor", { cell: text(c), kind: "none" }), text(c)) : text(c)),
  // Pivot tables (pivot.ts): the app's configuration as JSON, {sourceRange,
  // rowFields, colFields, valueField, aggFunction, outputCell, filters}.
  CREATEPIVOT: async (exec, id, json) => {
    const c = (typeof json === "string" ? JSON.parse(json) : json ?? {}) as Record<string, unknown>;
    const list = (xs: unknown) => (Array.isArray(xs) ? xs.map(text).join(",") : "");
    await must(exec, "pivot.create", {
      id: text(id), source: text(c.sourceRange), value: text(c.valueField), agg: text(c.aggFunction), output: text(c.outputCell),
      rows: list(c.rowFields), cols: list(c.colFields), filters: JSON.stringify(c.filters ?? {}),
    });
    return text(id);
  },
  UPDATEPIVOT: async (exec, id) => (await must(exec, "pivot.update", { id: text(id) }), text(id)),
  DELETEPIVOT: async (exec, id) => (await must(exec, "pivot.delete", { id: text(id) }), text(id)),
  GETPIVOT: async (exec, id) => {
    const o = await exec("pivot.get", { id: text(id) });
    return o.ok ? JSON.stringify(o.output) : "";
  },
  LISTPIVOTS: async (exec) => JSON.stringify((await must(exec, "pivot.list")).output),
  // A configuration for CREATEPIVOT, its field lists as JSON arrays.
  PIVOTCONFIG: (_, source, value, agg, output, rows, cols) => JSON.stringify({
    sourceRange: text(source), valueField: text(value), aggFunction: text(agg).toUpperCase(), outputCell: text(output),
    rowFields: rows ? JSON.parse(text(rows)) : [], colFields: cols ? JSON.parse(text(cols)) : [],
  }),
  HIDEROW: async (exec, n) => (await must(exec, "rows.hide", { range: rowsOf(row(n, "HIDEROW")) }), "OK"),
  UNHIDEROW: async (exec, n) => (await must(exec, "rows.unhide", { range: rowsOf(row(n, "UNHIDEROW")) }), "OK"),
  HIDECOLUMN: async (exec, c) => (await must(exec, "columns.hide", { range: `${col(c, "HIDECOLUMN")}1` }), "OK"),
  UNHIDECOLUMN: async (exec, c) => (await must(exec, "columns.unhide", { range: `${col(c, "UNHIDECOLUMN")}1` }), "OK"),
  ISROWHIDDEN: async (exec, n) => ((await state(exec)).hiddenRows.includes(row(n, "ISROWHIDDEN")) ? 1 : 0),
  ISCOLUMNHIDDEN: async (exec, c) => ((await state(exec)).hiddenCols.includes(col(c, "ISCOLUMNHIDDEN")) ? 1 : 0),
  ISROWVISIBLE: async (exec, n) => ((await state(exec)).hiddenRows.includes(row(n, "ISROWVISIBLE")) ? "0" : "1"),
  SETROWHEIGHT: async (exec, n, h) => (await must(exec, "rows.height", { range: rowsOf(row(n, "SETROWHEIGHT")), height: text(h) }), "OK"),
  SETCOLWIDTH: async (exec, c, w) => (await must(exec, "columns.width", { range: `${col(c, "SETCOLWIDTH")}1`, width: text(w) }), "OK"),
  // 0: the grid's own width (and height, below).
  GETCOLWIDTH: async (exec, c) => {
    const letter = col(c, "GETCOLWIDTH");
    return (await state(exec)).colWidths.find(([at]) => at === letter)?.[1] ?? 0;
  },
  GETFILTERCRITERIA: async (exec) => {
    const filter = (await state(exec)).filter;
    return filter ? JSON.stringify(filter) : "";
  },
  GETROWHEIGHT: async (exec, n) => {
    const r = row(n, "GETROWHEIGHT");
    return (await state(exec)).rowHeights.find(([at]) => at === r)?.[1] ?? 0;
  },
  DEFINENAMEDRANGE: async (exec, n, r) => (await must(exec, "names.define", { name: text(n), range: text(r) }), "OK"),
  DELETENAMEDRANGE: async (exec, n) => (await must(exec, "names.delete", { name: text(n) }), "OK"),
  GETNAMEDRANGE: async (exec, n) => named(await names(exec)).get(text(n).toUpperCase()) ?? "",
  GETALLNAMEDRANGES: async (exec) => {
    const all = [...named(await names(exec))];
    return Object.fromEntries([[0, all.length], ...all.flatMap(([n, r], i) => [[`${i + 1}.NAME`, n], [`${i + 1}.RANGE`, r]])]);
  },
  FREEZEPANES: async (exec, rows, cols) => {
    await must(exec, "view.freezeRows", { count: text(rows) || "0" });
    return (await must(exec, "view.freezeCols", { count: text(cols) || "0" }), "OK");
  },
  UNFREEZEPANES: async (exec) => (await must(exec, "view.freezeRows", { count: "0" }), await must(exec, "view.freezeCols", { count: "0" }), "OK"),
  GETFROZENPANES: async (exec) => {
    const st = await state(exec);
    return { ROWS: st.frozenRows, COLUMNS: st.frozenCols };
  },
  SETCELLVALIDATION: async (exec, c, type, ...args) => (await must(exec, "validation.set", { range: text(c), ...validation(text(type), args) }), "OK"),
  CLEARCELLVALIDATION: async (exec, c) => (await must(exec, "validation.clear", { range: text(c) }), "OK"),
  VALIDATECELL: async (exec, c, v) => {
    const rule = (await about(exec, c)).validation;
    return rule && check(rule, text(v), true) !== null ? 0 : 1;
  },
  UNDO: async (exec) => ((await state(exec)).canUndo ? (await must(exec, "edit.undo"), 1) : 0),
  REDO: async (exec) => ((await state(exec)).canRedo ? (await must(exec, "edit.redo"), 1) : 0),
  CANUNDO: async (exec) => (await state(exec)).canUndo,
  CANREDO: async (exec) => (await state(exec)).canRedo,
  ADDSHEET: async (exec, n) => (await must(exec, "sheet.add", { name: text(n) }), text(n)),
  DELETESHEET: async (exec, n) => (await must(exec, "sheet.delete", { name: text(n) }), ""),
  SETACTIVESHEET: async (exec, n) => (await must(exec, "sheet.activate", { name: text(n) }), text(n)),
  GETACTIVESHEET: async (exec) => (await state(exec)).name,
  GETSHEETNAMES: async (exec) => stem((await state(exec)).sheets),
  // The app filters every row by a substring; demo 18's AutoFilter keeps
  // the first row as headers and reads the criterion as COUNTIF does.
  APPLYROWFILTER: async (exec, c, criteria) => {
    const letter = col(c, "APPLYROWFILTER");
    return (await must(exec, "data.filter", { range: `A1:${colName(COLS - 1)}${ROWS}`, column: letter, criterion: text(criteria) }), "");
  },
  CLEARROWFILTER: async (exec) => (await must(exec, "data.clearFilter"), ""),
  MERGECELLS: async (exec, r) => {
    await must(exec, "cells.merge", { range: text(r) });
    return (await about(exec, text(r).split(":")[0])).merged.split(":")[0];
  },
  UNMERGECELLS: async (exec, c) => {
    const merged = (await about(exec, c)).merged;
    await must(exec, "cells.unmerge", { range: text(c) });
    return merged.split(":")[0];
  },
  GETMERGEDRANGE: async (exec, c) => (await about(exec, c)).merged,
  ISCELLMERGED: async (exec, c) => ((await about(exec, c)).merged ? "1" : "0"),
  DEFINETABLE: async (exec, n, r, headers) => {
    await must(exec, "table.define", { name: text(n), range: text(r), ...(headers === undefined || yes(headers) ? { headers: "1" } : {}) });
    return JSON.stringify(await table(exec, text(n)));
  },
  DELETETABLE: async (exec, n) => (await must(exec, "table.delete", { name: text(n) }), ""),
  GETTABLE: async (exec, n) => {
    const t = await table(exec, text(n));
    return t ? JSON.stringify(t) : "";
  },
  LISTTABLES: async (exec) => stem(await tableNames(exec)),
  // The app's table metadata: {range, columns: {name: letter}, hasHeader,
  // types, descriptions}, a table here with its columns named so.
  SET_TABLE_METADATA: async (exec, n, json) => {
    const meta = (typeof json === "string" ? JSON.parse(json) : json) as { range?: unknown; columns?: Record<string, unknown>; hasHeader?: unknown; types?: unknown; descriptions?: unknown };
    const r = parseRange(text(meta?.range));
    if (!r || !meta.columns || typeof meta.columns !== "object") throw new Error("Table metadata must include range and columns");
    const letters = Object.entries(meta.columns).map(([name, letter]) => [name, text(letter).toUpperCase()]);
    const names = Array.from({ length: Math.abs(r[0].col - r[1].col) + 1 }, (_, i) => colName(Math.min(r[0].col, r[1].col) + i))
      .map((letter) => letters.find(([, l]) => l === letter)?.[0] ?? letter);
    await must(exec, "table.define", {
      name: text(n), range: text(meta.range), headers: meta.hasHeader === false ? "0" : "1", columns: names.join(","),
      types: JSON.stringify(meta.types ?? {}), descriptions: JSON.stringify(meta.descriptions ?? {}),
    });
    return `Table metadata set for '${text(n)}'`;
  },
  GET_TABLE_METADATA: async (exec, n) => {
    const t = await table(exec, text(n));
    if (!t) return "";
    const left = parseRange(t.range)![0].col;
    return JSON.stringify({ range: t.range, columns: Object.fromEntries(t.columns.map((c, i) => [c, colName(left + i)])), hasHeader: t.headers, types: t.types, descriptions: t.descriptions });
  },
  DELETE_TABLE_METADATA: async (exec, n) => (await must(exec, "table.delete", { name: text(n) }), `Table metadata deleted for '${text(n)}'`),
  LIST_TABLES: async (exec) => JSON.stringify(await tableNames(exec)),
  // asc, desc, then original.
  SORTTABLEBYCOLUMN: async (exec, n, c) => (await must(exec, "table.sort", { name: text(n), column: col(c, "SORTTABLEBYCOLUMN") })).output,
  CONFIGURE_AUTO_ID: async (exec, c, start, prefix) => {
    if (c === null || c === undefined || text(c) === "" || text(c) === "null") return (await must(exec, "ids.configure"), "Auto-ID disabled");
    const from = start === undefined ? "1" : text(start), pre = prefix === undefined ? "" : text(prefix);
    await must(exec, "ids.configure", { column: col(c, "CONFIGURE_AUTO_ID"), start: from, prefix: pre });
    return `Auto-ID configured: column=${text(c)}, nextId=${from}, prefix="${pre}"`;
  },
  FIND_ROW_BY_ID: async (exec, id) => String((await must(exec, "ids.find", { id: text(id) })).output),
  // [{column, value}, ...] set in the row with that ID.
  UPDATE_ROW_BY_ID: async (exec, id, updates) => {
    const row = String((await must(exec, "ids.find", { id: text(id) })).output);
    if (!row) throw new Error(`Row with ID ${text(id)} not found`);
    for (const u of unstem(updates) as { column?: unknown; value?: unknown }[]) await must(exec, "cell.set", { cell: `${col(u.column, "UPDATE_ROW_BY_ID")}${row}`, input: text(u.value) });
    return row;
  },
  GET_NEXT_ID: async (exec) => {
    const auto = (await state(exec)).autoId;
    if (!auto) throw new Error("Auto-ID column is not configured for this sheet");
    return auto.next;
  },
};

const options = (matchCase: unknown, wholeCell: unknown, inFormulas: unknown): Record<string, string> =>
  Object.fromEntries(([["matchCase", matchCase], ["wholeCell", wholeCell], ["inFormulas", inFormulas]] as const).filter(([, v]) => yes(v)).map(([k]) => [k, "1"]));

// The app's validation types as demo 18's rules: list, number (min, max),
// text (a length from min to max, or a pattern third); and Excel's: whole,
// date and time (min, max), custom (a formula). Then, as name=value pairs
// split by ';', mode= stop, warn or info, prompt= and error=.
function validation(type: string, args: readonly unknown[]): Record<string, string> {
  const given = (a: unknown) => a !== undefined && text(a).trim() !== "";
  const limits = (kind: string) => ({ kind, ...(given(args[0]) ? { min: text(args[0]) } : {}), ...(given(args[1]) ? { max: text(args[1]) } : {}) });
  const more = (i: number) => (given(args[i]) ? pairs(args[i]) : {});
  switch (type.toLowerCase()) {
    case "list": return { kind: "list", values: text(args[0]), ...more(1) };
    case "number": return { ...limits("decimal"), ...more(2) };
    case "whole": return { ...limits("whole"), ...more(2) };
    case "date": return { ...limits("date"), ...more(2) };
    case "time": return { ...limits("time"), ...more(2) };
    case "custom": return { kind: "custom", formula: text(args[0]), ...more(1) };
    case "text":
      if (given(args[2])) return { kind: "pattern", pattern: text(args[2]), ...more(3) };
      return { ...limits("length"), ...more(3) };
    default: throw new Error(`SETCELLVALIDATION: ${type} is not list, number, whole, date, time, custom or text`);
  }
}

// names.list's "NAME A1:B2" lines.
const names = async (exec: Exec) => (await must(exec, "names.list")).output as string[];
const named = (lines: readonly string[]) => new Map(lines.map((l) => l.split(" ") as [string, string]));

// A table as table.get gives it, or null; the tables' names, from
// table.list's "Sales A1:D5 (region, amount)" lines.
interface Table {
  readonly name: string;
  readonly range: string;
  readonly headers: boolean;
  readonly columns: readonly string[];
  readonly types: Readonly<Record<string, string>>;
  readonly descriptions: Readonly<Record<string, string>>;
}
async function table(exec: Exec, name: string): Promise<Table | null> {
  const o = await exec("table.get", { name });
  return o.ok ? (o.output as Table) : null;
}
const tableNames = async (exec: Exec) => ((await must(exec, "table.list")).output as string[]).map((l) => l.split(" ")[0]);


type Fn = (...args: any[]) => unknown;

// Every control function, for a script: those above, LISTCOMMANDS and
// BATCH_EXECUTE. `others` are the functions a
// script has besides (SETCELL, GETCELL, the range functions), which
// LISTCOMMANDS lists and BATCH_EXECUTE calls too.
export function controlFunctions(exec: Exec, others: Readonly<Record<string, Fn>>): Record<string, Fn> {
  const fns: Record<string, Fn> = {};
  for (const [name, fn] of Object.entries(CONTROL)) fns[name] = (...args: any[]) => fn(exec, ...args);
  const listed = [...new Set([...Object.keys(others), ...Object.keys(CONTROL), "LISTCOMMANDS", "BATCH_EXECUTE"])];
  // LISTCOMMANDS() as the app's: a stem of the names a script may call.
  // LISTCOMMANDS(1): each described (described.ts), {name, signature,
  // params (a stem of {name, type, optional, variadic, about}), does, gives}.
  fns.LISTCOMMANDS = (described?: unknown) => {
    if (!yes(described)) return stem(listed);
    return stem(listed.map((name) => {
      const d = DESCRIBED[name] ?? { params: [], does: "", gives: "" };
      const params = d.params.map((x) => ({ name: x.name, type: x.type, optional: x.optional ? 1 : 0, variadic: x.variadic ? 1 : 0, about: x.about ?? "" }));
      return { name, signature: signature(name, d), params: stem(params), does: d.does, gives: d.gives };
    }));
  };
  // [{command, args}, ...]: each run in turn, by any of these names; what
  // each gave or why it failed, as the app's. `stop` (not the app's): the
  // first that fails is the whole call's error, so a control-bus request
  // made of it is taken back whole (commands.ts's spreadsheet).
  fns.BATCH_EXECUTE = async (commands: unknown, stop?: unknown) => {
    const results: Record<string, unknown>[] = [];
    const halt = yes(stop);
    for (const [index, c] of (unstem(commands) as { command?: unknown; args?: unknown }[]).entries()) {
      const command = text(c.command).toUpperCase();
      const fn = fns[command] ?? others[command];
      if (!fn || command === "BATCH_EXECUTE") {
        if (halt) throw new Error(`command ${index + 1}: Unknown command: ${command}`);
        results.push({ index, command, success: false, error: `Unknown command: ${command}` });
        continue;
      }
      try {
        results.push({ index, command, success: true, result: await fn(...(Array.isArray(c.args) ? c.args : [])) });
      } catch (e) {
        if (halt) throw new Error(`command ${index + 1} (${command}): ${(e as Error).message}`);
        results.push({ index, command, success: false, error: (e as Error).message });
      }
    }
    return JSON.stringify(results);
  };
  return fns;
}
