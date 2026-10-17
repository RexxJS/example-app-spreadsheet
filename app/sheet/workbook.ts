// The workbook: everything a spreadsheet holds, with no ui. Its sheets, in
// order, each a Sheet (inputs, values, the graph, undo) and what its cells
// carry besides their inputs (notes), one of them the active one; the two
// formula languages (the major setting, the same for every sheet), the
// file's name and setup script, the clipboard, and each conditional style
// of the active sheet as last worked out. A formula on one sheet can read
// another's cells (Costs!B2); after a change, settle() works out what reads
// them. Formula auditing's state is here too: the traces (precedents or
// dependents of a cell, a number of levels deep, across sheets), Show
// Formulas, and the watched cells. It answers what a cell shows, its style
// and its status line.
// Commands (commands.ts) change it; the page's panels read it.
import { drawn } from "./editors.ts";
import { COLS, ROWS, addrOf, cellAddr, keyOf, order, parseRange, range, splitKey } from "./address.ts";
import { movedBy } from "./shift.ts";
import { REFERENCE_COLOURS } from "./typing.ts";
import type { Addr } from "./address.ts";
import type { PivotFields } from "./pivot.ts";
import { EXCEL, fromR1C1, toR1C1 } from "./excel.ts";
import { setDateSystem as setDates } from "./dates.ts";
import type { DateSystem } from "./dates.ts";
import { describe, show } from "./format.ts";
import { formatStyle, formatted, lookStyle, toStyle } from "./formats.ts";
import type { Language, TableDef } from "./language.ts";
import { dataRows } from "./structured.ts";
import { CELL_STYLES, TABLE_STYLES, under } from "./styles.ts";
import { THEMES, resolved, resolvedStyle, setTheme, theme } from "./themes.ts";
import type { Theme } from "./themes.ts";
import type { Look } from "./formats.ts";
import { CellNotes } from "./notes.ts";
import type { Note } from "./notes.ts";
import { RexxLanguage } from "./rexx.ts";
import { AUTOMATIC, Sheet, Turns } from "./sheet.ts";
import type { Calc, Context, SavedSheet } from "./sheet.ts";
import { CellError } from "./values.ts";
import type { Value } from "./values.ts";
import { applied } from "./rules.ts";
import { chartData, layout } from "./charts.ts";
import type { ChartData, ChartKind, ChartOptions, Shape } from "./charts.ts";
import { rangeName } from "./rules.ts";
import { sheetPrefix } from "./address.ts";
import type { CfRule, Mark } from "./rules.ts";
import { check } from "./validation.ts";

export type LanguageName = "Excel" | "RexxJS";

// What was copied, from where, in which language (so a paste moves its
// formulas): as tab-separated text, and cell by cell (inputs, what they
// showed, their notes) for Paste Special; and, when it was cut, the sheet
// and range it is to be moved from. Pages have no system clipboard yet.
export interface Clip {
  readonly cut?: { readonly page: Page; readonly range: [Addr, Addr] };
  readonly tsv: string;
  readonly from: Addr;
  readonly language: Language;
  readonly inputs: readonly (readonly string[])[];
  readonly shown: readonly (readonly string[])[];
  readonly notes: readonly (readonly Note[])[];
}

// A chart on a sheet: what kind, of which range, under what title.
export interface ChartSpec {
  readonly kind: ChartKind;
  readonly range: [Addr, Addr];
  readonly title: string;
  readonly options?: ChartOptions;
  // A PivotChart's pivot (by its id): its range what the pivot last wrote,
  // its grand totals left out, so it follows the pivot.
  readonly pivot?: string;
}

// A chart as the page draws it, its shapes laid out in CHART_W x CHART_H.
export interface ChartView {
  readonly kind: ChartKind;
  readonly title: string;
  readonly range: string;
  readonly data: ChartData;
  readonly shapes: readonly Shape[];
}

export const CHART_W = 400, CHART_H = 240;

// A sheet of the workbook: its name as given, its cells and their notes, its
// conditional-format rules (rules.ts), in priority order, its charts, and
// how it is viewed: rows and columns hidden by hand (0-based), rows hidden
// by its AutoFilter (and the filter itself), how many rows are frozen, its
// merged cells, rows set to a height (0-based row to px) and columns set to
// a width (0-based column to px).
export interface Page {
  name: string;
  readonly sheet: Sheet;
  readonly notes: CellNotes;
  rules: CfRule[];
  charts: ChartSpec[];
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  filtered: Set<number> | null;
  filter: Filter | null;
  frozen: number;
  frozenCols: number;
  merges: [Addr, Addr][];
  heights: Map<number, number>;
  widths: Map<number, number>;
  autoId: AutoId | null;
  pivots: Map<string, Pivot>;
  // Its tab's colour (#rrggbb) or null, and whether the sheet is hidden.
  tabColor: string | null;
  hiddenSheet: boolean;
  // Its outline (outline.ts): each grouped row's (column's) level, 1 to 8.
  rowLevels: Map<number, number>;
  colLevels: Map<number, number>;
  // Its slicers: tiles of a table column's values, filtering the table's
  // rows to those picked (none picked: every row).
  slicers: Slicer[];
  // Its scenarios (Scenario Manager): named sets of values for some cells.
  scenarios: Scenario[];
}

export interface Scenario {
  readonly name: string;
  readonly cells: readonly string[];
  readonly values: readonly string[];
  readonly comment?: string;
}

export interface Slicer {
  readonly name: string;
  readonly table: string;
  readonly column: string;
  readonly picked: readonly string[];
}

// What a slicer shows: its tiles, each value with whether it is picked.
export interface SlicerView {
  readonly name: string;
  readonly caption: string;
  readonly tiles: readonly { readonly value: string; readonly picked: boolean }[];
}

// A pivot table (pivot.ts) by its id: its source range, the cell it is
// written from, and the cells it covered when last written (cleared when
// it is written again).
export interface Pivot extends PivotFields {
  readonly source: [Addr, Addr];
  readonly output: Addr;
  readonly written: [Addr, Addr] | null;
}

// The Tauri app's auto IDs: the column a row inserted is given the next ID
// in, as the prefix and a number counting up.
export interface AutoId {
  readonly column: number;
  readonly next: number;
  readonly prefix: string;
}

// An AutoFilter: its range (headers in the first row), the column it
// tests, and the criterion, as COUNTIF reads one.
export interface Filter {
  readonly range: [Addr, Addr];
  readonly column: number;
  readonly criterion: string;
}

export type Direction = "precedents" | "dependents";

interface Trace {
  readonly page: Page;
  readonly cell: string;
  readonly direction: Direction;
  levels: number;
}

// What a traced cell's background is: blue for what it reads, orange for
// what reads it.
const TRACED: Record<Direction, string> = { precedents: "background: #dbe9ff", dependents: "background: #ffe3c2" };

export interface Watch {
  readonly sheet: string;
  readonly cell: string;
  readonly input: string;
  readonly shown: string;
}

// A workbook as it was, to go back to: its sheets in order (the Page
// objects themselves, so a sheet deleted comes back as it was), each one's
// name, cells, notes, rules, charts, view and the rest, which was active,
// and the file's name and setup script. Not what is only looked at (the
// clipboard, traces, watches, Show Formulas, Circle Invalid), and not
// RexxJS's own state: what a setup script defined stays defined.
interface Saved {
  // What the change was, as Undo and Redo name it ("Merge Cells"); none for
  // the workbook saved before a change is known.
  readonly label?: string;
  readonly pages: readonly SavedPage[];
  readonly names: ReadonlyMap<string, BookName>;
  readonly active: Page;
  readonly name: string;
  readonly setupScript: string;
  readonly styles: ReadonlyMap<string, Look>;
  readonly theme: Theme;
}

interface SavedPage {
  readonly page: Page;
  readonly name: string;
  readonly sheet: SavedSheet;
  readonly notes: ReadonlyMap<string, Note>;
  readonly view: Omit<Page, "name" | "sheet" | "notes">;
}

// A page's own state, copied: each container new, what is in them shared
// (rules, charts, merges, filters and pivots are replaced, never changed
// in place).
const viewOf = (p: Omit<Page, "name" | "sheet" | "notes">): Omit<Page, "name" | "sheet" | "notes"> => ({
  rules: [...p.rules], charts: [...p.charts], hiddenRows: new Set(p.hiddenRows), hiddenCols: new Set(p.hiddenCols),
  filtered: p.filtered && new Set(p.filtered), filter: p.filter, frozen: p.frozen, frozenCols: p.frozenCols, merges: [...p.merges],
  heights: new Map(p.heights), widths: new Map(p.widths), autoId: p.autoId, pivots: new Map(p.pivots),
  tabColor: p.tabColor, hiddenSheet: p.hiddenSheet, rowLevels: new Map(p.rowLevels), colLevels: new Map(p.colLevels), slicers: [...p.slicers], scenarios: [...p.scenarios],
});

// A name for the whole workbook: a range of one of its sheets, or a
// formula (what follows its '=', a named LAMBDA). A sheet's own name of the
// same spelling hides it on that sheet, as in Excel.
export type BookName = { readonly page: Page; readonly range: [Addr, Addr] } | { readonly formula: string };

// Undo steps kept, as the Tauri spreadsheet keeps.
const HISTORY = 100;

// How many times settle() passes changes between sheets before it stops: a
// guard only, as a cycle through two sheets is found first (#crossCycles).
const ROUNDS = 20;

export interface Stats {
  readonly count: number;
  readonly numbers: number;
  readonly sum: number | null;
  readonly average: number | null;
  readonly min: number | null;
  readonly max: number | null;
}

// The status bar's line for a range's Stats: "Average: 2  Count: 3 ...".
export function statsLine(s: Stats): string {
  const n = (x: number | null) => show(x);
  return s.numbers ? `Average: ${n(s.average)}  Count: ${s.count}  Numerical Count: ${s.numbers}  Min: ${n(s.min)}  Max: ${n(s.max)}  Sum: ${n(s.sum)}` : `Count: ${s.count}`;
}

// A banded table's shade (Excel's light blue band), and the cells it
// shades: the first data row and every other one after it.
export const BAND = "#d9e1f2";

// A slicer's column's cells, by row, in its table's data rows (none if the
// table or column is gone).
export function slicerColumn(page: Page, sl: Slicer): [number, string][] {
  const t = page.sheet.table(sl.table);
  if (!t) return [];
  const left = Math.min(t.range[0].col, t.range[1].col), right = Math.max(t.range[0].col, t.range[1].col), top = Math.min(t.range[0].row, t.range[1].row);
  const names = Array.from({ length: right - left + 1 }, (_, i) => (t.columns ? t.columns[i] ?? "" : t.headers ? show(page.sheet.value(keyOf({ col: left + i, row: top }))) : String.fromCharCode(65 + left + i)));
  const at = names.findIndex((n) => n.trim().toLowerCase() === sl.column.trim().toLowerCase());
  if (at < 0) return [];
  const [d1, d2] = dataRows(t), out: [number, string][] = [];
  for (let r = d1; r <= d2; r++) out.push([r, show(page.sheet.value(keyOf({ col: left + at, row: r })))]);
  return out;
}

// Its values, each once (any case), in the order first seen.
export function slicerValues(page: Page, sl: Slicer): string[] {
  const out: string[] = [];
  for (const [, v] of slicerColumn(page, sl)) if (!out.some((x) => x.toLowerCase() === v.toLowerCase())) out.push(v);
  return out;
}

// The rows the page's slicers hide: a table's data row whose value is not
// among a slicer's picked ones (each slicer with some picked).
export function slicedRows(page: Page): Set<number> {
  const out = new Set<number>();
  for (const sl of page.slicers) {
    if (!sl.picked.length) continue;
    const picked = new Set(sl.picked.map((x) => x.toLowerCase()));
    for (const [r, v] of slicerColumn(page, sl)) if (!picked.has(v.toLowerCase())) out.add(r);
  }
  return out;
}
export function bandedKeys(t: TableDef): string[] {
  const [d1, d2] = dataRows(t), left = Math.min(t.range[0].col, t.range[1].col), right = Math.max(t.range[0].col, t.range[1].col);
  const out: string[] = [];
  for (let r = d1; r <= d2; r += 2) for (let c = left; c <= right; c++) out.push(keyOf({ col: c, row: r }));
  return out;
}

export class Workbook {
  // The cell styles by name (styles.ts's gallery, then the workbook's own
  // and their edits): a cell names one, drawn under its own look.
  styles = new Map<string, Look>(Object.entries(CELL_STYLES));
  // What the last file opened held that the app left out (file.report).
  fileReport: string[] = [];
  // The parts of the XLSX it was opened from that a file written from it
  // carries over (its document properties and theme).
  carried: [string, Uint8Array][] = [];
  // Its theme (themes.ts): the colour roles and fonts looks may name; the
  // page's, set as the workbook's is.
  get theme(): Theme { return theme(); }
  set theme(t: Theme) { setTheme(t); }
  readonly rexx = new RexxLanguage();
  readonly #turns = new Turns();
  // How Excel formulas are shown and typed (a view of them: they are
  // stored A1 either way, and RexxJS formulas are shown as written).
  referenceStyle: "A1" | "R1C1" = "A1";
  // How formulas are worked out: automatically or by hand, and whether a
  // circular reference is iterated (Calc); saved with the file.
  calculation: Calc = AUTOMATIC;
  #pages: Page[] = [];
  #active: Page;
  // Each sheet's cells (by its name in upper case) whose value changed
  // since the last settle().
  #pending = new Map<string, Set<string>>();
  name = "Untitled";
  setupScript = "";
  clip: Clip | null = null;
  #conditional = new Map<string, string>();
  // Circle Invalid: on, the cells whose input breaks their rule are marked.
  circling = false;
  #circled = new Set<string>();
  // The date system serials count from (1900, or 1904: dates.ts), saved
  // with the file. Set through setDateSystem, which works everything out.
  #dateSystem: DateSystem = 1900;
  get dateSystem(): DateSystem { return this.#dateSystem; }
  async setDateSystem(system: DateSystem): Promise<void> {
    this.#dateSystem = system;
    setDates(system);
    await this.#recalculate();
    await this.settle();
  }
  // The selection as a range ("B2:C4"), as the page last had it or a
  // command asked for (then `selecting`, for the page to make it so).
  selection = "A1";
  selecting = false;
  // How a URL is handed to the system (a hyperlink followed): "" when it
  // was, or why not. The app's view sets it to sae's shell.open.
  opener: (url: string) => string = () => "no way to open a URL here";
  // Split: a second pane over the sheet (a view).
  split = false;
  // The zoom, in percent (10 to 400, as Excel's): a view of the sheet.
  zoom = 100;
  // The panel open under the grid (menus.ts's PANELS), or null: a view.
  panel: string | null = null;
  // Show Formulas: the grid shows inputs rather than values.
  showFormulas = false;
  #traces: Trace[] = [];
  #traced = new Map<string, string>();
  #marks = new Map<string, Mark>();
  // A formula being typed: each cell its references cover, outlined.
  #typed = new Map<string, string>();
  #watches: { readonly page: Page; readonly cell: string }[] = [];
  // Undo: the workbook as it was before each change, and as it was before
  // each undo, to redo.
  readonly #undo: Saved[] = [];
  readonly #redo: Saved[] = [];
  // A group of changes made one step (a script): the workbook before it,
  // the lowest the undo stack went while it ran (an UNDO in it may take
  // back steps from before it), and the labels of the changes it made.
  #group: { depth: number; readonly before: Saved; low: number; readonly labels: string[]; undone: number } | null = null;
  #names = new Map<string, BookName>();

  constructor() {
    setDates(1900);
    this.#active = this.#make("Sheet1", EXCEL);
    this.#pages.push(this.#active);
  }

  // The active sheet, and its notes.
  get sheet(): Sheet { return this.#active.sheet; }
  get notes(): CellNotes { return this.#active.notes; }
  get active(): string { return this.#active.name; }

  sheetNames(): string[] { return this.#pages.map((p) => p.name); }

  // A cell's input as the reference style shows it.
  shownInput(key: string): string {
    const input = this.sheet.input(key);
    return this.referenceStyle === "R1C1" && this.sheet.language.name === "Excel" ? toR1C1(input, key) : input;
  }

  // What was typed into a cell, in the reference style, as it is stored.
  enteredInput(key: string, text: string): string {
    return this.referenceStyle === "R1C1" && this.sheet.language.name === "Excel" ? fromR1C1(text, key) : text;
  }
  pages(): readonly Page[] { return this.#pages; }

  // A sheet by its name, any case.
  page(name: string): Page | undefined {
    return this.#pages.find((p) => p.name.toUpperCase() === name.toUpperCase());
  }

  #make(name: string, language: Language): Page {
    // The page's name as it is now: a sheet can be renamed.
    let page: Page | null = null;
    const upper = () => (page?.name ?? name).toUpperCase();
    const context: Context = {
      get name() { return upper(); }, turns: this.#turns, value: (key) => this.#valueOf(key),
      named: (n) => {
        // Another sheet's own name: COSTS!QTY.
        if (n.includes("!")) {
          const { sheet, cell: name } = splitKey(n);
          const there = sheet === null ? undefined : this.page(sheet);
          const r = there?.sheet.name(name) ?? null;
          return there && r ? { range: r, sheet: there === page ? null : there.name.toUpperCase() } : null;
        }
        const b = this.#names.get(n);
        return b && "range" in b ? { range: b.range, sheet: b.page === page ? null : b.page.name.toUpperCase() } : null;
      },
      formula: (n) => {
        const b = this.#names.get(n);
        return b && "formula" in b ? b.formula : null;
      },
      input: (key) => {
        const { sheet, cell } = splitKey(key);
        return this.page(sheet ?? "")?.sheet.input(cell) ?? "";
      },
      sheets: () => this.sheetNames(),
      calc: () => this.calculation,
      hidden: (key) => {
        const { sheet, cell } = splitKey(key), p = this.page(sheet ?? ""), row = cellAddr(cell)?.row;
        if (!p || row === undefined) return null;
        return p.filtered?.has(row) ? "filter" : p.hiddenRows.has(row) ? "manual" : null;
      },
    };
    const changed = (key: string) => {
      let keys = this.#pending.get(upper());
      if (!keys) this.#pending.set(upper(), (keys = new Set()));
      keys.add(key);
    };
    page = { name, sheet: new Sheet({ changed }, language, context), notes: new CellNotes(), rules: [], charts: [],
      hiddenRows: new Set(), hiddenCols: new Set(), filtered: null, filter: null, frozen: 0, frozenCols: 0, merges: [], heights: new Map(), widths: new Map(), autoId: null, pivots: new Map(), tabColor: null, hiddenSheet: false, rowLevels: new Map(), colLevels: new Map(), slicers: [], scenarios: [] };
    return page;
  }

  // A cell on another sheet ("COSTS!B2"), or #REF! when there is no such sheet.
  #valueOf(key: string): Value {
    const { sheet, cell } = splitKey(key);
    const p = this.page(sheet ?? "");
    return p ? p.sheet.value(cell) : new CellError("#REF!", `no sheet ${sheet}`);
  }

  // A sheet added after the others, and made the active one; what read a
  // sheet of its name is worked out again.
  async add(name: string): Promise<void> {
    this.#active = this.#make(name, this.sheet.language);
    this.#pages.push(this.#active);
    await this.#recalculate();
  }

  // A copy of a sheet, put after it and made active: its cells, notes,
  // names and tables, and how it looks and is viewed.
  async copy(from: Page, name: string): Promise<void> {
    const saved = from.sheet.saved();
    this.#active = this.#make(name, from.sheet.language);
    this.#pages.splice(this.#pages.indexOf(from) + 1, 0, this.#active);
    const to = this.#active;
    Object.assign(to, viewOf(from), { hiddenSheet: false });
    to.notes.restore(from.notes.saved());
    await to.sheet.restore(saved);
    await this.#recalculate();
    await this.settle();
  }

  activate(name: string): void {
    this.#active = this.page(name) ?? this.#active;
  }

  // A sheet deleted (the one before it active, if it was); what read it
  // reads #REF!.
  // A sheet renamed, and every formula that read it by its old name made
  // to read it by the new one (in either language); what read a sheet of
  // the new name is worked out again. Refused, changing nothing, when a
  // reference cannot be written with the new name (RangeError).
  async rename(from: string, to: string): Promise<void> {
    const page = this.page(from)!, old = page.name.toUpperCase();
    const plan: [Page, [string, string][]][] = this.#pages.map((p) => [p, p.sheet.keys().flatMap((k): [string, string][] => {
      const input = p.sheet.input(k), out = p.sheet.language.renamed(input, old, to);
      if (out === null) throw new RangeError(`${p.name}!${k} reads ${page.name} as ${p.sheet.language.name} writes it, and ${to} cannot be written so`);
      return out === input ? [] : [[k, out]];
    })]);
    page.name = to;
    for (const [p, entries] of plan) if (entries.length) await p.sheet.setMany(entries);
    await this.#recalculate();
  }

  async remove(name: string): Promise<void> {
    const at = this.#pages.findIndex((p) => p.name.toUpperCase() === name.toUpperCase());
    if (at < 0) return;
    const [gone] = this.#pages.splice(at, 1);
    if (gone === this.#active) this.#active = this.#pages[Math.max(0, at - 1)];
    this.#pending.delete(gone.name.toUpperCase());
    await this.#recalculate();
  }

  // The workbook's names: a range as [NAME, "Sheet1!B2:B4"], a formula as
  // [NAME, "=LAMBDA(x, x*2)"].
  names(): [string, string][] {
    return [...this.#names].sort(([x], [y]) => (x < y ? -1 : 1))
      .map(([n, b]): [string, string] => [n, "formula" in b ? `=${b.formula}` : `${sheetPrefix(b.page.name)}${rangeName(b.range)}`]);
  }

  bookName(n: string): BookName | null {
    return this.#names.get(n.toUpperCase()) ?? null;
  }

  // A name for the whole workbook defined (or redefined), or deleted
  // (null): every sheet worked out again, as any formula may read it.
  // A sheet's own names changed: every other sheet compiled again, as a
  // formula there may name one (Costs!QTY).
  async ownNamesChanged(): Promise<void> {
    for (const p of this.#pages) if (p !== this.#active) await p.sheet.recalculate();
    await this.settle();
  }

  async defineName(n: string, b: BookName | null): Promise<void> {
    if (b) this.#names.set(n.toUpperCase(), b); else this.#names.delete(n.toUpperCase());
    await this.#recalculate();
  }

  // A sheet moved to `at` (0-based) among the others; what is active
  // stays. Every sheet worked out again: a formula may ask a sheet's
  // number (SHEET).
  async move(name: string, at: number): Promise<void> {
    const from = this.#pages.findIndex((p) => p.name.toUpperCase() === name.toUpperCase());
    if (from < 0) return;
    const [page] = this.#pages.splice(from, 1);
    this.#pages.splice(Math.max(0, Math.min(at, this.#pages.length)), 0, page);
    await this.#recalculate();
  }

  // One empty sheet, as a file opened starts; where undo stops.
  reset(name = "Sheet1"): void {
    this.#active = this.#make(name, this.sheet.language);
    this.#pages = [this.#active];
    this.#names = new Map();
    this.calculation = AUTOMATIC;
    this.#dateSystem = 1900;
    setDates(1900);
    this.styles = new Map(Object.entries(CELL_STYLES));
    this.theme = THEMES.Office;
    this.carried = [];
    this.#pending.clear();
    this.#undo.length = 0;
    this.#redo.length = 0;
  }

  // The workbook as it is now, for changed() once a change is made.
  saved(): Saved {
    return {
      pages: this.#pages.map((p) => ({ page: p, name: p.name, sheet: p.sheet.saved(), notes: p.notes.saved(), view: viewOf(p) })),
      names: new Map(this.#names),
      active: this.#active, name: this.name, setupScript: this.setupScript, styles: new Map(this.styles), theme: this.theme,
    };
  }

  // A change, `label`, was made since `before` was saved: one step to
  // undo, and nothing left to redo.
  changed(before: Saved, label: string): void {
    this.#undo.push({ ...before, label });
    if (this.#undo.length > HISTORY) this.#undo.shift();
    this.#redo.length = 0;
    this.#group?.labels.push(label);
  }

  // The changes made until endGroup() are one step of undo (a script): each
  // is still its own step while it runs, so an UNDO in the script takes
  // back the one before it; at the end they are folded into one, back to
  // the workbook before the script. Groups nest; the outermost counts.
  beginGroup(): void {
    if (this.#group) this.#group.depth++;
    else this.#group = { depth: 1, before: this.saved(), low: this.#undo.length, labels: [], undone: 0 };
  }

  // True when the outermost group ended folded into one step: it is the
  // newest step of undo now. A group that only undid and redid is not
  // folded: those stand as the steps they took, Redo with them, so UNDO()
  // sent twice goes back two steps rather than back and forth. `whole`
  // folds whatever it did, to be taken back as one (a failed request).
  endGroup(whole = false, always = false): boolean {
    const g = this.#group;
    if (!g || --g.depth > 0) return false;
    this.#group = null;
    if (!always && !g.labels.length && !g.undone) return false;
    if (!always && !g.labels.length && !whole) return false;
    this.#undo.length = Math.min(g.low, this.#undo.length);
    this.#undo.push({ ...g.before, label: g.labels.length === 1 && !g.undone ? g.labels[0] : "Script" });
    if (this.#undo.length > HISTORY) this.#undo.shift();
    this.#redo.length = 0;
    return true;
  }

  // The newest step of Redo let go of: what an Undo that only took back a
  // failed request leaves there, which is nothing to do again.
  forgetRedo(): void { this.#redo.pop(); }

  // A group whose runner was stopped before it could end it (sae's time
  // limit interrupts a page's work, and no finally runs): ended now, however
  // deep. `whole`: folded into one step back to the workbook before it
  // whatever it recorded, as a command can be stopped between changing a
  // cell and recording the step.
  endStoppedGroup(whole = false): boolean {
    if (!this.#group) return false;
    this.#group.depth = 1;
    return this.endGroup(whole, whole);
  }

  get canUndo(): boolean { return this.#undo.length > 0; }
  get canRedo(): boolean { return this.#redo.length > 0; }
  // What Undo and Redo would take back or do again ("Merge Cells"), or null.
  get undoLabel(): string | null { return this.#undo.at(-1)?.label ?? null; }
  get redoLabel(): string | null { return this.#redo.at(-1)?.label ?? null; }

  // The last change taken back, whatever it was and on whichever sheet (that
  // sheet active again), or done again. Each gives the active sheet's
  // keys whose shown value changed (all of them, if it is another sheet).
  undo(): Promise<string[]> { return this.#step(this.#undo, this.#redo); }
  redo(): Promise<string[]> { return this.#step(this.#redo, this.#undo); }

  async #step(from: Saved[], to: Saved[]): Promise<string[]> {
    const step = from.pop();
    if (!step) return [];
    to.push({ ...this.saved(), label: step.label });
    if (this.#group) {
      this.#group.undone++;
      this.#group.low = Math.min(this.#group.low, this.#undo.length);
    }
    const was = this.#active, before = this.sheet.keys();
    const ownNames = () => this.#pages.map((p) => JSON.stringify(p.sheet.names())).join("|");
    const namesBefore = ownNames();
    // Sheets added, deleted, moved or renamed: every sheet worked out again,
    // as what reads one by name may read another now.
    const sameNames = step.names.size === this.#names.size && [...step.names].every(([n, b]) => this.#names.get(n) === b);
    const moved = !sameNames || step.pages.length !== this.#pages.length || step.pages.some((s, i) => s.page !== this.#pages[i] || s.name !== s.page.name);
    this.#pages = step.pages.map((s) => s.page);
    this.#names = new Map(step.names);
    this.#active = step.active;
    this.name = step.name;
    this.setupScript = step.setupScript;
    this.styles = new Map(step.styles);
    this.theme = step.theme;
    let changed: string[] = [];
    for (const s of step.pages) {
      s.page.name = s.name;
      Object.assign(s.page, viewOf(s.view));
      s.page.notes.restore(s.notes);
      const c = await s.page.sheet.restore(s.sheet);
      if (s.page === this.#active) changed = c;
    }
    // A sheet's own names put back: other sheets may name them (Costs!QTY).
    if (moved || ownNames() !== namesBefore) {
      for (const p of this.#pages) {
        const c = await p.sheet.recalculate();
        if (p === this.#active) changed = [...new Set([...changed, ...c])];
      }
    }
    await this.settle();
    return this.#active === was ? changed : [...new Set([...before, ...this.sheet.keys()])];
  }

  // The major setting, for every sheet.
  async setLanguage(language: Language): Promise<string[]> {
    let changed: string[] = [];
    for (const p of this.#pages) {
      const c = await p.sheet.setLanguage(language);
      if (p === this.#active) changed = c;
    }
    return changed;
  }

  async #recalculate(): Promise<void> {
    for (const p of this.#pages) await p.sheet.recalculate();
  }

  // Every sheet's cells out of date under manual calculation, qualified.
  stale(): string[] {
    return this.#pages.flatMap((p) => p.sheet.stale().map((k) => `${p.name}!${k}`));
  }

  // Rows hidden or shown (by hand or a filter): every sheet worked out
  // again, as SUBTOTAL and AGGREGATE skip hidden rows.
  async hiddenChanged(): Promise<void> {
    await this.#recalculate();
  }

  language(name: LanguageName): Language {
    return name === "Excel" ? EXCEL : this.rexx;
  }

  // After a change: what reads, on other sheets, the cells that changed
  // worked out again (and what reads those, in turn), then the active
  // sheet's styles.
  async settle(): Promise<void> {
    await this.#crossCycles();
    // Under iteration a cycle through sheets is worked round by these
    // rounds, up to its most iterations, until nothing changes.
    const rounds = Math.max(ROUNDS, this.calculation.iterate?.max ?? 0);
    for (let round = 0; round < rounds && this.#pending.size; round++) {
      const pending = this.#pending;
      this.#pending = new Map();
      for (const [from, keys] of pending) {
        const touched = [...keys].map((k) => `${from}!${k}`);
        for (const p of this.#pages) if (p.name.toUpperCase() !== from) await p.sheet.touch(touched);
      }
    }
    this.#pending.clear();
    await this.#restyle();
  }

  // Cycles through more than one sheet (a sheet finds those within it):
  // the strongly connected cells of the workbook's graph, searched from the
  // cells that read another sheet (a cycle across sheets has one), each
  // marked #CYCLE! with a path round it, on its own sheet, and what reads it
  // worked out again. A cell already so marked is not marked again.
  async #crossCycles(): Promise<void> {
    const id = (p: Page, c: string) => `${p.name.toUpperCase()}!${c}`;
    const starts: [Page, string][] = this.#pages.flatMap((p) =>
      p.sheet.keys().filter((k) => p.sheet.reads(k).some((r) => r.includes("!"))).map((k): [Page, string] => [p, k]));
    if (!starts.length && !this.#pages.some((p) => p.sheet.cycling)) return;
    // Tarjan's strongly connected components over what each cell reads.
    const index = new Map<string, number>(), low = new Map<string, number>(), cell = new Map<string, [Page, string]>();
    const stack: string[] = [], onStack = new Set<string>(), groups: string[][] = [];
    let next = 0;
    const visit = (p: Page, c: string): void => {
      const v = id(p, c);
      index.set(v, next); low.set(v, next); next++;
      cell.set(v, [p, c]);
      stack.push(v); onStack.add(v);
      for (const [q, d] of this.#neighbours(p, c, "precedents")) {
        const w = id(q, d);
        if (!index.has(w)) {
          visit(q, d);
          low.set(v, Math.min(low.get(v)!, low.get(w)!));
        } else if (onStack.has(w)) low.set(v, Math.min(low.get(v)!, index.get(w)!));
      }
      if (low.get(v) !== index.get(v)) return;
      const group: string[] = [];
      for (let w = stack.pop()!; ; w = stack.pop()!) {
        onStack.delete(w);
        group.push(w);
        if (w === v) break;
      }
      if (new Set(group.map((w) => cell.get(w)![0])).size > 1) groups.push(group);
    };
    for (const [p, c] of starts) if (!index.has(id(p, c))) visit(p, c);
    // A path round the cycle from each cell, within its group.
    const marks = new Map<Page, Map<string, string>>();
    for (const group of this.calculation.iterate ? [] : groups) {
      const members = new Set(group);
      for (const v of group) {
        const [p, c] = cell.get(v)!;
        const path = this.#pathRound(v, members, cell);
        if (!marks.has(p)) marks.set(p, new Map());
        marks.get(p)!.set(c, `a circular reference through more than one sheet: ${path.join(" → ")}`);
      }
    }
    // Every sheet told which of its cells are on one now (none, for one
    // whose cycle was broken); what changed reaches the other sheets through
    // settle()'s rounds, as any change does.
    for (const p of this.#pages) if (marks.has(p) || p.sheet.cycling) await p.sheet.cycle(marks.get(p) ?? new Map());
  }

  // The cells from `v` round a cycle back to it, staying within `members`.
  #pathRound(v: string, members: ReadonlySet<string>, cell: ReadonlyMap<string, [Page, string]>): string[] {
    const path = [v], seen = new Set([v]);
    const walk = (w: string): boolean => {
      const [p, c] = cell.get(w)!;
      for (const [q, d] of this.#neighbours(p, c, "precedents")) {
        const u = `${q.name.toUpperCase()}!${d}`;
        if (u === v) { path.push(v); return true; }
        if (!members.has(u) || seen.has(u)) continue;
        seen.add(u); path.push(u);
        if (walk(u)) return true;
        path.pop();
      }
      return false;
    };
    walk(v);
    return path;
  }

  // A cell's neighbours in the graph, on any sheet: what it reads, or what
  // reads it.
  #neighbours(page: Page, cell: string, direction: Direction): [Page, string][] {
    if (direction === "precedents") {
      return page.sheet.reads(cell).flatMap((k): [Page, string][] => {
        const { sheet, cell: c } = splitKey(k);
        const p = sheet === null ? page : this.page(sheet);
        return p ? [[p, c]] : [];
      });
    }
    const here = page.name.toUpperCase();
    return [
      ...page.sheet.readBy(cell).map((k): [Page, string] => [page, k]),
      ...this.#pages.filter((p) => p !== page).flatMap((p) => p.sheet.readBy(`${here}!${cell}`).map((k): [Page, string] => [p, k])),
    ];
  }

  // The cells `levels` steps from a cell, a level at a time (each cell once).
  #levels(t: Trace): [Page, string][][] {
    const seen = new Set([`${t.page.name}!${t.cell}`]);
    const out: [Page, string][][] = [];
    let frontier: [Page, string][] = [[t.page, t.cell]];
    for (let i = 0; i < t.levels; i++) {
      const next = frontier.flatMap(([p, c]) => this.#neighbours(p, c, t.direction)).filter(([p, c]) => {
        const id = `${p.name}!${c}`;
        return !seen.has(id) && !!seen.add(id);
      });
      if (!next.length) break;
      out.push(next);
      frontier = next;
    }
    return out;
  }

  // A cell's precedents or dependents traced on the active sheet, a level
  // further than the last time it was; each level's cells, those on another
  // sheet named with it ("Costs!B1").
  trace(cell: string, direction: Direction): string[][] {
    let t = this.#traces.find((x) => x.page === this.#active && x.cell === cell && x.direction === direction);
    if (t) {
      if (this.#levels(t).length === t.levels) t.levels++;
    } else {
      this.#traces.push((t = { page: this.#active, cell, direction, levels: 1 }));
    }
    return this.#levels(t).map((level) => level.map(([p, c]) => (p === this.#active ? c : `${p.name}!${c}`)));
  }

  clearTraces(): void {
    this.#traces = [];
  }

  // A cell watched (the Watch Window), on its sheet; false if it already was.
  watch(page: Page, cell: string): boolean {
    if (this.#watches.some((w) => w.page === page && w.cell === cell)) return false;
    this.#watches.push({ page, cell });
    return true;
  }

  unwatch(page: Page, cell: string): boolean {
    const n = this.#watches.length;
    this.#watches = this.#watches.filter((w) => w.page !== page || w.cell !== cell);
    return this.#watches.length < n;
  }

  // The watched cells whose sheets are still here, as they are now.
  watches(): Watch[] {
    this.#watches = this.#watches.filter((w) => this.#pages.includes(w.page));
    return this.#watches.map(({ page, cell }) => ({ sheet: page.name, cell, input: page.sheet.input(cell), shown: this.#shown(page, cell) }));
  }

  // The active sheet's slicers, their tiles each a value of the column (in
  // the order first seen), picked or not.
  slicerViews(): SlicerView[] {
    return this.#active.slicers.map((sl) => ({ name: sl.name, caption: `${sl.name} (${sl.table}[${sl.column}])`,
      tiles: slicerValues(this.#active, sl).map((value) => ({ value, picked: sl.picked.some((x) => x.toLowerCase() === value.toLowerCase()) })) }));
  }

  // The active sheet's charts, laid out from its cells as they are now.
  charts(): ChartView[] {
    const page = this.#active;
    return page.charts.map((c) => {
      const written = c.pivot ? page.pivots.get(c.pivot)?.written : null, r = written ?? c.range;
      let data = chartData(r, (k) => page.sheet.value(k), c.options?.byRows);
      if (c.pivot) data = { ...data, labels: data.labels.filter((l) => l !== "Grand Total"), series: data.series.filter((x) => x.name !== "Grand Total").map((x) => ({ ...x, values: x.values.filter((_, i) => data.labels[i] !== "Grand Total") })) };
      return { kind: c.kind, title: c.title, range: c.pivot ? `pivot ${c.pivot}` : rangeName(r), data, shapes: layout(c.kind, data, c.title, CHART_W, CHART_H, c.options ?? {}) };
    });
  }

  // Every conditional style worked out again, in the sheet's language (it
  // may read any cell); what is traced, marked.
  async #restyle(): Promise<void> {
    const next = new Map<string, string>();
    for (const [key, expr] of this.notes.conditional()) next.set(key, toStyle(await this.sheet.evaluate(expr)));
    this.#conditional = next;
    this.#circled = new Set(this.circling ? this.invalid() : []);
    const page = this.#active;
    // A formula rule's formula, moved from the range's top-left to each
    // cell as a copy would be, true there or not.
    const hits = new Map<CfRule, Set<string>>();
    for (const rule of page.rules) {
      if (rule.kind !== "formula") continue;
      const [a, b] = rule.range, top = { col: Math.min(a.col, b.col), row: Math.min(a.row, b.row) }, yes = new Set<string>();
      for (const k of range(a, b)) {
        const at = addrOf(k)!, moved = movedBy(page.sheet.language, `=${rule.formula.replace(/^=/, "")}`, { cols: at.col - top.col, rows: at.row - top.row });
        const v = await page.sheet.evaluate(moved.slice(1));
        if (v === true || (typeof v === "number" && v !== 0)) yes.add(k);
      }
      hits.set(rule, yes);
    }
    this.#marks = applied(page.rules, (k) => page.sheet.value(k), hits);
    this.#traces = this.#traces.filter((t) => this.#pages.includes(t.page));
    this.#traced = new Map();
    for (const t of this.#traces) {
      for (const level of this.#levels(t)) for (const [p, c] of level) if (p === this.#active) this.#traced.set(c, TRACED[t.direction]);
    }
  }

  // The cells whose input breaks their rule, whenever it was typed, in sheet
  // order.
  invalid(): string[] {
    return this.sheet.keys().filter((k) => {
      const rule = this.notes.get(k).validation;
      return rule && check({ ...rule, when: "always" }, this.sheet.input(k), false) !== null;
    });
  }

  // What a cell shows: its value in its number format (after its icon, if a
  // rule gives it one), or with Show Formulas on, its input.
  textOf(key: string): string {
    if (this.showFormulas) return this.sheet.input(key);
    const mark = this.#marks.get(key), icon = mark?.icon, text = mark?.hide ? "" : this.#shown(this.#active, key);
    return icon ? (text ? `${icon} ${text}` : icon) : text;
  }

  // A cell's look: its cell style's under its own.
  lookOf(note: Note): Look | undefined {
    return under(note.cellStyle ? this.styles.get(note.cellStyle) : undefined, note.look);
  }

  #shown(page: Page, key: string): string {
    const v = page.sheet.value(key), note = page.notes.get(key);
    const shown = formatted(v, note.format || this.lookOf(note)?.number || "") ?? show(v);
    return note.editor ? drawn(note.editor, v, shown) : shown;
  }

  // Its style, each layer over the last: numbers to the right, its own
  // style and look, its format's style parts, its conditional style, its
  // rules', a trace, a circle (its input breaks its rule), an error in red,
  // and a comment's corner marker.
  styleOf(key: string): string {
    const v = this.sheet.value(key), n = this.notes.get(key);
    const table = this.#tableLook(key);
    const zoomed = this.zoom === 100 ? "" : `font-size: ${Math.round((n.look?.size ?? 13) * this.zoom / 100)}px`;
    return [typeof v === "number" ? "text-align: end" : "", table, toStyle(n.style), lookStyle(this.lookOf(n)), formatStyle(v, n.format ?? ""), this.#conditional.get(key) ?? "", resolvedStyle(this.#marks.get(key)?.style ?? ""), this.#traced.get(key) ?? "",
      this.#circled.has(key) ? "background: #ffd6d6" : "", v instanceof CellError ? "color: #c0392b" : "",
      n.comment ? "--corner: #d00000" : "", this.#typed.get(key) ?? "", zoomed].filter((p) => p).join("; ");
  }

  // What a table's style gives a cell of it: its banded rows shaded, a
  // totals row bold over a line.
  #tableLook(key: string): string {
    const at = addrOf(key);
    if (!at) return "";
    for (const t of this.sheet.tables()) {
      const [a, b] = t.range, left = Math.min(a.col, b.col), right = Math.max(a.col, b.col), top = Math.min(a.row, b.row), bottom = Math.max(a.row, b.row);
      if (at.col < left || at.col > right || at.row < top || at.row > bottom) continue;
      if (!t.banded && !t.totals && !t.style) return "";
      const st = TABLE_STYLES[t.style ?? "Light"] ?? TABLE_STYLES.Light, line = resolved(st.line), parts: string[] = [];
      const [d1, d2] = dataRows(t);
      if (t.headers && at.row === top && t.style) {
        if (st.header.fill) parts.push(`background: ${resolved(st.header.fill)}`);
        if (st.header.color) parts.push(`color: ${resolved(st.header.color)}`);
        parts.push("font-weight: bold", `border-bottom: 1px solid ${line}`);
      } else if (t.totals && at.row === bottom) parts.push("font-weight: bold", `border-top: 1px solid ${line}`);
      else if (at.row >= d1 && at.row <= d2) {
        if ((t.banded && (at.row - d1) % 2 === 0) || (t.bandedColumns && (at.col - left) % 2 === 0)) parts.push(`background: ${resolved(st.band)}`);
        if ((t.firstColumn && at.col === left) || (t.lastColumn && at.col === right)) parts.push("font-weight: bold");
      }
      return parts.join("; ");
    }
    return "";
  }

  // The active sheet's rows (by hand or by its filter) and columns out of
  // view, in order.
  hidden(): { rows: number[]; cols: number[] } {
    const p = this.#active, up = (a: number, b: number) => a - b;
    return { rows: [...new Set([...p.hiddenRows, ...(p.filtered ?? [])])].sort(up), cols: [...p.hiddenCols].sort(up) };
  }

  frozenRows(): number {
    return this.#active.frozen;
  }

  // Its columns frozen at the left (Freeze Panes, Freeze First Column).
  frozenCols(): number {
    return this.#active.frozenCols;
  }

  // The active sheet's merged ranges ("A1:C1"), in sheet order, and its
  // rows set to a height ([row, px], 0-based, in order).
  merges(): string[] {
    return this.#active.merges.map(rangeName).sort((a, b) => order(parseRange(a)![0]) - order(parseRange(b)![0]));
  }

  rowHeights(): [number, number][] {
    return [...this.#active.heights].sort(([a], [b]) => a - b);
  }

  // Its columns set to a width ([column, px], 0-based, in order).
  colWidths(): [number, number][] {
    return [...this.#active.widths].sort(([a], [b]) => a - b);
  }

  // As the grid draws them at the zoom: every column and row scaled from
  // its own size or the grid's (`col`, `row`), or, at 100%, those set.
  shownColWidths(col: number): [number, number][] {
    if (this.zoom === 100) return this.colWidths();
    return Array.from({ length: COLS }, (_, c): [number, number] => [c, Math.round((this.#active.widths.get(c) ?? col) * this.zoom / 100)]);
  }
  shownRowHeights(row: number): [number, number][] {
    if (this.zoom === 100) return this.rowHeights();
    return Array.from({ length: ROWS }, (_, r): [number, number] => [r, Math.round((this.#active.heights.get(r) ?? row) * this.zoom / 100)]);
  }

  // A formula being typed in the active sheet (null once it is entered or
  // not): each reference it names outlined in its colour, as Excel does.
  typing(text: string | null): void {
    this.#typed = new Map();
    if (!text?.startsWith("=")) return;
    this.sheet.language.references(text.slice(1)).forEach((ref, i) => {
      const r = parseRange(ref);
      if (!r) return;   // on another sheet, or off this one
      const colour = REFERENCE_COLOURS[i % REFERENCE_COLOURS.length];
      const [a, b] = r, top = Math.min(a.row, b.row), bottom = Math.max(a.row, b.row), left = Math.min(a.col, b.col), right = Math.max(a.col, b.col);
      for (let row = top; row <= bottom; row++) {
        for (let col = left; col <= right; col++) {
          const edges = [row === top ? "top" : "", col === right ? "right" : "", row === bottom ? "bottom" : "", col === left ? "left" : ""];
          this.#typed.set(keyOf({ col, row }), edges.filter((e) => e).map((e) => `border-${e}: 2px solid ${colour}`).join("; "));
        }
      }
    });
  }

  // What shows when the pointer rests on a cell: its comment.
  // A cell's tooltip: its comment, and its rule's input prompt.
  tooltipOf(key: string): string {
    const n = this.notes.get(key);
    const link = n.hyperlink ? n.hyperlink.tip || `${n.hyperlink.target.replace(/^#/, "")}\nOpen Hyperlink to follow it` : "";
    return [link, n.comment ?? "", n.validation?.prompt ?? ""].filter((x) => x).join("\n");
  }

  // The rows (0-based) with a wrapped cell, to be as tall as their text.
  wrapped(): number[] {
    return [...new Set(this.notes.keys().filter((k) => this.notes.get(k).look?.wrap).map((k) => cellAddr(k)!.row))].sort((a, b) => a - b);
  }

  // The status line for a cell.
  // What the status bar says of a range, as Excel's: the average, count
  // (cells not empty), numerical count, min, max and sum of its numbers;
  // nothing but the count when it has no numbers.
  // Several areas (a multiple selection): each cell once.
  stats(r: [Addr, Addr] | readonly [Addr, Addr][]): Stats {
    const areas = (Array.isArray(r[0]) ? r : [r]) as readonly [Addr, Addr][];
    const keys = new Set<string>();
    for (const [a, b] of areas) {
      const lo = { col: Math.min(a.col, b.col), row: Math.min(a.row, b.row) }, hi = { col: Math.max(a.col, b.col), row: Math.max(a.row, b.row) };
      for (let row = lo.row; row <= hi.row; row++) for (let col = lo.col; col <= hi.col; col++) keys.add(keyOf({ col, row }));
    }
    const values: Value[] = [...keys].map((k) => this.sheet.value(k));
    const nums = values.filter((v): v is number => typeof v === "number");
    const sum = nums.reduce((t, n) => t + n, 0);
    return {
      count: values.filter((v) => v !== null && v !== "").length, numbers: nums.length,
      sum: nums.length ? sum : null, average: nums.length ? sum / nums.length : null,
      min: nums.length ? Math.min(...nums) : null, max: nums.length ? Math.max(...nums) : null,
    };
  }

  // The status line for a selected cell, with its rule's input prompt.
  describe(key: string): string {
    const prompt = this.notes.get(key).validation?.prompt;
    return describe(key, this.sheet.input(key), this.sheet.value(key), this.notes.get(key).comment) + (prompt ? ` · ${prompt}` : "");
  }
}
