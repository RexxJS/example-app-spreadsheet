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
import { cellAddr, keyOf, order, parseRange, splitKey } from "./address.ts";
import { REFERENCE_COLOURS } from "./typing.ts";
import type { Addr } from "./address.ts";
import type { PivotFields } from "./pivot.ts";
import { EXCEL } from "./excel.ts";
import { describe, show } from "./format.ts";
import { formatStyle, formatted, lookStyle, toStyle } from "./formats.ts";
import type { Language } from "./language.ts";
import { CellNotes } from "./notes.ts";
import type { Note } from "./notes.ts";
import { RexxLanguage } from "./rexx.ts";
import { Sheet, Turns } from "./sheet.ts";
import type { Context, SavedSheet } from "./sheet.ts";
import { CellError } from "./values.ts";
import type { Value } from "./values.ts";
import { applied } from "./rules.ts";
import { chartData, layout } from "./charts.ts";
import type { ChartData, ChartKind, Shape } from "./charts.ts";
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
  merges: [Addr, Addr][];
  heights: Map<number, number>;
  widths: Map<number, number>;
  autoId: AutoId | null;
  pivots: Map<string, Pivot>;
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
  filtered: p.filtered && new Set(p.filtered), filter: p.filter, frozen: p.frozen, merges: [...p.merges],
  heights: new Map(p.heights), widths: new Map(p.widths), autoId: p.autoId, pivots: new Map(p.pivots),
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

export class Workbook {
  readonly rexx = new RexxLanguage();
  readonly #turns = new Turns();
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
    this.#active = this.#make("Sheet1", EXCEL);
    this.#pages.push(this.#active);
  }

  // The active sheet, and its notes.
  get sheet(): Sheet { return this.#active.sheet; }
  get notes(): CellNotes { return this.#active.notes; }
  get active(): string { return this.#active.name; }

  sheetNames(): string[] { return this.#pages.map((p) => p.name); }
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
    };
    const changed = (key: string) => {
      let keys = this.#pending.get(upper());
      if (!keys) this.#pending.set(upper(), (keys = new Set()));
      keys.add(key);
    };
    page = { name, sheet: new Sheet({ changed }, language, context), notes: new CellNotes(), rules: [], charts: [],
      hiddenRows: new Set(), hiddenCols: new Set(), filtered: null, filter: null, frozen: 0, merges: [], heights: new Map(), widths: new Map(), autoId: null, pivots: new Map() };
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
    this.#pending.clear();
    this.#undo.length = 0;
    this.#redo.length = 0;
  }

  // The workbook as it is now, for changed() once a change is made.
  saved(): Saved {
    return {
      pages: this.#pages.map((p) => ({ page: p, name: p.name, sheet: p.sheet.saved(), notes: p.notes.saved(), view: viewOf(p) })),
      names: new Map(this.#names),
      active: this.#active, name: this.name, setupScript: this.setupScript,
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

  language(name: LanguageName): Language {
    return name === "Excel" ? EXCEL : this.rexx;
  }

  // After a change: what reads, on other sheets, the cells that changed
  // worked out again (and what reads those, in turn), then the active
  // sheet's styles.
  async settle(): Promise<void> {
    await this.#crossCycles();
    for (let round = 0; round < ROUNDS && this.#pending.size; round++) {
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
    for (const group of groups) {
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

  // The active sheet's charts, laid out from its cells as they are now.
  charts(): ChartView[] {
    const page = this.#active;
    return page.charts.map((c) => {
      const data = chartData(c.range, (k) => page.sheet.value(k));
      return { kind: c.kind, title: c.title, range: rangeName(c.range), data, shapes: layout(c.kind, data, c.title, CHART_W, CHART_H) };
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
    this.#marks = applied(page.rules, (k) => page.sheet.value(k));
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
    const icon = this.#marks.get(key)?.icon, text = this.#shown(this.#active, key);
    return icon ? `${icon} ${text}` : text;
  }

  #shown(page: Page, key: string): string {
    const v = page.sheet.value(key), note = page.notes.get(key);
    const shown = formatted(v, note.format ?? "") ?? show(v);
    return note.editor ? drawn(note.editor, v, shown) : shown;
  }

  // Its style, each layer over the last: numbers to the right, its own
  // style and look, its format's style parts, its conditional style, its
  // rules', a trace, a circle (its input breaks its rule), an error in red,
  // and a comment's corner marker.
  styleOf(key: string): string {
    const v = this.sheet.value(key), n = this.notes.get(key);
    return [typeof v === "number" ? "text-align: end" : "", toStyle(n.style), lookStyle(n.look), formatStyle(v, n.format ?? ""), this.#conditional.get(key) ?? "", this.#marks.get(key)?.style ?? "", this.#traced.get(key) ?? "",
      this.#circled.has(key) ? "background: #ffd6d6" : "", v instanceof CellError ? "color: #c0392b" : "",
      n.comment ? "--corner: #d00000" : "", this.#typed.get(key) ?? ""].filter((p) => p).join("; ");
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
  tooltipOf(key: string): string {
    return this.notes.get(key).comment ?? "";
  }

  // The rows (0-based) with a wrapped cell, to be as tall as their text.
  wrapped(): number[] {
    return [...new Set(this.notes.keys().filter((k) => this.notes.get(k).look?.wrap).map((k) => cellAddr(k)!.row))].sort((a, b) => a - b);
  }

  // The status line for a cell.
  describe(key: string): string {
    return describe(key, this.sheet.input(key), this.sheet.value(key), this.notes.get(key).comment);
  }
}
