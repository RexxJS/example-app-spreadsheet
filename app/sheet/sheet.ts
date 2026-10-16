// The sheet: what each cell was given, what it came to, and who reads whom.
// Setting cells recomputes them and everything downstream, in dependency
// order, once each, however many cells were set together (a paste). A cell
// on a cycle comes to #CYCLE!, and what reads it inherits the error. The
// sheet tells a listener which cells now show something else; it never draws.
// What a formula means is its Language's (the sheet's major setting, Excel
// or RexxJS); evaluation may be asynchronous, so setting cells is a promise,
// and two settings take turns. In a workbook (workbook.ts) a formula can read
// another sheet's cells ("COSTS!B2"): the sheet's Context answers them, and
// touch() works out again what reads them when they change; the sheets of a
// workbook share one Turns.
//
// A formula that comes to an array spills it into the cells below and to
// its right (its area): they show its values and are read as any cell, and
// each reads, for ordering, the formula it shows. A cell with an input in
// the area, or the sheet's edge, makes the formula #SPILL!; setting any cell
// of an area works its formula out again, so clearing the way spills it.
//
// A volatile formula (its language says: one calling OFFSET or INDIRECT)
// is worked out again on every change, and the cells its run-time
// references came to are linked as its reads after each run, so what it
// read orders it next time (and the auditor sees it).
import { COLS, ROWS, addrOf, cellAddr, keyOf, order, splitKey } from "./address.ts";
import { CellError, isArray } from "./values.ts";
import type { ArrayValue, Result, Value } from "./values.ts";
import { EXCEL } from "./excel.ts";
import type { Formula, Formulas, Hidden, Language, Named, Names, TableDef, Tables } from "./language.ts";
import type { Addr } from "./address.ts";
import { show } from "./format.ts";

export interface SheetListener {
  changed(key: string, value: Value): void;
}

export interface Problem {
  readonly key: string;
  readonly message: string;
}

// What a cell's input compiles to: a formula (or a parse error) in the
// sheet's language, or a constant.
export type Compiled = Formula | { readonly kind: "constant"; readonly value: Value };

const NUMERIC = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/;

export function compile(input: string, language: Language = EXCEL, names?: Names, tables?: Tables, formulas?: Formulas, at?: string): Compiled {
  if (input.startsWith("=")) return language.compile(input.slice(1), names, tables, formulas, at);
  if (input === "") return { kind: "constant", value: null };
  return { kind: "constant", value: NUMERIC.test(input) ? Number(input) : input };
}

const getOrAdd = <K, V>(m: Map<K, V>, k: K, make: () => V): V => {
  let v = m.get(k);
  if (v === undefined) m.set(k, (v = make()));
  return v;
};

const byOrder = (a: string, b: string): number => order(cellAddr(a)!) - order(cellAddr(b)!);

// This sheet's keys in sheet order, then other sheets' by name.
const byKey = (a: string, b: string): number => {
  const x = a.includes("!"), y = b.includes("!");
  return x !== y ? (x ? 1 : -1) : x ? (a < b ? -1 : a > b ? 1 : 0) : byOrder(a, b);
};

// One change at a time, across the sheets that share it: a language's
// evaluation may be asynchronous, and RexxJS's holds one interpreter.
export class Turns {
  #turn: Promise<unknown> = Promise.resolve();
  #busy = false;

  // Working a sheet out now: a formula is running (so whatever it calls
  // must not wait for a sheet to change, which waits for it).
  get busy(): boolean { return this.#busy; }

  take<T>(work: () => Promise<T>): Promise<T> {
    const done = this.#turn.then(async () => {
      this.#busy = true;
      try {
        return await work();
      } finally {
        this.#busy = false;
      }
    });
    this.#turn = done.catch(() => {});
    return done;
  }
}

// Where a sheet is: its name (upper case, as keys have it), the turns it
// shares, and the other sheets' cells' values.
// A sheet's contents as saved for undo (Sheet.saved, Sheet.restore).
export interface SavedSheet {
  readonly language: Language;
  readonly inputs: ReadonlyMap<string, string>;
  readonly names: ReadonlyMap<string, [Addr, Addr]>;
  readonly formulas: ReadonlyMap<string, string>;
  readonly tables: ReadonlyMap<string, TableDef>;
}

export interface Context {
  readonly name: string;
  readonly turns: Turns;
  value(key: string): Value;
  // The workbook's names, for a name the sheet has none of its own for.
  named?(name: string): Named | null;
  formula?(name: string): string | null;
  // Another sheet's cell's input ("" for none), and the workbook's sheet
  // names as given, in tab order (a sheet alone is its own book).
  input?(key: string): string;
  sheets?(): readonly string[];
  // Whether a cell's row (its key qualified) is hidden by hand, by a
  // filter, or not.
  hidden?(key: string): Hidden;
  // How the workbook calculates (Calc), when it is not automatically and
  // without iteration.
  calc?(): Calc;
}

// How formulas are worked out (the workbook's calculation settings, as
// Excel's Calculation Options): `manual`, a change works out only the cells
// set (a formula typed), what reads them waiting, stale, for Calculate;
// `iterate`, a circular reference is worked round, from the values its
// cells have, up to `max` times or until no cell of it moves by more than
// `change`, rather than being #CYCLE!.
export interface Calc {
  readonly manual: boolean;
  readonly iterate: { readonly max: number; readonly change: number } | null;
}
export const AUTOMATIC: Calc = { manual: false, iterate: null };

// A sheet on its own: no other sheet to read.
export const alone = (): Context => ({
  name: "", turns: new Turns(),
  value: (key) => new CellError("#REF!", `no sheet ${splitKey(key).sheet}`),
});

export class Sheet {
  readonly #inputs = new Map<string, string>();
  readonly #compiled = new Map<string, Compiled>();
  readonly #values = new Map<string, Value>();
  readonly #reads = new Map<string, Set<string>>();     // a cell -> the cells it reads
  readonly #readBy = new Map<string, Set<string>>();    // a cell -> the cells that read it
  readonly #volatile = new Set<string>();               // the cells whose formula is volatile
  readonly #listener: SheetListener | undefined;
  #language: Language;
  readonly #context: Context;
  readonly #names = new Map<string, [Addr, Addr]>();
  // Named formulas (a named LAMBDA), by name: what follows the '='.
  readonly #formulas = new Map<string, string>();
  readonly #tables = new Map<string, TableDef>();
  // The cells on a cycle through other sheets (the workbook's #crossCycles),
  // each with what it says: not worked out while they are on it.
  #onCycle = new Map<string, string>();
  readonly #areas = new Map<string, string[]>();          // a formula -> the cells its array covers (or would)
  readonly #spills = new Map<string, ArrayValue>();       // a formula -> the array it spills
  readonly #spilledBy = new Map<string, string>();        // a cell -> the formula whose array it shows
  #spillChanged = new Set<string>();
  // Manual calculation: the cells a change has made out of date, until the
  // sheet is calculated.
  readonly #stale = new Set<string>();

  constructor(listener?: SheetListener, language: Language = EXCEL, context: Context = alone()) {
    this.#listener = listener;
    this.#language = language;
    this.#context = context;
  }

  get language(): Language { return this.#language; }

  // Working a sheet out now (see Turns).
  get busy(): boolean { return this.#context.turns.busy; }

  input(key: string): string { return this.#inputs.get(key) ?? ""; }

  // A cell's value; a key on another sheet is the context's to answer.
  value(key: string): Value {
    const local = this.#local(key);
    return local === null ? this.#context.value(key) : this.#values.get(local) ?? null;
  }

  // A key as this sheet's own ("SHEET1!B2" on Sheet1 is B2), or null for a
  // cell on another sheet.
  #local(key: string): string | null {
    const { sheet, cell } = splitKey(key);
    return sheet === null ? key : sheet === this.#context.name ? cell : null;
  }

  // The cells holding something, in sheet order.
  keys(): string[] { return [...this.#inputs.keys()].sort(byOrder); }

  // What a cell reads, and what reads it (a key on another sheet qualified,
  // after this sheet's own, which are in sheet order).
  reads(key: string): string[] { return [...(this.#reads.get(key) ?? [])].sort(byKey); }
  readBy(key: string): string[] { return [...(this.#readBy.get(key) ?? [])].sort(byKey); }

  // Every cell whose value is an error, in sheet order.
  problems(): Problem[] {
    return [...this.#values].filter(([, v]) => v instanceof CellError).map(([k]) => k).sort(byOrder)
      .map((key) => {
        const e = this.#values.get(key) as CellError;
        return { key, message: `${e.code} ${e.why}` };
      });
  }

  set(key: string, input: string): Promise<string[]> {
    return this.setMany([[key, input]]);
  }

  // Set several cells, then recompute once. Gives the keys whose shown value
  // changed, in sheet order (the listener heard each of them). A key off the
  // sheet is refused before anything is set.
  setMany(entries: Iterable<readonly [string, string]>): Promise<string[]> {
    const list = this.#checked(entries);
    return this.#inTurn(() => this.#apply(list));
  }

  // Entries with their keys normalised; a key off the sheet is refused.
  #checked(entries: Iterable<readonly [string, string]>): [string, string][] {
    return [...entries].map(([k, input]): [string, string] => {
      const a = addrOf(k);
      if (!a) throw new RangeError(`${k} is not a cell on the sheet`);
      return [keyOf(a), input];
    });
  }

  // What the sheet holds, to be put back (the workbook's undo): its
  // language, inputs, names and tables. Values are worked out again.
  saved(): SavedSheet {
    return { language: this.#language, inputs: new Map(this.#inputs), names: new Map(this.#names), formulas: new Map(this.#formulas), tables: new Map(this.#tables) };
  }

  // The sheet put back as it was saved: the cells whose input differs set
  // back, and everything compiled again if its language, names or tables
  // differ. Gives the keys whose shown value changed, in sheet order.
  restore(saved: SavedSheet): Promise<string[]> {
    return this.#inTurn(async () => {
      const keys = new Set([...this.#inputs.keys(), ...saved.inputs.keys()]);
      const list = [...keys].filter((k) => this.input(k) !== (saved.inputs.get(k) ?? "")).map((k): [string, string] => [k, saved.inputs.get(k) ?? ""]);
      const same = <V>(a: Map<string, V>, b: ReadonlyMap<string, V>) => a.size === b.size && [...b].every(([k, v]) => a.get(k) === v);
      const relink = this.#language !== saved.language || !same(this.#names, saved.names) || !same(this.#formulas, saved.formulas) || !same(this.#tables, saved.tables);
      this.#language = saved.language;
      this.#names.clear();
      for (const [n, r] of saved.names) this.#names.set(n, r);
      this.#formulas.clear();
      for (const [n, f] of saved.formulas) this.#formulas.set(n, f);
      this.#tables.clear();
      for (const [n, t] of saved.tables) this.#tables.set(n, t);
      const changed = await this.#apply(list);
      return relink ? [...new Set([...changed, ...(await this.#recomputeAll())])] : changed;
    });
  }

  #apply(list: readonly (readonly [string, string])[]): Promise<string[]> {
    for (const [key, input] of list) {
      if (input === "") this.#inputs.delete(key); else this.#inputs.set(key, input);
      this.#compile(key);
    }
    return this.#recompute(list.map(([key]) => key));
  }

  // Another language (the major setting): every formula compiled again in
  // it, and the whole sheet worked out again.
  setLanguage(language: Language): Promise<string[]> {
    return this.#inTurn(() => {
      this.#language = language;
      return this.#recomputeAll();
    });
  }

  // An expression in the sheet's language, worked out against the sheet
  // (a conditional style): its value, or the error it came to.
  evaluate(formula: string): Promise<Value> {
    return this.#inTurn(async () => {
      const f = this.#language.compile(formula, this.#named, (n) => this.#tables.get(n.toUpperCase()) ?? null, this.#formula);
      const r = f.kind === "broken" ? f.error : (await f.run(this.#env())) ?? null;
      return isArray(r) ? r.values[0] ?? null : r;
    });
  }

  // The names, by name: a range's as [NAME, "B2:B4"], a formula's as
  // [NAME, "=LAMBDA(x, x*2)"].
  names(): [string, string][] {
    const ranges = [...this.#names].map(([n, [a, b]]): [string, string] => [n, keyOf(a) === keyOf(b) ? keyOf(a) : `${keyOf(a)}:${keyOf(b)}`]);
    const formulas = [...this.#formulas].map(([n, f]): [string, string] => [n, `=${f}`]);
    return [...ranges, ...formulas].sort(([x], [y]) => (x < y ? -1 : 1));
  }

  name(n: string): [Addr, Addr] | null {
    return this.#names.get(n.toUpperCase()) ?? null;
  }

  // A named formula's text (what follows the '='), or null.
  // A name as a formula reads it: the sheet's own, else the workbook's.
  readonly #named = (n: string): Named | null => {
    // Another sheet's own name (COSTS!QTY) is the workbook's to answer.
    if (n.includes("!")) return this.#context.named?.(n.toUpperCase()) ?? null;
    const r = this.#names.get(n.toUpperCase());
    return r ? { range: r, sheet: null } : this.#formulas.has(n.toUpperCase()) ? null : this.#context.named?.(n.toUpperCase()) ?? null;
  };
  readonly #formula = (n: string): string | null =>
    this.#formulas.get(n.toUpperCase()) ?? (this.#names.has(n.toUpperCase()) ? null : this.#context.formula?.(n.toUpperCase()) ?? null);
  named(n: string): Named | null { return this.#named(n); }

  // The sheet's own named formula's text (what follows the '='), or null.
  formula(n: string): string | null { return this.#formulas.get(n.toUpperCase()) ?? null; }
  // Its own or the workbook's.
  formulaOf(n: string): string | null { return this.#formula(n); }

  // A name for a range defined (or redefined), or deleted (null), or (a
  // string: what follows the '=') for a formula; a name is one or the
  // other. Every formula compiled again, as any may read it.
  define(n: string, r: [Addr, Addr] | string | null): Promise<string[]> {
    return this.#inTurn(() => {
      const key = n.toUpperCase();
      this.#names.delete(key);
      this.#formulas.delete(key);
      if (typeof r === "string") this.#formulas.set(key, r); else if (r) this.#names.set(key, r);
      return this.#recomputeAll();
    });
  }

  // The tables, by name; one by its name (any case).
  tables(): TableDef[] {
    return [...this.#tables.values()].sort((x, y) => (x.name.toUpperCase() < y.name.toUpperCase() ? -1 : 1));
  }

  table(n: string): TableDef | null {
    return this.#tables.get(n.toUpperCase()) ?? null;
  }

  // A table defined (or redefined), or deleted (null): every formula
  // compiled again, as any may read it.
  defineTable(n: string, t: TableDef | null): Promise<string[]> {
    return this.#inTurn(() => {
      if (t) this.#tables.set(n.toUpperCase(), t); else this.#tables.delete(n.toUpperCase());
      return this.#recomputeAll();
    });
  }

  // Everything compiled and worked out again, as after a language's setup
  // script ran, or a sheet was added to the workbook or deleted from it.
  recalculate(): Promise<string[]> {
    return this.#inTurn(() => this.#recomputeAll());
  }

  // Cells on other sheets changed (qualified keys): what reads them worked
  // out again. Not a change to undo.
  touch(keys: readonly string[]): Promise<string[]> {
    return this.#inTurn(() => this.#recompute(keys));
  }

  // The cells on a cycle through other sheets, now (the workbook finds
  // those; a sheet finds its own): each is #CYCLE!, saying the path, and is
  // not worked out while it is on one. One no longer on a cycle is worked
  // out again, and so is what reads any of them. The cells that changed.
  cycle(marks: ReadonlyMap<string, string>): Promise<string[]> {
    return this.#inTurn(async () => {
      const gone = [...this.#onCycle.keys()].filter((k) => !marks.has(k));
      this.#onCycle = new Map(marks);
      for (const [k, why] of marks) {
        this.#values.set(k, new CellError("#CYCLE!", why));
        this.#listener?.changed(k, this.value(k));
      }
      const readers = [...marks.keys()].flatMap((k) => [...(this.#readBy.get(k) ?? [])]).filter((r) => !marks.has(r));
      const again = [...new Set([...gone, ...readers])];
      const changed = again.length ? await this.#recompute(again) : [];
      return [...new Set([...marks.keys(), ...changed])];
    });
  }

  // Whether any of its cells is on a cycle through other sheets.
  get cycling(): boolean { return this.#onCycle.size > 0; }

  // The cells out of date under manual calculation, in sheet order.
  stale(): string[] { return [...this.#stale].sort(byOrder); }

  #recomputeAll(): Promise<string[]> {
    const keys = this.keys();
    keys.forEach((k) => this.#compile(k));
    this.#stale.clear();
    return this.#recompute(keys, 0, true);
  }

  #compile(key: string): void {
    const c = compile(this.input(key), this.#language, this.#named, (n) => this.#tables.get(n.toUpperCase()) ?? null, this.#formula, key);
    this.#compiled.set(key, c);
    if (c.kind === "formula" && c.volatile) this.#volatile.add(key); else this.#volatile.delete(key);
    this.#link(key, new Set(c.kind === "formula" ? c.reads.map((k) => this.#local(k) ?? k) : []));
  }

  #inTurn<T>(work: () => Promise<T>): Promise<T> {
    return this.#context.turns.take(work);
  }

  #link(key: string, reads: Set<string>): void {
    for (const old of this.#reads.get(key) ?? []) this.#readBy.get(old)?.delete(key);
    this.#reads.set(key, reads);
    for (const r of reads) getOrAdd(this.#readBy, r, () => new Set()).add(key);
  }

  // What a cell reads, and what reads it, a spilled cell reading the
  // formula it shows.
  #readsOf(k: string): string[] {
    const by = this.#spilledBy.get(k);
    return [...(this.#reads.get(k) ?? []), ...(by ? [by] : [])];
  }

  #readersOf(k: string): string[] {
    return [...(this.#readBy.get(k) ?? []), ...(this.#spills.has(k) ? this.#areas.get(k) ?? [] : [])];
  }

  // What formulas see: the cells, and what a cell spills (A1#); `read`
  // hears the cells a run-time reference came to; a cell's input and the
  // workbook's sheets (FORMULATEXT, SHEET).
  #env(read?: (k: string) => void) {
    return {
      get: (k: string) => this.value(k),
      spill: (k: string) => this.#spills.get(this.#local(k) ?? "") ?? null,
      read,
      input: (k: string) => { const local = this.#local(k); return local === null ? this.#context.input?.(k) ?? "" : this.input(local); },
      book: () => {
        const sheets = this.#context.sheets?.() ?? [this.#context.name];
        return { sheets, here: Math.max(0, sheets.findIndex((s) => s.toUpperCase() === this.#context.name)) };
      },
      calc: () => this.#context.calc?.() ?? AUTOMATIC,
      hidden: (k: string) => { const local = this.#local(k); return this.#context.hidden?.(local === null ? k : `${this.#context.name}!${local}`) ?? null; },
    };
  }

  // The set cells and everything downstream of them.
  #downstream(from: readonly string[]): Set<string> {
    const seen = new Set<string>();
    const visit = (k: string): void => {
      if (seen.has(k)) return;
      seen.add(k);
      this.#readersOf(k).forEach(visit);
    };
    from.forEach(visit);
    return seen;
  }

  // Kahn's algorithm over the cells in `cells`, by the reads among them:
  // the ones it can order, and the ones left (on a cycle, or downstream of
  // one).
  #ordered(cells: Set<string>): { ordered: string[]; left: Set<string> } {
    // The edges as they are now: working a formula out may move a spill.
    const readers = new Map([...cells].map((k) => [k, this.#readersOf(k)]));
    const waiting = new Map<string, number>();
    for (const k of cells) waiting.set(k, new Set(this.#readsOf(k).filter((r) => cells.has(r))).size);
    const ready = [...cells].filter((k) => waiting.get(k) === 0).sort(byOrder);
    const ordered: string[] = [];
    for (let k = ready.shift(); k !== undefined; k = ready.shift()) {
      ordered.push(k);
      for (const d of new Set(readers.get(k))) {
        if (!cells.has(d)) continue;
        const n = waiting.get(d)! - 1;
        waiting.set(d, n);
        if (n === 0) ready.push(d);
      }
    }
    const done = new Set(ordered);
    return { ordered, left: new Set([...cells].filter((k) => !done.has(k))) };
  }

  // A cell in `left` is on a cycle when it reads its way back to itself.
  #cycleThrough(start: string, left: Set<string>): string[] | null {
    const path: string[] = [];
    const seen = new Set<string>();
    const walk = (k: string): boolean => {
      path.push(k);
      for (const r of this.#readsOf(k)) {
        if (r === start) return true;
        if (left.has(r) && !seen.has(r)) {
          seen.add(r);
          if (walk(r)) return true;
        }
      }
      path.pop();
      return false;
    };
    return walk(start) ? path : null;
  }

  // The set cells (and cells on other sheets, which are not worked out
  // here) and what reads them, worked out in order; a formula whose area a
  // set cell is in, and every volatile formula, too. When spills moved,
  // what reads the cells they left or took is worked out again.
  //
  // Under manual calculation (and not `all`) only the set cells, and a
  // formula whose area one is in, are worked out; what is downstream of
  // them is stale until the sheet is calculated. Under iteration a cycle's
  // cells are worked round (#iterate) instead of being #CYCLE!.
  async #recompute(set: readonly string[], depth = 0, all = false): Promise<string[]> {
    const calc = this.#context.calc?.() ?? AUTOMATIC;
    const before = new Map<string, string>();
    const anchors = [...this.#areas].filter(([, area]) => area.some((k) => set.includes(k))).map(([a]) => a);
    let affected = this.#downstream([...set, ...anchors, ...(calc.manual && !all ? [] : this.#volatile)]);
    if (calc.manual && !all) {
      const now = new Set([...set, ...anchors].filter((k) => !k.includes("!")));
      for (const k of affected) if (!now.has(k) && !k.includes("!")) this.#stale.add(k);
      for (const k of now) this.#stale.delete(k);
      affected = now;
    }
    this.#spillChanged = new Set();
    const mine = [...affected].filter((k) => !k.includes("!"));
    for (const k of mine) before.set(k, show(this.value(k)));

    const { ordered, left } = this.#ordered(affected);
    for (const k of ordered) if (!k.includes("!")) await this.#evaluate(k);
    if (left.size > 0) {
      const onCycle = new Set<string>();
      for (const k of left) {
        const path = this.#cycleThrough(k, left);
        if (!path) continue;
        onCycle.add(k);
        if (!calc.iterate) this.#values.set(k, new CellError("#CYCLE!", `a circular reference: ${[...path, k].join(" → ")}`));
      }
      if (calc.iterate) await this.#iterate([...onCycle].sort(byOrder), calc.iterate);
      // What is downstream of a cycle reads its error, in order.
      const rest = new Set([...left].filter((k) => !onCycle.has(k)));
      for (const k of this.#ordered(rest).ordered) await this.#evaluate(k);
    }

    let changed = mine.filter((k) => set.includes(k) || before.get(k) !== show(this.value(k))).sort(byOrder);
    for (const k of changed) this.#listener?.changed(k, this.value(k));
    const moved = [...this.#spillChanged];
    if (moved.length && depth < 4) changed = [...new Set([...changed, ...(await this.#recompute(moved, depth + 1, all))])].sort(byOrder);
    return changed;
  }

  // A cycle's cells worked round, in sheet order, each from the values the
  // others have now (Excel's iteration): from where they were (0 for an
  // empty one, or one that was #CYCLE!), up to `max` times, until no
  // number moved by more than `change` and nothing else changed.
  async #iterate(cells: readonly string[], { max, change }: { max: number; change: number }): Promise<number> {
    for (const k of cells) {
      const v = this.#values.get(k);
      if (v instanceof CellError && v.code === "#CYCLE!") this.#values.delete(k);
    }
    for (let i = 1; i <= max; i++) {
      let moved = 0;
      for (const k of cells) {
        const was = this.value(k);
        await this.#evaluate(k);
        const now = this.value(k);
        const d = typeof was === "number" && typeof now === "number" ? Math.abs(now - was)
          : (was ?? 0) === (now ?? 0) || show(was) === show(now) ? 0 : Infinity;
        moved = Math.max(moved, d);
      }
      if (moved <= change) return i;
    }
    return max;
  }

  // A formula's spill taken back: its cells empty again (but for any that
  // have an input of their own).
  #unspill(anchor: string): void {
    for (const k of this.#areas.get(anchor) ?? []) {
      if (this.#spilledBy.get(k) !== anchor) continue;
      this.#spilledBy.delete(k);
      if (!this.#inputs.has(k)) this.#values.delete(k);
    }
    this.#spills.delete(anchor);
  }

  // An array result spilled from `key`, or why not: the formula's own value.
  #spill(key: string, v: ArrayValue): Value {
    const a = addrOf(key)!, area: string[] = [];
    let offSheet = false;
    for (let i = 0; i < v.rows; i++) {
      for (let j = 0; j < v.cols; j++) {
        if (i === 0 && j === 0) continue;
        if (a.col + j >= COLS || a.row + i >= ROWS) offSheet = true; else area.push(keyOf({ col: a.col + j, row: a.row + i }));
      }
    }
    this.#areas.set(key, area);
    if (offSheet) return new CellError("#SPILL!", "the result does not fit on the sheet");
    const blocker = area.find((k) => this.#inputs.has(k) || (this.#spilledBy.has(k) && this.#spilledBy.get(k) !== key));
    if (blocker) return new CellError("#SPILL!", `${blocker} is in the way`);
    this.#spills.set(key, v);
    area.forEach((k, i) => {
      this.#spilledBy.set(k, key);
      this.#values.set(k, v.values[i + 1]);
    });
    return v.values[0];
  }

  async #evaluate(key: string): Promise<void> {
    // A cell another formula spills into shows that formula's value.
    if (!this.#inputs.has(key) && this.#spilledBy.has(key)) return;
    const cycle = this.#onCycle.get(key);
    if (cycle !== undefined) {
      this.#values.set(key, new CellError("#CYCLE!", cycle));
      return;
    }
    const c = this.#compiled.get(key);
    let r: Result = null;
    if (c?.kind === "formula") {
      const dynamic = new Set<string>();
      r = (await c.run(this.#env((k) => dynamic.add(this.#local(k) ?? k)))) ?? 0;
      if (c.volatile || dynamic.size) this.#link(key, new Set([...c.reads.map((k) => this.#local(k) ?? k), ...dynamic]));
    } else if (c?.kind === "broken") r = c.error;
    else if (c?.kind === "constant") r = c.value;
    const area = this.#areas.get(key) ?? [];
    const was = area.map((k) => show(this.value(k)));
    this.#unspill(key);
    this.#areas.delete(key);
    const v = isArray(r) ? this.#spill(key, r) : r;
    const now = this.#areas.get(key) ?? [];
    for (const [i, k] of area.entries()) if (show(this.value(k)) !== was[i]) this.#spillChanged.add(k);
    for (const k of now) if (!area.includes(k) && this.#spilledBy.get(k) === key) this.#spillChanged.add(k);
    if (v === null && !this.#inputs.has(key)) this.#values.delete(key); else this.#values.set(key, v);
  }
}

