// What the spreadsheet's panels share: the workbook and its commands (the
// seam), the view's own state (the selected cell's address, the formula
// bar, the status line, the problems, the sheets' names, the watched cells, the charts), the grid's handle and selection, and
// run(), the one way a panel changes anything: it runs a command, prints
// its outcome (what specs read), and shows the workbook again, in one batch.
import { COLS, addrOf, keyOf, parseRange } from "../address.ts";
import type { Addr } from "../address.ts";
import { Commands } from "../commands.ts";
import type { Outcome, Params } from "../commands.ts";
import type { Problem } from "../sheet.ts";
import { Workbook } from "../workbook.ts";
import { COLUMN_WIDTH, ROWS_WHEN_SPLIT, ROWS_WITH_PANEL, ROW_PX, buildPane } from "./grid.ts";

// aether-ui's datagrid's own row height, which a zoom scales.
const ROW_HEIGHT = 24;
import type { ChartView, DrawingView, SlicerView, Watch } from "../workbook.ts";
import { statsLine } from "../workbook.ts";
import type { CellEditor } from "../editors.ts";

export class SheetView {
  readonly wb = new Workbook();
  // Files are saved in the page's storage.
  readonly cmds = new Commands(this.wb, storage);
  // A link's URL opened by the system: the installed app's shell.open
  // (app.json grants https:, http: and mailto:); a web page has none.
  readonly #opener = this.wb.opener = (url: string): string => {
    if (typeof shell === "undefined") return "opening a URL needs the installed app (open-urls)";
    try { shell.open(url); return ""; } catch (e) { return String((e as Error)?.message ?? e); }
  };
  readonly address = state("A1");
  readonly input = state("");
  readonly status = state("");
  // The status bar: the selection's figures, the sheet's cells, the bus.
  readonly figures = state("");
  bus = false;
  // What Undo and Redo would do: "Undo Merge Cells · Redo Paste".
  readonly history = state("");
  readonly comment = state("");
  readonly format = state("");
  // The selected cell's editor (editors.ts), or null.
  readonly editor = state<CellEditor | null>(null);
  readonly formulas = state<string>("Excel");
  readonly setup = state("");
  readonly problems = state<Problem[]>([]);
  readonly sheets = state<string[]>(["Sheet1"]);
  // The tabs shown: the sheets not hidden, each with its colour.
  readonly tabs = state<{ name: string; color: string | null }[]>([{ name: "Sheet1", color: null }]);
  readonly active = state("Sheet1");
  readonly watches = state<Watch[]>([]);
  readonly charts = state<(ChartView & { n: number })[]>([]);
  readonly slicers = state<SlicerView[]>([]);
  readonly drawings = state<DrawingView[]>([]);
  grid = 0;
  // The split pane (grid.ts's splitPane), its box, and whether it shows.
  pane = 0;
  paneBox = 0;
  #shownSplit = false;
  // What each panel button (and a few keys) does, by its label: the menus
  // run the same (menubar.ts).
  readonly actions = new Map<string, () => unknown>();
  // The panels under the grid by name, the panel each action is on, the
  // one being built, and the one shown.
  readonly #panels = new Map<string, number>();
  readonly #panelOf = new Map<string, string>();
  #building: string | null = null;
  #shownPanel: string | null = null;
  // The rows made as tall as their wrapped text; the rows, columns (model)
  // and frozen rows the grid was last given.
  #fitted = new Set<number>();
  #shownHidden = { rows: new Set<number>(), cols: new Set<number>() };
  #frozen = 0;
  #frozenCols = 0;
  #merged: string[] = [];
  #sized = new Map<number, number>();

  // A panel's button, its action kept by its label for the menus.
  btn(label: string, fn: () => unknown): number {
    this.action(label, fn);
    return ui.btn(label, fn);
  }

  // An action the menus may run, by its label (the first one kept), and
  // the panel it is on.
  action(label: string, fn: () => unknown): void {
    if (this.actions.has(label)) return;
    this.actions.set(label, fn);
    if (this.#building) this.#panelOf.set(label, this.#building);
  }

  // A panel under the grid: its bars, hidden until it is opened.
  panel(name: string, build: () => void): void {
    this.#building = name;
    const h = ui.vstack(build);
    this.#building = null;
    ui.set_visible(h, false);
    this.#panels.set(name, h);
  }

  // An action run from a menu or the toolbar: its panel opened first, so
  // its fields are there to read and fill.
  async perform(label: string): Promise<unknown> {
    const panel = this.#panelOf.get(label);
    if (panel && this.wb.panel !== panel) await this.run("view.panel", { name: panel });
    return this.actions.get(label)?.();
  }

  // The selected cell's key (A1 when the address field names none).
  here(): string {
    return keyOf(addrOf(this.address.value) ?? { col: 0, row: 0 });
  }

  // The grid's selection as two sheet corners (the row numbers are not on
  // the sheet: a selection reaching them starts at column A), or null.
  selection(): [Addr, Addr] | null {
    const [ar, ac, r, c] = ui.datagrid_selection(this.grid);
    const corner = (row: number, col: number): Addr | null => {
      const m = Math.max(1, ui.datagrid_col_model(this.grid, col));
      return row >= 0 && m <= COLS ? { col: m - 1, row } : null;
    };
    const a = corner(ar, ac), b = corner(r, c);
    // Top-left first, however it was dragged.
    return a && b ? [{ col: Math.min(a.col, b.col), row: Math.min(a.row, b.row) }, { col: Math.max(a.col, b.col), row: Math.max(a.row, b.row) }] : null;
  }

  // Where a sheet column is on screen now (columns can be dragged).
  viewCol(modelCol: number): number {
    for (let c = 0; c <= COLS; c++) if (ui.datagrid_col_model(this.grid, c) === modelCol) return c;
    return modelCol;
  }

  // A range selected, its first cell the address.
  // The address first: following it selects its one cell, which the range
  // then replaces, anchored at its far corner so the cursor (the address,
  // what the grid scrolls to) is its first cell, as Excel's active cell.
  select([a, b]: [Addr, Addr]): void {
    this.address.set(keyOf(a));
    ui.datagrid_select_range(this.grid, b.row, this.viewCol(b.col + 1), a.row, this.viewCol(a.col + 1));
  }

  // The selection as a command's range parameter: "F4:G7", or the cell.
  // With several areas (a Ctrl/Cmd click added one), each, the current
  // one last: "A1:B2,D4:D9".
  range(): string {
    const s = this.selection();
    const one = s ? `${keyOf(s[0])}:${keyOf(s[1])}` : this.here();
    return [...this.areas().map(([a, b]) => `${keyOf(a)}:${keyOf(b)}`), one].join(",");
  }

  // The selection's areas before the current one, as sheet corners.
  areas(): [Addr, Addr][] {
    if (!this.grid) return [];
    const corner = (row: number, col: number): Addr | null => {
      const m = Math.max(1, ui.datagrid_col_model(this.grid, col));
      return row >= 0 && m <= COLS ? { col: m - 1, row } : null;
    };
    return String(ui.datagrid_areas(this.grid)).split(";").filter((x) => x).flatMap((x): [Addr, Addr][] => {
      const [ar, ac, r, c] = x.split(",").map(Number);
      const a = corner(ar, ac), b = corner(r, c);
      return a && b ? [[{ col: Math.min(a.col, b.col), row: Math.min(a.row, b.row) }, { col: Math.max(a.col, b.col), row: Math.max(a.row, b.row) }]] : [];
    });
  }

  // The selected cell's input, status line and notes in their fields.
  follow(key = this.here()): void {
    const note = this.wb.notes.get(key);
    if (this.grid && !this.wb.selecting) this.wb.selection = this.range();
    const sel = this.grid ? this.selection() : null;
    const others = this.areas();
    const many = sel && (others.length > 0 || sel[0].col !== sel[1].col || sel[0].row !== sel[1].row);
    const cells = this.wb.sheet.keys().length;
    const bar = [many ? statsLine(this.wb.stats([...others, sel!])) : "Ready", `${cells} cell${cells === 1 ? "" : "s"}`, `Zoom ${this.wb.zoom}%`, this.wb.calculation.manual && this.wb.stale().length ? "Calculate" : "", this.bus ? "Control bus" : ""];
    batch(() => {
      this.figures.set(bar.filter((x) => x).join("  ·  "));
      this.input.set(this.wb.shownInput(key));
      this.status.set(this.wb.describe(key));
      this.comment.set(note.comment ?? "");
      this.format.set(note.format ?? "");
      this.editor.set(note.editor ?? null);
    });
  }

  // A command run, its outcome printed, the workbook shown again.
  run(id: string, params: Params = {}): Promise<Outcome> {
    return this.#outcome(() => this.cmds.execute(id, params));
  }

  // The same, from a command line (the command bar, and specs through it).
  runLine(line: string): Promise<Outcome> {
    return this.#outcome(() => this.cmds.run(line));
  }

  // A command that throws (a bug, as against a refusal) is said too: a
  // button's handler that returns a promise has no one to tell.
  async #outcome(command: () => Promise<Outcome>): Promise<Outcome> {
    try {
      return this.#shown(await command());
    } catch (e) {
      return this.#shown({ ok: false, message: `failed: ${String((e as Error)?.message ?? e)}`, changed: [] });
    }
  }

  #shown(o: Outcome): Outcome {
    print(`sheet: ${o.ok ? "" : "refused: "}${o.message}`);
    this.refresh();
    if (!o.ok) this.status.set(o.message);
    return o;
  }

  // The workbook shown again, in one batch (after a command, or a script).
  refresh(): void {
    batch(() => {
      this.problems.set(this.wb.sheet.problems());
      this.sheets.set(this.wb.sheetNames());
      this.tabs.set(this.wb.pages().filter((p) => !p.hiddenSheet).map((p) => ({ name: p.name, color: p.tabColor })));
      this.active.set(this.wb.active);
      this.watches.set(this.wb.watches());
      this.charts.set(this.wb.charts().map((c, i) => ({ ...c, n: i + 1 })));
      this.slicers.set(this.wb.slicerViews());
      this.drawings.set(this.wb.drawingViews());
      this.formulas.set(this.wb.sheet.language.name);
      this.setup.set(this.wb.setupScript ? `Setup script: ${this.wb.setupScript}` : "");
      const u = this.wb.undoLabel, r = this.wb.redoLabel;
      this.history.set([u ? `Undo ${u}` : "", r ? `Redo ${r}` : ""].filter((x) => x).join(" · "));
      this.follow();
    });
    if (this.wb.panel !== this.#shownPanel || this.wb.split !== this.#shownSplit) {
      for (const [n, h] of this.#panels) ui.set_visible(h, n === this.wb.panel);
      if (this.paneBox && this.wb.split !== this.#shownSplit) {
        if (this.wb.split) buildPane(this);
        else {
          ui.clear(this.paneBox);
          this.pane = 0;
        }
      }
      // The grid takes the height back when no panel is open (and gives
      // the split pane room when split).
      const rows = this.wb.split ? ROWS_WHEN_SPLIT : this.wb.panel ? ROWS_WITH_PANEL : 0;
      if (this.grid) {
        const was = ui.datagrid_height(this.grid);
        ui.datagrid_set_height(this.grid, rows * ROW_PX);
        if (ui.datagrid_height(this.grid) !== was) print(`sheet: the grid ${Math.round(ui.datagrid_height(this.grid) / ROW_PX)} rows high`);
      }
      this.#shownPanel = this.wb.panel;
      this.#shownSplit = this.wb.split;
    }
    if (this.grid && this.wb.selecting) {
      this.wb.selecting = false;
      // Several areas: the first selected, each other added after it.
      const rs = this.wb.selection.split(",").map((x) => parseRange(x)).filter((x): x is [Addr, Addr] => !!x);
      const r = rs[0];
      // Selected, the grid has the keys (Delete clears it), as after the
      // Name Box.
      if (r) {
        this.select(r);
        for (const [a, b] of rs.slice(1)) ui.datagrid_add_area(this.grid, a.row, this.viewCol(a.col + 1), b.row, this.viewCol(b.col + 1));
        try { ui.focus(this.grid); } catch { /* a grid sae cannot focus: the keys stay where they were */ }
        this.follow();
      }
    }
    if (this.grid) {
      this.#showView();
      ui.datagrid_refresh(this.grid);
      if (this.pane) ui.datagrid_refresh(this.pane);
      this.#fitRows();
      this.#sizeCols();
    }
  }

  // Columns set to a width that width; columns no longer set the grid's.
  #sizeCols(): void {
    const set = new Map(this.wb.shownColWidths(COLUMN_WIDTH));
    for (const [c] of this.#sized) if (!set.has(c)) ui.datagrid_set_col_width(this.grid, this.viewCol(c + 1), COLUMN_WIDTH);
    for (const [c, w] of set) if (this.#sized.get(c) !== w) ui.datagrid_set_col_width(this.grid, this.viewCol(c + 1), w);
    this.#sized = set;
  }

  // The active sheet's hidden rows and columns and frozen rows, given to the
  // grid where they changed.
  #showView(): void {
    const { rows, cols } = this.wb.hidden(), was = this.#shownHidden;
    const want = { rows: new Set(rows), cols: new Set(cols) };
    for (const r of new Set([...was.rows, ...want.rows])) if (was.rows.has(r) !== want.rows.has(r)) ui.datagrid_hide_row(this.grid, r, want.rows.has(r) ? 1 : 0);
    for (const c of new Set([...was.cols, ...want.cols])) if (was.cols.has(c) !== want.cols.has(c)) ui.datagrid_hide_col(this.grid, this.viewCol(c + 1), want.cols.has(c) ? 1 : 0);
    this.#shownHidden = want;
    if (this.wb.frozenRows() !== this.#frozen) ui.datagrid_freeze_rows(this.grid, (this.#frozen = this.wb.frozenRows()));
    if (this.wb.frozenCols() !== this.#frozenCols) ui.datagrid_freeze_cols(this.grid, (this.#frozenCols = this.wb.frozenCols()));
    // Merges: those gone undone, those new made (at the columns on screen).
    const merges = this.wb.merges();
    for (const m of this.#merged) if (!merges.includes(m)) {
      const [a] = parseRange(m)!;
      ui.datagrid_unmerge(this.grid, a.row, this.viewCol(a.col + 1));
    }
    for (const m of merges) if (!this.#merged.includes(m)) {
      const [a, b] = parseRange(m)!;
      ui.datagrid_merge(this.grid, a.row, this.viewCol(a.col + 1), b.row - a.row + 1, b.col - a.col + 1);
    }
    this.#merged = merges;
  }

  // Rows set to a height that height; other rows with wrapped text as tall
  // as it; rows no longer either as tall as the grid's.
  #fitRows(): void {
    const set = new Map(this.wb.shownRowHeights(ROW_HEIGHT)), wrapped = new Set(this.wb.wrapped().filter((r) => !set.has(r)));
    for (const r of this.#fitted) if (!set.has(r) && !wrapped.has(r)) ui.datagrid_set_row_height(this.grid, r, 0);
    for (const [r, h] of set) ui.datagrid_set_row_height(this.grid, r, h);
    for (const r of wrapped) ui.datagrid_fit_row(this.grid, r);
    this.#fitted = new Set([...set.keys(), ...wrapped]);
  }
}
