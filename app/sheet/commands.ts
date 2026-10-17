// The seam under the page: every action on a workbook is a command, as in
// GridCraft (its engine.execute) and the Tauri spreadsheet's control bus. A
// command has an id ("data.sortDescending"), a label and string parameters
// ("range=F4:G7 by=G"), and gives an Outcome: ok or refused, a message, the
// cells whose shown value changed, and sometimes an output (what cell.get
// read, what edit.find found). The page's panels and keys run commands and
// nothing else; so do RexxJS scripts (ADDRESS SHEET, SETCELL, GETCELL:
// script.ts), and the specs, headless and in the browser (the page's
// command bar). The commands that change the workbook are journaled as
// command lines, which replay.
import { COLS, ROWS, cellAddr, colName, colOf, keyOf, order, parseRange, range, sheetPrefix, splitKey } from "./address.ts";
import { evaluationSteps } from "./audit.ts";
import type { Addr } from "./address.ts";
import { copied, filled, pasted } from "./copy.ts";
import type { FillWay } from "./copy.ts";
import { fromHtml, fromJson, toHtml } from "./tabular.ts";
import { readXlsx } from "./xlsx.ts";
import { readOds } from "./ods.ts";
import { readXls } from "./xls.ts";
import { decryptPackage, isEncrypted } from "./encrypted.ts";
import { PAPER, writePdf } from "./pdf.ts";
import type { PdfSheet } from "./pdf.ts";
import { writeXlsx } from "./xlsxwrite.ts";
import type { XlsxRead } from "./xlsx.ts";
import { DELIMITERS as CSV_DELIMITERS, decode, detectDelimiter, fromBase64, fromCsv, toBase64, toCsv } from "./csv.ts";
import { readSheetFile, writeSheetFile } from "./file.ts";
import type { SheetFile } from "./file.ts";
import { show } from "./format.ts";
import { EXAMPLE, FORMATTED, SAMPLE_BUDGET, TEST_SHEET } from "./samples.ts";
import { runScript, runSpreadsheet } from "./script.ts";
import { lambdaOf } from "./rexx.ts";
import { continued } from "./series.ts";
import { check, describeRule, errorOf, ruleOf } from "./validation.ts";
import { asApp, editorOf, refusal, toggled } from "./editors.ts";
import { ICONS, PERIODS, THRESHOLDS, describeRule as describeCf, rangeName } from "./rules.ts";
import type { CfRule, Comparison, IconSet, Threshold } from "./rules.ts";
import { borderLine, colour, toStyle } from "./formats.ts";
import { CATEGORIES, numberFormat } from "./numfmt.ts";
import { ROLES, THEMES, isRole, resolved } from "./themes.ts";
import type { Role, Theme } from "./themes.ts";
import type { Look } from "./formats.ts";
import { CHART_KINDS, LEGENDS, SPARKLINE_TYPES, sparklineOptions } from "./charts.ts";
import type { ChartOptions, Shape, Trendline } from "./charts.ts";
import { TREND_TYPES } from "./trend.ts";
import type { ChartKind } from "./charts.ts";
import type { ChartSpec, Clip } from "./workbook.ts";
import type { TableDef } from "./language.ts";
import { blockMoved, deleted, inserted, movedBy, restructured, swapped } from "./shift.ts";
import { TOTALS, dataRows, totalsInput, unstructured } from "./structured.ts";
import { TABLE_STYLE_NAMES, styleNamed } from "./styles.ts";
import { BAND, bandedKeys, slicedRows, slicerValues } from "./workbook.ts";
import type { Edit, Scope } from "./shift.ts";
import { CONSOLIDATE_FUNCTIONS, consolidate } from "./consolidate.ts";
import { solveLp } from "./solver.ts";
import type { LpRow, Relation } from "./solver.ts";
import { MOST_LEVELS, SUBTOTAL_FUNCTIONS, depth, detailOf, hiddenAtLevel, regrouped, runs } from "./outline.ts";
import { found, rearranged, rearrangedMoves, replaced, sortedRows, sortedRowsBy } from "./sortfind.ts";
import type { SortKey } from "./sortfind.ts";
import { distinctRows, secant, splitText } from "./whatif.ts";
import { criterion, recordTest } from "./functions.ts";
import { errorCheck } from "./errorcheck.ts";
import { PANELS, menuLines } from "./menus.ts";
import { complete, cycleRef } from "./complete.ts";
import { KEYS } from "./keys.ts";
import { infer, run as runFill } from "./flashfill.ts";

const resolvedColour = (c: string | null): string | null => (c === null ? null : resolved(c));
const NUMERIC_TEXT = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)\s*$/;
import { CellError } from "./values.ts";
import type { Value } from "./values.ts";
import type { Direction, Filter, LanguageName, Page, Pivot, Scenario, Workbook } from "./workbook.ts";
import { statsLine } from "./workbook.ts";
import { AGGREGATES, LAYOUTS, SHOW_AS, pivoted } from "./pivot.ts";
import type { PivotOptions } from "./pivot.ts";
import type { Aggregate } from "./pivot.ts";
import { FormulaError } from "./lexer.ts";

export type Params = Readonly<Record<string, string>>;

// Where files are saved: the page's `storage` (get, set, remove; a key is
// letters, digits, '.', '_' or '-'), or one in memory.
export interface Store {
  get(key: string): string | null;
  set(key: string, value: string): void;
  remove(key: string): void;
}

export function memoryStore(): Store {
  const m = new Map<string, string>();
  return { get: (k) => m.get(k) ?? null, set: (k, v) => { m.set(k, v); }, remove: (k) => { m.delete(k); } };
}

export interface Outcome {
  readonly ok: boolean;
  readonly message: string;
  readonly changed: readonly string[];
  readonly output?: unknown;
}

interface Spec {
  readonly id: string;
  readonly label: string;
  readonly params: readonly string[];       // each required
  readonly optional?: readonly string[];
  readonly journaled: boolean;              // it changes the workbook (or the clipboard)
  // Not a step to undo, though journaled: it changes only what is looked
  // at or copied, is undo itself, or opens a file (where undo stops).
  readonly unstepped?: true;
  run(wb: Workbook, p: Params, store: Store): Promise<Outcome>;
}

// The commands that run over each area of a multiple selection
// ("A1:B2,D4:D9") as one step.
const AREAS = new Set(["range.clear", "range.clearFormats", "range.clearComments", "range.clearAll", "format.font", "format.align", "format.fill",
  "format.color", "format.number", "format.border", "format.wrap", "validation.set", "validation.clear", "rows.hide", "rows.unhide", "columns.hide", "columns.unhide"]);

// The commands that take several areas themselves: one answer over all.
const OWN_AREAS = new Set(["range.stats", "view.select"]);

// A command that cannot run as asked: its message is the Outcome's.
class Refused extends Error {}
const refuse = (message: string): never => { throw new Refused(message); };

// ---- The command line: an id, then name=value pairs; a value with spaces,
// quotes or '=' is quoted, "" being a quote inside.

export function parseLine(line: string): { id: string; params: Record<string, string> } {
  const m = /^\s*(\S+)/.exec(line);
  const params: Record<string, string> = {};
  let rest = line.slice(m ? m[0].length : 0);
  for (;;) {
    const p = /^\s*([A-Za-z_][\w]*)=("(?:[^"]|"")*"|\S*)/.exec(rest);
    if (!p) break;
    params[p[1]] = p[2].startsWith('"') ? p[2].slice(1, -1).replace(/""/g, '"') : p[2];
    rest = rest.slice(p[0].length);
  }
  return { id: m ? m[1] : "", params };
}

export function lineOf(id: string, params: Params): string {
  const value = (v: string) => (v === "" || /[\s"=]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
  return [id, ...Object.entries(params).map(([k, v]) => `${k}=${value(v)}`)].join(" ");
}

// ---- Reading parameters

const cellOf = (s: string): string => {
  const r = parseRange(s);
  return r && keyOf(r[0]) === keyOf(r[1]) ? keyOf(r[0]) : refuse(`${s} is not a cell on the sheet`);
};
const rangeOf = (s: string): [Addr, Addr] => parseRange(s) ?? refuse(`${s} is not a range on the sheet`);
const named = ([a, b]: [Addr, Addr]) => `${keyOf(a)}:${keyOf(b)}`;
const flag = (s: string | undefined) => /^(1|true|yes|on)$/i.test(s ?? "");
const topLeft = ([a, b]: [Addr, Addr]): Addr => ({ col: Math.min(a.col, b.col), row: Math.min(a.row, b.row) });
const columnOf = (s: string): number => colOf(String(s)) ?? refuse(`${s} is not a column on the sheet`);
const countOf = (s: string | undefined): number => {
  const n = Number(s ?? "1");
  return Number.isInteger(n) && n >= 1 ? n : refuse(`${s} is not a count`);
};

// ---- What changing commands share

const inputs = (wb: Workbook) => (key: string) => wb.sheet.input(key);

async function setting(wb: Workbook, entries: readonly (readonly [string, string])[], what: string): Promise<Outcome> {
  const changed = await wb.sheet.setMany(entries);
  await wb.settle();
  return { ok: true, message: `${what}: ${new Set(entries.map(([k]) => k)).size} cells set, ${changed.length} changed`, changed };
}

// Every sheet compiled and worked out again: the active one's cells whose
// shown value changed.
async function recalculated(wb: Workbook): Promise<string[]> {
  // Worked out as if automatic, so what one sheet reads of another is
  // brought up to date too; manual again after.
  const calc = wb.calculation;
  wb.calculation = { ...calc, manual: false };
  try {
    return await everySheet(wb);
  } finally {
    wb.calculation = calc;
  }
}

async function everySheet(wb: Workbook): Promise<string[]> {
  let changed: string[] = [];
  for (const page of wb.pages()) {
    const keys = await page.sheet.recalculate();
    if (page.name === wb.active) changed = keys;
  }
  await wb.settle();
  return changed;
}

async function changedBy(wb: Workbook, changed: readonly string[], message: string): Promise<Outcome> {
  await wb.settle();
  return { ok: true, message, changed };
}

const SAMPLES: Record<string, unknown> = { budget: SAMPLE_BUDGET, test: TEST_SHEET, formatted: FORMATTED };

// A sheet file: the workbook emptied, its language, its setup script run,
// each sheet's cells set, its notes and names kept; where undo stops
// (reset()).
async function open(wb: Workbook, file: SheetFile): Promise<Outcome> {
  const before = wb.sheet.keys();
  wb.reset(file.sheets[0].name);
  if (file.dateSystem === 1904) await wb.setDateSystem(1904);
  for (const [n, l] of file.cellStyles) wb.styles.set(n, l);
  wb.theme = file.theme;
  // Worked out as it opens, iterating if the file does; manual after.
  wb.calculation = { ...file.calculation, manual: false };
  const failed = file.language === "RexxJS" ? await wb.rexx.setup(file.setupScript) : null;
  await wb.setLanguage(wb.language(file.language));
  for (const [i, s] of file.sheets.entries()) {
    if (i > 0) await wb.add(s.name);
    for (const [n, r] of s.names) {
      const corners = parseRange(r);
      if (corners || r.startsWith("=")) await wb.sheet.define(n, corners ?? r.slice(1));
    }
    await wb.sheet.setMany(s.entries);
    for (const [key, note] of s.notes) wb.notes.set(key, note);
    wb.page(s.name)!.rules = [...s.rules];
    const page = wb.page(s.name)!;
    page.charts = [...s.charts];
    page.hiddenRows = new Set(s.view.hiddenRows);
    page.hiddenCols = new Set(s.view.hiddenCols);
    page.frozen = s.view.frozenRows;
    page.frozenCols = s.view.frozenCols;
    page.tabColor = s.view.tabColor;
    page.hiddenSheet = s.view.hidden;
    page.heights = new Map(s.view.rowHeights);
    page.widths = new Map(s.view.colWidths);
    page.rowLevels = new Map(s.view.rowLevels);
    page.colLevels = new Map(s.view.colLevels);
    page.slicers = [...s.slicers];
    page.scenarios = [...s.scenarios];
    page.autoId = s.autoId;
    page.pivots = new Map(s.pivots);
    page.merges = s.merges.map((m) => parseRange(m)!);
    for (const t of s.tables) await wb.sheet.defineTable(t.name, t);
  }
  // The workbook's names, once its sheets are there: "Costs!B1:B3" or
  // "=LAMBDA(...)"; one naming no sheet of it is left out.
  for (const [n, r] of file.bookNames) {
    if (r.startsWith("=")) {
      await wb.defineName(n, { formula: r.slice(1) });
      continue;
    }
    const m = /^\s*(?:'([^']+)'|([^'!]+))!(.*)$/.exec(r), page = m ? wb.page((m[1] ?? m[2]).trim()) : null, range = m ? parseRange(m[3]) : null;
    if (page && range) await wb.defineName(n, { page, range });
  }
  wb.activate(file.sheets[0].name);
  wb.name = file.name;
  wb.setupScript = file.setupScript;
  await wb.settle();
  wb.calculation = file.calculation;
  // A filter is worked out once every sheet's cells are.
  for (const s of file.sheets) if (s.view.filter) filtering(wb.page(s.name)!, s.view.filter);
  // And a slicer's picks.
  for (const s of file.sheets) {
    const page = wb.page(s.name)!;
    if (page.slicers.some((x) => x.picked.length)) page.filtered = slicedRows(page);
  }
  if (file.sheets.some((s) => s.slicers.length)) await wb.hiddenChanged();
  const cells = file.sheets.reduce((n, s) => n + s.entries.length, 0);
  const message = `opened ${file.name}: ${cells} cells, ${file.language} formulas`;
  return { ok: true, message: failed ? `${message}; its setup script failed: ${failed.split("\n").pop()}` : message, changed: [...new Set([...before, ...wb.sheet.keys()])] };
}

// Saved files: each under "sheet.file.<name>" (any character of the name
// but a letter or digit as _ and its code: "Spec book" is Spec_0020book),
// and their names, in the order saved, under "sheet.files".
const FILES = "sheet.files";
const fileKey = (name: string) => `sheet.file.${name.replace(/[^A-Za-z0-9]/g, (c) => `_${c.charCodeAt(0).toString(16).padStart(4, "0")}`)}`;
const savedNames = (store: Store): string[] => {
  try {
    const v = JSON.parse(store.get(FILES) ?? "[]");
    return Array.isArray(v) ? v.map(String) : [];
  } catch {
    return [];
  }
};
const savedName = (store: Store, name: string): string =>
  savedNames(store).find((n) => n.toUpperCase() === name.trim().toUpperCase()) ?? refuse(`no saved file ${name}`);

// Rows of inputs set from a cell, refused when they do not fit.
function placed(wb: Workbook, at: Addr, rows: string[][], what: string): Promise<Outcome> {
  const cols = Math.max(0, ...rows.map((r) => r.length));
  if (at.col + cols > COLS) refuse(`${cols} columns from ${keyOf(at)} do not fit on the sheet`);
  if (at.row + rows.length > ROWS) refuse(`${rows.length} rows from ${keyOf(at)} do not fit on the sheet`);
  const entries = rows.flatMap((row, i) => row.map((input, j): [string, string] => [keyOf({ col: at.col + j, row: at.row + i }), input]));
  return setting(wb, entries, `imported ${rows.length} rows of ${what} at ${keyOf(at)}`);
}

// A delimiter by name (comma, tab, semicolon, bar, space) or as itself.
const delimiterOf = (s: string): string => CSV_DELIMITERS[s.trim().toLowerCase()] ?? (s.length === 1 && !/[A-Za-z0-9"\r\n]/.test(s) ? s : refuse(`${s} is not a delimiter (${Object.keys(CSV_DELIMITERS).join(", ")})`));

// The used part of the sheet: A1 to its last row and column.
function usedRange(wb: Workbook): [Addr, Addr] {
  const used = wb.sheet.keys().map((k) => parseRange(k)![0]);
  return [{ col: 0, row: 0 }, { col: Math.max(0, ...used.map((a) => a.col)), row: Math.max(0, ...used.map((a) => a.row)) }];
}

// Rows or columns inserted or deleted: every cell and formula moved, as one
// change (one undo); notes and conditional styles follow, and so do other
// sheets' references to this one (each sheet's own change).
function restructure(how: "insert" | "delete", axis: "row" | "col") {
  return async (wb: Workbook, p: Params): Promise<Outcome> => {
    const at = axis === "row" ? Number(p.at) - 1 : columnOf(p.at);
    if (!Number.isInteger(at) || at < 0 || at >= (axis === "row" ? ROWS : COLS)) refuse(`${p.at} is not a ${axis === "row" ? "row" : "column"} on the sheet`);
    const n = countOf(p.count);
    const entries = await reshaped(wb, axis, (moves) => (how === "insert" ? inserted(axis, at, n, moves) : deleted(axis, at, n, moves)));
    // The Tauri app's auto IDs: each row inserted is given the next one.
    const page = wb.page(wb.active)!;
    if (how === "insert" && axis === "row" && page.autoId) {
      for (let row = at; row < at + n; row++) {
        entries.push([keyOf({ col: page.autoId.column, row }), `${page.autoId.prefix}${page.autoId.next}`]);
        page.autoId = { ...page.autoId, next: page.autoId.next + 1 };
      }
    }
    const thing = `${n} ${axis === "row" ? "row" : "column"}${n === 1 ? "" : "s"}`;
    return setting(wb, entries, `${how === "insert" ? "inserted" : "deleted"} ${thing} at ${axis === "row" ? at + 1 : colName(at)}`);
  };
}

// The active sheet reshaped by a structural edit (rows or columns inserted,
// deleted or swapped, or cells moved: axis null): the references to it on
// other sheets rewritten, and its notes, rules, charts, merges, tables and
// filter moved with its cells, and for rows or columns its hidden lines,
// heights or widths and auto IDs too. Gives its cells' entries, for
// setting().
async function reshaped(wb: Workbook, axis: "row" | "col" | null, editOf: (moves: Scope) => Edit): Promise<[string, string][]> {
  const here = wb.active.toUpperCase();
  const edit = editOf((sheet) => sheet === null || sheet === here);
  const entries = restructured(wb.sheet.keys().map((k) => [k, wb.sheet.input(k)]), edit, wb.sheet.language);
  const there = editOf((sheet) => sheet === here);
  for (const page of wb.pages()) {
    if (page.sheet === wb.sheet) continue;
    const rewritten = page.sheet.keys().map((k): [string, string] => [k, page.sheet.language.rewrite(page.sheet.input(k), there)])
      .filter(([k, input]) => input !== page.sheet.input(k));
    if (rewritten.length) await page.sheet.setMany(rewritten);
  }
  wb.notes.remap((a) => edit.cell(a, { col: false, row: false }), (e) => wb.sheet.language.rewrite(`=${e}`, edit).slice(1));
  const page = wb.page(wb.active)!, still = { col: false, row: false };
  page.rules = page.rules.flatMap((r): CfRule[] => {
    const to = edit.range(r.range[0], r.range[1], still, still);
    return to ? [{ ...r, range: to }] : [];
  });
  page.charts = page.charts.flatMap((c): ChartSpec[] => {
    const to = edit.range(c.range[0], c.range[1], still, still);
    return to ? [{ ...c, range: to }] : [];
  });
  // Hidden rows (columns) move with the rest, and go with deleted ones.
  const moved = (set: Set<number>, row: boolean) => new Set([...set].flatMap((i) => {
    const to = edit.cell(row ? { col: 0, row: i } : { col: i, row: 0 }, still);
    return to ? [row ? to.row : to.col] : [];
  }));
  if (axis === "row") page.hiddenRows = moved(page.hiddenRows, true);
  if (axis === "col") page.hiddenCols = moved(page.hiddenCols, false);
  if (page.filtered && axis === "row") page.filtered = moved(page.filtered, true);
  page.merges = page.merges.flatMap((m): [Addr, Addr][] => {
    const to = edit.range(m[0], m[1], still, still);
    return to && keyOf(to[0]) !== keyOf(to[1]) ? [to] : [];
  });
  for (const t of wb.sheet.tables()) {
    const to = edit.range(t.range[0], t.range[1], still, still);
    if (!to || rangeName(to) !== rangeName(t.range) || axis) await wb.sheet.defineTable(t.name, to ? { ...t, range: to, sort: undefined } : null);
  }
  // Outline levels move with their rows (columns), and go with deleted ones.
  const levelsMoved = (m: Map<number, number>, row: boolean) => new Map([...m].flatMap(([i, l]): [number, number][] => {
    const to = edit.cell(row ? { col: 0, row: i } : { col: i, row: 0 }, still);
    return to ? [[row ? to.row : to.col, l]] : [];
  }));
  if (axis === "row") page.rowLevels = levelsMoved(page.rowLevels, true);
  if (axis === "col") page.colLevels = levelsMoved(page.colLevels, false);
  if (axis === "row") {
    page.heights = new Map([...page.heights].flatMap(([row, h]): [number, number][] => {
      const to = edit.cell({ col: 0, row }, still);
      return to ? [[to.row, h]] : [];
    }));
  } else if (axis === "col") {
    page.widths = new Map([...page.widths].flatMap(([col, w]): [number, number][] => {
      const to = edit.cell({ col, row: 0 }, still);
      return to ? [[to.col, w]] : [];
    }));
  }
  page.pivots = new Map([...page.pivots].flatMap(([id, pv]): [string, Pivot][] => {
    const source = edit.range(pv.source[0], pv.source[1], still, still), output = edit.cell(pv.output, still);
    const written = pv.written && edit.range(pv.written[0], pv.written[1], still, still);
    return source && output ? [[id, { ...pv, source, output, written }]] : [];
  }));
  if (page.autoId && axis) {
    const to = edit.cell({ col: page.autoId.column, row: 0 }, still);
    page.autoId = to ? { ...page.autoId, column: to.col } : null;
  }
  if (page.filter) {
    const r = edit.range(page.filter.range[0], page.filter.range[1], still, still);
    const col = axis ? edit.cell({ col: page.filter.column, row: 0 }, still) : r && { col: page.filter.column + r[0].col - page.filter.range[0].col, row: 0 };
    page.filter = r && col ? { ...page.filter, range: r, column: col.col } : null;
    if (!page.filter) page.filtered = null;
  }
  return entries;
}

// The clipboard for a range of the active sheet: as text, and cell by cell.
function clipOf(wb: Workbook, r: [Addr, Addr]): Clip {
  const from = topLeft(r);
  const rows = Math.abs(r[0].row - r[1].row) + 1, cols = Math.abs(r[0].col - r[1].col) + 1;
  const grid = <T>(f: (key: string) => T): T[][] =>
    Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => f(keyOf({ col: from.col + j, row: from.row + i }))));
  return {
    tsv: copied(inputs(wb), r[0], r[1]), from, language: wb.sheet.language,
    inputs: grid((k) => wb.sheet.input(k)), shown: grid((k) => show(wb.sheet.value(k))), notes: grid((k) => wb.notes.get(k)),
  };
}

// Cut cells moved to `at` (their top-left), on the sheet they were cut
// from: one structural edit (shift.ts's blockMoved), so one undo.
async function moving(wb: Workbook, cut: NonNullable<Clip["cut"]>, at: Addr): Promise<Outcome> {
  if (!wb.pages().includes(cut.page)) refuse("the sheet the cells were cut from is gone");
  if (cut.page !== wb.page(wb.active)) refuse(`cut cells are pasted on the sheet they were cut from, ${cut.page.name}`);
  const from = topLeft(cut.range), by = { cols: at.col - from.col, rows: at.row - from.row };
  const rows = Math.abs(cut.range[0].row - cut.range[1].row), cols = Math.abs(cut.range[0].col - cut.range[1].col);
  if (at.col + cols >= COLS || at.row + rows >= ROWS) refuse(`${named(cut.range)} moved to ${keyOf(at)} does not fit on the sheet`);
  wb.clip = null;
  if (by.cols === 0 && by.rows === 0) return { ok: true, message: `moved ${named(cut.range)} to where it was`, changed: [] };
  const entries = await reshaped(wb, null, (moves) => blockMoved(cut.range, by, moves));
  return setting(wb, entries, `moved ${named(cut.range)} to ${keyOf(at)}`);
}

// Fill down, right, up or left over a range (copy.ts), from its first
// `source` lines (1: Ctrl+D's first row).
function fill(way: FillWay) {
  return (wb: Workbook, p: Params): Promise<Outcome> => {
    const r = rangeOf(p.range), lines = countOf(p.source);
    return setting(wb, filled(inputs(wb), r[0], r[1], way, wb.sheet.language, lines), `filled ${named(r)} ${way}`);
  };
}

function sort(ascending: boolean) {
  return async (wb: Workbook, p: Params): Promise<Outcome> => {
    const r = rangeOf(p.range), by = columnOf(p.by);
    const left = Math.min(r[0].col, r[1].col), right = Math.max(r[0].col, r[1].col);
    if (by < left || by > right) refuse(`${p.by} is not a column of ${named(r)}`);
    const entries = rearranging(wb, r, sortedRows(wb.sheet, r[0], r[1], by - left, ascending));
    return setting(wb, entries, `sorted ${named(r)} by ${colName(by)}, ${ascending ? "ascending" : "descending"}`);
  };
}

const scenarioNamed = (wb: Workbook, n: string): Scenario =>
  wb.page(wb.active)!.scenarios.find((x) => x.name.toLowerCase() === n.trim().toLowerCase()) ?? refuse(`no scenario ${n}`);

// The rows the page's slicers hide, as its filter's (the AutoFilter's
// taken off: a table's slicers are its filter).
async function sliced(wb: Workbook, page: Page, said: string): Promise<Outcome> {
  const rows = slicedRows(page);
  page.filter = null;
  page.filtered = rows.size || page.slicers.some((x) => x.picked.length) ? rows : null;
  await wb.hiddenChanged();
  return changedBy(wb, wb.sheet.keys(), `${said} (${rows.size} row${rows.size === 1 ? "" : "s"} hidden)`);
}

// A sheet as pdf.ts lays it out: its used range (or the one asked), each
// cell as it shows, in its look; sizes in points (a pixel three quarters
// of one), hidden rows and columns left out.
function pdfSheet(wb: Workbook, page: Page, r: [Addr, Addr] | null): PdfSheet {
  const used = page.sheet.keys().map((k) => cellAddr(k)!);
  const [a, b] = r ?? [{ col: 0, row: 0 }, { col: Math.max(0, ...used.map((x) => x.col)), row: Math.max(0, ...used.map((x) => x.row)) }];
  const c0 = Math.min(a.col, b.col), r0 = Math.min(a.row, b.row);
  const cols = Array.from({ length: Math.abs(a.col - b.col) + 1 }, (_, i) => c0 + i).filter((c) => !page.hiddenCols.has(c));
  const rows = Array.from({ length: Math.abs(a.row - b.row) + 1 }, (_, i) => r0 + i).filter((x) => !page.hiddenRows.has(x) && !page.filtered?.has(x));
  const was = wb.active;
  return {
    name: page.name, rows: rows.length, cols: cols.length, firstRow: r0, firstCol: c0,
    colWidth: (c) => (page.widths.get(cols[c]) ?? PDF_COLUMN_PX) * 0.75, rowHeight: (x) => (page.heights.get(rows[x]) ?? PDF_ROW_PX) * 0.75,
    cells: (ri, ci) => {
      const key = keyOf({ col: cols[ci], row: rows[ri] }), note = page.notes.get(key), look = wb.lookOf(note) ?? {}, v = page.sheet.value(key);
      if (was !== page.name) wb.activate(page.name);
      const text = wb.textOf(key);
      if (was !== page.name) wb.activate(was);
      if (!text && !look.fill && !look.borders) return null;
      const line = (s?: string) => (s ? borderLine(s).split(" ").pop() : undefined);
      return { text, bold: look.bold, italic: look.italic, size: look.size ? look.size * 0.75 : undefined, color: look.color ? resolved(look.color) : v instanceof CellError ? "#c0392b" : undefined,
        fill: look.fill ? resolved(look.fill) : undefined, align: look.align ?? (typeof v === "number" ? "end" : "start"),
        borders: look.borders ? { top: line(look.borders.top), right: line(look.borders.right), bottom: line(look.borders.bottom), left: line(look.borders.left) } : undefined };
    },
  };
}
const PDF_COLUMN_PX = 80, PDF_ROW_PX = 24;

// A theme as words: its accents and fonts.
const themeSaid = (t: Theme): string => `accents ${["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((r) => t.colors[r as Role]).join(" ")}, headings ${t.fonts.major}, body ${t.fonts.minor}`;

// A look as words: "bold, size 15, fill accent1, bottom border thick accent1".
function describeLook(l: Look): string {
  const parts = [l.bold ? "bold" : "", l.italic ? "italic" : "", l.underline ? "underline" : "", l.strike ? "strikethrough" : "", l.size ? `size ${l.size}` : "",
    l.fill ? `fill ${l.fill}` : "", l.color ? `text ${l.color}` : "", l.number ? `number ${l.number}` : "",
    ...Object.entries(l.borders ?? {}).map(([side, line]) => `${side} border ${line}`)].filter((x) => x);
  return parts.length ? parts.join(", ") : "as Normal";
}

// ---- The outline (outline.ts).

// Rows, unless axis=columns (or the range is whole columns, "B:D").
const axisOf = (p: Params): boolean => {
  const a = (p.axis ?? "").trim().toLowerCase();
  if (a && !["rows", "row", "columns", "column", "cols"].includes(a)) refuse(`${p.axis} is not rows or columns`);
  return a ? a.startsWith("row") : !/^\s*\$?[A-Za-z]{1,3}\s*:\s*\$?[A-Za-z]{1,3}\s*$/.test(p.range ?? "");
};

async function outlined(wb: Workbook, p: Params, by: 1 | -1): Promise<Outcome> {
  const rows = axisOf(p), page = wb.page(wb.active)!, { at, said } = spanOf(rangeOf(p.range), rows);
  const was = rows ? page.rowLevels : page.colLevels;
  if (by > 0 && at.some((i) => (was.get(i) ?? 0) >= MOST_LEVELS)) refuse(`an outline has ${MOST_LEVELS} levels at most`);
  if (by < 0 && at.every((i) => !was.get(i))) refuse(`${said} ${at.length === 1 ? "is" : "are"} not grouped`);
  const next = regrouped(was, at[0], at[at.length - 1], by);
  if (rows) page.rowLevels = next; else page.colLevels = next;
  // Ungrouped lines no outline hides any more are shown.
  if (by < 0) for (const i of at) if (!next.get(i)) (rows ? page.hiddenRows : page.hiddenCols).delete(i);
  if (rows) await wb.hiddenChanged();
  return changedBy(wb, wb.sheet.keys(), `${by > 0 ? "grouped" : "ungrouped"} ${said}: ${depth(next)} level${depth(next) === 1 ? "" : "s"}`);
}

async function detail(wb: Workbook, p: Params, hide: boolean): Promise<Outcome> {
  const rows = axisOf(p), page = wb.page(wb.active)!, { at } = spanOf(rangeOf(p.range), rows);
  const lines = detailOf(rows ? page.rowLevels : page.colLevels, at[0]);
  if (!lines.length) refuse(`${rows ? `row ${at[0] + 1}` : `column ${colName(at[0])}`} is in no group`);
  const set = rows ? page.hiddenRows : page.hiddenCols;
  for (const i of lines) if (hide) set.add(i); else set.delete(i);
  if (rows) await wb.hiddenChanged();
  const name = (i: number) => (rows ? String(i + 1) : colName(i));
  return changedBy(wb, wb.sheet.keys(), `${hide ? "hid" : "showed"} the detail ${name(lines[0])} to ${name(lines[lines.length - 1])}`);
}

// A rectangle's rows put in an order (the row each place takes its cells
// from), their formulas moved and their notes going with them: the entries
// for setting().
function rearranging(wb: Workbook, r: [Addr, Addr], rows: readonly number[]): [string, string][] {
  const entries = rearranged(wb.sheet, r[0], r[1], rows, wb.sheet.language);
  wb.notes.remap(rearrangedMoves(r[0], r[1], rows),
    (e, from, to) => movedBy(wb.sheet.language, `=${e}`, { cols: 0, rows: to.row - from.row }).slice(1));
  return entries;
}

// A name for a range: a letter first, then letters, digits or _; not a
// cell, nor TRUE or FALSE.
// A named formula's text, what follows its '=': refused when it does not
// parse in the sheet's language (in RexxJS formulas, a LAMBDA whose body is
// RexxJS), or would hide a function of the same name.
function formulaOf(wb: Workbook, name: string, text: string): string {
  if (!text.startsWith("=")) refuse(`a named formula starts with =: ${text}`);
  if (wb.sheet.language.functions().includes(name)) refuse(`${name} is a function already`);
  if (wb.sheet.language.name === "Excel") {
    const c = wb.sheet.language.compile(text.slice(1));
    if (c.kind === "broken") refuse(c.error.why);
  } else {
    const l = lambdaOf(text.slice(1));
    if (typeof l === "string") refuse(l);
  }
  return text.slice(1);
}

const nameOf = (s: string): string => {
  if (/^[A-Za-z]+\d+$/.test(s)) refuse(`${s} is a cell, not a name`);
  if (!/^[A-Za-z_]\w*$/.test(s) || /^(TRUE|FALSE)$/i.test(s)) refuse(`${s} is not a name: a letter first, then letters, digits or _`);
  return s.toUpperCase();
};

// A sheet's name: letters, digits, spaces or _ (Excel's quotes take the
// spaces: 'My Costs'!B2).
const sheetNameOf = (s: string): string => {
  const t = String(s).trim();
  return /^[A-Za-z0-9_ ]+$/.test(t) ? t : refuse(`${s} is not a sheet name: letters, digits, spaces or _`);
};
const pageOf = (wb: Workbook, name: string) => wb.page(name) ?? refuse(`no sheet ${name}`);

const scopeOf = (s: string | undefined): "workbook" | "sheet" =>
  s === undefined || /^workbook$/i.test(s) ? "workbook" : /^sheet$/i.test(s) ? "sheet" : refuse(`${s} is not a scope: workbook or sheet`);

// A range, perhaps of another sheet ("Costs!B1:B3", "'My Costs'!B1"), and
// its sheet (the active one when it names none).
function sheetRangeOf(wb: Workbook, s: string): { page: Page; range: [Addr, Addr] } {
  const m = /^\s*(?:'([^']+)'|([^'!]+))!(.*)$/.exec(s);
  return m ? { page: pageOf(wb, (m[1] ?? m[2]).trim()), range: rangeOf(m[3]) } : { page: wb.page(wb.active)!, range: rangeOf(s) };
}

// What the grid shows changes wholesale when the active sheet does.
const shownKeys = (wb: Workbook, before: readonly string[]) => [...new Set([...before, ...wb.sheet.keys()])];

// ---- What-if and data tools (whatif.ts).

const numberOf = (s: string): number => {
  const n = Number(s);
  return s.trim() !== "" && Number.isFinite(n) ? n : refuse(`${s} is not a number`);
};

// A value as an input that comes back to it (a number in full).
const inputOf = (v: Value): string => (v === null ? "" : v instanceof CellError ? v.code : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v));

// Cells tried with each of `trials` in turn, then put back as they were;
// what `read` gave for each.
async function tried<T>(wb: Workbook, cells: readonly string[], trials: readonly (readonly string[])[], read: () => T): Promise<T[]> {
  const was = cells.map((k) => [k, wb.sheet.input(k)] as [string, string]);
  const out: T[] = [];
  try {
    for (const t of trials) {
      await wb.sheet.setMany(cells.map((k, i) => [k, t[i]] as [string, string]));
      out.push(read());
    }
  } finally {
    await wb.sheet.setMany(was);
  }
  return out;
}

const DELIMITERS: Record<string, string> = { comma: ",", tab: "\t", space: " ", semicolon: ";" };

// ---- Conditional-format rules (rules.ts), on the active sheet.

const COMPARISONS: Record<string, Comparison> = {
  greater: ">", less: "<", between: "between", equal: "=", notequal: "<>", greaterorequal: ">=", lessorequal: "<=", contains: "contains",
  notbetween: "notBetween", notcontains: "notContains", beginswith: "beginsWith", endswith: "endsWith",
  ">": ">", "<": "<", "=": "=", "<>": "<>", ">=": ">=", "<=": "<=",
};
// Excel's Light Red Fill with Dark Red Text.
const HIGHLIGHT = "background: #ffc7ce; color: #9c0006";
const styleParam = (p: Params): string => (p.style === undefined ? HIGHLIGHT : toStyle(p.style) || refuse(`${p.style} is not a style the grid can draw`));

// A rule added to the active sheet's, last (so the earlier ones win);
// stop=1 makes it Stop If True.
async function addRule(wb: Workbook, rule: CfRule, p: Params = {}): Promise<Outcome> {
  if (flag(p.stop)) rule = { ...rule, stop: true };
  wb.page(wb.active)!.rules.push(rule);
  return changedBy(wb, [...range(rule.range[0], rule.range[1])], describeCf(rule));
}

// ---- A cell's look (formats.ts's Look), a part at a time.

// Each cell of a range's look changed by `change`.
function relook(wb: Workbook, r: [Addr, Addr], change: (look: Look, key: string) => Look): string[] {
  const keys = [...range(r[0], r[1])];
  for (const k of keys) {
    const look = change(wb.notes.get(k).look ?? {}, k);
    wb.notes.set(k, { look: Object.values(look).some((v) => v !== undefined && v !== false) ? look : undefined });
  }
  return keys;
}

// A colour over a range: the look's fill or text colour.
function coloured(wb: Workbook, p: Params, part: "fill" | "color"): Promise<Outcome> {
  const r = rangeOf(p.range), off = /^(none|auto|automatic)$/i.test(p.color.trim());
  const c = off ? undefined : colour(p.color) ?? refuse(`${p.color} is not a colour (a name such as red, or #rrggbb)`);
  return changedBy(wb, relook(wb, r, (look) => ({ ...look, [part]: c })), `${rangeName(r)}: ${part === "fill" ? "fill" : "text"} ${c ?? "none"}`);
}

// Excel's Borders menu: the range's outline, every cell's every side, the
// lines inside it (both ways, or across or down only), one side of the
// range, or none (Clear borders, as the eraser over all of it).
const EDGES = ["outline", "all", "inside", "horizontal", "vertical", "top", "right", "bottom", "left", "none"];
// A line's style, as Excel's Line Style list (formats.ts draws its width).
const LINE_STYLES = ["thin", "medium", "thick", "double", "dashed", "dotted", "hair", "none"];

// ---- Rows and columns out of view.

// The rows (or columns) a range covers, 0-based, and as people count them.
function spanOf(r: [Addr, Addr], rows: boolean): { at: number[]; said: string } {
  const lo = rows ? Math.min(r[0].row, r[1].row) : Math.min(r[0].col, r[1].col), hi = rows ? Math.max(r[0].row, r[1].row) : Math.max(r[0].col, r[1].col);
  const name = (i: number) => (rows ? String(i + 1) : colName(i));
  const at = Array.from({ length: hi - lo + 1 }, (_, i) => lo + i);
  return { at, said: `${rows ? "row" : "column"}${at.length > 1 ? `s ${name(lo)} to ${name(hi)}` : ` ${name(lo)}`}` };
}

// A page's AutoFilter set, and the rows it hides worked out: under the
// range's first row (its headers), those whose value in the column fails
// the criterion, read as COUNTIF reads one (5, ">100", "<>x", "north").
function filtering(page: Page, filter: Filter): Set<number> {
  const [a, b] = filter.range, keep = criterion(NUMERIC_TEXT.test(filter.criterion) ? Number(filter.criterion) : filter.criterion);
  const hidden = new Set<number>();
  for (let row = Math.min(a.row, b.row) + 1; row <= Math.max(a.row, b.row); row++) {
    if (!keep(page.sheet.value(keyOf({ col: filter.column, row })))) hidden.add(row);
  }
  page.filter = filter;
  page.filtered = hidden;
  return hidden;
}

function hiding(rows: boolean, hide: boolean) {
  return async (wb: Workbook, p: Params): Promise<Outcome> => {
    const page = wb.page(wb.active)!, { at, said } = spanOf(rangeOf(p.range), rows);
    const set = rows ? page.hiddenRows : page.hiddenCols;
    for (const i of at) if (hide) set.add(i); else set.delete(i);
    if (rows) await wb.hiddenChanged();
    return changedBy(wb, wb.sheet.keys(), `${hide ? "hid" : "showed"} ${said}`);
  };
}

// A table's columns' names: given, or its first row's, or its letters.
// A pivot written from its output cell: what it covered before cleared,
// its cells set (one undo), and it kept with what it now covers.
async function writePivot(wb: Workbook, id: string, pivot: Pivot): Promise<Outcome> {
  const page = wb.page(wb.active)!, [a, b] = pivot.source, rows: Value[][] = [];
  for (let row = Math.min(a.row, b.row); row <= Math.max(a.row, b.row); row++) {
    const line: Value[] = [];
    for (let col = Math.min(a.col, b.col); col <= Math.max(a.col, b.col); col++) line.push(wb.sheet.value(keyOf({ col, row })));
    rows.push(line);
  }
  const grid = pivoted(rows, pivot), o = pivot.output;
  const end = { col: o.col + Math.max(...grid.map((r) => r.length)) - 1, row: o.row + grid.length - 1 };
  if (end.col >= COLS || end.row >= ROWS) refuse(`pivot ${id} is ${grid.length} rows by ${end.col - o.col + 1} columns: it does not fit from ${keyOf(o)}`);
  const sets = new Map<string, string>();
  grid.forEach((line, r) => line.forEach((v, c) => sets.set(keyOf({ col: o.col + c, row: o.row + r }), v)));
  const clears = pivot.written ? [...range(pivot.written[0], pivot.written[1])].filter((k) => !sets.has(k)).map((k): [string, string] => [k, ""]) : [];
  page.pivots.set(id, { ...pivot, written: [o, end] });
  return setting(wb, [...clears, ...sets], `pivot ${id} at ${rangeName([o, end])}`);
}

// A pivot as the app's GETPIVOT gives its configuration.
const pivotConfig = (pv: Pivot) => ({
  sourceRange: rangeName(pv.source), rowFields: [...pv.rows], colFields: [...pv.cols], valueField: pv.value, aggFunction: pv.agg, outputCell: keyOf(pv.output), filters: { ...pv.filters },
  ...(pv.options && Object.keys(pv.options).length ? { options: pv.options } : {}),
});

// A table's types or descriptions, as JSON: {"amount": "number"}.
function byName(json: string | undefined, what: string): Record<string, string> | undefined {
  if (json === undefined) return undefined;
  let o: unknown;
  try {
    o = JSON.parse(json);
  } catch {
    o = null;
  }
  if (!o || typeof o !== "object" || Array.isArray(o)) refuse(`${what} must be a JSON object, a column's name to text`);
  return Object.fromEntries(Object.entries(o as object).map(([k, v]) => [k, String(v)]));
}

function tableColumns(wb: Workbook, t: TableDef): string[] {
  const left = Math.min(t.range[0].col, t.range[1].col), right = Math.max(t.range[0].col, t.range[1].col);
  const cols = Array.from({ length: right - left + 1 }, (_, i) => left + i);
  if (t.columns) return [...t.columns];
  return cols.map((col) => (t.headers ? show(wb.sheet.value(keyOf({ col, row: t.range[0].row }))) : colName(col)));
}

// Two ranges that share a cell.
const overlap = ([a, b]: [Addr, Addr], [c, d]: [Addr, Addr]): boolean =>
  Math.max(Math.min(a.col, b.col), Math.min(c.col, d.col)) <= Math.min(Math.max(a.col, b.col), Math.max(c.col, d.col))
  && Math.max(Math.min(a.row, b.row), Math.min(c.row, d.row)) <= Math.min(Math.max(a.row, b.row), Math.max(c.row, d.row));

// Trace Precedents or Dependents: a level further each time.
function trace(direction: Direction) {
  const [verb, more] = direction === "precedents" ? ["reads", "they read"] : ["is read by", "they are read by"];
  return async (wb: Workbook, p: Params): Promise<Outcome> => {
    const key = cellOf(p.cell), levels = wb.trace(key, direction);
    await wb.settle();
    const said = levels.length ? levels.map((l, i) => `${i ? `, and ${more}` : verb} ${l.join(" ")}`).join("") : `${verb} nothing`;
    return { ok: true, message: `${key} ${said}`, changed: wb.sheet.keys(), output: levels };
  };
}

// A watched cell: "Costs!B2", or B2 on the active sheet.
const watchedOf = (wb: Workbook, s: string): [Page, string] => {
  const { sheet, cell } = splitKey(s.trim());
  return [sheet === null ? wb.page(wb.active)! : pageOf(wb, sheet), cellOf(cell)];
};

const findOptions = (p: Params) => ({ matchCase: flag(p.matchCase), wholeCell: flag(p.wholeCell), inFormulas: flag(p.inFormulas) });
const FIND_OPTIONS = ["matchCase", "wholeCell", "inFormulas"];

const SPECS: readonly Spec[] = [
  {
    id: "cell.set", label: "Set a cell", params: ["cell", "input"], journaled: true,
    run: async (wb, p) => {
      const key = cellOf(p.cell);
      // Its rule, if it has one: stop refuses what breaks it, warn says so.
      const rule = wb.notes.get(key).validation, was = wb.sheet.input(key);
      let broken = rule ? check(rule, p.input, was === "") : null;
      if (broken && rule!.mode === "stop") refuse(`${key} ${broken}`);
      // Its editor, if it has one, takes only what it offers.
      const editor = wb.notes.get(key).editor, unsuited = editor ? refusal(editor, p.input) : null;
      if (unsuited) refuse(`${key} ${unsuited}`);
      let changed = await wb.sheet.set(key, p.input);
      // A custom rule's formula, moved from the rule's top-left to this
      // cell, worked out with the input in; a stop puts the cell back.
      if (rule?.kind === "custom" && p.input.trim() !== "" && !p.input.trim().startsWith("=") && check(rule, p.input, was === "") === null
          && !((rule.when === "create" && was !== "") || (rule.when === "update" && was === ""))) {
        const from = cellAddr(rule.origin ?? key)!, at = cellAddr(key)!;
        const f = movedBy(wb.sheet.language, `=${rule.formula}`, { cols: at.col - from.col, rows: at.row - from.row }).slice(1);
        const v = await wb.sheet.evaluate(f);
        if (!(v === true || (typeof v === "number" && v !== 0))) {
          broken = errorOf(rule);
          if (rule.mode === "stop") {
            await wb.sheet.set(key, was);
            await wb.settle();
            refuse(`${key} ${broken}`);
          }
        }
        changed = [...new Set([...changed, key])];
      }
      const said = broken ? ` (${rule!.mode === "info" ? "note: " : ""}${key} ${broken})` : "";
      return changedBy(wb, changed, `${key} = ${show(wb.sheet.value(key))}${said}`);
    },
  },
  {
    id: "cell.get", label: "Read a cell", params: ["cell"], journaled: false,
    run: async (wb, p) => {
      const key = cellOf(p.cell), shown = show(wb.sheet.value(key));
      return { ok: true, message: `${key} = ${shown}`, changed: [], output: shown };
    },
  },
  {
    // All a script may ask of a cell (control.ts's GETFORMAT, GETCOMMENT,
    // GETMERGEDRANGE and the rest read it).
    id: "cell.about", label: "About a cell", params: ["cell"], journaled: false,
    run: async (wb, p) => {
      const key = cellOf(p.cell), note = wb.notes.get(key), at = cellAddr(key)!;
      const merge = wb.page(wb.active)!.merges.find((m) => overlap(m, [at, at]));
      const about = {
        cell: key, value: scripted(wb.sheet.value(key)), input: wb.sheet.input(key), format: note.format ?? "", comment: note.comment ?? "",
        wrap: note.look?.wrap ? 1 : 0, merged: merge ? rangeName(merge) : "", validation: note.validation ?? null, editor: note.editor ? asApp(note.editor) : null,
      };
      return { ok: true, message: `${key}: ${JSON.stringify(about)}`, changed: [], output: about };
    },
  },
  {
    // A cell's editor (editors.ts): a checkbox, a dropdown of options, a
    // date; kind=none takes it away.
    id: "cell.editor", label: "Cell Editor", params: ["cell", "kind"], optional: ["options", "format"], journaled: true,
    run: async (wb, p) => {
      const key = cellOf(p.cell);
      if (p.kind.toLowerCase() === "none") {
        if (!wb.notes.get(key).editor) refuse(`${key} has no editor`);
        wb.notes.set(key, { editor: undefined });
        return changedBy(wb, [key], `${key} has no editor`);
      }
      const editor = editorOf(p.kind, p.options, p.format);
      wb.notes.set(key, { editor });
      return changedBy(wb, [key], `${key} is a ${editor.kind}${editor.options ? ` of ${editor.options.join(", ")}` : ""}`);
    },
  },
  {
    // A checkbox ticked, or not.
    id: "cell.toggle", label: "Toggle", params: ["cell"], journaled: true,
    run: async (wb, p) => {
      const key = cellOf(p.cell);
      if (wb.notes.get(key).editor?.kind !== "checkbox") refuse(`${key} is not a checkbox`);
      const changed = await wb.sheet.set(key, toggled(wb.sheet.value(key)));
      return changedBy(wb, changed, `${key} = ${show(wb.sheet.value(key))}`);
    },
  },
  {
    id: "cell.comment", label: "Comment", params: ["cell", "text"], journaled: true,
    run: async (wb, p) => {
      const key = cellOf(p.cell);
      wb.notes.set(key, { comment: p.text });
      return changedBy(wb, [key], `${key} comment ${p.text}`);
    },
  },
  {
    id: "cell.format", label: "Number format or style", params: ["cell", "format"], journaled: true,
    run: async (wb, p) => {
      const key = cellOf(p.cell);
      wb.notes.set(key, { format: p.format });
      return changedBy(wb, [key], `${key} format ${p.format}`);
    },
  },
  {
    id: "range.clear", label: "Clear contents", params: ["range"], journaled: true,
    run: (wb, p) => {
      const r = rangeOf(p.range);
      return setting(wb, [...range(r[0], r[1])].map((k) => [k, ""] as const), `cleared ${named(r)}`);
    },
  },
  {
    // A rule over a range: list, whole, decimal, length, date, time (min=
    // and max= as written or as serials), custom (formula= for the
    // range's top-left), pattern (pattern= a regular expression); mode=
    // stop, warn or info; prompt= the input message, error= its own.
    id: "validation.set", label: "Data Validation", params: ["range", "kind"], optional: ["values", "min", "max", "formula", "pattern", "mode", "when", "prompt", "error"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), keys = [...range(r[0], r[1])];
      let rule = ruleOf(p);
      if (rule.kind === "custom") {
        const compiled = wb.sheet.language.compile(rule.formula!);
        if (compiled.kind === "broken") refuse(`${p.formula} is not a formula: ${compiled.error.why}`);
        rule = { ...rule, origin: keyOf(topLeft(r)) };
      }
      for (const k of keys) wb.notes.set(k, { validation: rule });
      return changedBy(wb, keys, `${named(r)} ${describeRule(rule)}`);
    },
  },
  {
    id: "validation.clear", label: "Clear Validation", params: ["range"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), keys = [...range(r[0], r[1])];
      for (const k of keys) wb.notes.set(k, { validation: undefined });
      return changedBy(wb, keys, `${named(r)} has no rule`);
    },
  },
  {
    id: "validation.circle", label: "Circle Invalid Data", params: [], optional: ["off"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      wb.circling = !flag(p.off);
      const invalid = wb.invalid();
      await wb.settle();
      return { ok: true, message: wb.circling ? `${invalid.length} invalid: ${invalid.join(" ") || "none"}` : "circles cleared", changed: invalid, output: invalid };
    },
  },
  {
    id: "range.clearFormats", label: "Clear Formats", params: ["range"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), keys = [...range(r[0], r[1])];
      for (const k of keys) wb.notes.set(k, { format: "", style: "", styleExpression: "", look: undefined });
      return changedBy(wb, keys, `cleared the formats of ${named(r)}`);
    },
  },
  {
    id: "range.clearComments", label: "Clear Comments", params: ["range"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), keys = [...range(r[0], r[1])];
      for (const k of keys) wb.notes.set(k, { comment: "" });
      return changedBy(wb, keys, `cleared the comments of ${named(r)}`);
    },
  },
  {
    id: "range.clearAll", label: "Clear All", params: ["range"], journaled: true,
    run: (wb, p) => {
      const r = rangeOf(p.range), keys = [...range(r[0], r[1])];
      for (const k of keys) wb.notes.set(k, { comment: "", format: "", style: "", styleExpression: "", look: undefined });
      return setting(wb, keys.map((k) => [k, ""] as const), `cleared all of ${named(r)}`);
    },
  },
  {
    id: "sheet.clear", label: "Clear all", params: [], journaled: true,
    run: (wb) => {
      wb.notes.clear();
      return setting(wb, wb.sheet.keys().map((k) => [k, ""] as const), "cleared");
    },
  },
  {
    id: "edit.copy", label: "Copy", params: ["range"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range);
      wb.clip = clipOf(wb, r);
      return { ok: true, message: `copied ${named(r)}`, changed: [] };
    },
  },
  {
    // Cut: copied, and marked to be moved by the next paste (on the same
    // sheet), as a spreadsheet's cut is.
    id: "edit.cut", label: "Cut", params: ["range"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range);
      wb.clip = { ...clipOf(wb, r), cut: { page: wb.page(wb.active)!, range: r } };
      return { ok: true, message: `cut ${named(r)}`, changed: [] };
    },
  },
  {
    // What was copied pasted with its first cell at `at`; when `at` is a
    // range as many times the copy's size each way, into each part of it (a
    // cell copied into every cell). What was cut is moved instead: its
    // cells, notes and what reads them (here or on another sheet) go to
    // `at`'s first cell, what was under them is gone, and the cut is used up.
    id: "edit.paste", label: "Paste", params: ["at"], journaled: true,
    run: (wb, p) => {
      const r = rangeOf(p.at), at = topLeft(r);
      const clip = wb.clip ?? refuse("nothing copied");
      if (clip.cut) return moving(wb, clip.cut, at);
      const h = clip.inputs.length, w = clip.inputs[0]?.length ?? 1;
      const rows = Math.abs(r[0].row - r[1].row) + 1, cols = Math.abs(r[0].col - r[1].col) + 1;
      const tiled = rows % h === 0 && cols % w === 0 && rows * cols > h * w;
      const origins = tiled ? Array.from({ length: (rows / h) * (cols / w) }, (_, i) => ({ col: at.col + (i % (cols / w)) * w, row: at.row + Math.floor(i / (cols / w)) * h })) : [at];
      const entries = origins.flatMap((o) => pasted(clip.tsv, o, { from: clip.from, language: clip.language }));
      return setting(wb, entries, tiled ? `pasted into ${named(r)}` : `pasted at ${keyOf(at)}`);
    },
  },
  {
    // Text from another app (the system clipboard): tab-separated rows
    // placed at `at`'s first cell, taken as written (its formulas are not
    // moved: they were never this sheet's).
    id: "edit.pasteText", label: "Paste text", params: ["at", "text"], journaled: true,
    run: (wb, p) => {
      const at = topLeft(rangeOf(p.at));
      return setting(wb, pasted(p.text, at), `pasted text at ${keyOf(at)}`);
    },
  },
  {
    // Paste Special, as Excel's dialog: what (all, formulas, values,
    // formats, comments), an operation with what is there (add, subtract,
    // multiply, divide: numbers worked, a formula there wrapped as
    // =(formula)+n), skip blanks (an empty copied cell leaves its target),
    // and transpose.
    id: "edit.pasteSpecial", label: "Paste Special", params: ["at", "what"], optional: ["transpose", "operation", "skipBlanks"], journaled: true,
    run: async (wb, p) => {
      const at = rangeOf(p.at)[0];
      const clip = wb.clip ?? refuse("nothing copied");
      const what = p.what.toLowerCase();
      if (!["values", "formulas", "formats", "comments", "all"].includes(what)) refuse(`${p.what} is not all, formulas, values, formats or comments`);
      const op = (p.operation ?? "none").toLowerCase();
      const ops: Record<string, (a: number, b: number) => number> = { add: (a, b) => a + b, subtract: (a, b) => a - b, multiply: (a, b) => a * b, divide: (a, b) => a / b };
      const sign: Record<string, string> = { add: "+", subtract: "-", multiply: "*", divide: "/" };
      if (op !== "none" && !ops[op]) refuse(`${p.operation} is not an operation (none, add, subtract, multiply, divide)`);
      if (op !== "none" && (what === "formats" || what === "comments")) refuse(`an operation pastes values or formulas, not ${what}`);
      const turned = flag(p.transpose), skip = flag(p.skipBlanks);
      const entries: [string, string][] = [];
      clip.inputs.forEach((row, i) => row.forEach((input, j) => {
        const to = { col: at.col + (turned ? i : j), row: at.row + (turned ? j : i) };
        if (to.col >= COLS || to.row >= ROWS) return;
        if (skip && input === "") return;
        const by = { cols: to.col - (clip.from.col + j), rows: to.row - (clip.from.row + i) };
        const key = keyOf(to);
        let pasting = what === "values" ? clip.shown[i][j] : what === "formulas" || what === "all" ? movedBy(clip.language, input, by) : null;
        if (pasting !== null && op !== "none") {
          // The operation with the copied cell's value and the target's.
          const src = clip.shown[i][j], n = src === "" ? 0 : Number(src), there = wb.sheet.input(key), v = wb.sheet.value(key);
          if (!isNaN(n) && there.startsWith("=")) pasting = `=(${there.slice(1)})${sign[op]}${n}`;
          else if (!isNaN(n) && (v === null || typeof v === "number")) pasting = String(ops[op]((v as number | null) ?? 0, n));
        }
        if (pasting !== null) entries.push([key, pasting]);
        const n = clip.notes[i][j];
        if (what === "formats" || what === "all") {
          wb.notes.set(key, {
            format: n.format ?? "", style: n.style ?? "", look: n.look,
            styleExpression: n.styleExpression ? movedBy(clip.language, `=${n.styleExpression}`, by).slice(1) : "",
          });
        }
        if (what === "comments" || what === "all") wb.notes.set(key, { comment: n.comment ?? "" });
      }));
      const how = [what, op !== "none" ? op : "", skip ? "skipping blanks" : "", turned ? "transposed" : ""].filter((x) => x).join(", ");
      return setting(wb, entries, `pasted ${how} at ${keyOf(at)}`);
    },
  },
  {
    // Flash Fill: a column's empty cells filled by the pattern its filled
    // ones show (flashfill.ts), from the columns to its left that hold
    // something on the first example's row.
    id: "data.flashFill", label: "Flash Fill", params: ["range"], journaled: true,
    run: (wb, p) => {
      const r = rangeOf(p.range), tl = topLeft(r);
      if (r[0].col !== r[1].col) refuse("Flash Fill fills one column");
      const rows = Array.from({ length: Math.abs(r[0].row - r[1].row) + 1 }, (_, i) => tl.row + i);
      const at = (col: number, row: number) => show(wb.sheet.value(keyOf({ col, row })));
      const given = rows.filter((row) => wb.sheet.input(keyOf({ col: tl.col, row })) !== "");
      if (!given.length) refuse(`type an example or two in ${rangeName(r)} first`);
      const cols: number[] = [];
      for (let c = tl.col - 1; c >= 0 && at(c, given[0]) !== ""; c--) cols.unshift(c);
      if (!cols.length) refuse(`nothing to the left of ${colName(tl.col)} to fill from`);
      const srcs = (row: number) => cols.map((c) => at(c, row));
      const program = infer(given.map((row) => [srcs(row), at(tl.col, row)] as const)) ?? refuse(`no pattern found that gives every example in ${rangeName(r)}`);
      const entries: [string, string][] = [];
      for (const row of rows) {
        if (given.includes(row) || srcs(row).every((x) => x === "")) continue;
        const out = runFill(program, srcs(row));
        if (out !== null) entries.push([keyOf({ col: tl.col, row }), out]);
      }
      return setting(wb, entries, `flash filled ${rangeName(r)} from ${given.length} example${given.length === 1 ? "" : "s"}`);
    },
  },
  {
    id: "edit.fillSeries", label: "Series", params: ["range"], journaled: true,
    run: (wb, p) => {
      const r = rangeOf(p.range), tl = topLeft(r);
      const rows = Math.abs(r[0].row - r[1].row) + 1, cols = Math.abs(r[0].col - r[1].col) + 1;
      // Down each column, or across each row when the selection is one row tall.
      const down = rows > 1;
      const lines = down ? cols : rows, length = down ? rows : cols;
      const at = (line: number, i: number) => keyOf(down ? { col: tl.col + line, row: tl.row + i } : { col: tl.col + i, row: tl.row + line });
      const entries: [string, string][] = [];
      for (let line = 0; line < lines; line++) {
        const first = [wb.sheet.input(at(line, 0))];
        if (!first[0]) continue;
        if (length > 2 && wb.sheet.input(at(line, 1))) first.push(wb.sheet.input(at(line, 1)));
        // A formula is moved by its offset from the cell it repeats, as Fill Down moves it.
        continued(first, length - first.length).forEach((input, i) => {
          const n = first.length + i, src = n % first.length;
          const off = n - src;
          entries.push([at(line, n), movedBy(wb.sheet.language, input, down ? { cols: 0, rows: off } : { cols: off, rows: 0 })]);
        });
      }
      return setting(wb, entries, `filled ${named(r)} as a series`);
    },
  },
  {
    id: "formulas.autoSum", label: "AutoSum", params: ["cell"], journaled: true,
    run: async (wb, p) => {
      const key = cellOf(p.cell), a = parseRange(key)![0];
      const isNumber = (col: number, row: number) => col >= 0 && row >= 0 && typeof wb.sheet.value(keyOf({ col, row })) === "number";
      let from: Addr | null = null;
      if (isNumber(a.col, a.row - 1)) {
        let row = a.row - 1;
        while (isNumber(a.col, row - 1)) row--;
        from = { col: a.col, row };
      } else if (isNumber(a.col - 1, a.row)) {
        let col = a.col - 1;
        while (isNumber(col - 1, a.row)) col--;
        from = { col, row: a.row };
      }
      if (!from) refuse(`no numbers above or to the left of ${key}`);
      const to = from!.col === a.col ? { col: a.col, row: a.row - 1 } : { col: a.col - 1, row: a.row };
      const span = `${keyOf(from!)}:${keyOf(to)}`;
      const formula = wb.sheet.language.name === "Excel" ? `=SUM(${span})` : `=SUM_RANGE("${span}")`;
      const changed = await wb.sheet.set(key, formula);
      return changedBy(wb, changed, `${key} = ${show(wb.sheet.value(key))}`);
    },
  },
  { id: "edit.fillDown", label: "Fill Down", params: ["range"], optional: ["source"], journaled: true, run: fill("down") },
  { id: "edit.fillRight", label: "Fill Right", params: ["range"], optional: ["source"], journaled: true, run: fill("right") },
  { id: "edit.fillUp", label: "Fill Up", params: ["range"], optional: ["source"], journaled: true, run: fill("up") },
  { id: "edit.fillLeft", label: "Fill Left", params: ["range"], optional: ["source"], journaled: true, run: fill("left") },
  {
    id: "edit.undo", label: "Undo", params: [], journaled: true, unstepped: true,
    run: async (wb) => {
      const label = wb.undoLabel;
      if (label === null) return changedBy(wb, [], "nothing to undo");
      const changed = await wb.undo();
      return changedBy(wb, changed, `undid ${label}: ${changed.length} changed`);
    },
  },
  {
    id: "edit.redo", label: "Redo", params: [], journaled: true, unstepped: true,
    run: async (wb) => {
      const label = wb.redoLabel;
      if (label === null) return changedBy(wb, [], "nothing to redo");
      const changed = await wb.redo();
      return changedBy(wb, changed, `redid ${label}: ${changed.length} changed`);
    },
  },
  {
    id: "edit.find", label: "Find", params: ["text"], optional: FIND_OPTIONS, journaled: false,
    run: async (wb, p) => {
      const keys = found(wb.sheet, p.text, findOptions(p));
      return { ok: true, message: `find ${p.text}: ${keys.join(" ") || "none"}`, changed: [], output: keys };
    },
  },
  {
    id: "edit.replace", label: "Replace All", params: ["text", "with"], optional: FIND_OPTIONS, journaled: true,
    run: async (wb, p) => {
      const entries = replaced(wb.sheet, p.text, p.with, findOptions(p));
      return { ...await setting(wb, entries, `replaced ${p.text}`), output: entries.length };
    },
  },
  {
    // The Tauri app's auto IDs: a column each inserted row is given the
    // next ID in (the prefix, then a number from `start`); no column, none.
    id: "ids.configure", label: "Auto IDs", params: [], optional: ["column", "start", "prefix"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!;
      if (!p.column) {
        page.autoId = null;
        return { ok: true, message: "no auto IDs", changed: [] };
      }
      const start = Number(p.start ?? "1");
      if (!Number.isInteger(start)) refuse(`${p.start} is not a whole number to start from`);
      page.autoId = { column: columnOf(p.column), next: start, prefix: p.prefix ?? "" };
      return { ok: true, message: `auto IDs in column ${colName(page.autoId.column)}, next ${page.autoId.prefix}${start}`, changed: [] };
    },
  },
  {
    // The row (from 1) whose auto-ID cell shows the ID; "" for none.
    id: "ids.find", label: "Find by ID", params: ["id"], journaled: false,
    run: async (wb, p) => {
      const auto = wb.page(wb.active)!.autoId ?? refuse("no auto-ID column on this sheet");
      for (let row = 0; row < ROWS; row++) {
        if (show(wb.sheet.value(keyOf({ col: auto.column, row }))) === p.id) return { ok: true, message: `${p.id} is row ${row + 1}`, changed: [], output: row + 1 };
      }
      return { ok: true, message: `no row ${p.id}`, changed: [], output: "" };
    },
  },
  { id: "rows.insert", label: "Insert Sheet Rows", params: ["at"], optional: ["count"], journaled: true, run: restructure("insert", "row") },
  { id: "rows.delete", label: "Delete Sheet Rows", params: ["at"], optional: ["count"], journaled: true, run: restructure("delete", "row") },
  { id: "columns.insert", label: "Insert Sheet Columns", params: ["at"], optional: ["count"], journaled: true, run: restructure("insert", "col") },
  { id: "columns.delete", label: "Delete Sheet Columns", params: ["at"], optional: ["count"], journaled: true, run: restructure("delete", "col") },
  {
    // A column swapped with the one beside it, as the Tauri app moves one:
    // its cells, what reads them, and its notes, width and the rest go too.
    id: "columns.move", label: "Move Column", params: ["at", "way"], journaled: true,
    run: async (wb, p) => {
      const at = columnOf(p.at), way = p.way.toLowerCase();
      if (way !== "left" && way !== "right") refuse(`${p.way} is not left or right`);
      const to = way === "left" ? at - 1 : at + 1;
      if (to < 0 || to >= COLS) refuse(`${colName(at)} is the ${way === "left" ? "first" : "last"} column: it cannot move ${way}`);
      const entries = await reshaped(wb, "col", (moves) => swapped("col", at, to, moves));
      return setting(wb, entries, `moved column ${colName(at)} ${way}`);
    },
  },
  {
    // A multilevel sort, as Excel's Sort dialog: keys="B desc, C asc" (or
    // "C fill=#ffff00": that cell colour on top; "C font=#ff0000": that
    // text colour), each deciding where the ones before tie; headers=1
    // keeps the first row where it is.
    id: "data.sort", label: "Custom Sort", params: ["range", "keys"], optional: ["headers"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), left = Math.min(r[0].col, r[1].col), right = Math.max(r[0].col, r[1].col);
      const keys = p.keys.split(",").map((k) => k.trim()).filter((k) => k).map((k): SortKey => {
        const m = /^\$?([A-Za-z]{1,3})(?:\s+(asc|ascending|desc|descending|(fill|font)\s*=\s*(\S+)))?$/i.exec(k) ?? refuse(`${k} is not a sort key (a column, then asc, desc, fill=colour or font=colour)`);
        const col = columnOf(m[1]);
        if (col < left || col > right) refuse(`${m[1]} is not a column of ${named(r)}`);
        if (m[3]) {
          const c = colour(m[4]) ?? refuse(`${m[4]} is not a colour`);
          return { col: col - left, ascending: true, colour: { of: m[3].toLowerCase() as "fill" | "font", is: c } };
        }
        return { col: col - left, ascending: !/^desc/i.test(m[2] ?? "") };
      });
      if (!keys.length) refuse("a sort needs a key");
      if (keys.length > 64) refuse("a sort takes 64 keys at most, as Excel's");
      const look = (k: string, of: "fill" | "font") => (of === "fill" ? wb.notes.get(k).look?.fill : wb.notes.get(k).look?.color) ?? "";
      const entries = rearranging(wb, r, sortedRowsBy(wb.sheet, r[0], r[1], keys, look, flag(p.headers)));
      const said = keys.map((k) => `${colName(left + k.col)} ${k.colour ? `${k.colour.of} ${k.colour.is} on top` : k.ascending ? "ascending" : "descending"}`).join(", then ");
      return setting(wb, entries, `sorted ${named(r)} by ${said}`);
    },
  },
  {
    // Advanced Filter: the range's records (its first row their fields'
    // names) that a criteria range picks (as the D* functions read one: a
    // row's conditions all hold, any row will do), shown in place (the rest
    // hidden, as a filter's) or copied to copyTo= with their header row,
    // unique=1 each record once.
    id: "data.advancedFilter", label: "Advanced Filter", params: ["range", "criteria"], optional: ["copyTo", "unique"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), c = rangeOf(p.criteria), page = wb.page(wb.active)!;
      const rowOf = (row: number, [a, b]: [Addr, Addr]) => Array.from({ length: Math.abs(a.col - b.col) + 1 }, (_, i) => wb.sheet.value(keyOf({ col: Math.min(a.col, b.col) + i, row })));
      const top = Math.min(r[0].row, r[1].row), bottom = Math.max(r[0].row, r[1].row), left = Math.min(r[0].col, r[1].col), width = Math.abs(r[0].col - r[1].col) + 1;
      const cTop = Math.min(c[0].row, c[1].row), cBottom = Math.max(c[0].row, c[1].row);
      const test = recordTest(rowOf(top, r), Array.from({ length: cBottom - cTop + 1 }, (_, i) => rowOf(cTop + i, c)));
      if (typeof test === "string") refuse(test);
      const seen = new Set<string>(), picked: number[] = [];
      for (let row = top + 1; row <= bottom; row++) {
        const rec = rowOf(row, r);
        if (!(test as (rec: readonly Value[]) => boolean)(rec)) continue;
        const said = JSON.stringify(rec.map(show));
        if (flag(p.unique) && seen.has(said)) continue;
        seen.add(said);
        picked.push(row);
      }
      if (p.copyTo === undefined) {
        const keep = new Set(picked), hidden = new Set<number>();
        for (let row = top + 1; row <= bottom; row++) if (!keep.has(row)) hidden.add(row);
        page.filter = null;
        page.filtered = hidden;
        await wb.hiddenChanged();
        return changedBy(wb, wb.sheet.keys(), `filtered ${named(r)} by ${named(c)}: ${picked.length} record${picked.length === 1 ? "" : "s"} shown, ${hidden.size} hidden`);
      }
      const to = cellAddr(cellOf(p.copyTo))!;
      if (to.row + picked.length >= ROWS || to.col + width > COLS) refuse(`${picked.length} records do not fit at ${p.copyTo}`);
      // The values, as they show, under the header row.
      const entries: [string, string][] = [];
      [top, ...picked].forEach((row, i) => {
        for (let j = 0; j < width; j++) {
          const from = keyOf({ col: left + j, row }), v = wb.sheet.value(from);
          entries.push([keyOf({ col: to.col + j, row: to.row + i }), v === null ? "" : typeof v === "string" && (NUMERIC_TEXT.test(v) || v.startsWith("=") || v.startsWith("'")) ? `'${v}` : show(v)]);
        }
      });
      return setting(wb, entries, `copied ${picked.length} record${picked.length === 1 ? "" : "s"} of ${named(r)} to ${keyOf(to)}`);
    },
  },
  {
    // Group: the rows (columns) of a range a level deeper in the outline,
    // up to 8; axis=columns for columns ("B:D" is columns of itself).
    id: "outline.group", label: "Group", params: ["range"], optional: ["axis"], journaled: true,
    run: (wb, p) => outlined(wb, p, 1),
  },
  { id: "outline.ungroup", label: "Ungroup", params: ["range"], optional: ["axis"], journaled: true, run: (wb, p) => outlined(wb, p, -1) },
  {
    // The outline's buttons: show levels 1 to n, the deeper rows (columns)
    // hidden; a level past the deepest shows them all.
    id: "outline.level", label: "Show Outline Level", params: ["level"], optional: ["axis"], journaled: true,
    run: async (wb, p) => {
      const rows = axisOf(p), page = wb.page(wb.active)!, levels = rows ? page.rowLevels : page.colLevels, n = Number(p.level);
      if (!Number.isInteger(n) || n < 1 || n > MOST_LEVELS) refuse(`${p.level} is not an outline level (1 to ${MOST_LEVELS})`);
      const set = rows ? page.hiddenRows : page.hiddenCols, hide = new Set(hiddenAtLevel(levels, n));
      for (const i of levels.keys()) if (hide.has(i)) set.add(i); else set.delete(i);
      if (rows) await wb.hiddenChanged();
      return changedBy(wb, wb.sheet.keys(), `outline level ${n}: ${hide.size} ${rows ? "row" : "column"}${hide.size === 1 ? "" : "s"} hidden`);
    },
  },
  {
    // Hide Detail and Show Detail: the group the range's first row (column)
    // is in, at its own level.
    id: "outline.hideDetail", label: "Hide Detail", params: ["range"], optional: ["axis"], journaled: true, run: (wb, p) => detail(wb, p, true),
  },
  { id: "outline.showDetail", label: "Show Detail", params: ["range"], optional: ["axis"], journaled: true, run: (wb, p) => detail(wb, p, false) },
  {
    // Clear Outline: no levels left, what they hid shown.
    id: "outline.clear", label: "Clear Outline", params: [], optional: ["axis"], journaled: true,
    run: async (wb, p) => {
      const rows = axisOf(p), page = wb.page(wb.active)!, levels = rows ? page.rowLevels : page.colLevels, set = rows ? page.hiddenRows : page.hiddenCols;
      for (const i of levels.keys()) set.delete(i);
      if (rows) page.rowLevels = new Map(); else page.colLevels = new Map();
      if (rows) await wb.hiddenChanged();
      return changedBy(wb, wb.sheet.keys(), `${rows ? "rows'" : "columns'"} outline cleared`);
    },
  },
  {
    // Excel's Subtotal: a sorted list (its first row the header) gains a
    // row under each run of one value in by=, "<value> Total" and
    // SUBTOTAL(function, the run) in each of columns=, then a Grand Total;
    // the outline in three levels (show level 2: the totals only).
    id: "data.subtotal", label: "Subtotal", params: ["range", "by", "columns"], optional: ["function"], journaled: true,
    run: async (wb, p) => {
      if (wb.sheet.language.name !== "Excel") refuse("Subtotal writes SUBTOTAL formulas: Excel formulas only");
      const r = rangeOf(p.range), page = wb.page(wb.active)!;
      const left = Math.min(r[0].col, r[1].col), right = Math.max(r[0].col, r[1].col), top = Math.min(r[0].row, r[1].row), bottom = Math.max(r[0].row, r[1].row);
      const by = columnOf(p.by), fn = (p.function ?? "sum").trim().toLowerCase();
      if (by < left || by > right) refuse(`${p.by} is not a column of ${named(r)}`);
      if (!(fn in SUBTOTAL_FUNCTIONS)) refuse(`${p.function} is not a subtotal function (${Object.keys(SUBTOTAL_FUNCTIONS).join(", ")})`);
      const cols = p.columns.split(",").map((c) => c.trim()).filter((c) => c).map(columnOf);
      for (const c of cols) if (c < left || c > right) refuse(`${colName(c)} is not a column of ${named(r)}`);
      if (top >= bottom) refuse(`${named(r)} has no rows under its header`);
      const groupsOf = runs(top + 1, bottom, (row) => show(wb.sheet.value(keyOf({ col: by, row }))));
      if (bottom + groupsOf.length + 1 >= ROWS) refuse(`${named(r)}'s subtotals do not fit on the sheet`);
      // From the last run up, so the rows above stay where they are.
      for (const [, last] of [...groupsOf].reverse()) await wb.sheet.setMany(await reshaped(wb, "row", (moves) => inserted("row", last + 1, 1, moves)));
      const entries: [string, string][] = [];
      const levels = new Map(page.rowLevels);
      let shift = 0;
      for (const [first, last, value] of groupsOf) {
        const a = first + shift, b = last + shift, at = b + 1;
        entries.push([keyOf({ col: by, row: at }), `${value} Total`]);
        for (const c of cols) entries.push([keyOf({ col: c, row: at }), `=SUBTOTAL(${SUBTOTAL_FUNCTIONS[fn]},${keyOf({ col: c, row: a })}:${keyOf({ col: c, row: b })})`]);
        for (let row = a; row <= b; row++) levels.set(row, 2);
        levels.set(at, 1);
        shift++;
      }
      const end = bottom + shift, grand = end + 1;
      entries.push([keyOf({ col: by, row: grand }), "Grand Total"]);
      for (const c of cols) entries.push([keyOf({ col: c, row: grand }), `=SUBTOTAL(${SUBTOTAL_FUNCTIONS[fn]},${keyOf({ col: c, row: top + 1 })}:${keyOf({ col: c, row: end })})`]);
      page.rowLevels = levels;
      return setting(wb, entries, `subtotals of ${named(r)} at each change in ${colName(by)}: ${groupsOf.length} groups, ${fn} of ${cols.map(colName).join(", ")}`);
    },
  },
  {
    // Excel's Consolidate: sources= ranges of any sheets ("Jan!A1:C4,
    // Feb!A1:C4"), their numbers put together at to= by function= (sum by
    // default), by position or by the labels in their top row and left
    // column (labels=top, left or both); links=1 writes each number as a
    // formula over the cells it came from, so it follows them.
    id: "data.consolidate", label: "Consolidate", params: ["to", "sources"], optional: ["function", "labels", "links"], journaled: true,
    run: async (wb, p) => {
      const fn = (p.function ?? "sum").trim().toLowerCase();
      if (!(fn in CONSOLIDATE_FUNCTIONS)) refuse(`${p.function} is not a consolidation function (${Object.keys(CONSOLIDATE_FUNCTIONS).join(", ")})`);
      const labels = (p.labels ?? "").trim().toLowerCase();
      if (!["", "none", "top", "left", "both", "top,left", "left,top"].includes(labels.replace(/\s+/g, ""))) refuse(`${p.labels} is not top, left or both`);
      const top = /top|both/.test(labels), left = /left|both/.test(labels);
      const parts = p.sources.split(/,(?=(?:[^']*'[^']*')*[^']*$)/).map((x) => x.trim()).filter((x) => x);
      if (!parts.length) refuse("Consolidate needs a source range");
      const sources = parts.map((part) => {
        const { page, range: [a, b] } = sheetRangeOf(wb, part);
        const r0 = Math.min(a.row, b.row), c0 = Math.min(a.col, b.col);
        const values = Array.from({ length: Math.abs(a.row - b.row) + 1 }, (_, r) => Array.from({ length: Math.abs(a.col - b.col) + 1 }, (_, c) => page.sheet.value(keyOf({ col: c0 + c, row: r0 + r }))));
        const prefix = page === wb.page(wb.active) ? "" : sheetPrefix(page.name);
        return { values, cell: (r: number, c: number) => `${prefix}${keyOf({ col: c0 + c, row: r0 + r })}` };
      });
      const out = consolidate(sources, fn, top, left), to = cellAddr(cellOf(p.to))!;
      if (to.row + out.rows.length > ROWS || to.col + Math.max(0, ...out.rows.map((r) => r.length)) > COLS) refuse(`the consolidation does not fit at ${p.to}`);
      const entries: [string, string][] = [];
      out.rows.forEach((row, r) => row.forEach((v, c) => {
        const key = keyOf({ col: to.col + c, row: to.row + r }), from = out.from[r][c];
        const linked = flag(p.links) && from.length && typeof v !== "string";
        entries.push([key, linked ? (wb.sheet.language.name === "Excel" ? `=${CONSOLIDATE_FUNCTIONS[fn]}(${from.join(",")})` : `=${v}`) : v === null ? "" : typeof v === "string" ? v : String(v)]);
      }));
      return setting(wb, entries, `consolidated ${parts.length} range${parts.length === 1 ? "" : "s"} at ${keyOf(to)} by ${fn}${top || left ? `, by the labels in the ${top && left ? "top row and left column" : top ? "top row" : "left column"}` : ""}${flag(p.links) ? ", linked" : ""}`);
    },
  },
  { id: "data.sortAscending", label: "Sort A to Z", params: ["range", "by"], journaled: true, run: sort(true) },
  { id: "data.sortDescending", label: "Sort Z to A", params: ["range", "by"], journaled: true, run: sort(false) },
  {
    // A name for a range, or for a formula (a named LAMBDA:
    // formula="=LAMBDA(x, x*2)", then =DOUBLE(4) in any formula). Its scope
    // is the workbook, as in Excel (a range of any sheet: range=Costs!B1:B3),
    // or scope=sheet: the active sheet's own, which hides a workbook name of
    // the same spelling there.
    id: "names.define", label: "Define Name", params: ["name"], optional: ["range", "formula", "scope"], journaled: true,
    run: async (wb, p) => {
      const n = nameOf(p.name), own = scopeOf(p.scope) === "sheet";
      if ((p.range === undefined) === (p.formula === undefined)) refuse("names.define needs a range or a formula");
      if (p.formula !== undefined) {
        const f = formulaOf(wb, n, p.formula);
        if (own) await wb.sheet.define(n, f); else await wb.defineName(n, { formula: f });
        if (own) await wb.ownNamesChanged();
      } else {
        const { page, range: r } = sheetRangeOf(wb, p.range!);
        if (own && page !== wb.page(wb.active)) refuse(`a sheet's own name is for a range of it: ${p.range}`);
        if (own) await wb.sheet.define(n, r); else await wb.defineName(n, { page, range: r });
        if (own) await wb.ownNamesChanged();
      }
      const listed = (own ? wb.sheet.names() : wb.names()).find(([x]) => x === n)![1];
      return changedBy(wb, wb.sheet.keys(), `${own ? `${sheetPrefix(wb.active)}${n}` : n} is ${listed}`);
    },
  },
  {
    // The active sheet's own name of that spelling, or else the workbook's
    // (scope says which).
    id: "names.delete", label: "Delete Name", params: ["name"], optional: ["scope"], journaled: true,
    run: async (wb, p) => {
      const n = p.name.toUpperCase();
      const mine = wb.sheet.name(n) !== null || wb.sheet.formula(n) !== null;
      const own = p.scope === undefined ? mine : scopeOf(p.scope) === "sheet";
      if (own && !mine) refuse(`no name ${n} of ${wb.active}'s own`);
      if (!own && !wb.bookName(n)) refuse(`no name ${n}`);
      if (own) {
        await wb.sheet.define(n, null);
        await wb.ownNamesChanged();
      } else await wb.defineName(n, null);
      return changedBy(wb, wb.sheet.keys(), `${own ? `${sheetPrefix(wb.active)}${n}` : n} deleted`);
    },
  },
  {
    // The workbook's names, then each sheet's own (Costs!RATE).
    id: "names.list", label: "Name Manager", params: [], journaled: false,
    run: async (wb) => {
      const names = [...wb.names().map(([n, r]) => `${n} ${r}`),
        ...wb.pages().flatMap((page) => page.sheet.names().map(([n, r]) => `${sheetPrefix(page.name)}${n} ${r}`))];
      return { ok: true, message: names.length ? names.join(", ") : "no names", changed: [], output: names };
    },
  },
  { id: "formulas.tracePrecedents", label: "Trace Precedents", params: ["cell"], journaled: false, run: trace("precedents") },
  { id: "formulas.traceDependents", label: "Trace Dependents", params: ["cell"], journaled: false, run: trace("dependents") },
  {
    id: "formulas.removeArrows", label: "Remove Arrows", params: [], journaled: false,
    run: async (wb) => {
      wb.clearTraces();
      return changedBy(wb, wb.sheet.keys(), "arrows removed");
    },
  },
  {
    id: "formulas.showFormulas", label: "Show Formulas", params: [], optional: ["off"], journaled: false,
    run: async (wb, p) => {
      wb.showFormulas = !flag(p.off);
      return changedBy(wb, wb.sheet.keys(), wb.showFormulas ? "showing formulas" : "showing values");
    },
  },
  {
    id: "formulas.evaluate", label: "Evaluate Formula", params: ["cell"], journaled: false,
    run: async (wb, p) => {
      const key = cellOf(p.cell), input = wb.sheet.input(key);
      if (!input.startsWith("=")) refuse(`${key} is not a formula`);
      if (wb.sheet.language.name !== "Excel") refuse("Evaluate Formula steps through Excel formulas");
      let steps: string[] = [];
      try {
        steps = evaluationSteps(input.slice(1), { get: (k) => wb.sheet.value(k) }, (n) => wb.sheet.named(n), (n) => wb.sheet.formulaOf(n));
      } catch (e) {
        if (!(e instanceof FormulaError)) throw e;
        refuse(`${key}: parse error at ${e.at}: ${e.message}`);
      }
      return { ok: true, message: `${key}: ${steps.join(" → ")}`, changed: [], output: steps };
    },
  },
  {
    id: "formulas.watch", label: "Add Watch", params: ["cell"], journaled: false,
    run: async (wb, p) => {
      const [page, cell] = watchedOf(wb, p.cell);
      return { ok: true, message: `${wb.watch(page, cell) ? "watching" : "already watching"} ${page.name}!${cell}`, changed: [] };
    },
  },
  {
    id: "formulas.unwatch", label: "Delete Watch", params: ["cell"], journaled: false,
    run: async (wb, p) => {
      const [page, cell] = watchedOf(wb, p.cell);
      if (!wb.unwatch(page, cell)) refuse(`not watching ${page.name}!${cell}`);
      return { ok: true, message: `stopped watching ${page.name}!${cell}`, changed: [] };
    },
  },
  {
    id: "format.highlight", label: "Highlight Cells Rules", params: ["range", "op", "value"], optional: ["stop", "value2", "style"], journaled: true,
    run: (wb, p) => {
      const op = COMPARISONS[p.op.toLowerCase()] ?? refuse(`${p.op} is not a comparison (greater, less, between, notBetween, equal, notEqual, greaterOrEqual, lessOrEqual, contains, notContains, beginsWith, endsWith)`);
      if ((op === "between" || op === "notBetween") && p.value2 === undefined) refuse(`${p.op} needs value2`);
      return addRule(wb, { kind: "cellValue", range: rangeOf(p.range), op, value: p.value, ...(op === "between" || op === "notBetween" ? { value2: p.value2 } : {}), style: styleParam(p) }, p);
    },
  },
  {
    id: "format.topBottom", label: "Top/Bottom Rules", params: ["range"], optional: ["stop", "n", "bottom", "percent", "style"], journaled: true,
    run: (wb, p) => addRule(wb, { kind: "top", range: rangeOf(p.range), n: countOf(p.n ?? "10"), bottom: flag(p.bottom), percent: flag(p.percent), style: styleParam(p) }, p),
  },
  {
    id: "format.aboveAverage", label: "Above Average", params: ["range"], optional: ["stop", "below", "style"], journaled: true,
    run: (wb, p) => addRule(wb, { kind: "average", range: rangeOf(p.range), below: flag(p.below), style: styleParam(p) }, p),
  },
  {
    id: "format.duplicates", label: "Duplicate Values", params: ["range"], optional: ["stop", "unique", "style"], journaled: true,
    run: (wb, p) => addRule(wb, { kind: "duplicates", range: rangeOf(p.range), unique: flag(p.unique), style: styleParam(p) }, p),
  },
  {
    id: "format.colorScale", label: "Color Scales", params: ["range"], optional: ["colors"], journaled: true,
    run: (wb, p) => {
      const colors = (p.colors ?? "#f8696b,#ffeb84,#63be7b").split(",").map((c) => c.trim().toLowerCase());
      if (colors.length < 2 || colors.length > 3 || colors.some((c) => !/^#[0-9a-f]{6}$/.test(c))) refuse("a colour scale takes 2 or 3 colours (#rrggbb)");
      return addRule(wb, { kind: "colorScale", range: rangeOf(p.range), colors });
    },
  },
  {
    // An icon set: its icons, the two thresholds (high, middle: percent of
    // the way from lowest to highest, 67 and 33, by Excel's default) and
    // how they read (by= percent, percentile or number), reversed, icons
    // only (showValue=0).
    id: "format.iconSet", label: "Icon Sets", params: ["range"], optional: ["icons", "thresholds", "by", "reverse", "showValue", "stop"], journaled: true,
    run: (wb, p) => {
      const icons = (p.icons ?? "arrows").toLowerCase();
      if (!(icons in ICONS)) refuse(`${p.icons} is not an icon set (${Object.keys(ICONS).join(", ")})`);
      const by = (p.by ?? "percent").trim().toLowerCase() as Threshold;
      if (!THRESHOLDS.includes(by)) refuse(`${p.by} is not a threshold type (${THRESHOLDS.join(", ")})`);
      const ts = p.thresholds === undefined ? undefined : p.thresholds.split(",").map((t) => Number(t.trim()));
      if (ts && (ts.length !== 2 || ts.some((t) => !isFinite(t)) || ts[0] < ts[1])) refuse(`${p.thresholds} is not two thresholds, the higher first (67,33)`);
      return addRule(wb, { kind: "iconSet", range: rangeOf(p.range), icons: icons as IconSet, ...(ts ? { thresholds: [ts[0], ts[1]] as [number, number] } : {}),
        ...(by !== "percent" ? { by } : {}), ...(flag(p.reverse) ? { reverse: true } : {}), ...(p.showValue !== undefined && !flag(p.showValue) ? { iconOnly: true } : {}) }, p);
    },
  },
  {
    // A formula's rule (Use a formula to determine which cells to format):
    // written for the range's top-left, moved for each cell as a copy is,
    // in the sheet's language; the cells it is true (or not 0) for.
    id: "format.formulaRule", label: "Formula Rule", params: ["range", "formula"], optional: ["style", "stop"], journaled: true,
    run: (wb, p) => {
      const formula = p.formula.trim().replace(/^=/, "");
      if (!formula) refuse("a formula rule needs a formula");
      const compiled = wb.sheet.language.compile(formula);
      if (compiled.kind === "broken") refuse(`${p.formula} is not a formula: ${compiled.error.why}`);
      return addRule(wb, { kind: "formula", range: rangeOf(p.range), formula, style: styleParam(p) }, p);
    },
  },
  {
    // Format only cells that are blank (or, not=1, that are not); that
    // hold an error (or not).
    id: "format.blanks", label: "Blanks Rule", params: ["range"], optional: ["not", "style", "stop"], journaled: true,
    run: (wb, p) => addRule(wb, { kind: "blanks", range: rangeOf(p.range), not: flag(p.not), style: styleParam(p) }, p),
  },
  {
    id: "format.errors", label: "Errors Rule", params: ["range"], optional: ["not", "style", "stop"], journaled: true,
    run: (wb, p) => addRule(wb, { kind: "errors", range: rangeOf(p.range), not: flag(p.not), style: styleParam(p) }, p),
  },
  {
    // A Date Occurring: yesterday, today, tomorrow, last7days, lastWeek,
    // thisWeek, nextWeek, lastMonth, thisMonth, nextMonth.
    id: "format.datePeriod", label: "A Date Occurring", params: ["range", "period"], optional: ["style", "stop"], journaled: true,
    run: (wb, p) => {
      const period = PERIODS.find((x) => x.toLowerCase() === p.period.trim().toLowerCase()) ?? refuse(`${p.period} is not a period (${PERIODS.join(", ")})`);
      return addRule(wb, { kind: "date", range: rangeOf(p.range), period, style: styleParam(p) }, p);
    },
  },
  {
    // Manage Rules' Move Up and Move Down: rule `index` (from 1) to place
    // `to`, the first winning.
    id: "format.moveRule", label: "Move Rule", params: ["index", "to"], journaled: true,
    run: async (wb, p) => {
      const rules = wb.page(wb.active)!.rules, i = Number(p.index) - 1, j = Number(p.to) - 1;
      if (!Number.isInteger(i) || !rules[i]) refuse(`no rule ${p.index}`);
      if (!Number.isInteger(j) || j < 0 || j >= rules.length) refuse(`${p.to} is not a place in the list (1 to ${rules.length})`);
      const [r] = rules.splice(i, 1);
      rules.splice(j, 0, r);
      return changedBy(wb, [...range(r.range[0], r.range[1])], `rule ${i + 1} is now rule ${j + 1}: ${describeCf(r)}`);
    },
  },
  {
    // Stop If True on rule `index` (off=1 takes it off).
    id: "format.stopIfTrue", label: "Stop If True", params: ["index"], optional: ["off"], journaled: true,
    run: async (wb, p) => {
      const rules = wb.page(wb.active)!.rules, i = Number(p.index) - 1;
      if (!Number.isInteger(i) || !rules[i]) refuse(`no rule ${p.index}`);
      const { stop: _was, ...rest } = rules[i];
      rules[i] = (flag(p.off) ? rest : { ...rest, stop: true }) as CfRule;
      return changedBy(wb, [...range(rules[i].range[0], rules[i].range[1])], describeCf(rules[i]));
    },
  },
  {
    id: "format.rules", label: "Manage Rules", params: [], journaled: false,
    run: async (wb) => {
      const listed = wb.page(wb.active)!.rules.map((r, i) => `${i + 1}. ${describeCf(r)}`);
      return { ok: true, message: listed.length ? listed.join("; ") : "no rules", changed: [], output: listed };
    },
  },
  {
    id: "format.deleteRule", label: "Delete Rule", params: ["index"], journaled: true,
    run: async (wb, p) => {
      const rules = wb.page(wb.active)!.rules, i = Number(p.index) - 1;
      if (!Number.isInteger(i) || !rules[i]) refuse(`no rule ${p.index}`);
      const [gone] = rules.splice(i, 1);
      return changedBy(wb, [...range(gone.range[0], gone.range[1])], `deleted ${describeCf(gone)}`);
    },
  },
  {
    id: "format.clearRules", label: "Clear Rules", params: ["range"], journaled: true,
    run: async (wb, p) => {
      const [a, b] = rangeOf(p.range), page = wb.page(wb.active)!;
      const over = (r: CfRule) => {
        const [c, d] = r.range;
        return Math.max(Math.min(a.col, b.col), Math.min(c.col, d.col)) <= Math.min(Math.max(a.col, b.col), Math.max(c.col, d.col))
          && Math.max(Math.min(a.row, b.row), Math.min(c.row, d.row)) <= Math.min(Math.max(a.row, b.row), Math.max(c.row, d.row));
      };
      const n = page.rules.filter(over).length;
      page.rules = page.rules.filter((r) => !over(r));
      return changedBy(wb, wb.sheet.keys(), `cleared ${n} rule${n === 1 ? "" : "s"}`);
    },
  },
  {
    id: "format.font", label: "Font", params: ["range"], optional: ["bold", "italic", "underline", "strike", "size"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range);
      const size = p.size === undefined ? undefined : Number(p.size) > 0 ? Number(p.size) : refuse(`${p.size} is not a size`);
      const parts = (["bold", "italic", "underline", "strike"] as const).filter((k) => p[k] !== undefined);
      const keys = relook(wb, r, (look) => ({ ...look, ...Object.fromEntries(parts.map((k) => [k, flag(p[k])])), ...(size ? { size } : {}) }));
      const said = [...parts.map((k) => (flag(p[k]) ? k : `not ${k}`)), ...(size ? [`size ${size}`] : [])].join(", ");
      return changedBy(wb, keys, `${rangeName(r)}: ${said || "no change"}`);
    },
  },
  {
    id: "format.border", label: "Borders", params: ["range", "edge"], optional: ["color", "style"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), edge = p.edge.toLowerCase(), color = colour(p.color ?? "#000000") ?? refuse(`${p.color} is not a colour (a name such as red, or #rrggbb)`);
      const style = (p.style ?? "thin").trim().toLowerCase();
      if (!EDGES.includes(edge)) refuse(`${p.edge} is not an edge (${EDGES.join(", ")})`);
      if (!LINE_STYLES.includes(style)) refuse(`${p.style} is not a line style (${LINE_STYLES.join(", ")})`);
      const line = style === "none" ? undefined : style === "thin" ? color : `${style} ${color}`;
      const [a, b] = r, top = Math.min(a.row, b.row), bottom = Math.max(a.row, b.row), left = Math.min(a.col, b.col), right = Math.max(a.col, b.col);
      const keys = relook(wb, r, (look, k) => {
        const at = parseRange(k)![0];
        if (edge === "none") return { ...look, borders: undefined };
        const outer = { top: at.row === top, bottom: at.row === bottom, left: at.col === left, right: at.col === right };
        const on = (e: "top" | "right" | "bottom" | "left") => {
          const across = e === "top" || e === "bottom";
          switch (edge) {
            case "all": return true;
            case "outline": return outer[e];
            case "inside": return !outer[e];
            case "horizontal": return across && !outer[e];
            case "vertical": return !across && !outer[e];
            default: return edge === e && outer[e];
          }
        };
        const borders: Record<string, string | undefined> = { ...look.borders };
        for (const e of ["top", "right", "bottom", "left"] as const) if (on(e)) borders[e] = line;
        for (const e of Object.keys(borders)) if (borders[e] === undefined) delete borders[e];
        return { ...look, borders: Object.keys(borders).length ? borders : undefined };
      });
      const what = edge === "none" ? "no borders" : `${edge === "all" || edge === "outline" || edge === "inside" ? `${edge} borders` : `${edge} border`}${style === "thin" ? "" : `, ${style}`}`;
      return changedBy(wb, keys, `${rangeName(r)}: ${what}`);
    },
  },
  {
    // Format Painter: a cell's look, number format and style (not its
    // value, comment or rule) onto a range; a range's, tiled over it.
    id: "format.painter", label: "Format Painter", params: ["from", "to"], journaled: true,
    run: async (wb, p) => {
      const f = rangeOf(p.from), t = rangeOf(p.to), ftl = topLeft(f), ttl = topLeft(t);
      const fr = Math.abs(f[0].row - f[1].row) + 1, fc = Math.abs(f[0].col - f[1].col) + 1;
      const keys = [...range(t[0], t[1])];
      const src = new Map(keys.map((k) => {
        const a = cellAddr(k)!, from = keyOf({ col: ftl.col + ((a.col - ttl.col) % fc), row: ftl.row + ((a.row - ttl.row) % fr) });
        return [k, wb.notes.get(from)];
      }));
      for (const [k, n] of src) wb.notes.set(k, { format: n.format ?? "", style: n.style ?? "", look: n.look, styleExpression: n.styleExpression ?? "" });
      return changedBy(wb, keys, `painted ${rangeName(f)}'s format over ${rangeName(t)}`);
    },
  },
  {
    // Align Left, Center and Right, as the toolbar's; general: by the value
    // (numbers to the right), as Excel's default.
    id: "format.align", label: "Alignment", params: ["range", "align"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range);
      const align = ({ left: "start", center: "center", centre: "center", right: "end", general: undefined } as Record<string, "start" | "center" | "end" | undefined>);
      if (!(p.align.toLowerCase() in align)) refuse(`${p.align} is not an alignment (left, center, right, general)`);
      return changedBy(wb, relook(wb, r, (look) => ({ ...look, align: align[p.align.toLowerCase()] })), `${rangeName(r)}: aligned ${p.align.toLowerCase()}`);
    },
  },
  {
    // Fill Color and Font Color: a colour's name or #rrggbb; none takes it off.
    id: "format.fill", label: "Fill Color", params: ["range", "color"], journaled: true,
    run: async (wb, p) => coloured(wb, p, "fill"),
  },
  {
    id: "format.color", label: "Font Color", params: ["range", "color"], journaled: true,
    run: async (wb, p) => coloured(wb, p, "color"),
  },
  {
    // Number format over the selection, as the toolbar's Comma, Currency
    // and Percent Style: number, currency, percent, an Excel pattern, or
    // general (none).
    id: "format.number", label: "Number Format", params: ["range", "format"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), f = p.format.trim();
      const pattern = ({ general: "", number: "#,##0.00", comma: "#,##0.00", currency: "$#,##0.00", percent: "0%" } as Record<string, string>)[f.toLowerCase()] ?? CATEGORIES[f.toLowerCase()] ?? f;
      if (pattern && !/[#0?@]|general/i.test(pattern) && !/^[^A-Za-z]*(?:(?:[ymdhs]+|AM\/PM)[^A-Za-z]*)+$/i.test(pattern)) refuse(`${p.format} is not a number format (number, currency, percent, accounting, fraction, scientific, date, time, text, general, or a pattern such as 0.00)`);
      return changedBy(wb, relook(wb, r, (look) => ({ ...look, number: pattern || undefined })), `${rangeName(r)}: ${pattern || "general"}`);
    },
  },
  {
    // A value in a number format, as a cell in it would show it, and as
    // TEXT() gives it: a category's name (accounting, fraction ...) or a
    // code. Text that reads as a number is the number.
    id: "format.preview", label: "Format a value", params: ["value", "format"], journaled: false,
    run: async (_wb, p) => {
      const code = CATEGORIES[p.format.trim().toLowerCase()] ?? p.format;
      const v: Value = NUMERIC_TEXT.test(p.value) || /^\s*[-+]?\d*\.?\d+[eE][-+]?\d+\s*$/.test(p.value) ? Number(p.value) : p.value;
      const out = numberFormat(v, code);
      const text = out?.text ?? show(v);
      return { ok: true, message: out?.color ? `${text} (${out.color})` : text, changed: [], output: text };
    },
  },
  {
    id: "format.wrap", label: "Wrap Text", params: ["range"], optional: ["off"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), on = !flag(p.off);
      return changedBy(wb, relook(wb, r, (look) => ({ ...look, wrap: on || undefined })), `${rangeName(r)}: ${on ? "wrapped" : "not wrapped"}`);
    },
  },
  {
    // Data bars: a colour, a solid fill (fill=solid) or Excel's gradient
    // (the grid draws both solid), the bar only (showValue=0), and the
    // shortest and longest bar's numbers (min=, max=; automatic otherwise).
    id: "format.dataBar", label: "Data Bars", params: ["range"], optional: ["color", "fill", "showValue", "min", "max", "stop"], journaled: true,
    run: (wb, p) => {
      const color = (p.color ?? "#638ec6").toLowerCase();
      if (!/^#[0-9a-f]{6}$/.test(color)) refuse(`${p.color} is not a colour (#rrggbb)`);
      const fill = (p.fill ?? "gradient").trim().toLowerCase();
      if (fill !== "gradient" && fill !== "solid") refuse(`${p.fill} is not a fill (gradient, solid)`);
      const end = (v: string | undefined, what: string) => (v === undefined || /^\s*(auto|automatic)?\s*$/i.test(v) ? undefined : isFinite(Number(v)) ? Number(v) : refuse(`${v} is not a number for the ${what}`));
      const min = end(p.min, "shortest bar"), max = end(p.max, "longest bar");
      if (min !== undefined && max !== undefined && min >= max) refuse(`the shortest bar's number (${min}) must be below the longest's (${max})`);
      return addRule(wb, { kind: "dataBar", range: rangeOf(p.range), color, ...(fill === "solid" ? { solid: true } : {}), ...(p.showValue !== undefined && !flag(p.showValue) ? { barOnly: true } : {}),
        ...(min !== undefined ? { min } : {}), ...(max !== undefined ? { max } : {}) }, p);
    },
  },
  { id: "rows.hide", label: "Hide Rows", params: ["range"], journaled: true, run: hiding(true, true) },
  { id: "rows.unhide", label: "Unhide Rows", params: ["range"], journaled: true, run: hiding(true, false) },
  { id: "columns.hide", label: "Hide Columns", params: ["range"], journaled: true, run: hiding(false, true) },
  { id: "columns.unhide", label: "Unhide Columns", params: ["range"], journaled: true, run: hiding(false, false) },
  {
    id: "data.filter", label: "AutoFilter", params: ["range", "column", "criterion"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), col = columnOf(p.column);
      if (col < Math.min(r[0].col, r[1].col) || col > Math.max(r[0].col, r[1].col)) refuse(`${p.column} is not a column of ${named(r)}`);
      const hidden = filtering(wb.page(wb.active)!, { range: r, column: col, criterion: p.criterion });
      await wb.hiddenChanged();
      return changedBy(wb, wb.sheet.keys(), `filtered ${named(r)} by ${colName(col)} ${p.criterion}: ${hidden.size} row${hidden.size === 1 ? "" : "s"} hidden`);
    },
  },
  {
    id: "data.clearFilter", label: "Clear Filter", params: [], journaled: true,
    run: async (wb) => {
      const page = wb.page(wb.active)!;
      if (!page.filtered) refuse("no filter");
      const n = [...page.filtered!].filter((r) => !page.hiddenRows.has(r)).length;
      page.filtered = null;
      page.filter = null;
      await wb.hiddenChanged();
      return changedBy(wb, wb.sheet.keys(), `filter cleared: ${n} row${n === 1 ? "" : "s"} shown`);
    },
  },
  {
    id: "view.freezeRows", label: "Freeze Top Rows", params: ["count"], journaled: true,
    run: async (wb, p) => {
      const n = Number(p.count);
      if (!Number.isInteger(n) || n < 0 || n >= ROWS) refuse(`${p.count} is not a number of rows to freeze`);
      wb.page(wb.active)!.frozen = n;
      return changedBy(wb, [], n ? `froze ${n} row${n === 1 ? "" : "s"}` : "unfroze rows");
    },
  },
  {
    // Split: a second pane over the sheet under the first, each scrolling
    // on its own (Excel's View > Split); off=1 removes it. A view.
    id: "view.split", label: "Split", params: [], optional: ["off"], journaled: false,
    run: async (wb, p) => {
      wb.split = !flag(p.off);
      return { ok: true, message: wb.split ? "split in two panes" : "one pane", changed: [], output: wb.split ? 1 : 0 };
    },
  },
  {
    // Freeze the first columns (Freeze First Column, Freeze Panes' columns);
    // 0 frees them.
    id: "view.freezeCols", label: "Freeze Columns", params: ["count"], journaled: true,
    run: async (wb, p) => {
      const n = Number(p.count);
      if (!Number.isInteger(n) || n < 0 || n >= 9) refuse(`${p.count} is not a number of columns to freeze (0 to 8: the grid shows ten)`);
      wb.page(wb.active)!.frozenCols = n;
      return changedBy(wb, [], n ? `froze ${n} column${n === 1 ? "" : "s"}` : "unfroze columns");
    },
  },
  {
    // Freeze Panes at a cell, as Excel's: the rows above it and the columns
    // left of it; at A1 everything unfrozen.
    id: "view.freezePanes", label: "Freeze Panes", params: ["cell"], journaled: true,
    run: async (wb, p) => {
      const a = rangeOf(p.cell)[0], page = wb.page(wb.active)!;
      if (a.col >= 9) refuse(`${p.cell}: at most 8 columns freeze (the grid shows ten)`);
      page.frozen = a.row;
      page.frozenCols = a.col;
      return changedBy(wb, [], a.row || a.col ? `froze ${a.row} row${a.row === 1 ? "" : "s"} and ${a.col} column${a.col === 1 ? "" : "s"}` : "unfroze panes");
    },
  },
  {
    id: "cells.merge", label: "Merge Cells", params: ["range"], optional: ["center"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), tl = topLeft(r), page = wb.page(wb.active)!;
      if (keyOf(r[0]) === keyOf(r[1])) refuse(`${keyOf(r[0])} is one cell: nothing to merge`);
      const corners: [Addr, Addr] = [tl, { col: Math.max(r[0].col, r[1].col), row: Math.max(r[0].row, r[1].row) }];
      page.merges = [...page.merges.filter((m) => !overlap(m, corners)), corners];
      // As Excel does: the top-left's value is kept, the others' cleared.
      const cleared = [...range(r[0], r[1])].filter((k) => k !== keyOf(tl) && wb.sheet.input(k) !== "");
      if (cleared.length) await wb.sheet.setMany(cleared.map((k) => [k, ""] as const));
      if (flag(p.center)) wb.notes.set(keyOf(tl), { look: { ...wb.notes.get(keyOf(tl)).look, align: "center" } });
      return changedBy(wb, [...range(r[0], r[1])], `merged ${rangeName(corners)}${cleared.length ? ` (${cleared.length} cell${cleared.length === 1 ? "" : "s"} cleared)` : ""}`);
    },
  },
  {
    id: "cells.unmerge", label: "Unmerge Cells", params: ["range"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), page = wb.page(wb.active)!;
      const gone = page.merges.filter((m) => overlap(m, r));
      if (!gone.length) refuse(`no merge in ${rangeName(r)}`);
      page.merges = page.merges.filter((m) => !overlap(m, r));
      return changedBy(wb, gone.flatMap((m) => [...range(m[0], m[1])]), `unmerged ${gone.length}`);
    },
  },
  {
    id: "rows.height", label: "Row Height", params: ["range", "height"], journaled: true,
    run: async (wb, p) => {
      const h = Number(p.height);
      if (!Number.isInteger(h) || h < 0) refuse(`${p.height} is not a height (0 for the grid's)`);
      const page = wb.page(wb.active)!, { at, said } = spanOf(rangeOf(p.range), true);
      for (const row of at) if (h) page.heights.set(row, h); else page.heights.delete(row);
      return changedBy(wb, [], h ? `${said} ${at.length > 1 ? "are" : "is"} ${h} high` : `${said} ${at.length > 1 ? "are" : "is"} as tall as the grid's`);
    },
  },
  {
    id: "columns.width", label: "Column Width", params: ["range", "width"], journaled: true,
    run: async (wb, p) => {
      const w = Number(p.width);
      if (!Number.isInteger(w) || w < 0) refuse(`${p.width} is not a width (0 for the grid's)`);
      const page = wb.page(wb.active)!, { at, said } = spanOf(rangeOf(p.range), false);
      for (const col of at) if (w) page.widths.set(col, w); else page.widths.delete(col);
      return changedBy(wb, [], w ? `${said} ${at.length > 1 ? "are" : "is"} ${w} wide` : `${said} ${at.length > 1 ? "are" : "is"} as wide as the grid's`);
    },
  },
  {
    // AutoFit Column Width: each column as wide as its widest shown text,
    // by an estimate of the grid's font (7.5 px a character at 13 px, more
    // for a larger size) and its padding; an empty column the grid's.
    id: "columns.autofit", label: "AutoFit Column Width", params: ["range"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, { at, said } = spanOf(rangeOf(p.range), false), widths: number[] = [];
      for (const col of at) {
        let w = 0;
        for (let row = 0; row < ROWS; row++) {
          const key = keyOf({ col, row }), t = wb.textOf(key);
          if (!t) continue;
          const size = wb.notes.get(key).look?.size ?? 13;
          w = Math.max(w, Math.ceil(Math.max(...t.split("\n").map((l) => l.length)) * 7.5 * size / 13) + 14);
        }
        if (w) page.widths.set(col, Math.min(600, Math.max(24, w))); else page.widths.delete(col);
        widths.push(page.widths.get(col) ?? 0);
      }
      return { ...await changedBy(wb, [], `${said} fitted: ${widths.map((w) => w || "the grid's").join(", ")}`), output: widths };
    },
  },
  {
    // AutoFit Row Height: each row as tall as its tallest text, by its
    // lines and font size; a row of one line at the grid's size the grid's.
    id: "rows.autofit", label: "AutoFit Row Height", params: ["range"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, { at, said } = spanOf(rangeOf(p.range), true), heights: number[] = [];
      for (const row of at) {
        let h = 0;
        for (let col = 0; col < COLS; col++) {
          const key = keyOf({ col, row }), t = wb.textOf(key);
          if (!t) continue;
          const size = wb.notes.get(key).look?.size ?? 13, lines = t.split("\n").length;
          if (lines > 1 || size > 13) h = Math.max(h, Math.ceil(lines * size * 1.45 + 8));
        }
        if (h) page.heights.set(row, Math.min(400, h)); else page.heights.delete(row);
        heights.push(page.heights.get(row) ?? 0);
      }
      return { ...await changedBy(wb, [], `${said} fitted: ${heights.map((h) => h || "the grid's").join(", ")}`), output: heights };
    },
  },
  {
    id: "table.define", label: "Define Table", params: ["name", "range"], optional: ["headers", "columns", "types", "descriptions"], journaled: true,
    run: async (wb, p) => {
      const name = p.name.trim();
      if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(name)) refuse(`${p.name} is not a table name: a letter first, then letters, digits or _`);
      const r = rangeOf(p.range), width = Math.abs(r[0].col - r[1].col) + 1;
      const columns = p.columns === undefined ? null : p.columns.split(",").map((c) => c.trim());
      if (columns && columns.length !== width) refuse(`${name} has ${width} columns, and ${columns.length} name${columns.length === 1 ? "" : "s"}`);
      const headers = p.headers === undefined ? !columns : flag(p.headers);
      const corners: [Addr, Addr] = [topLeft(r), { col: Math.max(r[0].col, r[1].col), row: Math.max(r[0].row, r[1].row) }];
      const changed = await wb.sheet.defineTable(name, { name, range: corners, headers, columns, types: byName(p.types, "types"), descriptions: byName(p.descriptions, "descriptions") });
      return changedBy(wb, changed, `${name} is ${rangeName(corners)}, with columns ${tableColumns(wb, wb.sheet.table(name)!).join(", ")}`);
    },
  },
  {
    id: "table.delete", label: "Delete Table", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.name) ?? refuse(`no table ${p.name}`);
      const changed = await wb.sheet.defineTable(t.name, null);
      return changedBy(wb, changed, `deleted ${t.name}`);
    },
  },
  {
    // A table's totals row (Table Design's Total Row): the row under it
    // taken in, which must be empty, each column's aggregate by its name
    // (functions="Qty=sum,Price=average"; the first column "Total" and the
    // last a sum by default, as Excel's); on=0 takes it off, clearing it.
    id: "table.totals", label: "Total Row", params: ["name"], optional: ["on", "functions"], journaled: true,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.name) ?? refuse(`no table ${p.name}`), cols = tableColumns(wb, t);
      const left = Math.min(t.range[0].col, t.range[1].col), right = Math.max(t.range[0].col, t.range[1].col), bottom = Math.max(t.range[0].row, t.range[1].row);
      const top = Math.min(t.range[0].row, t.range[1].row);
      const row = (r: number) => Array.from({ length: right - left + 1 }, (_, i) => keyOf({ col: left + i, row: r }));
      if (p.on !== undefined && !flag(p.on)) {
        if (!t.totals) refuse(`${t.name} has no totals row`);
        const { totals: _gone, ...rest } = t;
        await wb.sheet.defineTable(t.name, { ...rest, range: [{ col: left, row: top }, { col: right, row: bottom - 1 }] });
        return setting(wb, row(bottom).map((k) => [k, ""]), `${t.name}: no totals row`);
      }
      const given: Record<string, string> = p.functions?.trim().startsWith("{") ? byName(p.functions, "functions") ?? {}
        : Object.fromEntries((p.functions ?? "").split(",").map((x) => x.trim()).filter((x) => x).map((x) => (x.includes("=") ? [x.slice(0, x.indexOf("=")).trim(), x.slice(x.indexOf("=") + 1).trim()] : refuse(`${x} is not column=total`))));
      const known = [...Object.keys(TOTALS), "none", "label"];
      for (const [c, f] of Object.entries(given)) {
        if (!cols.some((x) => x.toUpperCase() === c.toUpperCase())) refuse(`${t.name} has no column ${c}`);
        if (!known.includes(f.toLowerCase())) refuse(`${f} is not a total (${known.join(", ")})`);
      }
      const at = t.totals ? bottom : bottom + 1;
      if (at >= ROWS) refuse(`${t.name} has no row under it for totals`);
      if (!t.totals && row(at).some((k) => wb.sheet.input(k) !== "")) refuse(`the row under ${t.name} (${at + 1}) is not empty`);
      const functions: Record<string, string> = {};
      cols.forEach((c, i) => {
        const asked = Object.entries(given).find(([k]) => k.toUpperCase() === c.toUpperCase())?.[1].toLowerCase();
        functions[c] = asked ?? t.totals?.[c] ?? (i === cols.length - 1 && cols.length > 1 ? "sum" : i === 0 ? "label" : "none");
      });
      const def: TableDef = { ...t, range: [{ col: left, row: top }, { col: right, row: at }], totals: functions };
      const [d1, d2] = dataRows(def), excel = wb.sheet.language.name === "Excel";
      await wb.sheet.defineTable(t.name, def);
      const entries = cols.map((c, i): [string, string] => [keyOf({ col: left + i, row: at }),
        totalsInput(functions[c], t.name, c, `${keyOf({ col: left + i, row: d1 })}:${keyOf({ col: left + i, row: d2 })}`, excel)]);
      return setting(wb, entries, `${t.name}: totals in row ${at + 1} (${cols.map((c) => `${c} ${functions[c]}`).join(", ")})`);
    },
  },
  {
    // A calculated column: one formula in every data row of a column,
    // written for the first and moved down as a fill is ([@Qty] reads each
    // row's own).
    id: "table.calculatedColumn", label: "Calculated Column", params: ["name", "column", "formula"], journaled: true,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.name) ?? refuse(`no table ${p.name}`), cols = tableColumns(wb, t);
      const i = cols.findIndex((c) => c.toUpperCase() === p.column.trim().toUpperCase());
      if (i < 0) refuse(`${t.name} has no column ${p.column}`);
      // Excel's shorthand inside a table, [@Qty], is the table's own.
      const written = (p.formula.trim().startsWith("=") ? p.formula.trim() : `=${p.formula.trim()}`).replace(/(^|[^A-Za-z0-9_.\]'])\[@/g, `$1${t.name}[@`);
      const formula = written;
      const [d1, d2] = dataRows(t), col = Math.min(t.range[0].col, t.range[1].col) + i;
      if (d1 > d2) refuse(`${t.name} has no data rows`);
      const entries = Array.from({ length: d2 - d1 + 1 }, (_, r): [string, string] => [keyOf({ col, row: d1 + r }), movedBy(wb.sheet.language, formula, { cols: 0, rows: r })]);
      return setting(wb, entries, `${t.name}[${cols[i]}] = ${formula}`);
    },
  },
  {
    // Banded Rows: every other data row shaded (rows=0 takes it off).
    id: "table.banding", label: "Banded Rows", params: ["name"], optional: ["rows"], journaled: true,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.name) ?? refuse(`no table ${p.name}`), on = p.rows === undefined || flag(p.rows);
      const { banded: _was, ...rest } = t;
      await wb.sheet.defineTable(t.name, on ? { ...rest, banded: true } : rest);
      return changedBy(wb, [...range(t.range[0], t.range[1])], `${t.name}: ${on ? "banded rows" : "no banded rows"}`);
    },
  },
  {
    // Convert to Range: the table gone, its cells and looks kept (its
    // banding as the cells' fill), each structured reference to it on the
    // sheet written as the cells it came to ($B$2:$B$4; [@Qty] as $B2).
    id: "table.toRange", label: "Convert to Range", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.name) ?? refuse(`no table ${p.name}`);
      const header = (a: Addr) => show(wb.sheet.value(keyOf(a)));
      const entries = wb.sheet.keys().flatMap((k): [string, string][] => {
        const input = wb.sheet.input(k), now = unstructured(input, t, header, cellAddr(k)!.row);
        return now === input ? [] : [[k, now]];
      });
      if (t.banded) for (const k of bandedKeys(t)) wb.notes.set(k, { look: { ...wb.notes.get(k).look, fill: BAND } });
      await wb.sheet.defineTable(t.name, null);
      return { ...await setting(wb, entries, `${t.name} is a range again`), changed: [...range(t.range[0], t.range[1]), ...entries.map(([k]) => k)] };
    },
  },
  {
    // A table style (styles.ts's gallery: Light, Medium, Dark, Green,
    // Orange, Grey; none for plain) and Table Style Options: banded rows
    // and columns, first and last column bold.
    id: "table.style", label: "Table Style", params: ["name"], optional: ["style", "bandedRows", "bandedColumns", "firstColumn", "lastColumn"], journaled: true,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.name) ?? refuse(`no table ${p.name}`);
      const style = p.style === undefined ? t.style : /^none$/i.test(p.style.trim()) ? undefined : styleNamed(TABLE_STYLE_NAMES, p.style) ?? refuse(`${p.style} is not a table style (${TABLE_STYLE_NAMES.join(", ")}, none)`);
      const opt = (v: string | undefined, was: boolean | undefined) => (v === undefined ? was : flag(v) || undefined);
      const { style: _s, banded: _b, bandedColumns: _c, firstColumn: _f, lastColumn: _l, ...rest } = t;
      const def: TableDef = { ...rest, ...(style ? { style } : {}), ...(opt(p.bandedRows, t.banded ?? (style !== undefined && t.style === undefined ? true : undefined)) ? { banded: true } : {}),
        ...(opt(p.bandedColumns, t.bandedColumns) ? { bandedColumns: true } : {}), ...(opt(p.firstColumn, t.firstColumn) ? { firstColumn: true } : {}), ...(opt(p.lastColumn, t.lastColumn) ? { lastColumn: true } : {}) };
      await wb.sheet.defineTable(t.name, def);
      const options = [def.banded ? "banded rows" : "", def.bandedColumns ? "banded columns" : "", def.firstColumn ? "first column" : "", def.lastColumn ? "last column" : ""].filter((x) => x);
      return changedBy(wb, [...range(t.range[0], t.range[1])], `${t.name}: ${def.style ?? "no"} style${options.length ? `, ${options.join(", ")}` : ""}`);
    },
  },
  {
    // Page Layout's Themes: a theme by name (Office, Greyscale, Harbour,
    // Ember), its colour roles and fonts; whatever names a role (a look's
    // fill=accent1, a cell style's, a table style's) re-coloured.
    id: "theme.set", label: "Themes", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const name = styleNamed(Object.keys(THEMES), p.name) ?? refuse(`${p.name} is not a theme (${Object.keys(THEMES).join(", ")})`);
      wb.theme = THEMES[name];
      return changedBy(wb, wb.sheet.keys(), `theme ${name}: ${themeSaid(wb.theme)}`);
    },
  },
  {
    // Customize Colors: one role's colour (a #rrggbb), the theme renamed
    // Custom; and Customize Fonts: its major (headings) and minor (body).
    id: "theme.color", label: "Theme Colors", params: ["role", "color"], journaled: true,
    run: async (wb, p) => {
      const role = (ROLES as readonly string[]).find((r) => r === p.role.trim().toLowerCase()) ?? refuse(`${p.role} is not a theme colour (${ROLES.join(", ")})`);
      const c = colour(p.color);
      if (!c || isRole(c)) refuse(`${p.color} is not a colour (#rrggbb, or a name)`);
      wb.theme = { ...wb.theme, name: "Custom", colors: { ...wb.theme.colors, [role]: c } };
      return changedBy(wb, wb.sheet.keys(), `theme Custom: ${role} ${c}`);
    },
  },
  {
    id: "theme.fonts", label: "Theme Fonts", params: [], optional: ["major", "minor"], journaled: true,
    run: async (wb, p) => {
      const font = (f: string | undefined, was: string) => (f === undefined ? was : f.trim() || refuse("a font has a name"));
      wb.theme = { ...wb.theme, name: wb.theme.name === "Custom" || p.major !== undefined || p.minor !== undefined ? "Custom" : wb.theme.name, fonts: { major: font(p.major, wb.theme.fonts.major), minor: font(p.minor, wb.theme.fonts.minor) } };
      return changedBy(wb, wb.sheet.keys(), `theme ${wb.theme.name}: headings ${wb.theme.fonts.major}, body ${wb.theme.fonts.minor}`);
    },
  },
  {
    id: "theme.get", label: "About the Theme", params: [], journaled: false,
    run: async (wb) => ({ ok: true, message: `theme ${wb.theme.name}: ${themeSaid(wb.theme)}`, changed: [], output: wb.theme }),
  },
  {
    // A slicer for a table's column (Insert Slicer): its values as tiles;
    // picking some shows only the table's rows holding one (the table's
    // filter, as the sheet's AutoFilter is), none picked shows them all.
    id: "slicer.add", label: "Insert Slicer", params: ["table", "column"], optional: ["name"], journaled: true,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.table) ?? refuse(`no table ${p.table}`), page = wb.page(wb.active)!;
      const column = tableColumns(wb, t).find((c) => c.toLowerCase() === p.column.trim().toLowerCase()) ?? refuse(`${t.name} has no column ${p.column}`);
      const name = (p.name ?? column).trim() || column;
      if (page.slicers.some((x) => x.name.toLowerCase() === name.toLowerCase())) refuse(`there is a slicer ${name} already`);
      page.slicers = [...page.slicers, { name, table: t.name, column, picked: [] }];
      return changedBy(wb, [], `slicer ${name} for ${t.name}[${column}]: ${slicerValues(page, page.slicers[page.slicers.length - 1]).join(", ")}`);
    },
  },
  {
    // Its tiles picked: values= (comma-separated) picks them alone; add=1
    // adds them (Multi-Select), toggle=1 picks or unpicks each.
    id: "slicer.select", label: "Slicer", params: ["name", "values"], optional: ["add", "toggle"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, i = page.slicers.findIndex((x) => x.name.toLowerCase() === p.name.trim().toLowerCase());
      if (i < 0) refuse(`no slicer ${p.name}`);
      const sl = page.slicers[i], known = slicerValues(page, sl);
      const asked = p.values.split(",").map((v) => v.trim()).filter((v) => v).map((v) => known.find((k) => k.toLowerCase() === v.toLowerCase()) ?? refuse(`${v} is not a value of ${sl.name} (${known.join(", ")})`));
      const has = (xs: readonly string[], v: string) => xs.some((x) => x.toLowerCase() === v.toLowerCase());
      const picked = flag(p.toggle) ? [...sl.picked.filter((x) => !has(asked, x)), ...asked.filter((v) => !has(sl.picked, v))] : flag(p.add) ? [...sl.picked, ...asked.filter((v) => !has(sl.picked, v))] : asked;
      page.slicers = page.slicers.map((x, j) => (j === i ? { ...x, picked } : x));
      return sliced(wb, page, `${sl.name}: ${picked.length ? picked.join(", ") : "all"}`);
    },
  },
  {
    // Clear Filter: every tile picked again, the rows shown.
    id: "slicer.clear", label: "Clear Slicer", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, sl = page.slicers.find((x) => x.name.toLowerCase() === p.name.trim().toLowerCase()) ?? refuse(`no slicer ${p.name}`);
      page.slicers = page.slicers.map((x) => (x === sl ? { ...x, picked: [] } : x));
      return sliced(wb, page, `${sl.name}: all`);
    },
  },
  {
    id: "slicer.delete", label: "Delete Slicer", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, sl = page.slicers.find((x) => x.name.toLowerCase() === p.name.trim().toLowerCase()) ?? refuse(`no slicer ${p.name}`);
      page.slicers = page.slicers.filter((x) => x !== sl);
      return sliced(wb, page, `deleted slicer ${sl.name}`);
    },
  },
  {
    id: "slicer.list", label: "Slicers", params: [], journaled: false,
    run: async (wb) => {
      const views = wb.slicerViews(), listed = views.map((x) => `${x.caption}: ${x.tiles.map((t) => (t.picked ? `[${t.value}]` : t.value)).join(" ")}`);
      return { ok: true, message: listed.length ? listed.join("; ") : "no slicers", changed: [], output: listed };
    },
  },
  {
    // Solver, Simplex LP: set objective= to max, min or value= by changing
    // changing= (cells or ranges, comma-separated) subject to constraints=
    // ("B5<=40;B6>=B7;B1:B3=int;B4=bin", a range against a number or a
    // range of the same size); the sheet's formulas must be linear in the
    // changing cells (each read off by setting them, as Excel's Simplex
    // does). nonneg=0 frees the unconstrained ones. The solution kept (one
    // step of undo) unless restore=1; report=1 writes Excel's Answer
    // Report on a sheet of its own.
    id: "solver.solve", label: "Solver", params: ["objective", "changing"], optional: ["goal", "value", "constraints", "nonneg", "restore", "report"], journaled: true,
    run: async (wb, p) => {
      const obj = cellOf(p.objective), goal = (p.goal ?? "max").trim().toLowerCase();
      if (!["max", "min", "value"].includes(goal)) refuse(`${p.goal} is not max, min or value`);
      if (goal === "value" && !isFinite(Number(p.value))) refuse("goal=value needs value= a number");
      const vars = [...new Set(p.changing.split(",").flatMap((x) => [...range(...rangeOf(x.trim()))]))];
      if (!vars.length || vars.length > 200) refuse("Solver changes 1 to 200 cells, as Excel's");
      for (const k of vars) if (wb.sheet.input(k).startsWith("=")) refuse(`${k} holds a formula, not a value Solver can change`);
      const was = vars.map((k): [string, string] => [k, wb.sheet.input(k)]);
      const num = (k: string) => { const v = wb.sheet.value(k); return typeof v === "number" ? v : v === null || v === "" ? 0 : NaN; };
      const at = async (x: readonly number[]) => { await wb.sheet.setMany(vars.map((k, i): [string, string] => [k, String(x[i])])); await wb.settle(); };
      // The constraints: each a cell's (or each of a range's cells') relation.
      const ints = new Set<string>(), bins = new Set<string>();
      const cons: { lhs: string; rel: Relation; rhs: string | number; said: string }[] = [];
      for (const part of (p.constraints ?? "").split(";").map((x) => x.trim()).filter((x) => x)) {
        const m = /^(.+?)\s*(<=|>=|=)\s*(.+)$/.exec(part) ?? refuse(`${part} is not a constraint (B5<=40, B1:B3=int)`);
        const lhs = [...range(...rangeOf(m[1].trim()))];
        if (/^(int|bin)$/i.test(m[3].trim()) && m[2] === "=") { for (const k of lhs) (m[3].trim().toLowerCase() === "int" ? ints : bins).add(k); continue; }
        const rhsCells = parseRange(m[3].trim()) ? [...range(...rangeOf(m[3].trim()))] : null;
        if (rhsCells && rhsCells.length !== 1 && rhsCells.length !== lhs.length) refuse(`${part}: the two sides differ in size`);
        if (!rhsCells && !isFinite(Number(m[3]))) refuse(`${part}: ${m[3]} is not a number or a cell`);
        lhs.forEach((k, i) => cons.push({ lhs: k, rel: m[2] as Relation, rhs: rhsCells ? rhsCells[rhsCells.length === 1 ? 0 : i] : Number(m[3]), said: `${k} ${m[2]} ${rhsCells ? rhsCells[rhsCells.length === 1 ? 0 : i] : m[3]}` }));
      }
      for (const k of [...ints, ...bins]) if (!vars.includes(k)) refuse(`${k} is int or bin but not a changing cell`);
      const origValues = vars.map(num), origObj = num(obj);
      // Each formula's coefficients, read off: its value with every
      // changing cell 0, then each one 1; then a check at another point.
      const watched = [obj, ...cons.flatMap((c) => [c.lhs, ...(typeof c.rhs === "string" ? [c.rhs] : [])])];
      const read = () => watched.map(num);
      try {
        await at(vars.map(() => 0));
        const base = read(), coef = watched.map(() => vars.map(() => 0));
        for (let j = 0; j < vars.length; j++) {
          await at(vars.map((_, i) => (i === j ? 1 : 0)));
          read().forEach((v, w) => (coef[w][j] = v - base[w]));
        }
        const probe = vars.map((_, i) => i + 2);
        await at(probe);
        const now = read();
        const linear = watched.every((_, w) => { const want = base[w] + coef[w].reduce((s, c, j) => s + c * probe[j], 0); return Math.abs(now[w] - want) <= 1e-7 * (1 + Math.abs(want)); });
        if (watched.some((_, w) => !isFinite(base[w]))) refuse(`${watched.find((_, w) => !isFinite(base[w]))} is not a number with the changing cells at 0`);
        if (!linear) refuse(`the problem is not linear in ${vars.join(", ")}: Simplex LP needs linear formulas`);
        const wi = (k: string) => watched.indexOf(k);
        const rows: LpRow[] = cons.map((c) => {
          const l = coef[wi(c.lhs)], lb = base[wi(c.lhs)];
          if (typeof c.rhs === "number") return { a: l, rel: c.rel, b: c.rhs - lb };
          const r = coef[wi(c.rhs)], rb = base[wi(c.rhs)];
          return { a: l.map((v, j) => v - r[j]), rel: c.rel, b: rb - lb };
        });
        for (const k of bins) rows.push({ a: vars.map((v) => (v === k ? 1 : 0)), rel: "<=", b: 1 });
        if (goal === "value") rows.push({ a: coef[0], rel: "=", b: Number(p.value) - base[0] });
        const free = vars.map(() => p.nonneg !== undefined && !flag(p.nonneg));
        const result = solveLp({ c: goal === "value" ? vars.map(() => 0) : coef[0], maximise: goal !== "min", rows, free, integer: vars.map((k) => ints.has(k) || bins.has(k)) });
        if (result.status !== "optimal") {
          await wb.sheet.setMany(was);
          await wb.settle();
          refuse(result.status === "infeasible" ? "Solver could not find a feasible solution" : "the objective's cells do not converge: the problem is unbounded");
        }
        const x = (result as { x: number[] }).x.map((v) => Number(v.toPrecision(12)));
        if (flag(p.restore)) { await wb.sheet.setMany(was); await wb.settle(); }
        else await at(x);
        const finalObj = base[0] + coef[0].reduce((s, c, j) => s + c * x[j], 0);
        if (flag(p.report)) {
          const source = wb.active;
          let name = "Answer Report 1";
          for (let i = 2; wb.page(name); i++) name = `Answer Report ${i}`;
          await wb.add(name);
          const lines: string[][] = [["Answer Report"], [], ["Objective Cell (" + (goal === "value" ? `Value Of ${p.value}` : goal === "max" ? "Max" : "Min") + ")"], ["Cell", "Original Value", "Final Value"], [`${source}!${obj}`, show(origObj), String(Number(finalObj.toPrecision(12)))], [],
            ["Variable Cells"], ["Cell", "Original Value", "Final Value", "Integer"], ...vars.map((k, i) => [`${source}!${k}`, show(origValues[i]), String(x[i]), ints.has(k) ? "Integer" : bins.has(k) ? "Binary" : "Contin"]), [],
            ["Constraints"], ["Cell", "Cell Value", "Formula", "Status", "Slack"],
            ...cons.map((c, i) => {
              const lhs = base[wi(c.lhs)] + coef[wi(c.lhs)].reduce((s, v, j) => s + v * x[j], 0), slackV = Math.abs(rows[i].b - rows[i].a.reduce((s, v, j) => s + v * x[j], 0));
              return [`${source}!${c.lhs}`, String(Number(lhs.toPrecision(12))), c.said, slackV < 1e-7 || c.rel === "=" ? "Binding" : "Not Binding", String(Number(slackV.toPrecision(12)))];
            })];
          await wb.sheet.setMany(lines.flatMap((l, r) => l.map((v, c): [string, string] => [keyOf({ col: c, row: r }), /^[=']/.test(v) || /^[-+]?\d/.test(v) && !NUMERIC_TEXT.test(v) ? `'${v}` : v])));
          wb.activate(source);
        }
        await wb.settle();
        return { ok: true, message: `Solver found a solution: ${obj} = ${Number(finalObj.toPrecision(12))} at ${vars.map((k, i) => `${k}=${x[i]}`).join(", ")}${flag(p.restore) ? " (the cells as they were)" : ""}`, changed: wb.sheet.keys(), output: x };
      } catch (e) {
        await wb.sheet.setMany(was);
        await wb.settle();
        throw e;
      }
    },
  },
  {
    // Scenario Manager's Add: a name, its changing cells ("B1,B2" or
    // "B1:B3"), their values (comma-separated; the cells' own when left
    // out); the same name again edits it.
    id: "scenario.add", label: "Add Scenario", params: ["name", "cells"], optional: ["values", "comment"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, name = p.name.trim() || refuse("a scenario has a name");
      const cells = p.cells.split(",").flatMap((x) => [...range(...rangeOf(x.trim()))]);
      if (!cells.length || cells.length > 32) refuse("a scenario changes 1 to 32 cells, as Excel's");
      const values = p.values === undefined ? cells.map((k) => wb.sheet.input(k)) : p.values.split(",").map((x) => x.trim());
      if (values.length !== cells.length) refuse(`${cells.length} cells and ${values.length} values`);
      const sc: Scenario = { name, cells, values, ...(p.comment ? { comment: p.comment } : {}) };
      const i = page.scenarios.findIndex((x) => x.name.toLowerCase() === name.toLowerCase());
      page.scenarios = i < 0 ? [...page.scenarios, sc] : page.scenarios.map((x, j) => (j === i ? sc : x));
      return changedBy(wb, [], `scenario ${name}: ${cells.map((k, j) => `${k}=${values[j]}`).join(", ")}`);
    },
  },
  {
    // Show: its values put in its cells (one step of undo).
    id: "scenario.show", label: "Show Scenario", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const sc = scenarioNamed(wb, p.name);
      return setting(wb, sc.cells.map((k, i): [string, string] => [k, sc.values[i]]), `showed scenario ${sc.name}`);
    },
  },
  {
    id: "scenario.delete", label: "Delete Scenario", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, sc = scenarioNamed(wb, p.name);
      page.scenarios = page.scenarios.filter((x) => x !== sc);
      return changedBy(wb, [], `deleted scenario ${sc.name}`);
    },
  },
  {
    id: "scenario.list", label: "Scenarios", params: [], journaled: false,
    run: async (wb) => {
      const listed = wb.page(wb.active)!.scenarios.map((x) => `${x.name}: ${x.cells.map((k, j) => `${k}=${x.values[j]}`).join(", ")}`);
      return { ok: true, message: listed.length ? listed.join("; ") : "no scenarios", changed: [], output: listed };
    },
  },
  {
    // Summary: a sheet "Scenario Summary" (or to=), a row a changing cell
    // and a result cell (results=), a column for the current values and
    // one a scenario, each scenario's results worked out with its values
    // in, the sheet as it was after.
    id: "scenario.summary", label: "Scenario Summary", params: ["results"], optional: ["to"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, scs = page.scenarios;
      if (!scs.length) refuse("no scenarios");
      const results = p.results.split(",").flatMap((x) => [...range(...rangeOf(x.trim()))]);
      const changing = [...new Set(scs.flatMap((x) => x.cells))];
      const was = changing.map((k): [string, string] => [k, wb.sheet.input(k)]);
      const read = (ks: readonly string[]) => ks.map((k) => wb.sheet.value(k));
      const columns: { name: string; inputs: Value[]; results: Value[] }[] = [{ name: "Current Values", inputs: read(changing), results: read(results) }];
      for (const sc of scs) {
        await wb.sheet.setMany(sc.cells.map((k, i): [string, string] => [k, sc.values[i]]));
        await wb.settle();
        columns.push({ name: sc.name, inputs: read(changing), results: read(results) });
        await wb.sheet.setMany(was);
        await wb.settle();
      }
      const target = (p.to ?? "Scenario Summary").trim(), source = wb.active;
      if (wb.page(target)) refuse(`${target} is a sheet already`);
      await wb.add(target);
      const cell = (c: number, r: number) => keyOf({ col: c, row: r });
      const text = (v: Value) => (v === null ? "" : typeof v === "number" ? String(v) : show(v));
      const entries: [string, string][] = [[cell(0, 0), "Scenario Summary"], [cell(1, 1), "Cell"], ...columns.map((c, j): [string, string] => [cell(2 + j, 1), c.name]), [cell(0, 2), "Changing Cells:"]];
      changing.forEach((k, i) => { entries.push([cell(1, 3 + i), `${source}!${k}`]); columns.forEach((c, j) => entries.push([cell(2 + j, 3 + i), text(c.inputs[i])])); });
      const r0 = 4 + changing.length;
      entries.push([cell(0, r0 - 1), "Result Cells:"]);
      results.forEach((k, i) => { entries.push([cell(1, r0 + i), `${source}!${k}`]); columns.forEach((c, j) => entries.push([cell(2 + j, r0 + i), text(c.results[i])])); });
      wb.activate(target);
      return setting(wb, entries.map(([k, v]): [string, string] => [k, /^[=']/.test(v) ? `'${v}` : v]), `scenario summary on ${target}`);
    },
  },
  {
    // A cell style defined or edited: its look from a cell's (from=A1, as
    // Excel's New Cell Style by example), or another style's (basedOn=),
    // and these parts over it: bold, italic, underline, strike (1 or 0),
    // size, fill, color, number. Every cell of it is drawn anew.
    id: "style.define", label: "New Cell Style", params: ["name"], optional: ["from", "basedOn", "bold", "italic", "underline", "strike", "size", "fill", "color", "number"], journaled: true,
    run: async (wb, p) => {
      const name = p.name.trim();
      if (!name || name.length > 60) refuse(`${p.name} is not a style name`);
      const known = styleNamed(wb.styles.keys(), name) ?? name;
      let look: Look = p.from !== undefined ? wb.lookOf(wb.notes.get(cellOf(p.from))) ?? {}
        : p.basedOn !== undefined ? wb.styles.get(styleNamed(wb.styles.keys(), p.basedOn) ?? refuse(`no cell style ${p.basedOn}`))! : wb.styles.get(known) ?? {};
      for (const part of ["bold", "italic", "underline", "strike"] as const) if (p[part] !== undefined) look = { ...look, [part]: flag(p[part]) || undefined };
      if (p.size !== undefined) {
        const n = Number(p.size);
        if (!(n > 0 && n <= 400)) refuse(`${p.size} is not a size`);
        look = { ...look, size: n };
      }
      for (const part of ["fill", "color"] as const) if (p[part] !== undefined) look = { ...look, [part]: /^none$/i.test(p[part]!.trim()) ? undefined : colour(p[part]!) ?? refuse(`${p[part]} is not a colour`) };
      if (p.number !== undefined) look = { ...look, number: CATEGORIES[p.number.trim().toLowerCase()] ?? (p.number.trim() || undefined) };
      wb.styles.set(known, JSON.parse(JSON.stringify(look)));
      const keys = wb.sheet.keys().filter((k) => wb.notes.get(k).cellStyle === known);
      for (const k of wb.notes.keys()) if (wb.notes.get(k).cellStyle === known && !keys.includes(k)) keys.push(k);
      return changedBy(wb, keys, `cell style ${known}: ${describeLook(look)}`);
    },
  },
  {
    // A cell style on cells (Normal takes it off).
    id: "style.apply", label: "Cell Styles", params: ["range", "name"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), name = styleNamed(wb.styles.keys(), p.name) ?? refuse(`no cell style ${p.name} (${[...wb.styles.keys()].join(", ")})`);
      const keys = [...range(r[0], r[1])];
      for (const k of keys) wb.notes.set(k, { cellStyle: name === "Normal" ? undefined : name });
      return changedBy(wb, keys, `${named(r)}: ${name}`);
    },
  },
  {
    // A style deleted: its cells Normal again (Normal itself stays).
    id: "style.delete", label: "Delete Cell Style", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const name = styleNamed(wb.styles.keys(), p.name) ?? refuse(`no cell style ${p.name}`);
      if (name === "Normal") refuse("Normal cannot be deleted");
      wb.styles.delete(name);
      const keys = wb.notes.keys().filter((k) => wb.notes.get(k).cellStyle === name);
      for (const k of keys) wb.notes.set(k, { cellStyle: undefined });
      return changedBy(wb, keys, `deleted cell style ${name}`);
    },
  },
  {
    id: "style.list", label: "Cell Style Gallery", params: [], journaled: false,
    run: async (wb) => {
      const listed = [...wb.styles].map(([n, l]) => `${n}: ${describeLook(l)}`);
      return { ok: true, message: [...wb.styles.keys()].join(", "), changed: [], output: listed };
    },
  },
  {
    // A table as the app's metadata has it (GETTABLE, GET_TABLE_METADATA).
    id: "table.get", label: "About a table", params: ["name"], journaled: false,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.name) ?? refuse(`no table ${p.name}`);
      const about = {
        name: t.name, range: rangeName(t.range), headers: t.headers, columns: tableColumns(wb, t), types: t.types ?? {}, descriptions: t.descriptions ?? {},
        sort: t.sort ? { column: colName(t.sort.column), direction: t.sort.direction } : null, totals: t.totals ?? null, banded: !!t.banded,
        ...(t.style ? { style: t.style } : {}), ...(t.bandedColumns ? { bandedColumns: true } : {}), ...(t.firstColumn ? { firstColumn: true } : {}), ...(t.lastColumn ? { lastColumn: true } : {}),
      };
      return { ok: true, message: JSON.stringify(about), changed: [], output: about };
    },
  },
  {
    // The app's sort cycle: by a new column ascending; by the same one,
    // descending, then the rows as they were before the first sort.
    id: "table.sort", label: "Sort Table", params: ["name", "column"], journaled: true,
    run: async (wb, p) => {
      const t = wb.sheet.table(p.name) ?? refuse(`no table ${p.name}`), col = columnOf(p.column);
      const left = Math.min(t.range[0].col, t.range[1].col), right = Math.max(t.range[0].col, t.range[1].col);
      if (col < left || col > right) refuse(`${colName(col)} is not a column of ${t.name}`);
      const [top, bottom] = dataRows(t);
      if (top > bottom) refuse(`${t.name} has no rows to sort`);
      const a = { col: left, row: top }, b = { col: right, row: bottom }, height = bottom - top + 1;
      const origin = t.sort?.origin ?? Array.from({ length: height }, (_, i) => top + i);
      const next = t.sort?.column !== col ? "asc" : t.sort.direction === "asc" ? "desc" : "original";
      const rows = next === "original" ? Array.from({ length: height }, (_, i) => top + origin.indexOf(top + i)) : sortedRows(wb.sheet, a, b, col - left, next === "asc");
      const entries = rearranging(wb, [a, b], rows);
      await wb.sheet.defineTable(t.name, { ...t, sort: next === "original" ? undefined : { column: col, direction: next, origin: rows.map((from) => origin[from - top]) } });
      return { ...await setting(wb, entries, `${t.name} by ${colName(col)}: ${next}`), output: next };
    },
  },
  {
    // A pivot table (pivot.ts) written from a cell: rows= and cols= name
    // the fields grouped down and across (comma-separated), value= the one
    // aggregated by agg=, filters= a JSON object of a field's value a row
    // must have. The same id again replaces it.
    id: "pivot.create", label: "PivotTable", params: ["id", "source", "value", "agg", "output"], optional: ["rows", "cols", "filters"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, agg = p.agg.toUpperCase() as Aggregate;
      if (!AGGREGATES.includes(agg)) refuse(`${p.agg} is not ${AGGREGATES.join(", ")}`);
      const list = (s?: string) => (s ?? "").split(",").map((x) => x.trim()).filter((x) => x);
      const pivot: Pivot = {
        source: rangeOf(p.source), output: cellAddr(cellOf(p.output))!, rows: list(p.rows), cols: list(p.cols), value: p.value, agg,
        filters: byName(p.filters, "filters") ?? {}, written: page.pivots.get(p.id)?.written ?? null,
        // The same id again keeps its options (PIVOTOPTIONS).
        ...(page.pivots.get(p.id)?.options ? { options: page.pivots.get(p.id)!.options } : {}),
      };
      return writePivot(wb, p.id, pivot);
    },
  },
  {
    // Excel's PivotTable options (pivot.ts's PivotOptions), the pivot written
    // again: layout= tabular, outline or compact; subtotals= and
    // grandTotals= 1 or 0; sort= asc, desc, valueAsc, valueDesc or none;
    // showAs= normal, percentOfTotal, percentOfRow, percentOfColumn,
    // runningTotal, rank; group= a date field, by= year, quarter or month
    // (group=none takes it off); calculated="Tax=amount*0.2" (adds or
    // replaces one; "Tax=" removes it); collapse= and expand= an outer row
    // group.
    id: "pivot.options", label: "PivotTable Options", params: ["id"], optional: ["layout", "subtotals", "grandTotals", "sort", "showAs", "group", "by", "calculated", "collapse", "expand"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, pv = page.pivots.get(p.id) ?? refuse(`no pivot ${p.id}`);
      const o: Record<string, unknown> = { ...pv.options };
      const set = (k: string, v: unknown) => { if (v === undefined || v === false || v === "none" || v === "normal") delete o[k]; else o[k] = v; };
      if (p.layout !== undefined) { const l = p.layout.trim().toLowerCase(); if (!LAYOUTS.includes(l)) refuse(`${p.layout} is not a layout (${LAYOUTS.join(", ")})`); set("layout", l === "tabular" ? undefined : l); }
      if (p.subtotals !== undefined) set("subtotals", flag(p.subtotals));
      if (p.grandTotals !== undefined) set("grandTotals", flag(p.grandTotals));
      if (p.sort !== undefined) { const v = ["asc", "desc", "valueAsc", "valueDesc", "none"].find((x) => x.toLowerCase() === p.sort!.trim().toLowerCase()) ?? refuse(`${p.sort} is not a sort (asc, desc, valueAsc, valueDesc, none)`); set("sort", v); }
      if (p.showAs !== undefined) { const v = SHOW_AS.find((x) => x.toLowerCase() === p.showAs!.trim().toLowerCase()) ?? refuse(`${p.showAs} is not a way to show values (${SHOW_AS.join(", ")})`); set("showAs", v); }
      if (p.group !== undefined) {
        if (/^none$/i.test(p.group.trim())) set("dateGroup", undefined);
        else { const by = (p.by ?? "month").trim().toLowerCase(); if (!["year", "quarter", "month"].includes(by)) refuse(`${p.by} is not year, quarter or month`); o.dateGroup = { field: p.group.trim(), by }; }
      }
      if (p.calculated !== undefined) {
        const i = p.calculated.indexOf("="), name = p.calculated.slice(0, i).trim(), formula = p.calculated.slice(i + 1).trim();
        if (i < 1 || !name) refuse(`${p.calculated} is not Name=formula`);
        const calc: Record<string, string> = { ...(o.calculated as Record<string, string> | undefined) };
        if (formula) calc[name] = formula; else delete calc[name];
        set("calculated", Object.keys(calc).length ? calc : undefined);
      }
      const coll = new Set((o.collapsed as string[] | undefined) ?? []);
      if (p.collapse !== undefined) coll.add(p.collapse);
      if (p.expand !== undefined) coll.delete(p.expand);
      set("collapsed", coll.size ? [...coll] : undefined);
      const next = { ...pv, options: o as PivotOptions };
      try { return await writePivot(wb, p.id, next); } catch (e) { if (e instanceof RangeError) refuse(e.message); throw e; }
    },
  },
  {
    // A PivotChart: a chart of a pivot, following it as it is written
    // again (its grand totals left out).
    id: "chart.pivot", label: "PivotChart", params: ["id"], optional: ["kind", "title"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, pv = page.pivots.get(p.id) ?? refuse(`no pivot ${p.id}`);
      const kind = CHART_KINDS.find((k) => k.toLowerCase() === (p.kind ?? "bar").trim().toLowerCase()) ?? refuse(`${p.kind} is not a chart (${CHART_KINDS.join(", ")})`);
      page.charts.push({ kind, range: pv.written ?? [pv.output, pv.output], title: p.title ?? `${pv.agg} of ${pv.value}`, pivot: p.id });
      return changedBy(wb, [], `inserted a ${kind} PivotChart of pivot ${p.id}`);
    },
  },
  {
    // Written again from its source (the app's UPDATEPIVOT).
    id: "pivot.update", label: "Refresh PivotTable", params: ["id"], journaled: true,
    run: async (wb, p) => writePivot(wb, p.id, wb.page(wb.active)!.pivots.get(p.id) ?? refuse(`no pivot ${p.id}`)),
  },
  {
    // Forgotten; the cells it wrote stay, as the app leaves them.
    id: "pivot.delete", label: "Delete PivotTable", params: ["id"], journaled: true,
    run: async (wb, p) => {
      if (!wb.page(wb.active)!.pivots.delete(p.id)) refuse(`no pivot ${p.id}`);
      return { ok: true, message: `deleted pivot ${p.id}`, changed: [] };
    },
  },
  {
    id: "pivot.get", label: "About a PivotTable", params: ["id"], journaled: false,
    run: async (wb, p) => {
      const config = pivotConfig(wb.page(wb.active)!.pivots.get(p.id) ?? refuse(`no pivot ${p.id}`));
      return { ok: true, message: JSON.stringify(config), changed: [], output: config };
    },
  },
  {
    id: "pivot.list", label: "PivotTables", params: [], journaled: false,
    run: async (wb) => {
      const all = [...wb.page(wb.active)!.pivots].map(([id, pv]) => ({ id, config: pivotConfig(pv) }));
      return { ok: true, message: all.length ? all.map((x) => x.id).join(", ") : "no pivots", changed: [], output: all };
    },
  },
  {
    id: "table.list", label: "Tables", params: [], journaled: false,
    run: async (wb) => {
      const listed = wb.sheet.tables().map((t) => `${t.name} ${rangeName(t.range)} (${tableColumns(wb, t).join(", ")})`);
      return { ok: true, message: listed.length ? listed.join("; ") : "no tables", changed: [], output: listed };
    },
  },
  {
    id: "chart.insert", label: "Insert Chart", params: ["kind", "range"], optional: ["title"], journaled: true,
    run: async (wb, p) => {
      const kind = CHART_KINDS.find((k) => k.toLowerCase() === p.kind.trim().toLowerCase()) ?? refuse(`${p.kind} is not a chart (${CHART_KINDS.join(", ")})`);
      const page = wb.page(wb.active)!, r = rangeOf(p.range);
      page.charts.push({ kind, range: r, title: p.title ?? `Chart ${page.charts.length + 1}` });
      return changedBy(wb, [], `inserted a ${kind} chart of ${named(r)}`);
    },
  },
  {
    // Chart Elements: title=, xTitle=, yTitle= (none takes one off),
    // legend= bottom, top, right, left or none, labels= (data labels),
    // gridlines=, byRows= (Switch Row/Column): each 1 or 0.
    // secondary= the series (numbers from 1, comma-separated; none) a
    // combo draws on its secondary axis.
    id: "chart.option", label: "Chart Elements", params: ["index"], optional: ["title", "xTitle", "yTitle", "legend", "labels", "gridlines", "byRows", "secondary"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, i = Number(p.index) - 1, c = page.charts[i] ?? refuse(`no chart ${p.index}`);
      const o: Record<string, unknown> = { ...c.options };
      const textOpt = (k: "xTitle" | "yTitle") => { if (p[k] !== undefined) { if (/^none$/i.test(p[k]!.trim()) || !p[k]!.trim()) delete o[k]; else o[k] = p[k]; } };
      textOpt("xTitle"); textOpt("yTitle");
      if (p.legend !== undefined) {
        const l = p.legend.trim().toLowerCase();
        if (!LEGENDS.includes(l)) refuse(`${p.legend} is not a legend's place (${LEGENDS.join(", ")})`);
        if (l === "bottom") delete o.legend; else o.legend = l;
      }
      for (const k of ["labels", "byRows"] as const) if (p[k] !== undefined) { if (flag(p[k])) o[k] = true; else delete o[k]; }
      if (p.gridlines !== undefined) { if (flag(p.gridlines)) delete o.gridlines; else o.gridlines = false; }
      if (p.secondary !== undefined) {
        const ns = /^\s*(none)?\s*$/i.test(p.secondary) ? [] : p.secondary.split(",").map((x) => Number(x.trim()) - 1);
        if (ns.some((n) => !Number.isInteger(n) || n < 0)) refuse(`${p.secondary} is not series numbers (1, 2 ...)`);
        if (ns.length) o.secondary = ns; else delete o.secondary;
      }
      page.charts = page.charts.map((x, j) => (j === i ? { ...x, ...(p.title !== undefined ? { title: p.title } : {}), options: o as ChartOptions } : x));
      const said = [p.title !== undefined ? `title ${p.title || "none"}` : "", ...Object.entries(o).map(([k, v]) => `${k} ${v}`)].filter((x) => x).join(", ");
      return changedBy(wb, [], `chart ${i + 1}: ${said || "as it was"}`);
    },
  },
  {
    // A trendline on a series (by its number from 1, or its name): type=
    // linear, exponential, logarithmic, polynomial (order= 2 to 6), power,
    // movingAverage (period=), or none to take it off; equation=1, r2=1
    // label it.
    id: "chart.trendline", label: "Trendline", params: ["index", "series", "type"], optional: ["order", "period", "equation", "r2"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, i = Number(p.index) - 1, c = page.charts[i] ?? refuse(`no chart ${p.index}`);
      const names = wb.charts()[i].data.series.map((x) => x.name);
      const si = /^\d+$/.test(p.series.trim()) ? Number(p.series) - 1 : names.findIndex((n) => n.toLowerCase() === p.series.trim().toLowerCase());
      if (si < 0 || si >= names.length) refuse(`${p.series} is not a series of chart ${i + 1} (${names.join(", ")})`);
      const kind = p.type.trim().toLowerCase(), others = (c.options?.trendlines ?? []).filter((t) => t.series !== si);
      let said: string;
      if (kind === "none") { page.charts = page.charts.map((x, j) => (j === i ? { ...x, options: { ...x.options, trendlines: others } } : x)); said = "no trendline"; }
      else {
        const type = TREND_TYPES.find((t) => t.toLowerCase() === kind) ?? refuse(`${p.type} is not a trendline (${TREND_TYPES.join(", ")}, none)`);
        const order = p.order === undefined ? undefined : Number(p.order), period = p.period === undefined ? undefined : Number(p.period);
        if (order !== undefined && !(Number.isInteger(order) && order >= 2 && order <= 6)) refuse(`${p.order} is not an order (2 to 6)`);
        if (period !== undefined && !(Number.isInteger(period) && period >= 2)) refuse(`${p.period} is not a period (2 or more)`);
        const t: Trendline = { series: si, type, ...(type === "polynomial" ? { order: order ?? 2 } : {}), ...(type === "movingAverage" ? { period: period ?? 2 } : {}), ...(flag(p.equation) ? { equation: true } : {}), ...(flag(p.r2) ? { r2: true } : {}) };
        page.charts = page.charts.map((x, j) => (j === i ? { ...x, options: { ...x.options, trendlines: [...others, t] } } : x));
        said = `a ${type === "polynomial" ? `polynomial (order ${t.order})` : type === "movingAverage" ? `${t.period}-point moving average` : type} trendline`;
      }
      const shown = wb.charts()[i].shapes.filter((x): x is Extract<Shape, { kind: "text" }> => x.kind === "text" && /^y = |R² = /.test(x.text)).map((x) => x.text);
      return { ...await changedBy(wb, [], `chart ${i + 1}, ${names[si]}: ${said}${shown.length ? `: ${shown.join("; ")}` : ""}`), output: shown };
    },
  },
  {
    // Select Data: the chart's range; and Change Chart Type: its kind.
    id: "chart.data", label: "Select Data", params: ["index"], optional: ["range", "kind"], journaled: true,
    run: async (wb, p) => {
      const page = wb.page(wb.active)!, i = Number(p.index) - 1, c = page.charts[i] ?? refuse(`no chart ${p.index}`);
      const kind = p.kind === undefined ? c.kind : CHART_KINDS.find((k) => k.toLowerCase() === p.kind!.trim().toLowerCase()) ?? refuse(`${p.kind} is not a chart (${CHART_KINDS.join(", ")})`);
      const r = p.range === undefined ? c.range : rangeOf(p.range);
      page.charts = page.charts.map((x, j) => (j === i ? { ...x, kind, range: r } : x));
      return changedBy(wb, [], `chart ${i + 1}: a ${kind} chart of ${named(r)}`);
    },
  },
  {
    id: "chart.list", label: "Charts", params: [], journaled: false,
    run: async (wb) => {
      const listed = wb.charts().map((c, i) => `${i + 1}. ${c.title}: ${c.kind} chart of ${c.range}`);
      return { ok: true, message: listed.length ? listed.join("; ") : "no charts", changed: [], output: listed };
    },
  },
  {
    id: "chart.delete", label: "Delete Chart", params: ["index"], journaled: true,
    run: async (wb, p) => {
      const charts = wb.page(wb.active)!.charts, i = Number(p.index) - 1;
      if (!Number.isInteger(i) || !charts[i]) refuse(`no chart ${p.index}`);
      const [gone] = charts.splice(i, 1);
      return changedBy(wb, [], `deleted ${gone.title}`);
    },
  },
  {
    // A sparkline at a cell over a range: line, column or winloss, with
    // Excel's Show options (high, low, first, last; axis), as a formula.
    id: "insert.sparkline", label: "Sparkline", params: ["range", "at"], optional: ["type", "options"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), key = cellOf(p.at), type = (p.type ?? "line").trim().toLowerCase(), options = (p.options ?? "").trim();
      if (!(SPARKLINE_TYPES as readonly string[]).includes(type)) refuse(`${p.type} is not a sparkline type (${SPARKLINE_TYPES.join(", ")})`);
      if (!sparklineOptions(options)) refuse(`${p.options} is not sparkline options (high, low, first, last, axis)`);
      const extra = type === "line" && !options ? "" : options ? `, "${type}", "${options}"` : `, "${type}"`;
      const formula = wb.sheet.language.name === "Excel" ? `=SPARKLINE(${named(r)}${extra})` : `=SPARKLINE_RANGE("${named(r)}"${extra})`;
      const changed = await wb.sheet.set(key, formula);
      return changedBy(wb, changed, `${key} = ${show(wb.sheet.value(key))}`);
    },
  },
  {
    id: "data.goalSeek", label: "Goal Seek", params: ["set", "to", "by"], journaled: true,
    run: async (wb, p) => {
      const target = cellOf(p.set), by = cellOf(p.by), to = numberOf(p.to);
      if (!wb.sheet.input(target).startsWith("=")) refuse(`${target} is not a formula`);
      if (wb.sheet.input(by).startsWith("=")) refuse(`${by} holds a formula, not a value to change`);
      const start = Number(wb.sheet.input(by)) || 0, was = wb.sheet.input(by);
      let found: { x: number } | null = null;
      try {
        found = await secant(async (x) => {
          await wb.sheet.setMany([[by, String(x)]]);
          const v = wb.sheet.value(target);
          return typeof v === "number" ? v - to : NaN;
        }, start);
      } finally {
        await wb.sheet.setMany([[by, was]]);
      }
      if (!found) return refuse(`${target} cannot reach ${p.to} by changing ${by}`);
      const changed = await wb.sheet.set(by, String(Number(found.x.toPrecision(12))));
      return changedBy(wb, changed, `${target} = ${show(wb.sheet.value(target))} with ${by} = ${show(wb.sheet.value(by))}`);
    },
  },
  {
    id: "data.table", label: "Data Table", params: ["range"], optional: ["row", "column"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), tl = topLeft(r);
      const rows = Math.abs(r[0].row - r[1].row) + 1, cols = Math.abs(r[0].col - r[1].col) + 1;
      const at = (i: number, j: number) => keyOf({ col: tl.col + j, row: tl.row + i });
      const formula = (k: string) => wb.sheet.input(k).startsWith("=");
      const row = p.row === undefined ? null : cellOf(p.row), column = p.column === undefined ? null : cellOf(p.column);
      const entries: [string, string][] = [];
      if (row && column) {
        // Two inputs: the formula at the corner, the row's values across the
        // top, the column's down the left.
        if (!formula(at(0, 0))) refuse(`${at(0, 0)} has no formula`);
        const pairs = Array.from({ length: (rows - 1) * (cols - 1) }, (_, n) => [Math.floor(n / (cols - 1)) + 1, (n % (cols - 1)) + 1]);
        const got = await tried(wb, [row, column], pairs.map(([i, j]) => [wb.sheet.input(at(0, j)), wb.sheet.input(at(i, 0))]), () => inputOf(wb.sheet.value(at(0, 0))));
        pairs.forEach(([i, j], n) => entries.push([at(i, j), got[n]]));
      } else if (row || column) {
        // One input: its values along one edge, formulas along the other.
        const down = !!column, input = (column ?? row)!;
        const formulas = Array.from({ length: down ? cols - 1 : rows - 1 }, (_, n) => (down ? at(0, n + 1) : at(n + 1, 0))).filter(formula);
        if (!formulas.length) refuse(`${named(r)} has no formula in its ${down ? "top row" : "left column"}`);
        const values = Array.from({ length: down ? rows - 1 : cols - 1 }, (_, n) => (down ? at(n + 1, 0) : at(0, n + 1)));
        const got = await tried(wb, [input], values.map((k) => [wb.sheet.input(k)]), () => formulas.map((f) => inputOf(wb.sheet.value(f))));
        values.forEach((_, n) => formulas.forEach((f, m) => {
          const a = parseRange(f)![0];
          entries.push([down ? keyOf({ col: a.col, row: tl.row + n + 1 }) : keyOf({ col: tl.col + n + 1, row: a.row }), got[n][m]]);
        }));
      } else refuse("a data table needs a row input, a column input, or both");
      const changed = await wb.sheet.setMany(entries);
      return changedBy(wb, changed, `filled the data table ${named(r)}: ${entries.length} results`);
    },
  },
  {
    id: "data.removeDuplicates", label: "Remove Duplicates", params: ["range"], optional: ["columns", "headers"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range), tl = topLeft(r);
      const first = flag(p.headers) ? 1 : 0;
      const rows = Math.abs(r[0].row - r[1].row) + 1, cols = Math.abs(r[0].col - r[1].col) + 1;
      const at = (i: number, j: number) => keyOf({ col: tl.col + j, row: tl.row + i });
      const by = p.columns === undefined ? Array.from({ length: cols }, (_, j) => j)
        : p.columns.split(",").map((c) => {
          const j = columnOf(c) - tl.col;
          return j >= 0 && j < cols ? j : refuse(`${c.trim()} is not a column of ${named(r)}`);
        });
      const body = Array.from({ length: rows - first }, (_, i) => Array.from({ length: cols }, (_, j) => show(wb.sheet.value(at(i + first, j)))));
      const kept = distinctRows(body, by);
      const entries: [string, string][] = [];
      // The rows kept move up, their formulas with them; the rest are cleared.
      for (let i = 0; i < rows - first; i++) {
        for (let j = 0; j < cols; j++) {
          const from = kept[i];
          entries.push([at(i + first, j), from === undefined ? "" : movedBy(wb.sheet.language, wb.sheet.input(at(from + first, j)), { cols: 0, rows: i - from })]);
        }
      }
      const gone = body.length - kept.length;
      const changed = await wb.sheet.setMany(entries);
      return changedBy(wb, changed, `removed ${gone} duplicate row${gone === 1 ? "" : "s"} of ${named(r)}; ${kept.length} left`);
    },
  },
  {
    id: "data.textToColumns", label: "Text to Columns", params: ["range"], optional: ["delimiter", "replace"], journaled: true,
    run: async (wb, p) => {
      const r = rangeOf(p.range);
      if (r[0].col !== r[1].col) refuse("Text to Columns splits one column");
      const d = p.delimiter === undefined ? "," : DELIMITERS[p.delimiter.toLowerCase()] ?? (p.delimiter.length === 1 ? p.delimiter : refuse(`${p.delimiter} is not a delimiter (one character, or comma, tab, space, semicolon)`));
      const keys = [...range(r[0], r[1])];
      const split = keys.map((k) => splitText(show(wb.sheet.value(k)), d));
      const width = Math.max(...split.map((s) => s.length));
      const col = r[0].col;
      if (col + width > COLS) refuse(`${width} columns from ${colName(col)} do not fit on the sheet`);
      const entries: [string, string][] = [];
      split.forEach((fields, i) => {
        const row = parseRange(keys[i])![0].row;
        fields.forEach((field, j) => {
          const key = keyOf({ col: col + j, row });
          if (j > 0 && !flag(p.replace) && wb.sheet.input(key) !== "") refuse(`${key} is not empty`);
          entries.push([key, field]);
        });
      });
      const changed = await wb.sheet.setMany(entries);
      return changedBy(wb, changed, `split ${rangeName(r)} into ${width} columns`);
    },
  },
  {
    // A formula worked out against the sheet without a cell to hold it: the
    // app's EVALUATE.
    id: "formulas.calculate", label: "Calculate", params: ["formula"], journaled: false,
    run: async (wb, p) => {
      const text = p.formula.trim().replace(/^=/, "");
      const v = await wb.sheet.evaluate(text);
      return { ok: !(v instanceof CellError), message: `=${text} → ${show(v)}${v instanceof CellError && v.why ? ` (${v.why})` : ""}`, changed: [], output: scripted(v) };
    },
  },
  {
    // Every sheet compiled and worked out again; how many formulas it took.
    id: "formulas.recalculate", label: "Calculate Now", params: [], journaled: false,
    run: async (wb) => {
      const changed = await recalculated(wb);
      const n = wb.pages().reduce((t, page) => t + page.sheet.keys().filter((k) => page.sheet.input(k).startsWith("=")).length, 0);
      return { ...await changedBy(wb, changed, `${n} formula${n === 1 ? "" : "s"} worked out`), output: n };
    },
  },
  {
    // The 1904 date system (Excel's "Use 1904 date system"), or 1900: the
    // serials stay, so a date typed as a number shows four years and a day
    // on; what DATE and the other date functions give follows it.
    id: "workbook.dateSystem", label: "Date System", params: ["system"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      const sys = Number(p.system);
      if (sys !== 1900 && sys !== 1904) refuse(`${p.system} is not a date system (1900, 1904)`);
      await wb.setDateSystem(sys as 1900 | 1904);
      return changedBy(wb, wb.sheet.keys(), `the ${sys} date system`);
    },
  },
  {
    // Zoom: the sheet drawn at a percentage, 10 to 400, as Excel's; or in
    // or out a step (Excel's 10, 25, 50, 75, 100, 125, 150, 200, 400). A
    // view, not a step to undo, not saved.
    id: "view.zoom", label: "Zoom", params: [], optional: ["percent", "by"], journaled: false,
    run: async (wb, p) => {
      const steps = [10, 25, 50, 75, 100, 125, 150, 200, 400];
      let z = wb.zoom;
      if (p.percent !== undefined) {
        z = Number(String(p.percent).replace(/%$/, ""));
        if (!Number.isInteger(z) || z < 10 || z > 400) refuse(`${p.percent} is not a zoom (10 to 400)`);
      } else if (p.by === "in") z = steps.find((s) => s > wb.zoom) ?? 400;
      else if (p.by === "out") z = [...steps].reverse().find((s) => s < wb.zoom) ?? 10;
      else refuse("view.zoom needs percent= or by=in|out");
      wb.zoom = z;
      return { ok: true, message: `zoom ${z}%`, changed: wb.sheet.keys(), output: z };
    },
  },
  {
    id: "view.selection", label: "Selection", params: [], journaled: false,
    run: async (wb) => ({ ok: true, message: wb.selection, changed: [], output: wb.selection }),
  },
  {
    // Select a range: a cell, a rectangle, whole columns ("B:B") or rows
    // ("3:5"), or all ("all"); the page selects it. A view, not a step.
    id: "view.select", label: "Select", params: ["range"], journaled: false,
    run: async (wb, p) => {
      const rs = /^\s*all\s*$/i.test(p.range) ? [[{ col: 0, row: 0 }, { col: COLS - 1, row: ROWS - 1 }] as [Addr, Addr]] : p.range.split(",").map((a) => rangeOf(a.trim()));
      wb.selection = rs.map(rangeName).join(",");
      wb.selecting = true;
      return { ok: true, message: `selected ${wb.selection}`, changed: [], output: wb.selection };
    },
  },
  {
    // Go To Special: the cells of a kind, in sheet order, over a range or
    // the used range (A1 to Ctrl+End's cell): formulas, constants (either
    // narrowed by type=numbers|text|logicals|errors), blanks, errors,
    // comments, validation, conditional formats.
    id: "edit.goToSpecial", label: "Go To Special", params: ["what"], optional: ["range", "type"], journaled: false,
    run: async (wb, p) => {
      const kinds = ["formulas", "constants", "blanks", "errors", "comments", "validation", "conditional"];
      const what = p.what.trim().toLowerCase().replace(/^conditional formats?$/, "conditional");
      if (!kinds.includes(what)) refuse(`${p.what} is not one of ${kinds.join(", ")}`);
      const types = ["numbers", "text", "logicals", "errors"], type = p.type?.trim().toLowerCase();
      if (type !== undefined && !types.includes(type)) refuse(`${p.type} is not a type (${types.join(", ")})`);
      const keys = wb.sheet.keys();
      const r: [Addr, Addr] = p.range ? rangeOf(p.range) : [{ col: 0, row: 0 }, { col: Math.max(0, ...keys.map((k) => cellAddr(k)!.col)), row: Math.max(0, ...keys.map((k) => cellAddr(k)!.row)) }];
      const page = wb.page(wb.active)!;
      const ofType = (v: Value) => !type || (type === "numbers" ? typeof v === "number" : type === "text" ? typeof v === "string" : type === "logicals" ? typeof v === "boolean" : v instanceof CellError);
      const is = (k: string): boolean => {
        const input = wb.sheet.input(k), v = wb.sheet.value(k), n = wb.notes.get(k), a = cellAddr(k)!;
        switch (what) {
          case "formulas": return input.startsWith("=") && ofType(v);
          case "constants": return input !== "" && !input.startsWith("=") && ofType(v);
          case "blanks": return input === "" && (v === null || v === "");
          case "errors": return v instanceof CellError;
          case "comments": return !!n.comment;
          case "validation": return !!n.validation;
          default: return page.rules.some((rule) => a.col >= Math.min(rule.range[0].col, rule.range[1].col) && a.col <= Math.max(rule.range[0].col, rule.range[1].col)
            && a.row >= Math.min(rule.range[0].row, rule.range[1].row) && a.row <= Math.max(rule.range[0].row, rule.range[1].row));
        }
      };
      const found = [...range(r[0], r[1])].filter(is).sort((x, y) => order(cellAddr(x)!) - order(cellAddr(y)!));
      return { ok: true, message: found.length ? `${found.length} ${what}: ${found.join(", ")}` : `no ${what} in ${rangeName(r)}`, changed: [], output: found };
    },
  },
  {
    // Ctrl+End's cell: the last row used and the last column used.
    id: "sheet.lastCell", label: "Last Cell", params: [], journaled: false,
    run: async (wb) => {
      const at = wb.sheet.keys().map((k) => cellAddr(k)!);
      const last = keyOf({ col: Math.max(0, ...at.map((a) => a.col)), row: Math.max(0, ...at.map((a) => a.row)) });
      return { ok: true, message: last, changed: [], output: last };
    },
  },
  {
    // The keyboard map (keys.ts), for a script to read.
    id: "keys.list", label: "Keyboard Shortcuts", params: [], journaled: false,
    run: async () => ({ ok: true, message: `${KEYS.length} keys`, changed: [], output: KEYS.map((k) => `${k.combo} | ${k.does} | ${k.id}`) }),
  },
  {
    // The status bar's figures for a range: average, count, numerical
    // count, min, max, sum.
    id: "range.stats", label: "Status Bar figures", params: ["range"], journaled: false,
    run: async (wb, p) => {
      const st = wb.stats(p.range.split(",").map((a) => rangeOf(a.trim())));
      return { ok: true, message: statsLine(st), changed: [], output: st };
    },
  },
  {
    // The panel open under the grid, one at a time (Edit, Looks ... Lists:
    // menus.ts's PANELS), or none. A view, not a step.
    id: "view.panel", label: "Show Panel", params: ["name"], journaled: false,
    run: async (wb, p) => {
      const n = p.name.trim().toLowerCase();
      const panel = n === "none" ? null : PANELS.find((x) => x.name === n || x.label.toLowerCase() === n)?.name ?? refuse(`${p.name} is not a panel (${PANELS.map((x) => x.name).join(", ")}, none)`);
      wb.panel = panel;
      return { ok: true, message: panel ? `the ${panel} panel` : "no panel", changed: [], output: panel ?? "" };
    },
  },
  {
    // F4 on a formula being typed: the reference at the caret (the end by
    // default) to its next form, A1, $A$1, A$1, $A1.
    id: "formulas.cycleRef", label: "Cycle Reference", params: ["text"], optional: ["caret"], journaled: false,
    run: async (_wb, p) => {
      const r = cycleRef(p.text, p.caret === undefined ? p.text.length : Number(p.caret));
      return { ok: true, message: r.text, changed: [], output: r.text };
    },
  },
  {
    // The function names that complete a formula typed so far, and the
    // call its end is in (Formula AutoComplete and the argument hint).
    id: "formulas.complete", label: "Formula AutoComplete", params: ["text"], journaled: false,
    run: async (wb, p) => {
      const c = complete(p.text, wb.sheet.language.functions());
      const hint = c.call ? `${c.call}: argument ${c.arg}` : "";
      return { ok: true, message: [c.names.join(", ") || "no names", hint].filter((x) => x).join("; "), changed: [], output: c };
    },
  },
  {
    // The menu bar's items, each "File > Save | file.save | Ctrl+S": for a
    // script to find a command by where it is on the menus.
    id: "menus.list", label: "Menus", params: [], journaled: false,
    run: async () => {
      const lines = menuLines();
      return { ok: true, message: `${lines.length} menu items`, changed: [], output: lines };
    },
  },
  {
    // Error Checking over the active sheet: formulas come to an error,
    // inconsistent with those beside them, omitting a number next to their
    // range, and numbers stored as text (errorcheck.ts); a cell ignored is
    // passed over.
    id: "formulas.errorCheck", label: "Error Checking", params: [], journaled: false,
    run: async (wb) => {
      const found = errorCheck(wb.sheet, new Set(wb.notes.keys().filter((k) => wb.notes.get(k).ignoreError)));
      const message = found.length ? `${found.length} to check: ${found.map((f) => `${f.cell} ${f.check}`).join(", ")}` : "no errors found";
      return { ok: true, message, changed: [], output: found };
    },
  },
  {
    // Ignore Error: Error Checking passes over the cell from now on; off,
    // it checks it again.
    id: "formulas.ignoreError", label: "Ignore Error", params: ["cell"], optional: ["off"], journaled: true,
    run: async (wb, p) => {
      const key = cellOf(p.cell);
      wb.notes.set(key, { ignoreError: flag(p.off) ? undefined : true });
      return changedBy(wb, [key], `${key}: ${flag(p.off) ? "checked again" : "error ignored"}`);
    },
  },
  {
    // Calculate Sheet (Shift+F9): the active sheet worked out again; under
    // manual calculation, what reads it on other sheets is stale then.
    id: "formulas.calculateSheet", label: "Calculate Sheet", params: [], journaled: false,
    run: async (wb) => {
      const changed = await wb.sheet.recalculate();
      const n = wb.sheet.keys().filter((k) => wb.sheet.input(k).startsWith("=")).length;
      return { ...await changedBy(wb, changed, `${wb.active}: ${n} formula${n === 1 ? "" : "s"} worked out`), output: n };
    },
  },
  {
    // Calculation Options: automatic, or manual (a change works out only
    // the cells set; Calculate Now or Sheet the rest). Back to automatic,
    // everything is worked out. Saved with the file, not a step to undo.
    id: "formulas.calculation", label: "Calculation Options", params: ["mode"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      const manual = ({ automatic: false, manual: true } as Record<string, boolean>)[p.mode.toLowerCase()] ?? refuse(`${p.mode} is not a calculation mode (automatic, manual)`);
      wb.calculation = { ...wb.calculation, manual };
      const changed = manual ? [] : await recalculated(wb);
      return changedBy(wb, changed, `${manual ? "manual" : "automatic"} calculation`);
    },
  },
  {
    // Enable iterative calculation: a circular reference worked round, up
    // to `max` times (100) or until no cell of it moves by more than
    // `change` (0.001), instead of #CYCLE!; off, it is #CYCLE! again.
    id: "formulas.iteration", label: "Iterative Calculation", params: [], optional: ["off", "max", "change"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      const max = Number(p.max ?? wb.calculation.iterate?.max ?? 100), change = Number(p.change ?? wb.calculation.iterate?.change ?? 0.001);
      if (!flag(p.off) && !(Number.isInteger(max) && max >= 1 && max <= 32767)) refuse(`${p.max} is not a number of iterations (1 to 32767)`);
      if (!flag(p.off) && !(change >= 0)) refuse(`${p.change} is not a maximum change (0 or more)`);
      wb.calculation = { ...wb.calculation, iterate: flag(p.off) ? null : { max, change } };
      const changed = wb.calculation.manual ? [] : await recalculated(wb);
      return changedBy(wb, changed, flag(p.off) ? "no iteration" : `iteration: at most ${max} times, to a change of ${change}`);
    },
  },
  {
    // The setup script (RexxJS: what it LETs, every formula reads), as the
    // file keeps it; set, and run (everything is then worked out again).
    id: "setup.get", label: "Setup script", params: [], journaled: false,
    run: async (wb) => ({ ok: true, message: wb.setupScript || "no setup script", changed: [], output: wb.setupScript }),
  },
  {
    id: "setup.set", label: "Set the setup script", params: ["text"], journaled: true,
    run: async (wb, p) => {
      wb.setupScript = p.text;
      return { ok: true, message: `setup script: ${p.text.split("\n").length} line${p.text.split("\n").length === 1 ? "" : "s"}`, changed: [] };
    },
  },
  {
    id: "setup.run", label: "Run the setup script", params: [], journaled: true, unstepped: true,
    run: async (wb) => {
      const failed = await wb.rexx.setup(wb.setupScript);
      if (failed) refuse(`the setup script failed: ${failed.split("\n").pop()}`);
      const changed = await recalculated(wb);
      return changedBy(wb, changed, "ran the setup script");
    },
  },
  {
    id: "formulas.use", label: "Formula language", params: ["language"], journaled: true,
    run: async (wb, p) => {
      const name = ({ excel: "Excel", rexxjs: "RexxJS" } as Record<string, LanguageName>)[p.language.toLowerCase()]
        ?? refuse(`${p.language} is not a formula language (Excel, RexxJS)`);
      const changed = await wb.setLanguage(wb.language(name));
      return changedBy(wb, changed, `${name} formulas`);
    },
  },
  {
    // How Excel formulas are shown and typed on the page: A1 or R1C1. A
    // view of them: what is stored, and what cell.set takes, stays A1.
    id: "formulas.style", label: "Reference style", params: ["style"], journaled: false,
    run: async (wb, p) => {
      const style = ({ a1: "A1", r1c1: "R1C1" } as Record<string, "A1" | "R1C1">)[p.style.toLowerCase()] ?? refuse(`${p.style} is not a reference style (A1, R1C1)`);
      wb.referenceStyle = style;
      return changedBy(wb, wb.sheet.keys(), `${style} references`);
    },
  },
  {
    id: "sheet.add", label: "Insert Sheet", params: [], optional: ["name"], journaled: true,
    run: async (wb, p) => {
      let name = p.name;
      if (name === undefined) for (let i = wb.pages().length + 1; !name || wb.page(name); i++) name = `Sheet${i}`;
      name = sheetNameOf(name);
      if (wb.page(name)) refuse(`${name} is a sheet already`);
      const before = wb.sheet.keys();
      await wb.add(name);
      return changedBy(wb, shownKeys(wb, before), `added ${name}`);
    },
  },
  {
    // All a script may ask of the active sheet, rows counted from 1 and
    // columns by letter (control.ts's ISROWHIDDEN, GETSHEETNAMES, CANUNDO ...).
    id: "sheet.state", label: "About the sheet", params: [], journaled: false,
    run: async (wb) => {
      const hidden = wb.hidden(), page = wb.page(wb.active)!;
      const state = {
        name: wb.active, sheets: wb.sheetNames(), hiddenSheets: wb.pages().filter((x) => x.hiddenSheet).map((x) => x.name),
        tabColors: Object.fromEntries(wb.pages().filter((x) => x.tabColor).map((x) => [x.name, x.tabColor])), hiddenRows: hidden.rows.map((r) => r + 1), hiddenCols: hidden.cols.map(colName),
        frozenRows: wb.frozenRows(), frozenCols: wb.frozenCols(), rowHeights: wb.rowHeights().map(([r, h]) => [r + 1, h]), colWidths: wb.colWidths().map(([c, w]) => [colName(c), w]),
        filter: page.filter && { range: rangeName(page.filter.range), column: colName(page.filter.column), criterion: page.filter.criterion },
        canUndo: wb.canUndo ? 1 : 0, canRedo: wb.canRedo ? 1 : 0,
        autoId: page.autoId && { column: colName(page.autoId.column), next: `${page.autoId.prefix}${page.autoId.next}` },
        language: wb.sheet.language.name, referenceStyle: wb.referenceStyle,
        panel: wb.panel ?? "", split: wb.split ? 1 : 0, zoom: wb.zoom, dateSystem: wb.dateSystem, calculation: wb.calculation.manual ? "Manual" : "Automatic", iteration: wb.calculation.iterate, stale: wb.stale(),
      };
      return { ok: true, message: JSON.stringify(state), changed: [], output: state };
    },
  },
  {
    id: "sheet.activate", label: "Go to a sheet", params: ["name"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      const page = pageOf(wb, p.name), before = wb.sheet.keys();
      wb.activate(page.name);
      return changedBy(wb, shownKeys(wb, before), page.name);
    },
  },
  {
    // What read the sheet by its old name reads it by the new one.
    id: "sheet.rename", label: "Rename Sheet", params: ["name", "to"], journaled: true,
    run: async (wb, p) => {
      const page = pageOf(wb, p.name), to = sheetNameOf(p.to);
      if (wb.page(to) && wb.page(to) !== page) refuse(`${to} is a sheet already`);
      const was = page.name, before = wb.sheet.keys();
      await wb.rename(was, to);
      return changedBy(wb, shownKeys(wb, before), `${was} is now ${to}`);
    },
  },
  {
    // A sheet moved among the tabs: to=N (1 is first), or way=left/right.
    id: "sheet.move", label: "Move Sheet", params: ["name"], optional: ["to", "way"], journaled: true,
    run: async (wb, p) => {
      const page = pageOf(wb, p.name), from = wb.pages().indexOf(page), last = wb.pages().length - 1;
      if ((p.to === undefined) === (p.way === undefined)) refuse("sheet.move needs to or way");
      const at = p.to !== undefined ? countOf(p.to) - 1
        : /^left$/i.test(p.way!) ? from - 1 : /^right$/i.test(p.way!) ? from + 1 : refuse(`${p.way} is not left or right`);
      if (at < 0 || at > last) refuse(`${page.name} cannot move ${p.to !== undefined ? `to ${p.to}` : p.way}: there are ${last + 1} sheets`);
      await wb.move(page.name, at);
      return changedBy(wb, [], `${page.name} is sheet ${at + 1} of ${last + 1}`);
    },
  },
  {
    // A sheet's tab colour: a colour's name or #rrggbb; none takes it off.
    id: "sheet.color", label: "Tab Color", params: ["name", "color"], journaled: true,
    run: async (wb, p) => {
      const page = pageOf(wb, p.name), off = /^(none|auto|automatic)$/i.test(p.color.trim());
      page.tabColor = off ? null : resolvedColour(colour(p.color)) ?? refuse(`${p.color} is not a colour (a name such as red, or #rrggbb)`);
      return changedBy(wb, [], `${page.name}'s tab ${page.tabColor ?? "uncoloured"}`);
    },
  },
  {
    // Hide a sheet (never the last one shown; the active one hidden, the
    // next one shown is made active), or unhide it.
    id: "sheet.hide", label: "Hide Sheet", params: ["name"], optional: ["off"], journaled: true,
    run: async (wb, p) => {
      const page = pageOf(wb, p.name), hide = !flag(p.off), before = wb.sheet.keys();
      if (hide && !page.hiddenSheet && wb.pages().filter((x) => !x.hiddenSheet).length === 1) refuse(`${page.name} is the only sheet shown`);
      page.hiddenSheet = hide;
      if (hide && wb.active === page.name) {
        const pages = wb.pages(), i = pages.indexOf(page);
        wb.activate((pages.slice(i + 1).find((x) => !x.hiddenSheet) ?? [...pages.slice(0, i)].reverse().find((x) => !x.hiddenSheet))!.name);
      }
      return changedBy(wb, shownKeys(wb, before), `${page.name} ${hide ? "hidden" : "shown"}`);
    },
  },
  {
    // Move or Copy's "create a copy": the sheet with its cells and looks,
    // after it, named "Name copy" unless named (a sheet name has no
    // brackets here, so not Excel's "Name (2)"); the copy made active.
    id: "sheet.copy", label: "Copy Sheet", params: ["name"], optional: ["to"], journaled: true,
    run: async (wb, p) => {
      const page = pageOf(wb, p.name), before = wb.sheet.keys();
      let to = p.to;
      if (to === undefined) for (let i = 1; !to || wb.page(to); i++) to = i === 1 ? `${page.name} copy` : `${page.name} copy ${i}`;
      to = sheetNameOf(to);
      if (wb.page(to)) refuse(`${to} is a sheet already`);
      await wb.copy(page, to);
      return changedBy(wb, shownKeys(wb, before), `copied ${page.name} to ${to}`);
    },
  },
  {
    id: "sheet.delete", label: "Delete Sheet", params: ["name"], journaled: true,
    run: async (wb, p) => {
      const page = pageOf(wb, p.name);
      if (wb.pages().length === 1) refuse(`${page.name} is the only sheet`);
      const before = wb.sheet.keys();
      await wb.remove(page.name);
      return changedBy(wb, shownKeys(wb, before), `deleted ${page.name}`);
    },
  },
  {
    id: "file.open", label: "Open", params: [], optional: ["name", "sample"], journaled: true, unstepped: true,
    run: (wb, p, store) => {
      if (p.name !== undefined) return open(wb, readSheetFile(JSON.parse(store.get(fileKey(savedName(store, p.name)))!)));
      if (p.sample !== undefined) return open(wb, readSheetFile(SAMPLES[p.sample.toLowerCase()] ?? refuse(`no sample ${p.sample} (budget, test, formatted)`)));
      return refuse("file.open needs a name or a sample");
    },
  },
  {
    // The workbook as the file's JSON (what file.save keeps), and a file's
    // JSON opened: the app's EXPORT and IMPORT.
    id: "file.export", label: "Export the file", params: [], journaled: false,
    run: async (wb) => {
      const text = JSON.stringify(writeSheetFile(wb));
      return { ok: true, message: `${text.length} characters`, changed: [], output: text };
    },
  },
  {
    id: "file.import", label: "Import a file", params: ["text"], journaled: true, unstepped: true,
    run: (wb, p) => {
      let data: unknown;
      try {
        data = JSON.parse(p.text);
      } catch (e) {
        return refuse(`not a sheet file: ${(e as Error).message}`);
      }
      return open(wb, readSheetFile(data as SheetFile));
    },
  },
  {
    id: "file.save", label: "Save", params: [], optional: ["name"], journaled: false,
    run: async (wb, p, store) => {
      const name = (p.name ?? wb.name).trim() || refuse("a file needs a name");
      const file = writeSheetFile(wb);
      store.set(fileKey(name), JSON.stringify({ ...file, name }));
      const names = savedNames(store).filter((n) => n.toUpperCase() !== name.toUpperCase());
      store.set(FILES, JSON.stringify([...names, name]));
      wb.name = name;
      const cells = wb.pages().reduce((n, page) => n + page.sheet.keys().length, 0);
      return { ok: true, message: `saved ${name}: ${wb.pages().length} sheet${wb.pages().length === 1 ? "" : "s"}, ${cells} cells`, changed: [] };
    },
  },
  {
    id: "file.list", label: "Saved files", params: [], journaled: false,
    run: async (_wb, _p, store) => {
      const names = savedNames(store);
      return { ok: true, message: names.length ? `saved: ${names.join(", ")}` : "no saved files", changed: [], output: names };
    },
  },
  {
    id: "file.delete", label: "Delete a saved file", params: ["name"], journaled: false,
    run: async (_wb, p, store) => {
      const name = savedName(store, p.name);
      store.remove(fileKey(name));
      store.set(FILES, JSON.stringify(savedNames(store).filter((n) => n !== name)));
      return { ok: true, message: `deleted ${name}`, changed: [] };
    },
  },
  {
    id: "file.new", label: "New", params: [], journaled: true, unstepped: true,
    run: async (wb) => {
      const before = wb.sheet.keys();
      wb.reset();
      wb.name = "Untitled";
      wb.setupScript = "";
      await wb.settle();
      return { ok: true, message: "a new workbook", changed: before };
    },
  },
  {
    // What cells show, as CSV: delimiter= comma (the default), tab (TSV),
    // semicolon or bar, a field holding it quoted.
    id: "file.exportCsv", label: "Export CSV", params: [], optional: ["range", "delimiter"], journaled: false,
    run: async (wb, p) => {
      const r = p.range === undefined ? usedRange(wb) : rangeOf(p.range), tl = topLeft(r);
      const d = delimiterOf(p.delimiter ?? "comma");
      const rows = Math.abs(r[0].row - r[1].row) + 1, cols = Math.abs(r[0].col - r[1].col) + 1;
      const csv = wb.sheet.keys().length || p.range !== undefined ? toCsv(Array.from({ length: rows }, (_, i) =>
        Array.from({ length: cols }, (_, j) => wb.textOf(keyOf({ col: tl.col + j, row: tl.row + i })))), d) : "";
      return { ok: true, message: `${named(r)} as ${d === "\t" ? "TSV" : "CSV"}: ${csv ? rows : 0} rows`, changed: [], output: csv };
    },
  },
  {
    // CSV or TSV in at a cell: text=, or base64= bytes in encoding= (auto:
    // a byte-order mark, else UTF-8, else Windows-1252); delimiter= comma,
    // tab, semicolon, bar, space or auto (the default: detected).
    id: "file.importCsv", label: "Import CSV", params: ["at"], optional: ["text", "base64", "encoding", "delimiter"], journaled: true,
    run: (wb, p) => {
      if ((p.text === undefined) === (p.base64 === undefined)) refuse("Import CSV takes text= or base64=");
      let text: string;
      try {
        text = p.text ?? decode(fromBase64(p.base64!), p.encoding ?? "auto");
      } catch (e) {
        return refuse((e as Error).message);
      }
      const d = (p.delimiter ?? "auto").trim().toLowerCase() === "auto" ? detectDelimiter(text) : delimiterOf(p.delimiter!);
      const at = rangeOf(p.at)[0], rows = fromCsv(text, d);
      const cols = Math.max(0, ...rows.map((r) => r.length));
      if (at.col + cols > COLS) refuse(`${cols} columns from ${keyOf(at)} do not fit on the sheet`);
      if (at.row + rows.length > ROWS) refuse(`${rows.length} rows from ${keyOf(at)} do not fit on the sheet`);
      const entries = rows.flatMap((row, i) => row.map((input, j): [string, string] => [keyOf({ col: at.col + j, row: at.row + i }), input]));
      return setting(wb, entries, `imported ${rows.length} rows at ${keyOf(at)}${d === "," ? "" : `, split by ${Object.entries(CSV_DELIMITERS).find(([, x]) => x === d)?.[0] ?? d}`}`);
    },
  },
  {
    // An XLSX file opened (base64= its bytes): its sheets, cells, formulas,
    // number formats, merges, sizes, frozen panes and names (xlsx.ts), as
    // a workbook opened is; what it held that the app does not, reported.
    // An encrypted one (Agile or Standard AES) is opened by its password=.
    id: "file.importXlsx", label: "Open XLSX", params: ["base64"], optional: ["name", "password"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      let read: XlsxRead;
      try {
        let bytes = fromBase64(p.base64);
        if (isEncrypted(bytes)) {
          if (!p.password) refuse(`${p.name ?? "the file"} is encrypted: open it with its password`);
          bytes = decryptPackage(bytes, p.password!);
        }
        read = readXlsx(bytes, p.name ?? "Workbook");
      } catch (e) {
        if (e instanceof Refused) throw e;
        return refuse((e as Error).message);
      }
      wb.fileReport = read.report;
      const o = await open(wb, readSheetFile(read.file));
      wb.carried = read.kept;
      return { ...o, message: `${o.message}${read.report.length ? `; left out: ${read.report.join("; ")}` : ""}`, output: read.report };
    },
  },
  {
    // An ODS file opened (base64= its bytes): its tables, their cells'
    // cached values (ods.ts); what was left out, reported.
    id: "file.importOds", label: "Open ODS", params: ["base64"], optional: ["name"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      let read: { file: Record<string, unknown>; report: string[] };
      try {
        read = readOds(fromBase64(p.base64), p.name ?? "Workbook");
      } catch (e) {
        return refuse((e as Error).message);
      }
      wb.fileReport = read.report;
      const o = await open(wb, readSheetFile(read.file));
      wb.carried = [];
      return { ...o, message: `${o.message}${read.report.length ? `; ${read.report.join("; ")}` : ""}`, output: read.report };
    },
  },
  {
    // An XLS file (Excel 97 to 2003, BIFF8) opened (base64= its bytes):
    // its sheets, cells and formulas (xls.ts); what was left out, reported.
    id: "file.importXls", label: "Open XLS", params: ["base64"], optional: ["name"], journaled: true, unstepped: true,
    run: async (wb, p) => {
      let read: { file: Record<string, unknown>; report: string[] };
      try {
        read = readXls(fromBase64(p.base64), p.name ?? "Workbook");
      } catch (e) {
        return refuse((e as Error).message);
      }
      wb.fileReport = read.report;
      const o = await open(wb, readSheetFile(read.file));
      wb.carried = [];
      return { ...o, message: `${o.message}${read.report.length ? `; ${read.report.join("; ")}` : ""}`, output: read.report };
    },
  },
  {
    // The sheet (or every sheet shown, sheets=all) as a PDF (pdf.ts): its
    // used range or range=, on paper= a4 or letter, landscape=1, with
    // gridlines=1 and headings=1 when asked; the bytes as base64.
    id: "file.exportPdf", label: "Export PDF", params: [], optional: ["range", "sheets", "paper", "landscape", "gridlines", "headings", "margin"], journaled: false,
    run: async (wb, p) => {
      const paper = (p.paper ?? "a4").trim().toLowerCase();
      if (!(paper in PAPER)) refuse(`${p.paper} is not a paper size (${Object.keys(PAPER).join(", ")})`);
      const margin = p.margin === undefined ? 36 : Number(p.margin);
      if (!(margin >= 0 && margin <= 144)) refuse(`${p.margin} is not a margin (0 to 144 points)`);
      const all = (p.sheets ?? "").trim().toLowerCase() === "all";
      if (all && p.range !== undefined) refuse("a range is the active sheet's: not with sheets=all");
      const pages = all ? wb.pages().filter((x) => !x.hiddenSheet) : [wb.page(wb.active)!];
      const sheets = pages.map((page) => pdfSheet(wb, page, p.range === undefined ? null : rangeOf(p.range)));
      const { bytes, pages: n } = writePdf(sheets, { paper: paper as "a4" | "letter", landscape: flag(p.landscape), margin, gridlines: flag(p.gridlines), headings: flag(p.headings) });
      return { ok: true, message: `${wb.name}.pdf: ${n} page${n === 1 ? "" : "s"}, ${bytes.length} bytes`, changed: [], output: toBase64(bytes) };
    },
  },
  {
    // The workbook as an XLSX file (xlsxwrite.ts): its bytes as base64 in
    // the output; what it could not write, reported.
    id: "file.exportXlsx", label: "Save as XLSX", params: [], journaled: false,
    run: async (wb) => {
      const { bytes, report } = writeXlsx(wb);
      wb.fileReport = report;
      return { ok: true, message: `${wb.name}.xlsx: ${bytes.length} bytes, ${wb.pages().length} sheet${wb.pages().length === 1 ? "" : "s"}${report.length ? `; ${report.join("; ")}` : ""}`, changed: [], output: toBase64(bytes) };
    },
  },
  {
    // What the last file opened held that the app left out (FILEREPORT).
    id: "file.report", label: "File Report", params: [], journaled: false,
    run: async (wb) => ({ ok: true, message: wb.fileReport.length ? wb.fileReport.join("; ") : "nothing left out", changed: [], output: wb.fileReport }),
  },
  {
    // JSON rows in at a cell: an array of objects (their keys the header
    // row), or of arrays.
    id: "file.importJson", label: "Import JSON", params: ["at", "text"], journaled: true,
    run: (wb, p) => placed(wb, rangeOf(p.at)[0], fromJson(p.text), "JSON"),
  },
  {
    // An HTML page's table (table= its number, the first by default) in at
    // a cell.
    id: "file.importHtml", label: "Import HTML", params: ["at", "text"], optional: ["table"], journaled: true,
    run: (wb, p) => placed(wb, rangeOf(p.at)[0], fromHtml(p.text, countOf(p.table ?? "1")), "HTML"),
  },
  {
    // The cells as an HTML table, as they show, their looks inline;
    // header=1 makes the first row th.
    id: "file.exportHtml", label: "Export HTML", params: [], optional: ["range", "header"], journaled: false,
    run: async (wb, p) => {
      const r = p.range === undefined ? usedRange(wb) : rangeOf(p.range), tl = topLeft(r);
      const rows = Math.abs(r[0].row - r[1].row) + 1, cols = Math.abs(r[0].col - r[1].col) + 1;
      const html = toHtml(Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => {
        const k = keyOf({ col: tl.col + j, row: tl.row + i });
        return { text: wb.textOf(k), style: wb.styleOf(k) };
      })), flag(p.header));
      return { ok: true, message: `${named(r)} as HTML: ${rows} rows`, changed: [], output: html };
    },
  },
  {
    id: "example.paste", label: "Paste the Excel example", params: [], journaled: true,
    run: async (wb) => {
      if (wb.sheet.language.name !== "Excel") await wb.setLanguage(wb.language("Excel"));
      return setting(wb, pasted(EXAMPLE, { col: 0, row: 0 }), "pasted the example");
    },
  },
];

const BY_ID = new Map(SPECS.map((s) => [s.id, s]));

// A cell's value as a script sees it: '' for an empty one, an error's code.
const scripted = (v: Value): unknown => (v === null ? "" : v instanceof CellError ? v.code : typeof v === "boolean" ? (v ? 1 : 0) : v);

export class Commands {
  readonly #wb: Workbook;
  readonly #store: Store;
  readonly #journal: string[] = [];

  constructor(wb: Workbook, store: Store = memoryStore()) {
    this.#wb = wb;
    this.#store = store;
  }

  // Every command whose id or label holds `search` (any case).
  catalogue(search = ""): { id: string; label: string; params: string }[] {
    const s = search.toLowerCase();
    return SPECS.filter((c) => c.id.toLowerCase().includes(s) || c.label.toLowerCase().includes(s))
      .map((c) => ({ id: c.id, label: c.label, params: [...c.params, ...(c.optional ?? []).map((o) => `${o}?`)].join(", ") }));
  }

  async execute(id: string, params: Params = {}): Promise<Outcome> {
    const spec = BY_ID.get(id);
    try {
      if (!spec) refuse(`no command ${id}`);
      for (const name of spec!.params) if (params[name] === undefined) refuse(`${id} needs ${name}`);
      // The workbook as it was, kept as a step to undo once the command has
      // changed it.
      const before = spec!.journaled && !spec!.unstepped ? this.#wb.saved() : null;
      // Several areas ("A1:B2,D4:D9", a multiple selection): the command
      // over each, every one read first, as one step and one line.
      const areas = AREAS.has(id) && params.range?.includes(",") ? params.range.split(",").map((a) => a.trim()) : null;
      if (areas) for (const a of areas) rangeOf(a);
      if (!AREAS.has(id) && !OWN_AREAS.has(id) && params.range?.includes(",")) refuse(`${spec!.label} cannot be used on several areas (${params.range})`);
      const outcome = areas ? await this.#overAreas(spec!, params, areas) : await spec!.run(this.#wb, params, this.#store);
      if (outcome.ok && before) this.#wb.changed(before, spec!.label);
      if (outcome.ok && spec!.journaled) this.#journal.push(lineOf(id, params));
      return outcome;
    } catch (e) {
      if (e instanceof Refused || e instanceof RangeError) return { ok: false, message: e.message, changed: [] };
      throw e;
    }
  }

  async #overAreas(spec: Spec, params: Params, areas: readonly string[]): Promise<Outcome> {
    const outs: Outcome[] = [];
    for (const range of areas) {
      const o = await spec.run(this.#wb, { ...params, range }, this.#store);
      if (!o.ok) return o;
      outs.push(o);
    }
    return { ok: true, message: outs.map((o) => o.message).join("; "), changed: [...new Set(outs.flatMap((o) => o.changed))] };
  }

  run(line: string): Promise<Outcome> {
    const { id, params } = parseLine(line);
    return this.execute(id, params);
  }

  // The command lines that changed the workbook, in order.
  journal(): string[] {
    return [...this.#journal];
  }

  // A RexxJS script, driving these commands (script.ts); `say` gets its
  // SAY lines.
  // The whole script is one step of undo.
  script(text: string, say: (line: string) => void): Promise<{ ok: boolean; message: string }> {
    return this.#grouped(() => runScript(text, this.#scriptHost(), say));
  }

  // One control-bus string, as the Tauri app's (script.ts's runSpreadsheet).
  // One step of undo, however many commands it runs; and all or nothing: a
  // request that fails part way (SETCELL("A1", "x"); BOGUS()) is taken back,
  // so the program that sent it, told it failed, finds the sheet as it was.
  async spreadsheet(code: string, say: (line: string) => void): Promise<{ ok: boolean; message: string; output?: unknown }> {
    // One the bus gave up on, still open: ended, not taken back, as what the
    // person did since may be in it.
    if (this.#request) {
      this.#wb.endStoppedGroup();
      this.#request = null;
    }
    const request = { cancelled: false };
    this.#request = request;
    this.#wb.beginGroup();
    let o: { ok: boolean; message: string; output?: unknown } = { ok: false, message: "the request did not finish" };
    let stepped = false;
    try {
      o = await runSpreadsheet(code, this.#scriptHost(request), say);
    } finally {
      if (request.cancelled) o = { ok: false, message: "the request was stopped" };
      // Folded whole when it failed, to be taken back as one.
      if (this.#request === request) {
        stepped = this.#wb.endGroup(!o.ok);
        this.#request = null;
      }
    }
    if (!o.ok && stepped) return { ...o, message: `${o.message}${await this.#takeBack()}` };
    return o;
  }

  // The request in flight, which sae's control bus answered without it
  // (bus.serve's stopped): "interrupted", the page's time limit stopped it
  // and it will not come back, so it is taken back now; "deadline", it was
  // still running at the bus's 10 s, so it is cancelled: its next command
  // fails, and it is taken back as any failed request is.
  async stopped(reason: string): Promise<string> {
    const request = this.#request;
    if (!request) return "";
    request.cancelled = true;
    if (reason !== "interrupted") return "cancelled";
    this.#request = null;
    return this.#wb.endStoppedGroup(true) ? `taken back${await this.#takeBack()}` : "nothing to take back";
  }

  #request: { cancelled: boolean } | null = null;

  // The step a failed request folded into, undone, with nothing to redo.
  async #takeBack(): Promise<string> {
    await this.execute("edit.undo", {});
    this.#wb.forgetRedo();
    return " (nothing it did was kept)";
  }

  async #grouped<T>(work: () => Promise<T>): Promise<T> {
    this.#wb.beginGroup();
    try {
      return await work();
    } finally {
      this.#wb.endGroup();
    }
  }

  // What a script reaches the sheet by; for a control-bus request, each
  // command refused once the bus has stopped it.
  #scriptHost(request?: { cancelled: boolean }) {
    const live = () => {
      if (request?.cancelled) throw new Error("the request was stopped");
    };
    return {
      run: (line: string) => (live(), this.run(line)),
      execute: (id: string, params: Record<string, string>) => (live(), this.execute(id, params)),
      set: (cell: string, input: string) => (live(), this.execute("cell.set", { cell, input })),
      get: (cell: string) => (live(), scripted(this.#wb.sheet.value(cell))),
    };
  }
}
