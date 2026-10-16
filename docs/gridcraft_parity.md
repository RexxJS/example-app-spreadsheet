# Toward GridCraft, in small steps, each reachable from REXX

Started 2026-10-15. The yardstick is `gridcraft_feature_reference.md`, a
summary of GridCraft taken on trust (see `AGENTS.md`: its repository is
never read, it has a different licence). This file is the tracker: where
the spreadsheet stands against each section of that reference, and the
next steps in the reference's own kickoff order (its section 14), each
small enough to land with its tests in one commit.

## The rules every step follows

1. **A step is one feature, with its tests, in one commit.** A headless
   test in `tests/headless/` (the model, no window) and, where the ui
   shows it, an `it` in a spec. The reference's own measure: presence,
   discoverability, command access, correctness, undo/redo, persistence,
   scale. A step says which of those it proves.
2. **Everything is a command** (`commands.ts`: id, label, params, undo),
   as GridCraft's are, so the ui, scripts, the bus and the MCP server all
   reach it the same way, and a headless test can drive it without a
   dialog.
3. **Everything is REXX.** GridCraft has no REXX; this app does, and that
   is its edge. Each step adds (or extends) a control function in
   `control.ts`, described in `described.ts`, so it is callable from a
   script, from `ADDRESS SPREADSHEET`, over the control bus, and from the
   MCP server's `spreadsheet_call` and `spreadsheet_run_program` with no
   further work. `sheet_described` fails if a function is not described.
   A step whose feature a REXX program could not drive is not done.
4. **Excel formulas and RexxJS formulas both.** A new worksheet function
   lands in `functions.ts` (Excel) and, where it makes sense, as a REXX
   range function (`rexx.ts`), and is tested in both languages.
5. **Nothing from GridCraft but the list.** Behaviour comes from Excel's
   documented behaviour and this app's own tests, not from anyone's code.

## Where the spreadsheet stands (2026-10-15)

Evidence is this repository: `app/sheet/*.ts`, `commands.ts`'s 113
command ids, `functions.ts`'s 121 functions, `described.ts`'s 131 REXX
functions, 44 headless test files and 89 spec `it`s.

| Reference section | Have | Partial | Missing |
|---|---|---|---|
| 1 Product surface | Native macOS, Linux (GTK4), Windows (Win32) through aether-ui; blank, number, text, boolean, error values; formulas; date serials; one workbook open, several sheets | 26 x 100 cells a sheet (`address.ts`), not 16,384 x 1,048,576 | Browser (wasm) build; pictures in cells; 1904 dates |
| 2 Look and feel | Panels (bars) over the grid, a command bar, Name Box (`cell` field), formula bar, sheet tabs (rename, move); dark mode as aether-ui's theme; coloured references and point mode while typing a formula; frozen rows | Sheet tabs: no colour, hide, drag (an aether-ui ask); frozen columns need aether-ui | Ribbon, status-bar statistics, zoom, split views, selection glide, autocomplete and argument hints |
| 3 Editing | Cell and range selection, keyboard movement, in-cell editor, formula bar; copy, cut, paste, the system clipboard, paste as text, paste transpose; fill down/up/left/right and series; insert/delete rows, columns, sheets; hide/unhide; row heights, column widths; find and replace (3 options); undo/redo as snapshots, one step a command, labelled, grouped | Paste Special only transposes; no F4 cycling; shortcuts: a working set, not Excel's map | Multi-area and whole row/column selection, Go To Special, Flash Fill, Format Painter, AutoFit, R1C1 |
| 4 Formula language | Two languages (Excel, RexxJS); references, absolute/relative, cross-sheet, names (workbook and sheet scoped, qualified `Costs!QTY`), structured table references, dynamic arrays and spilling, `A1#`, `LET`, `LAMBDA`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY` (evaluator forms); dependency recalculation, cycles found across sheets (Tarjan); auditing: trace precedents/dependents, show formulas, evaluate formula (`audit.ts`), watch | Implicit intersection `@`: not seen | 3-D spans, R1C1, iterative circular calculation, manual calculation mode, `TRIMRANGE`, Error Checking, `#CALC!` |
| 5 Functions | 219 + `IF`: aggregation and criteria, logic, lookup (`XLOOKUP`, `XMATCH`, `INDEX`/`MATCH`, `VLOOKUP`, `HLOOKUP`, `LOOKUP`), references (`ROW`, `COLUMN`, `ADDRESS`, `OFFSET`, `INDIRECT`: the last two volatile; `FORMULATEXT`, `ISFORMULA`, `ISREF`, `SHEET`, `SHEETS`), arrays (`FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `TAKE`, `DROP`, `VSTACK`, `HSTACK`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `CHOOSECOLS`, `CHOOSEROWS`, `EXPAND`, `TRANSPOSE`), text (with `SEARCH`, `REPLACE`, `FIXED`, `DOLLAR`, `NUMBERVALUE`, `UNICODE`/`UNICHAR`, `ARRAYTOTEXT`/`VALUETOTEXT`, `REGEXTEST`/`REGEXEXTRACT`/`REGEXREPLACE`), dates (with `WEEKNUM`, `ISOWEEKNUM`, `YEARFRAC`, `DAYS360`, `TIMEVALUE`, the `.INTL` pair), math (rounding to multiples, `TRUNC`, `EXP`/`LN`/`LOG`, `GCD`/`LCM`, `FACT`/`COMBIN`/`PERMUT`, trig and hyperbolic, `RAND`/`RANDBETWEEN`/`RANDARRAY` volatile), basic statistics, aggregates (`SUBTOTAL` and `AGGREGATE` skipping hidden rows, `GROUPBY`, `PIVOTBY`, the `A` forms), information (the `IS*` family, `N`, `T`, `TYPE`, `ERROR.TYPE`, `INFO` in part), `XOR`, `FV PV PMT NPER RATE NPV IRR` | Against 505: about a third | statistics beyond the basics; most financial; engineering; database |
| 6 Formatting | Bold, italic, underline, strikethrough, size, colour, background, alignment, wrap, merge, borders; number formats (number, currency, percent, date, Excel patterns, the Tauri dialect); conditional rules: between, contains, top/bottom, above/below average, duplicates, unique, highlight, colour scales, data bars, icon sets; sparklines (line, bar); checkbox, dropdown and date editors | Rules: no formula rule, begins/ends, blanks, errors, date periods; fonts: no family, double underline, super/subscript; alignment: no indent, rotation, vertical | Cell and table style galleries, themes, Format Painter, win/loss sparklines |
| 7 Tables, filters, validation | Tables (define, columns, metadata, sort cycle), structured references; sort ascending/descending by a column; AutoFilter by a criterion (`data.filter`); text to columns; remove duplicates; validation: list, whole, decimal, length, invalid cells circled | Table totals row and calculated columns: to check; multilevel sort: no | Advanced Filter, slicers, Flash Fill, consolidation, outlines/subtotals, custom-formula and pattern validation, input prompts, JSON/HTML import |
| 8 Pivots and analysis | Pivot tables (rows, columns, value, SUM/COUNT/AVERAGE/MIN/MAX, filters, refresh); Goal Seek; data tables (`whatif.ts`) | Pivot layouts, grouping, show-values-as, calculated fields: no | Scenario Manager, Solver, Forecast Sheet, the 13 analysis tools, PivotCharts |
| 9 Charts and objects | Bar, line, pie, from a range, drawn in `vg`; sparklines in cells; chart panel, drag-out | 3 kinds of 25; titles, legend, axes: to check | Area, scatter, doughnut, stacked and the rest; trendlines; Select Data; pictures, shapes, text boxes, hyperlinks, ink |
| 10 Review, protection, print | Comments (one per cell, `notes.ts`); show formulas; freeze rows | | Threads and replies, protection and locked cells, print, PDF, page layout, zoom, split, gridline/heading toggles |
| 11 Files | Own JSON and `.rexxsheet` (the Tauri format), CSV in and out, through sae's `documents` | | XLSX (read, write, preservation), XLS, ODS, TSV, HTML, JSON import, encrypted files |
| 12 Automation | 113 commands with ids, labels, params and undo; a command bar; REXX scripts (`ADDRESS SHEET`, `ADDRESS SPREADSHEET`, setup scripts); 131 control functions, described (`LISTCOMMANDS(1)`); the sealed control bus; `sae-bus` as the CLI; an MCP server (6 tools, JSON or REXX, whole programs, sandboxed, one step of undo a request, all or nothing) | | Recording actions as a script (Automate) and replaying it; a UI-driving channel (keys, pointer, screenshots: sae-driver has one for tests); randomized command sequences; headless-vs-interactive comparison |
| 13 Localization | | | Interface languages; localized formula aliases and separators |

## The steps

The things to do are `TODO.md` at the repository's root: about ninety
steps in the reference's own kickoff order (its section 14), each one
commit with its tests, each with the REXX it gains, and the look-and-feel
plan (a menu bar and a toolbar over the grid, the panels shown on demand).
A step struck through there is done; this table is brought up to date
when a section of it changes.

## Deliberately not

GridCraft's own stubs and gaps need no chasing first: cube functions,
`WEBSERVICE`/`FILTERXML`, VBA/Python execution, `IMAGE`, `RTD`,
`STOCKHISTORY`, coauthoring, 3-D charts, SmartArt. And nothing of
GridCraft's code or docs beyond the reference is consulted (`AGENTS.md`).
