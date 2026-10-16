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

## Where the spreadsheet stands (2026-10-16)

Evidence is this repository: `app/sheet/*.ts`, `commands.ts`'s 142
command ids, 441 worksheet functions and `IF` (section 5), `described.ts`'s
173 REXX functions, 63 headless test files and 111 spec `it`s; `menus.ts`'s
nine menus, 137 items.

| Reference section | Have | Partial | Missing |
|---|---|---|---|
| 1 Product surface | Native macOS, Linux (GTK4), Windows (Win32) through aether-ui; blank, number, text, boolean, error values; formulas; date serials, 1900 or 1904; one workbook open, several sheets | 26 x 100 cells a sheet (`address.ts`), not 16,384 x 1,048,576 | Browser (wasm) build; pictures in cells |
| 2 Look and feel | A native menu bar as old Excel's (File, Edit, View, Insert, Format, Data, Formulas, Review, Script: every command on one, shortcuts shown) in the installed app; one toolbar row as old Excel's (undo to chart, tooltips), the Name Box and formula bar the row under it; fourteen panels on demand under the grid, one open at a time (tabs, View's items, `SHOWPANEL`); a status bar with the selection's figures (`SELECTIONSTATS`) and the zoom (`SETZOOM`); a command bar, Name Box (`cell` field), formula bar, sheet tabs (rename, move by Move left/right or by dragging a tab, colour, hide, copy); dark mode as aether-ui's theme; frozen rows and columns, Freeze Panes, Split; coloured references and point mode while typing a formula; frozen rows | | Ribbon (old Excel's toolbar instead, by choice), split views, selection glide, argument names in the hints |
| 3 Editing | Cell and range selection, whole rows and columns (`B:B`, `3:5`, `SELECTRANGE`), several areas (Ctrl/Cmd+click), keyboard movement, in-cell editor, formula bar; copy, cut, paste, the system clipboard, paste as text, Paste Special (values, formulas, formats, comments; add, subtract, multiply, divide; skip blanks; transpose); Format Painter; Go To Special; Flash Fill (words, initials, joins, case); fill down/up/left/right and series; insert/delete rows, columns, sheets; hide/unhide; row heights, column widths, AutoFit; find and replace (3 options); undo/redo as snapshots, one step a command, labelled, grouped | F4 cycling and AutoComplete in the formula bar, not the in-cell editor; Excel's keyboard map but Ctrl+Enter and Alt+Enter | Ctrl+drag (the grid has no drag-select) |
| 4 Formula language | Two languages (Excel, RexxJS); references, absolute/relative, cross-sheet, 3-D (`Jan:Mar!A1`), R1C1 as a display and entry mode (stored A1), names (workbook and sheet scoped, qualified `Costs!QTY`), structured table references, dynamic arrays and spilling, `A1#`, `LET`, `LAMBDA`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY` (evaluator forms), optional LAMBDA parameters and `ISOMITTED`, trim refs and `TRIMRANGE`, implicit intersection `@`; dependency recalculation, cycles found across sheets (Tarjan); manual calculation (stale cells, Calculate Now, Calculate Sheet) and iterative circular calculation (most iterations, maximum change, across sheets too), saved with the file; auditing: trace precedents/dependents, show formulas, evaluate formula (`audit.ts`), watch, Error Checking (errors, inconsistent formulas, numbers as text, omitted adjacent cells; Ignore Error) | A name shaped like a cell past column Z lexes as a reference (tracker 1.4) | `#CALC!` in every place Excel gives it |
| 5 Functions | 441 + `IF`: aggregation and criteria, logic, lookup (`XLOOKUP`, `XMATCH`, `INDEX`/`MATCH`, `VLOOKUP`, `HLOOKUP`, `LOOKUP`), references (`ROW`, `COLUMN`, `ADDRESS`, `OFFSET`, `INDIRECT`: the last two volatile; `FORMULATEXT`, `ISFORMULA`, `ISREF`, `SHEET`, `SHEETS`), arrays (`FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `TAKE`, `DROP`, `VSTACK`, `HSTACK`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `CHOOSECOLS`, `CHOOSEROWS`, `EXPAND`, `TRANSPOSE`), text (with `SEARCH`, `REPLACE`, `FIXED`, `DOLLAR`, `NUMBERVALUE`, `UNICODE`/`UNICHAR`, `ARRAYTOTEXT`/`VALUETOTEXT`, `REGEXTEST`/`REGEXEXTRACT`/`REGEXREPLACE`), dates (with `WEEKNUM`, `ISOWEEKNUM`, `YEARFRAC`, `DAYS360`, `TIMEVALUE`, the `.INTL` pair), math (rounding to multiples, `TRUNC`, `EXP`/`LN`/`LOG`, `GCD`/`LCM`, `FACT`/`COMBIN`/`PERMUT`, trig and hyperbolic, `RAND`/`RANDBETWEEN`/`RANDARRAY` volatile), descriptive statistics (spread, means, `SKEW`/`KURT`, correlation and covariance, the least-squares line with `TREND`/`GROWTH`/`LINEST`, `FREQUENCY`, ranks, `PERCENTRANK`, `MODE.MULT`), the distributions (normal, lognormal, t, χ², F, binomial, Poisson, exponential, gamma, beta, Weibull, hypergeometric, negative binomial: densities, tails, inverses) and the tests (`Z.TEST`, `T.TEST`, `F.TEST`, `CHISQ.TEST`), with the compatibility names, aggregates (`SUBTOTAL` and `AGGREGATE` skipping hidden rows, `GROUPBY`, `PIVOTBY`, the `A` forms), information (the `IS*` family, `N`, `T`, `TYPE`, `ERROR.TYPE`, `INFO` in part), `XOR`, financial (`FV PV PMT NPER RATE NPV IRR`, `IPMT`/`PPMT` and the cumulative pair, `XNPV`/`XIRR`/`MIRR`, depreciation `SLN SYD DB DDB VDB`, `EFFECT`/`NOMINAL`, `RRI`, `PDURATION`, `FVSCHEDULE`, `ISPMT`, `DOLLARDE`/`DOLLARFR`; bonds and bills: `PRICE`, `YIELD`, `DURATION`, the `COUP*`, `TBILL*`, `ACCRINT*`, discount and at-maturity securities, the `ODD*` four, on the five day-count bases), engineering (bases, bits, `CONVERT`, `ERF`, complex numbers, Bessel), the database functions (`D*` over a criteria range) | Against 505: about three quarters | computed criteria for the database functions |
| 6 Formatting | Bold, italic, underline, strikethrough, size, colour, background, alignment, wrap, merge, borders; number formats (number, currency, percent, date, Excel patterns, the Tauri dialect); conditional rules: between, contains, top/bottom, above/below average, duplicates, unique, highlight, colour scales, data bars, icon sets; sparklines (line, bar); checkbox, dropdown and date editors | Rules: no formula rule, begins/ends, blanks, errors, date periods; fonts: no family, double underline, super/subscript; alignment: no indent, rotation, vertical | Cell and table style galleries, themes, win/loss sparklines |
| 7 Tables, filters, validation | Tables (define, columns, metadata, sort cycle), structured references; sort ascending/descending by a column; AutoFilter by a criterion (`data.filter`); text to columns; remove duplicates; validation: list, whole, decimal, length, invalid cells circled | Table totals row and calculated columns: to check; multilevel sort: no | Advanced Filter, slicers, consolidation, outlines/subtotals, custom-formula and pattern validation, input prompts, JSON/HTML import |
| 8 Pivots and analysis | Pivot tables (rows, columns, value, SUM/COUNT/AVERAGE/MIN/MAX, filters, refresh); Goal Seek; data tables (`whatif.ts`) | Pivot layouts, grouping, show-values-as, calculated fields: no | Scenario Manager, Solver, Forecast Sheet, the 13 analysis tools, PivotCharts |
| 9 Charts and objects | Bar, line, pie, from a range, drawn in `vg`; sparklines in cells; chart panel, drag-out | 3 kinds of 25; titles, legend, axes: to check | Area, scatter, doughnut, stacked and the rest; trendlines; Select Data; pictures, shapes, text boxes, hyperlinks, ink |
| 10 Review, protection, print | Comments (one per cell, `notes.ts`); show formulas; freeze rows | | Threads and replies, protection and locked cells, print, PDF, page layout, zoom, gridline/heading toggles |
| 11 Files | Own JSON and `.rexxsheet` (the Tauri format), CSV in and out, through sae's `documents` | | XLSX (read, write, preservation), XLS, ODS, TSV, HTML, JSON import, encrypted files |
| 12 Automation | 142 commands with ids, labels, params and undo; a command bar; REXX scripts (`ADDRESS SHEET`, `ADDRESS SPREADSHEET`, setup scripts); 173 control functions, described (`LISTCOMMANDS(1)`); the sealed control bus; `sae-bus` as the CLI; an MCP server (6 tools, JSON or REXX, whole programs, sandboxed, one step of undo a request, all or nothing) | | Recording actions as a script (Automate) and replaying it; a UI-driving channel (keys, pointer, screenshots: sae-driver has one for tests); randomized command sequences; headless-vs-interactive comparison |
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
