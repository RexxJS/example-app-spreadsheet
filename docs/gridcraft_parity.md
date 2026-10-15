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
| 5 Functions | 121 + `IF`: aggregation and criteria, logic, lookup (`XLOOKUP`, `XMATCH`, `INDEX`/`MATCH`, `VLOOKUP`, `HLOOKUP`), arrays (`FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `TAKE`, `DROP`, `VSTACK`, `HSTACK`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `CHOOSECOLS`, `CHOOSEROWS`, `EXPAND`, `TRANSPOSE`), text, dates, basic math and statistics, `FV PV PMT NPER RATE NPV IRR` | Against 505: about a quarter | `OFFSET`, `INDIRECT`, `ADDRESS`, `ROW`, `COLUMN`, `LOOKUP`; most math (`CEILING`, `FLOOR`, `TRUNC`, `EXP`, `LN`, trig); information (`ISERR`, `ISNA`, `ISEVEN`, `N`, `TYPE`); statistics beyond the basics; most financial; engineering; database; `REGEX*`; `SUBTOTAL`, `AGGREGATE`, `GROUPBY`, `PIVOTBY`; `RAND`, `RANDBETWEEN`, `RANDARRAY` |
| 6 Formatting | Bold, italic, underline, strikethrough, size, colour, background, alignment, wrap, merge, borders; number formats (number, currency, percent, date, Excel patterns, the Tauri dialect); conditional rules: between, contains, top/bottom, above/below average, duplicates, unique, highlight, colour scales, data bars, icon sets; sparklines (line, bar); checkbox, dropdown and date editors | Rules: no formula rule, begins/ends, blanks, errors, date periods; fonts: no family, double underline, super/subscript; alignment: no indent, rotation, vertical | Cell and table style galleries, themes, Format Painter, win/loss sparklines |
| 7 Tables, filters, validation | Tables (define, columns, metadata, sort cycle), structured references; sort ascending/descending by a column; AutoFilter by a criterion (`data.filter`); text to columns; remove duplicates; validation: list, whole, decimal, length, invalid cells circled | Table totals row and calculated columns: to check; multilevel sort: no | Advanced Filter, slicers, Flash Fill, consolidation, outlines/subtotals, custom-formula and pattern validation, input prompts, JSON/HTML import |
| 8 Pivots and analysis | Pivot tables (rows, columns, value, SUM/COUNT/AVERAGE/MIN/MAX, filters, refresh); Goal Seek; data tables (`whatif.ts`) | Pivot layouts, grouping, show-values-as, calculated fields: no | Scenario Manager, Solver, Forecast Sheet, the 13 analysis tools, PivotCharts |
| 9 Charts and objects | Bar, line, pie, from a range, drawn in `vg`; sparklines in cells; chart panel, drag-out | 3 kinds of 25; titles, legend, axes: to check | Area, scatter, doughnut, stacked and the rest; trendlines; Select Data; pictures, shapes, text boxes, hyperlinks, ink |
| 10 Review, protection, print | Comments (one per cell, `notes.ts`); show formulas; freeze rows | | Threads and replies, protection and locked cells, print, PDF, page layout, zoom, split, gridline/heading toggles |
| 11 Files | Own JSON and `.rexxsheet` (the Tauri format), CSV in and out, through sae's `documents` | | XLSX (read, write, preservation), XLS, ODS, TSV, HTML, JSON import, encrypted files |
| 12 Automation | 113 commands with ids, labels, params and undo; a command bar; REXX scripts (`ADDRESS SHEET`, `ADDRESS SPREADSHEET`, setup scripts); 131 control functions, described (`LISTCOMMANDS(1)`); the sealed control bus; `sae-bus` as the CLI; an MCP server (6 tools, JSON or REXX, whole programs, sandboxed, one step of undo a request, all or nothing) | | Recording actions as a script (Automate) and replaying it; a UI-driving channel (keys, pointer, screenshots: sae-driver has one for tests); randomized command sequences; headless-vs-interactive comparison |
| 13 Localization | | | Interface languages; localized formula aliases and separators |

## The steps, in the reference's kickoff order

Each is a commit with its tests; the REXX column is what a script or
`spreadsheet_run_program` gains. Numbers are for reference, not strict
order within a stage. Steps already done are struck through when they
land.

### Stage 1: core workbook

| # | Step | Tests prove | REXX |
|---|---|---|---|
| 1.1 | Whole row and column selection (click a header), and select all | spec: a header click selects the column, Delete clears it, undo | `SELECTRANGE("B:B")` |
| 1.2 | Multi-area selection (Ctrl/Cmd+click, Ctrl+drag), commands over every area | headless: `range.clear` over two areas is one step; spec | `SELECTRANGE("A1:B2,D4:D9")` |
| 1.3 | Go To Special: formulas, constants, blanks, errors, comments, validation, conditional formats | headless: each category's cells, on a sample sheet | `GOTOSPECIAL("blanks")` gives the stem of cells |
| 1.4 | A bigger sheet: sparse cells up to 16,384 x 1,048,576 as an address space (storage already sparse? check `sheet.ts`), the grid virtual over it (aether-ui datagrid: an ask if it pages) | headless: `XFD1048576` sets and reads; spec: scroll far, type, formula back | `SETCELL("ZZ100000", 1)` |
| 1.5 | Sheet tab colour and hiding, move/copy sheet (copy a sheet with its cells and looks) | headless: `sheet.copy`, undo restores; spec: the tab | `SETSHEETCOLOR`, `HIDESHEET`, `COPYSHEET` |
| 1.6 | 1904 date system as a workbook setting, honoured by the serial conversions | headless: the same serial, two dates | `SETDATESYSTEM(1904)` |

### Stage 2: calculation

| # | Step | Tests prove | REXX |
|---|---|---|---|
| 2.1a | ~~`ROW`, `COLUMN` (of a reference, over a range, or of the formula's own cell), `ADDRESS` (all four mixes, R1C1, a sheet), `LOOKUP` (vector and array forms)~~ (done 2026-10-15: `ROW`/`COLUMN` are evaluator forms, the formula's cell passed to `evaluate`; a reference past the 26 x 100 sheet is `#REF!` as everywhere, until 1.4) | `sheet_functions`: 11 checks | in formulas, `EVALUATE()`, and `GETROW`/`GETCOL` already |
| 2.1b | `OFFSET`, `INDIRECT`: references made at run time, so the formula is recalculated on every change (volatile, as Excel's are) | `sheet_functions`; `sheet_book`: a change at the pointed-at cell shows | in formulas and `EVALUATE()` |
| 2.1c | `FORMULATEXT`, `ISFORMULA`, `ISREF`, `SHEET`, `SHEETS` (the evaluator told a cell's input and the workbook's sheets) | `sheet_functions`, `sheet_book` | |
| 2.2 | Math batch: `CEILING`/`FLOOR` (+ `.MATH`, `.PRECISE`), `TRUNC`, `MROUND`, `EVEN`, `ODD`, `EXP`, `LN`, `LOG`, `LOG10`, `QUOTIENT`, `GCD`, `LCM`, `FACT`, `COMBIN`, `PERMUT`, `SUMSQ`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, trig and hyperbolic | headless, with volatile `RAND*` recalculating on `formulas.recalculate` | `RECALCULATE()` already |
| 2.3 | Information and logic: `ISERR`, `ISNA`, `ISEVEN`, `ISODD`, `ISLOGICAL`, `ISNONTEXT`, `N`, `T`, `TYPE`, `ERROR.TYPE`, `XOR`, `TRUE`, `FALSE`, `INFO` (what applies) | headless | |
| 2.4 | Text batch: `CODE`, `CHAR`, `UNICODE`, `UNICHAR`, `CLEAN`, `FIXED`, `DOLLAR`, `SEARCH`, `REPLACE`, `CONCATENATE`, `NUMBERVALUE`, `ARRAYTOTEXT`, `VALUETOTEXT`, `REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE` | headless, regex on QuickJS's engine | REXX's own string functions already cover much; the Excel ones for formulas |
| 2.5 | Dates batch: `WEEKNUM`, `ISOWEEKNUM`, `YEARFRAC`, `DAYS360`, `TIMEVALUE`, `NETWORKDAYS.INTL`, `WORKDAY.INTL` | headless: against documented examples | |
| 2.6 | Aggregates: `SUBTOTAL` (with hidden-row skipping), `AGGREGATE`, `MAXA`/`MINA`/`AVERAGEA`, `COUNTA` over mixed, `GROUPBY`, `PIVOTBY` | headless; `SUBTOTAL` with `rows.hide` | |
| 2.7 | Statistics batch 1 (descriptive): `AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`, `TRIMMEAN`, `KURT`, `SKEW`, `STDEV.S/.P`, `VAR.S/.P`, `CORREL`, `PEARSON`, `COVARIANCE.*`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, `FORECAST.LINEAR`, `TREND`, `GROWTH`, `LINEST`, `FREQUENCY`, `RANK.EQ/.AVG`, `PERCENTILE.INC/.EXC`, `QUARTILE.INC/.EXC`, `PERCENTRANK.*`, `MODE.SNGL/.MULT`, `STANDARDIZE` | headless; numbers against a worked oracle in the test | |
| 2.8 | Statistics batch 2 (distributions): `NORM.*`, `T.*`, `CHISQ.*`, `F.*`, `BINOM.*`, `POISSON.DIST`, `EXPON.DIST`, `GAMMA*`, `BETA.*`, `LOGNORM.*`, `WEIBULL.DIST`, `HYPGEOM.DIST`, `NEGBINOM.DIST`, `CONFIDENCE.*`, `Z.TEST`, `T.TEST`, `F.TEST`, `CHISQ.TEST`, `FISHER*`, `GAUSS`, `PHI`, `PROB`, and the compatibility aliases | headless against published values (tolerances stated) | |
| 2.9 | Financial batch: `IPMT`, `PPMT`, `CUMIPMT`, `CUMPRINC`, `XNPV`, `XIRR`, `MIRR`, `SLN`, `SYD`, `DB`, `DDB`, `VDB`, `EFFECT`, `NOMINAL`, `RRI`, `PDURATION`, `FVSCHEDULE`, `ISPMT`, `DOLLARDE/FR`; then bonds and bills (`PRICE`, `YIELD`, `DURATION`, `COUP*`, `TBILL*`, `ACCRINT*`, `ODD*`) | headless; the iterative ones converge on documented examples | `finance.ts` is where the TVM ones are |
| 2.10 | Engineering: base conversion (`DEC2BIN` ...), bits (`BITAND` ...), `CONVERT`, `DELTA`, `GESTEP`, `ERF*`, complex numbers (`COMPLEX`, `IM*`), Bessel | headless | |
| 2.11 | Database functions `D*` over a criteria range | headless on a sample table | |
| 2.12 | `ISOMITTED`, named lambdas (a name whose value is a `LAMBDA`: check `sheet_let`), `TRIMRANGE` and trim refs, implicit intersection `@` (the array helpers are there already) | headless `sheet_arrays` | a REXX `LAMBDA`-free equivalent: `RUN_PROGRAM` loops |
| 2.13 | 3-D references `Sheet1:Sheet3!A1` in `SUM` and friends | headless `sheet_book` | |
| 2.14 | R1C1 as a display and entry mode (stored A1) | headless: the same formula shown both ways; spec: the toggle | `SETREFERENCESTYLE("R1C1")` |
| 2.15 | Manual calculation mode and calculate sheet/all; iterative circular calculation (max iterations, max change) as workbook settings | headless: a cycle converges under iteration, is `#CYCLE!` without | `SETCALCULATION("manual")`, `SETITERATION(100, 0.001)` |
| 2.16 | Error Checking: a pane listing cells with errors, inconsistent formulas in a row/column, numbers stored as text | headless: the list for a sample sheet | `ERRORCHECK()` gives a stem |

### Stage 3: desktop workflow

| # | Step | Tests prove | REXX |
|---|---|---|---|
| 3.1 | Paste Special: values, formulas, formats, comments, transpose, operations (add, multiply), skip blanks | headless `sheet_copy`; spec | `PASTESPECIAL("B1", "values")` |
| 3.2 | Format Painter (copy a cell's look to a range) | headless; spec | `COPYFORMAT("A1", "B1:B9")` |
| 3.3 | F4 cycling of a reference while editing; autocomplete of function names and argument hints in the formula bar | spec | |
| 3.4 | Flash Fill: pattern inference from an example column (split, join, case) | headless: the examples the reference names (split names, case) | `FLASHFILL("B1:B9")` |
| 3.5 | AutoFit column width and row height to content | spec (geometry through the driver) | `AUTOFIT("B")` |
| 3.6 | Keyboard map: Excel's navigation (Ctrl+arrows to the edge, Ctrl+Home/End, Page Up/Down, Shift+Space, Ctrl+Space), editing (F2, Ctrl+Enter, Alt+Enter), formatting (Ctrl+B/I/U, Ctrl+Shift+$,%), Ctrl+Z/Y, Ctrl+F/H, Ctrl+S | spec: each key's effect | |
| 3.7 | Status bar: sum, average, count, min, max of the selection; zoom | spec | `SELECTIONSTATS()` |
| 3.8 | Split views and frozen columns (aether-ui asks as needed) | spec | `FREEZEPANES(rows, cols)` for real |
| 3.9 | Selection glide (the outline animates to the new selection; a setting turns it off) | spec: the setting; the glide itself is a `vg`/aether-ui matter | `SETANIMATIONS(0)` |
| 3.10 | Several workbooks open at once (windows), as sae apps allow | spec | `OPENWORKBOOK`, `LISTWORKBOOKS` |

### Stage 4: presentation

| # | Step | Tests prove | REXX |
|---|---|---|---|
| 4.1 | Number formats: General with Excel's rules, accounting, fractions, scientific, elapsed time `[h]:mm`, custom codes with sections (`;`), conditions, colours and literals; `TEXT()` on the same formatter | headless `sheet_numfmt`: a table of codes and values | `SETFORMAT` already; `FORMATVALUE(v, code)` |
| 4.2 | Fonts: family, double underline, super/subscript; alignment: vertical, indent, rotation | headless `sheet_look`; spec: drawn | `SETFONT("A1", "family=Georgia;size=14")` |
| 4.3 | Borders: per-side styles and colours, draw-border, clear | headless; spec | `SETBORDER("A1:C3", "outline", "thin")` |
| 4.4 | Conditional rules: formula rules, begins/ends with, blanks/nonblanks, errors/nonerrors, date periods (today, last 7 days, this month ...), top/bottom percent; rule order and stop-if-true | headless `sheet_notes`/rules | `ADDRULE("A1:A9", "formula", "=A1>B1", "bold;color:red")` |
| 4.5 | Data bar and icon set options: gradient/solid, show value or not, thresholds, reversal | headless; spec: drawn | `ADDRULE(..., "databar", ...)` |
| 4.6 | Cell styles (a named look applied to cells, edited in one place) and table styles with banding and header/total emphasis | headless; spec | `DEFINESTYLE`, `APPLYSTYLE`, `SETTABLESTYLE` |
| 4.7 | Themes: colour and font roles the looks refer to | headless: a theme change re-colours | `SETTHEME` |
| 4.8 | Win/loss sparklines; sparkline options (markers, axis) | headless `sheet_charts`; spec | `INSERTSPARKLINE(..., "winloss")` |
| 4.9 | Table totals row with its aggregate per column, calculated columns (a formula filled to the column), banded rows, convert to range | headless `sheet_structured`; spec | `SETTABLETOTALS`, `TABLETORANGE` |
| 4.10 | Multilevel sort (several keys, each a direction; by colour) and Advanced Filter (a criteria range) | headless `sheet_sortfind` | `SORTRANGE` with a stem of keys; `ADVANCEDFILTER` |
| 4.11 | Validation: date, time, custom formula, pattern; input prompts and error styles | headless `sheet_validation` | `SETCELLVALIDATION` extended |
| 4.12 | Slicers for a table column | spec | `ADDSLICER("Sales", "region")`, `SLICE("Sales", "region", "north")` |
| 4.13 | Outlines: group rows/columns, show/hide detail, `data.subtotal` | headless; spec | `GROUPROWS(2, 9)`, `SUBTOTALS(...)` |
| 4.14 | Consolidation across ranges/sheets | headless | `CONSOLIDATE` |

### Stage 5: files

| # | Step | Tests prove | REXX |
|---|---|---|---|
| 5.1 | TSV and delimiter detection, BOM and encodings on CSV import; CSV export of displayed values keeping the separator | headless `sheet_csv` | `IMPORTCSV`, `EXPORTCSV` already? extend |
| 5.2 | JSON (rows of objects) and HTML table import; HTML export | headless | `IMPORTJSON`, `EXPORTHTML` |
| 5.3 | XLSX read: cells, formulas (stored text), number formats, merges, names, sheets (a zip + XML reader in the app or an aether `contrib`) | headless on fixture files made by hand (not GridCraft's) | `OPENFILE("x.xlsx")` through documents |
| 5.4 | XLSX write of the same, and round-trip tests | headless: write, read, equal | `SAVEAS("x.xlsx")` |
| 5.5 | XLSX: styles, conditional formatting, validation, comments, tables, sparklines, charts (what we draw), pivots; preservation of parts we do not model (pass-through), with a report of what was dropped | headless per part | `FILEREPORT()` |
| 5.6 | XLS (BIFF8) import of values and formulas; ODS import of cached values | headless on hand-made fixtures | |
| 5.7 | Encrypted XLSX opening (Agile/Standard AES), with sae's password prompt | spec | |
| 5.8 | PDF export and printing through sae/aether-ui (an ask: a print capability) | spec: a file appears | `EXPORTPDF()` |

### Stage 6: analysis and graphics

| # | Step | Tests prove | REXX |
|---|---|---|---|
| 6.1 | Chart kinds: column/bar clustered, stacked, 100%; line with markers, stacked line; area, stacked area; doughnut; scatter, scatter with lines, bubble | headless `sheet_charts` (the model); spec: drawn, pixels sampled | `INSERTCHART(range, kind)` already? extend |
| 6.2 | Chart titles, axis titles, legend position, data labels, gridlines, switch row/column, Select Data | headless; spec | `SETCHARTOPTION(id, "title", ...)` |
| 6.3 | Trendlines: linear, exponential, logarithmic, polynomial, power, moving average; equation and R² | headless (the fit against known data); spec: drawn | `ADDTRENDLINE(id, series, "linear")` |
| 6.4 | Radar, histogram, waterfall, funnel, treemap, sunburst, box-and-whisker, stock, combo (secondary axis) | per kind, a step each | |
| 6.5 | Pivots: compact/outline/tabular layouts, subtotals and grand totals, sorting and filtering, date grouping, show-values-as, calculated fields, collapse/expand | headless `sheet_pivot` | `PIVOTCONFIG` extended |
| 6.6 | PivotCharts bound to a pivot | headless; spec | |
| 6.7 | Scenario Manager (named sets of input values, a summary) | headless | `ADDSCENARIO`, `SHOWSCENARIO`, `SCENARIOSUMMARY` |
| 6.8 | Solver: simplex LP first (objective, changing cells, constraints, Answer report); nonlinear and evolutionary later | headless on textbook problems | `SOLVE(...)` from a REXX program: the natural way to describe a model |
| 6.9 | Data Analysis tools, one or two a step: descriptive statistics, correlation, covariance, histogram, moving average, exponential smoothing, rank and percentile, regression, t-tests, F-test, ANOVA | headless, each writing values | `ANALYZE("regression", ...)` |
| 6.10 | Forecast Sheet (`FORECAST.ETS` family, then the sheet and chart) | headless | |
| 6.11 | Pictures: floating and in cells (aether-ui image widgets), hyperlinks, text boxes, basic shapes | spec | `INSERTPICTURE`, `SETHYPERLINK` |
| 6.12 | Threaded comments: replies, resolve, a comments pane; notes converted to comments | headless `sheet_notes`; spec | `ADDCOMMENT`, `REPLYCOMMENT`, `RESOLVECOMMENT` |
| 6.13 | Protection: locked cells, sheet and workbook protection, Allow Edit Ranges (a password hash kept, as a file setting) | headless: a locked cell refuses `cell.set` with protection on | `PROTECTSHEET`, `LOCKCELLS` |
| 6.14 | Views: gridlines, headings, formula bar toggles; zoom; Page Layout and page breaks (with 5.8) | spec | `SETVIEW("gridlines", 0)` |

### Stage 7: automation and hardening

| # | Step | Tests prove | REXX |
|---|---|---|---|
| 7.1 | **Record a session as a REXX script** (GridCraft's Automate, in REXX): every command run from the ui becomes a line of a script (its control function with its arguments), editable, replayable, savable with the workbook | headless: a journal of commands renders to REXX that, run, reproduces the workbook; spec: Record, act, Stop, Replay | this is the edge: the recording IS a REXX program |
| 7.2 | Replay with parameters (a recorded script with `ARG`), and scheduled/triggered scripts (on open, on change of a range) | headless; spec | `ONCHANGE("A1:A9", script)` |
| 7.3 | A `sae-bus` CLI that covers what GridCraft's CLI does: evaluate a formula, inspect a workbook, convert files, run a command, list commands and functions | a shell test per subcommand | |
| 7.4 | MCP parity: every command reachable (`spreadsheet_call` already), plus workbook inspection tools (sheets, names, tables, charts), and a headless mode (the server starts the app itself, no window) | `mcp` tests | |
| 7.5 | A UI-driving channel: keys, pointer, screenshots over the bus (sae-driver's for tests, opened to the bus with a grant) | spec | |
| 7.6 | Malformed-input tests: every command with missing, wrong-typed and huge parameters refuses cleanly; a fuzz of random command sequences leaves the workbook consistent (undo all the way back equals the start) | headless `sheet_fuzz`, seeded | `RUN_PROGRAM` of a random script |
| 7.7 | Headless-vs-interactive: the same workload through the bus and through the ui compared workbook to workbook | spec | |
| 7.8 | Large-file limits: a sheet of 100,000 cells opens, recalculates and scrolls within stated times; limits refused with a message | headless timings asserted | |
| 7.9 | Localization: interface strings in a table; Spanish formula aliases, decimal comma and `;` arguments in the editor (stored canonical) | headless `sheet_language`; spec | `SETLOCALE("es")` |
| 7.10 | Accessibility: the grid and panels readable by a screen reader (aether-ui asks); keyboard-only operation of every panel | spec via the driver | |

## Deliberately not

GridCraft's own stubs and gaps need no chasing first: cube functions,
`WEBSERVICE`/`FILTERXML`, VBA/Python execution, `IMAGE`, `RTD`,
`STOCKHISTORY`, coauthoring, 3-D charts, SmartArt. And nothing of
GridCraft's code or docs beyond the reference is consulted (`AGENTS.md`).
