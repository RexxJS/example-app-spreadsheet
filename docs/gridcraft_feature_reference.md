# GridCraft feature reference (the summary this app tracks against)

Kept here as given by Paul on 2026-10-15, as the yardstick for
`gridcraft_parity.md`. It is taken on trust: nobody working on this app
reads GridCraft's repository (a different licence, MIT OR Apache-2.0),
not even to check a detail here. Where this summary says "no runtime
conformance ... testing was performed", the same holds for us.

---

As of 2026-10-15, source revision `e6b6b34`. A starting brief for an independent implementation and a checklist for feature comparison. Describes the checked-out implementation; a registered command or function does not establish complete Excel compatibility. This is a source/documentation review, with the existing sales-dashboard screenshot inspected; no runtime conformance or performance testing was performed for this document.

GridCraft is an open-source, Rust-native desktop and browser spreadsheet with an Excel-style workflow, XLSX as its main workbook format, and first-class command/API control. It is currently alpha. Its strengths are broad formula coverage, familiar editing and formatting, tables, charts, and automation. Advanced chart fidelity, external-data systems, accessibility, collaboration, and complete international typography remain substantial gaps.

## 1. Product surface

| Dimension | Current scope |
| --- | --- |
| Platforms | Native macOS, Windows, Linux and BSD; browser through WebAssembly. Platform verification and packaging maturity vary. |
| Application | Multiple open workbooks managed by a session; worksheets with Excel-sized address space: 1,048,576 rows × 16,384 columns, through XFD. This is an address limit, not a practical memory-capacity claim. |
| Data | Blank, numeric, text, Boolean and error values; formulas; 1900/1904 date systems; static pictures inside cells. |
| Distribution | MIT OR Apache-2.0; local operation; no subscription or licence server. |
| Interaction | Ribbon, keyboard shortcuts, direct grid manipulation, dialogs, task panes, CLI, MCP and a local JSON control channel. |
| Maturity | Feature breadth exceeds depth. File preservation, calculation edge cases and platform behavior need individual comparison tests. |

## 2. Look and feel

The default screen is a dense desktop spreadsheet: title/quick-access strip at the top; ribbon tabs and grouped controls below; Name Box and formula bar; a large white worksheet; sheet tabs and horizontal scrolling below it; a bottom status bar with selection statistics, view buttons and zoom. Right-side task panes appear for relevant workflows. The worksheet remains the dominant surface.

- **Ribbon:** Home, Insert, Draw, Page Layout, Formulas, Data, Review, View and Automate, with File backstage navigation. Contextual controls include Table Design, Chart Design, chart Format, PivotTable and Header & Footer workflows. Groups combine large icon buttons, smaller buttons, split/dropdown controls, palettes and galleries.
- **Visual language:** restrained grey chrome, fine separators, flat original icons, white inputs, subtle hover/pressed fills, green active-tab accents and green cell-selection outlines. Headers distinguish selected rows and columns. Floating charts use white plotting surfaces and contrasting categorical series colors.
- **Light tokens:** window `#F3F3F3`, ribbon/grid `#FFFFFF`, main text `#242424`, gridlines `#E1E1E1`, headers `#F8F8F8`, accent/selection border `#107C41`, soft accent `#DFF1E6`. Selection fill is translucent green.
- **Dark tokens:** window `#292929`, ribbon `#333333`, main text `#EEEEEE`, accent `#5CC587`. The worksheet itself retains a white background and black cell text; dark mode principally changes application chrome.
- **Typography:** platform UI fonts; installed workbook-font matching and fallback for cells; regular, bold, italic and bold-italic roles. Layout varies with installed fonts. Font files are an optional external build input, never bundled from proprietary applications.
- **Motion:** selection border, translucent fill and fill handle glide to a new selection over approximately 120 ms with deceleration. Logical selection changes immediately. Scrolling and zoom keep the outline attached to cells; cross-sheet/offscreen jumps appear directly. View → Animations disables the glide. Charts do not have equivalent data-change animation.
- **Density and navigation:** virtualized grid, smooth scrolling, zoom, adjustable interface size, resizable rows/columns, frozen panes and split views. Sheet tabs support rename, reorder, color, hiding and scrolling.
- **Feedback:** colored formula references, source-range outlines for selected charts, resize handles, note/comment indicators, validation messages and import-warning notices.

Visual examples: sales dashboard, dark Insert ribbon, budget formatting, formula entry, PivotTable fields, chart formatting, page breaks (repository images, not reproduced here). These are repository examples, not a guarantee of current pixel geometry.

## 3. Workbook editing and navigation

| Area | Features and behavior |
| --- | --- |
| Selection | Active cell, rectangular ranges and multiple areas; row/column/all-sheet selection; mouse drag, Shift extension, keyboard movement, Name Box and Go To. |
| Entry | In-cell editor and formula bar; Enter/Tab commit and move; Escape cancels; multiline text; autocomplete and function argument hints; point at cells/ranges while constructing formulas. |
| References while editing | Colored reference outlines, A1/R1C1 modes, F4 cycling of relative/absolute/mixed references. |
| Clipboard | Copy, cut, paste; internal structured clipboard plus text/HTML exchange; Paste Special for values, formulas, formats, comments and transpose; Format Painter. |
| Fill | Drag fill handle, directional fill, numeric/date series and Flash Fill pattern inference. Formula references adjust on copy/fill. |
| Structural edits | Insert/delete cells, rows, columns and sheets; shift affected cells; resize/AutoFit; hide/unhide; sheet rename, move/copy, reorder and tab colors. |
| Search | Find/Replace, Go To, Go To Special for selected cell categories. |
| History | Undo/redo using copy-on-write workbook snapshots; commands distinguish mutations from queries. |
| Keyboard | Common spreadsheet navigation, editing, formatting, file, print and function shortcuts, with platform-specific mappings. Alt/F10 keytips cover an initial subset of Home and legacy paths; full keyboard parity is incomplete. |

Do not equate window-management command names with a complete native multiwindow workspace. Independent windows, pane cycling and some advanced selection modes remain comparison items.

## 4. Formula language and calculation

- Operators, literals, nested expressions, functions, ranges, whole rows/columns, sheet-qualified references, 3-D sheet spans and array constants.
- A1 and R1C1 addresses, absolute/relative/mixed references, workbook- and sheet-scoped names, Name Manager and names created from selection labels.
- Structured table references, including column references, totals and current-row forms.
- Dynamic arrays and spilling, spill-range references (`A1#`), implicit intersection (`@`) and legacy array formulas. Occupied cells, including picture cells, block spills.
- `LET`, first-class `LAMBDA`, named lambdas, omitted arguments and array helpers: `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY`.
- Trim references for leading/trailing empty edges, backed by `TRIMRANGE`.
- Dependency-graph calculation, incremental dirty propagation, volatile recalculation, automatic/manual calculation controls, calculate-sheet/all and cycle detection.
- Optional iterative circular calculation, bounded by maximum iterations and maximum change; persisted in XLSX calculation settings.
- Formula auditing: precedent/dependent arrows, Show Formulas, Error Checking, Evaluate Formula and Watch Window.
- Error values include classic spreadsheet errors and dynamic-array errors such as `#SPILL!` and `#CALC!`.

Calculation is single-threaded. Broad name coverage does not establish numerical or coercion equivalence: floating-point comparisons, distributions, iterative financial solvers and structural-reference edge cases require an oracle suite. External workbook link metadata can be preserved, but links are not followed or refreshed.

## 5. Worksheet function coverage

A static extraction of library declarations plus evaluator special forms finds **514 distinct names, including legacy `USDOLLAR`**. Nine are deliberate stubs: seven `CUBE*` functions return `#N/A`; `WEBSERVICE` and `FILTERXML` return `#VALUE!`. Thus **505 names have implementations beyond those stubs**; this counts presence, not validated correctness. The full extracted inventory is in the appendix. Some older parity-doc tables contain conflicting counts and should not be used as a release measurement.

| Family | Representative functions / intended workloads |
| --- | --- |
| Aggregation and criteria | `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `COUNTBLANK`, `SUMIF/SUMIFS`, `COUNTIF/COUNTIFS`, `AVERAGEIF/AVERAGEIFS`, `SUBTOTAL`, `AGGREGATE`, `SUMPRODUCT`, `PERCENTOF`. |
| Logic and errors | `IF`, `IFS`, `SWITCH`, `CHOOSE`, `AND`, `OR`, `NOT`, `XOR`, `IFERROR`, `IFNA`, type/error predicates. |
| Lookup and reference | `XLOOKUP`, `XMATCH`, `VLOOKUP`, `HLOOKUP`, `LOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `INDIRECT`, address and reference introspection. |
| Arrays and reshaping | `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`, `TAKE`, `DROP`, `VSTACK`, `HSTACK`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `CHOOSECOLS`, `CHOOSEROWS`, `EXPAND`, `GROUPBY`, `PIVOTBY`. |
| Text | Concatenation/join, extraction/search/replace, case/trim/clean, number/text conversion, `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`, regular-expression test/extract/replace. |
| Dates and times | Date/time constructors, serial conversion, parts, date differences, month ends, week numbers, workdays/network days and international weekend variants. |
| Math | Rounding, trigonometry/hyperbolic functions, logarithms, powers, random values, integer/combinatorial arithmetic, matrix operations and sums of powers/products. |
| Statistics | Descriptive statistics, percentiles/ranks, distributions and inverses, hypothesis tests, correlation/covariance, regression/trends and the `FORECAST.ETS` seasonal forecasting family. |
| Financial | Time value of money, amortization, NPV/IRR/XIRR, bond pricing/yields/coupons, depreciation, treasury bills and odd first/last coupon periods. |
| Engineering | Complex arithmetic, base conversion, bit operations, unit conversion, Bessel and error functions. |
| Database | Criteria-table operations such as `DSUM`, `DAVERAGE`, `DCOUNT`, `DGET`, database extrema and variance/deviation. |
| Compatibility | Older function names retained for imported workbooks; legacy statistical aliases and `USDOLLAR`; `EUROCONVERT` is in Add-in/Automation. |
| Web/cube | `ENCODEURL` works. Remote fetch/XML and cube functions above are stubs. |

Not provided: `GETPIVOTDATA`, `IMAGE`, `RTD`, `CALL`, `REGISTER.ID`, `STOCKHISTORY`, `TRANSLATE`, `DETECTLANGUAGE`, `COPILOT` and `PY`. Embedded picture insertion is separate from `IMAGE()`. No VBA, Python, COM or Office Scripts execution environment.

## 6. Formatting and visual effects

| Area | Supported surface |
| --- | --- |
| Fonts | Family/size, bold, italic, underline/double underline, strike, super/subscript, font color. Exact rendering depends on font availability and supported faces. |
| Alignment | Horizontal/vertical alignment, indent, text rotation/orientation, wrapping, merged cells, Merge Across and Merge & Center. |
| Borders and fills | Cell fills; border sides/styles/colors, draw-border workflow, clear borders. |
| Number formats | General, decimal, currency/accounting, percent, date/time, elapsed time, fractions, scientific and custom codes. Multiple sections, conditions, colors, literals, padding/fill and `TEXT()` share the formatter. |
| Styles | Built-in original cell-style and table-style galleries, custom cell styles, Format Painter, table header/totals/first/last-column emphasis and row/column banding. |
| Themes | Workbook theme colors and major/minor fonts. A registered Theme Effects command currently returns success without applying effects; treat it as a placeholder. |
| Conditional rules | Value comparison, formula expressions, contains/begins/ends text, blanks/nonblanks, errors/nonerrors, duplicates/uniques, top/bottom count or percent, above/below average and date periods. |
| Conditional visuals | Two/three-stop color scales, data bars including gradient/value visibility options, icon sets with thresholds/reversal/value visibility. Rule management and removal. |
| In-cell graphics | Line, column and win/loss sparklines; Boolean checkboxes with click/Space toggling. |

"Effects" here primarily means data-driven cell visuals and chart styling. A full Office-style shadow, glow, reflection, bevel, 3-D material and per-object effects stack is not established. Compare individual rendered properties rather than counting the Effects label.

## 7. Tables, cleaning, filters and validation

- Structured tables: creation from a range, naming/resizing, calculated columns, structured references, filter headers, total-row aggregates, banded styles, conversion back to a normal range and PivotTable creation.
- Ascending/descending and multilevel custom sorting; AutoFilter, criteria editing, clearing/reapplying and Advanced Filter.
- Table slicers: floating value tiles for one table column, single/multiple selection, clear filter and delete. They use the table's row-filter machinery. Pivot timelines and full slicer XLSX compatibility are separate gaps.
- Text to Columns, Remove Duplicates, Flash Fill, consolidation, outline grouping/ungrouping, show/hide detail and subtotals.
- Validation for numeric/date/time/text-length/list/custom-formula constraints, input prompts, error styles and invalid-data marking.
- CSV/TSV, JSON and HTML-table data import. Some query/connection commands exist; this is not a Power Query transformation editor or database connector suite.

## 8. PivotTables and analysis

**PivotTables:** insert/recommended pivots; a fields pane with Filters, Columns, Rows and Values; field add/remove/reorder; sum/count and other aggregate functions; compact/outline/tabular layouts; grand totals/subtotals; sorting/filtering; date grouping; collapse/expand; show-values-as; calculated fields; source changes and refresh. Pivot caches/tables are read/written in XLSX. PivotCharts bind to a report and update on refresh/layout changes, with limited category/field-button support.

**What-if:** Goal Seek, Scenario Manager and one/two-variable data tables.

**Forecast Sheet:** generates a table and chart from timeline/value data using the ETS forecasting implementation. Excel's exact optimizer behavior is not guaranteed.

**Solver:** Simplex LP, local nonlinear/GRG-style and Evolutionary methods; objective max/min/target, changing cells, bounds/relational/integer/binary constraints, saved model and Answer/Limits reports. It is an independent solver; do not assume Excel's optimizer or identical convergence. No Sensitivity report.

**Data Analysis:** 13 static report tools: single-factor ANOVA, correlation, covariance, descriptive statistics, exponential smoothing, F-test, histogram, moving average, rank/percentile, regression, paired t-test, equal-variance t-test and unequal-variance t-test. Reports write values, not live formulas. This is a subset of Analysis ToolPak.

Advanced PivotTable calculated items, data-model pivots, relationships/DAX and complete pivot filtering/grouping depth remain comparison gaps.

## 9. Charts and inserted objects

The chart model has **25 kinds**:

| Family | Kinds |
| --- | --- |
| Column / bar | Clustered, stacked and 100% stacked for each orientation. |
| Line / area | Line, line with markers, stacked line; area and stacked area. |
| Circular | Pie, doughnut. |
| XY | Scatter, scatter with lines, bubble. |
| Other | Radar, histogram, waterfall, funnel, treemap, sunburst, box-and-whisker, stock, combo. |

Chart workflows include Insert/Recommended Charts with previews, live worksheet data, row/column switching, Select Data, chart-type changes, title and axis-title text, legend position, label/gridline toggles, palettes/styles, chart-level Format pane, movement and resizing. Combo supports a secondary-axis line series. Source data is outlined when a chart is selected.

Trendlines include linear, exponential, logarithmic, polynomial, power and moving average, with forecast extension, optional intercept, equation/R² labels and line styling. Supported for applicable 2-D series; no complete Format Trendline pane.

Chart depth is limited: no full axis-scale/options system, error bars, chart data tables, per-point formatting or full label formatting; no 3-D/surface/contour/map/Pareto chart coverage. Not all modeled kinds/properties have equivalent XLSX preservation. Existing chart sheets can be kept unchanged under supported conditions but are not shown/editable. Chart templates are absent.

Other objects: floating PNG/JPEG pictures, pictures fitted inside cells, basic shapes, text boxes, original icons, symbols, hyperlinks and ink. Object commands cover placement, size, stacking, alignment, grouping and rotation to varying depth. Draw includes pen/highlighter, eraser, lasso and Ink to Shape. Supported ink polylines survive XLSX save/reopen. Pen pressure/tilt and a full touch mode are absent.

Cell pictures preserve aspect ratio, follow sizing/sort/filter/structural edits and internal clipboard operations, support alt text and XLSX rich-value storage, and render in PDF. Formula references to picture values fall back to `#VALUE!`; no remote image fetching. SmartArt, WordArt, equations, signatures, embedded-object editing and 3-D models are absent.

## 10. Review, protection, printing and views

- Threaded comments with replies/resolution, notes and visibility controls; comments pane and navigation; conversion of notes to comments.
- Spelling, thesaurus, workbook statistics and a registered accessibility-checking workflow. These do not imply complete linguistic coverage or screen-reader accessibility.
- Sheet/workbook protection, locked cells and Allow Edit Ranges, including compatible password-hash metadata. Sheet protection is distinct from encrypted-file saving.
- Normal, Page Layout and Page Break Preview workflows; print areas/titles, page breaks, margins, orientation, paper size, scaling, headers/footers and print gridlines/headings. Page Layout depth should be compared separately from its command's presence.
- PDF export; desktop printer discovery/submission through craft-print with printer selection, copies, duplex and grayscale; browser printing/export. Physical output fidelity and driver-specific options need validation.
- PDF currently uses the standard PDF fonts without font embedding, limiting non-Latin output.
- View controls for gridlines, headings, formula bar, zoom/selection zoom, frozen panes, splits, interface language/size, dark mode and animations.

No real-time coauthoring, shared change history or cloud collaboration backend. A Share control is not evidence of those systems. Screen-reader navigation, complex text shaping and complete bidirectional UI remain gaps.

## 11. Files and interoperability

| Format / workflow | Support and limits |
| --- | --- |
| XLSX | Main read/write format: cells/formulas, styles/themes, tables, conditional formatting, validation, names, comments/notes, links, pictures/shapes, charts, sparklines, pivots, protection, print and calculation settings. Broad, imperfect round-trip fidelity. |
| Strict XLSX | Read; saved as transitional Open XML. |
| XLSM / XLTM | VBA project preserved byte-for-byte when saving to macro-enabled types; never executed. Saving to a nonmacro type drops it. |
| XLTX | Template content type supported; full new-from-template workflow incomplete. |
| XLS / XLT | Partial BIFF8 import: values, many formulas, number formats, names, merges, dimensions and calculation settings. Rich styles/objects missing; no XLS write. |
| XLSB / ODS | Worksheet-data import using cached formula results; no formula engine round trip or rich formatting/object fidelity. No write support. |
| CSV / TSV / text | Delimiter detection and common encoding/BOM handling; displayed-value export, preserving an imported CSV separator. Formatting, objects and formulas are flattened to text/results. |
| HTML / JSON | HTML-table import/export; JSON import and internal workbook/debug representation. Not general web-query or lossless interchange substitutes. |
| Encrypted Open XML | Password opening for supported Agile/Standard AES encryption. No encrypted saving; unsupported encryption types rejected. |
| PDF | Static export/printing, subject to typography limits above. |

Selected unmodeled XLSX content can pass through unchanged: custom XML/properties, external-link metadata, connections, some query/control/OLE/background parts and chart sheets. This does not enable editing or execution of that content. Structural edits and relationships to rewritten parts can prevent preservation. Unsupported chart properties/extensions and data-model/slicer/timeline content can still be lost. Large workbooks are constrained by in-memory import and defensive allocation limits.

## 12. Automation and implementation contracts

All meaningful edits go through named commands with metadata: ID, label, menu/ribbon path, shortcut, parameter documentation, enabled predicate and execution function. Examples include `cell.set`, `range.setValues`, `selection.set`, `home.bold`, `insert.chart` and `data.sortAscending`.

- UI dispatches commands to a toolkit-independent engine; headless calls never require a dialog. Missing parameters produce defaults, queries or structured errors.
- CLI covers evaluation, workbook inspection/conversion, command execution, command/function listing and MCP startup.
- MCP runs over stdio, headless or bridged to a running desktop app.
- Optional localhost JSON-lines channel executes commands, inspects state and drives UI keyboard/pointer/screenshot actions.
- Automate records actions into editable GridCraft command scripts and replays them. These are independent of VBA and Office Scripts.
- Engine errors use results; input-sized work is bounded; undo uses copy-on-write snapshots. A session guard restores workbook state if a panic escapes. These are safety design measures, not proof that every operation is defect-free.

Example control-channel request:

```json
{"id":1,"method":"engine.execute","params":{"command":"cell.set","params":{"cell":"B2","input":"=SUM(A1:A9)"}}}
```

Architecture: core addresses/values and number formats → formula parser/function library → workbook model/calculation/file I/O → chart layout → command/session engine → interchangeable UI/MCP → desktop, CLI and web apps. Native UI uses egui/eframe and wgpu; calc is single-threaded; visible-cell virtualization and incremental recalc are the primary performance techniques. No measured speed advantage is claimed here.

## 13. Localization

Eleven interface-language choices are present: English, Spanish, Simplified Chinese, Japanese, Korean, Russian, Brazilian Portuguese, German, Czech, French and Serbian. Translation completeness varies and missing strings fall back to English.

Spanish cell/formula-bar entry supports localized aliases, decimal commas and semicolon arguments; for example `=SI(VERDADERO;SUMA(1,5;2,25);0)`. Other interface-language choices currently retain English formula syntax. Stored formulas, APIs, CLI and clipboard use canonical English syntax. Localized editor spelling does not change workbook semantics.

System-font fallback supports additional scripts. Right-to-left sheet mode mirrors grid navigation/layout, but full bidirectional text shaping, RTL chrome/printing, IME verification and non-Latin PDF support are incomplete.

## 14. Suggested kickoff and comparison checklist

For an independent implementation, build in this order; each stage produces a usable vertical slice:

1. **Core workbook:** sparse cells, typed values, sheets, address/range rules, selection, command dispatch, undo/redo and error handling. Acceptance: edit/undo/redo a multi-sheet workbook without losing values or styles.
2. **Calculation:** parser, references, dependency graph and core arithmetic/aggregation/logic/lookup functions. Acceptance: edit a precedent, update only affected results, copy formulas with correct reference adjustment, report cycles.
3. **Desktop workflow:** virtual grid, ribbon, formula bar, keyboard navigation, clipboard, fill and structural edits. Acceptance: create a budget using both keyboard and mouse, including cross-sheet formulas.
4. **Presentation:** number formats, typography, borders, merged/wrapped cells, styles, conditional formatting and table/filter/validation workflows. Acceptance: reproduce the budget/gradebook examples using commands.
5. **Files:** XLSX read/write and CSV import/export before advanced objects. Acceptance: reopen saved files with equivalent values/formulas/styles; report unsupported content and test preservation explicitly.
6. **Analysis and graphics:** dynamic arrays/LAMBDA, charts/sparklines, pivots, what-if, pictures/shapes, print/PDF. Acceptance: reproduce the sales dashboard and pivot report, then verify edit/refresh/save/reopen behavior.
7. **Automation and hardening:** CLI/MCP/control parity, script replay, malformed-input tests, randomized command sequences, large-file limits, localization and accessibility. Acceptance: run the same workload headlessly and interactively and compare the resulting workbook.

For every feature, compare **presence, UI discoverability, command/API access, calculation/rendering correctness, undo/redo, persistence, preservation of unsupported content, scale and platform behavior** separately. A command count cannot substitute for these checks.

## Appendix: declared function names

Extracted from library declarations plus evaluator special forms; test-only `ADD2` excluded. Evaluator special forms overlap some library declarations, so section totals must not be added. Categories retain source registry names. Cube and two Web entries are stubs as described above.

### AddIn

`EUROCONVERT`.

### Compatibility

`BETADIST`, `BETAINV`, `BINOMDIST`, `CEILING`, `CHIDIST`, `CHIINV`, `CHITEST`, `CONFIDENCE`, `COVAR`, `CRITBINOM`, `EXPONDIST`, `FDIST`, `FINV`, `FLOOR`, `FORECAST`, `FTEST`, `GAMMADIST`, `GAMMAINV`, `HYPGEOMDIST`, `LOGINV`, `LOGNORMDIST`, `MODE`, `NEGBINOMDIST`, `NORMDIST`, `NORMINV`, `NORMSDIST`, `NORMSINV`, `PERCENTILE`, `PERCENTRANK`, `POISSON`, `QUARTILE`, `RANK`, `STDEV`, `STDEVP`, `TDIST`, `TINV`, `TTEST`, `USDOLLAR`, `VAR`, `VARP`, `WEIBULL`, `ZTEST`.

### Cube

`CUBEKPIMEMBER`, `CUBEMEMBER`, `CUBEMEMBERPROPERTY`, `CUBERANKEDMEMBER`, `CUBESET`, `CUBESETCOUNT`, `CUBEVALUE`.

### Database

`DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DSUM`, `DVAR`, `DVARP`.

### DateTime

`DATE`, `DATEDIF`, `DATEVALUE`, `DAY`, `DAYS`, `DAYS360`, `EDATE`, `EOMONTH`, `HOUR`, `ISOWEEKNUM`, `MINUTE`, `MONTH`, `NETWORKDAYS`, `NETWORKDAYS.INTL`, `NOW`, `SECOND`, `TIME`, `TIMEVALUE`, `TODAY`, `WEEKDAY`, `WEEKNUM`, `WORKDAY`, `WORKDAY.INTL`, `YEAR`, `YEARFRAC`.

### Engineering

`BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`, `BIN2DEC`, `BIN2HEX`, `BIN2OCT`, `BITAND`, `BITLSHIFT`, `BITOR`, `BITRSHIFT`, `BITXOR`, `COMPLEX`, `CONVERT`, `DEC2BIN`, `DEC2HEX`, `DEC2OCT`, `DELTA`, `ERF`, `ERF.PRECISE`, `ERFC`, `ERFC.PRECISE`, `GESTEP`, `HEX2BIN`, `HEX2DEC`, `HEX2OCT`, `IMABS`, `IMAGINARY`, `IMARGUMENT`, `IMCONJUGATE`, `IMCOS`, `IMCOSH`, `IMCOT`, `IMCSC`, `IMCSCH`, `IMDIV`, `IMEXP`, `IMLN`, `IMLOG10`, `IMLOG2`, `IMPOWER`, `IMPRODUCT`, `IMREAL`, `IMSEC`, `IMSECH`, `IMSIN`, `IMSINH`, `IMSQRT`, `IMSUB`, `IMSUM`, `IMTAN`, `OCT2BIN`, `OCT2DEC`, `OCT2HEX`.

### Financial

`ACCRINT`, `ACCRINTM`, `AMORDEGRC`, `AMORLINC`, `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `CUMIPMT`, `CUMPRINC`, `DB`, `DDB`, `DISC`, `DOLLARDE`, `DOLLARFR`, `DURATION`, `EFFECT`, `FV`, `FVSCHEDULE`, `INTRATE`, `IPMT`, `IRR`, `ISPMT`, `MDURATION`, `MIRR`, `NOMINAL`, `NPER`, `NPV`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PDURATION`, `PMT`, `PPMT`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `PV`, `RATE`, `RECEIVED`, `RRI`, `SLN`, `SYD`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `VDB`, `XIRR`, `XNPV`, `YIELD`, `YIELDDISC`, `YIELDMAT`.

### Information

`ERROR.TYPE`, `HYPERLINK`, `INFO`, `ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISTEXT`, `N`, `NA`, `TYPE`.

### Logical

`AND`, `FALSE`, `IF`, `IFERROR`, `IFNA`, `IFS`, `NOT`, `OR`, `SWITCH`, `TRUE`, `XOR`.

### Lookup

`ADDRESS`, `CHOOSE`, `CHOOSECOLS`, `CHOOSEROWS`, `COLUMNS`, `DROP`, `EXPAND`, `FILTER`, `GROUPBY`, `HLOOKUP`, `HSTACK`, `INDEX`, `LOOKUP`, `MATCH`, `PIVOTBY`, `ROWS`, `SORT`, `SORTBY`, `TAKE`, `TOCOL`, `TOROW`, `TRANSPOSE`, `UNIQUE`, `VLOOKUP`, `VSTACK`, `WRAPCOLS`, `WRAPROWS`, `XLOOKUP`, `XMATCH`.

### MathTrig

`ABS`, `ACOS`, `ACOSH`, `ACOT`, `ACOTH`, `ARABIC`, `ASIN`, `ASINH`, `ATAN`, `ATAN2`, `ATANH`, `BASE`, `CEILING.MATH`, `CEILING.PRECISE`, `COMBIN`, `COMBINA`, `COS`, `COSH`, `COT`, `COTH`, `CSC`, `CSCH`, `DECIMAL`, `DEGREES`, `EVEN`, `EXP`, `FACT`, `FACTDOUBLE`, `FLOOR.MATH`, `FLOOR.PRECISE`, `GCD`, `INT`, `ISO.CEILING`, `LCM`, `LN`, `LOG`, `LOG10`, `MDETERM`, `MINVERSE`, `MMULT`, `MOD`, `MROUND`, `MULTINOMIAL`, `MUNIT`, `ODD`, `PERCENTOF`, `PI`, `POWER`, `PRODUCT`, `QUOTIENT`, `RADIANS`, `RAND`, `RANDARRAY`, `RANDBETWEEN`, `ROMAN`, `ROUND`, `ROUNDDOWN`, `ROUNDUP`, `SEC`, `SECH`, `SEQUENCE`, `SERIESSUM`, `SIGN`, `SIN`, `SINH`, `SQRT`, `SQRTPI`, `SUM`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `SUMSQ`, `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `TAN`, `TANH`, `TRUNC`.

### Statistical

`AVEDEV`, `AVERAGE`, `AVERAGEA`, `AVERAGEIF`, `AVERAGEIFS`, `BETA.DIST`, `BETA.INV`, `BINOM.DIST`, `BINOM.DIST.RANGE`, `BINOM.INV`, `CHISQ.DIST`, `CHISQ.DIST.RT`, `CHISQ.INV`, `CHISQ.INV.RT`, `CHISQ.TEST`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, `CORREL`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `COVARIANCE.P`, `COVARIANCE.S`, `DEVSQ`, `EXPON.DIST`, `F.DIST`, `F.DIST.RT`, `F.INV`, `F.INV.RT`, `F.TEST`, `FISHER`, `FISHERINV`, `FORECAST.ETS`, `FORECAST.ETS.CONFINT`, `FORECAST.ETS.SEASONALITY`, `FORECAST.ETS.STAT`, `FORECAST.LINEAR`, `FREQUENCY`, `GAMMA`, `GAMMA.DIST`, `GAMMA.INV`, `GAMMALN`, `GAMMALN.PRECISE`, `GAUSS`, `GEOMEAN`, `GROWTH`, `HARMEAN`, `HYPGEOM.DIST`, `INTERCEPT`, `KURT`, `LARGE`, `LINEST`, `LOGEST`, `LOGNORM.DIST`, `LOGNORM.INV`, `MAX`, `MAXA`, `MAXIFS`, `MEDIAN`, `MIN`, `MINA`, `MINIFS`, `MODE.MULT`, `MODE.SNGL`, `NEGBINOM.DIST`, `NORM.DIST`, `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `PEARSON`, `PERCENTILE.EXC`, `PERCENTILE.INC`, `PERCENTRANK.EXC`, `PERCENTRANK.INC`, `PERMUT`, `PERMUTATIONA`, `PHI`, `POISSON.DIST`, `PROB`, `QUARTILE.EXC`, `QUARTILE.INC`, `RANK.AVG`, `RANK.EQ`, `RSQ`, `SKEW`, `SKEW.P`, `SLOPE`, `SMALL`, `STANDARDIZE`, `STDEV.P`, `STDEV.S`, `STDEVA`, `STDEVPA`, `STEYX`, `T.DIST`, `T.DIST.2T`, `T.DIST.RT`, `T.INV`, `T.INV.2T`, `T.TEST`, `TREND`, `TRIMMEAN`, `VAR.P`, `VAR.S`, `VARA`, `VARPA`, `WEIBULL.DIST`, `Z.TEST`.

### Text

`ARRAYTOTEXT`, `ASC`, `BAHTTEXT`, `CHAR`, `CLEAN`, `CODE`, `CONCAT`, `CONCATENATE`, `DBCS`, `DOLLAR`, `EXACT`, `FIND`, `FINDB`, `FIXED`, `JIS`, `LEFT`, `LEFTB`, `LEN`, `LENB`, `LOWER`, `MID`, `MIDB`, `NUMBERVALUE`, `PHONETIC`, `PROPER`, `REGEXEXTRACT`, `REGEXREPLACE`, `REGEXTEST`, `REPLACE`, `REPLACEB`, `REPT`, `RIGHT`, `RIGHTB`, `SEARCH`, `SEARCHB`, `SUBSTITUTE`, `T`, `TEXTAFTER`, `TEXTBEFORE`, `TEXTJOIN`, `TEXTSPLIT`, `TRIM`, `UNICHAR`, `UNICODE`, `UPPER`, `VALUE`, `VALUETOTEXT`.

### Web

`ENCODEURL`, `FILTERXML`, `WEBSERVICE`.

### Evaluator special forms

`AGGREGATE`, `AREAS`, `BYCOL`, `BYROW`, `CELL`, `CHOOSE`, `COLUMN`, `COLUMNS`, `FORMULATEXT`, `IF`, `IFERROR`, `IFNA`, `IFS`, `INDEX`, `INDIRECT`, `ISFORMULA`, `ISOMITTED`, `ISREF`, `LAMBDA`, `LET`, `MAKEARRAY`, `MAP`, `OFFSET`, `REDUCE`, `ROW`, `ROWS`, `SCAN`, `SHEET`, `SHEETS`, `SUBTOTAL`, `SWITCH`, `TEXT`, `TRIMRANGE`.
