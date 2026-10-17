// Every function a script and the control bus may call (LISTCOMMANDS()),
// described: its parameters, what it does and what it gives. Read by
// LISTCOMMANDS(1) (control.ts), so a program driving the app (the MCP server
// in mcp/, a REXX script) learns how to call each one, not just its name.
// tests/headless/sheet_described.ts holds this to the functions themselves:
// every name listed, and each one's parameters as many as it takes.

// What an argument is, as the functions read it (control.ts's text, yes,
// row, col and cell):
//   cell     a cell reference, "A1"
//   range    a range, "A1:C5" (one cell, "B2", is a range of one)
//   row      a row number, counted from 1
//   column   a column, by its letter ("B") or its number (2)
//   text     any text
//   number   a number
//   flag     1 for yes, 0 for no
//   json     JSON, as text
//   stem     a REXX stem (list.0 = how many, list.1, list.2, ...)
//   query    a query, piped in: RANGE("A1:D9") |> WHERE("amount > 100")
export type ParamType = "cell" | "range" | "row" | "column" | "text" | "number" | "flag" | "json" | "stem" | "query";

export interface Param {
  readonly name: string;
  readonly type: ParamType;
  readonly optional?: true;
  readonly variadic?: true;   // the rest of the arguments
  readonly about?: string;
}

export interface Described {
  readonly params: readonly Param[];
  readonly does: string;
  readonly gives: string;
}

const p = (name: string, type: ParamType, about?: string): Param => (about ? { name, type, about } : { name, type });
const opt = (name: string, type: ParamType, about?: string): Param => ({ ...p(name, type, about), optional: true });
const fn = (params: Param[], does: string, gives: string): Described => ({ params, does, gives });

const CELL = p("cell", "cell");
const RANGE = p("range", "range");
const ROW = p("row", "row");
const COLUMN = p("column", "column");
const NAME = p("name", "text");
const QUERY = p("query", "query");
const FIELD = p("column", "text", "a column of the query, by its header or column_C");

export const DESCRIBED: Readonly<Record<string, Described>> = {
  // ---- cells
  SETCELL: fn([CELL, p("value", "text", "a value, or a formula starting with =")], "Sets a cell, as typing it does", "the value set"),
  GETCELL: fn([CELL], "Reads a cell", "what it shows (a formula's result)"),
  GETEXPRESSION: fn([CELL], "Reads a cell's formula", "the formula without its =, or \"\" for none"),
  CLEARCELL: fn([p("range", "range", "a cell or a range")], "Empties a cell or a range", "\"\""),
  GETCELLS: fn([RANGE], "Reads every cell of a range, row by row", "a stem of {ref, value, expression, comment, format}"),
  SETCELLS: fn([RANGE, p("values", "stem")], "Sets a range's cells, row by row, from a stem of values", "\"OK\""),
  BATCH_SET_CELLS: fn([p("updates", "stem", "of {address, value}")], "Sets each cell it can; one that fails does not stop the rest", "JSON {total, success, errors, errorDetails}"),
  CELL: fn([p("ref", "cell", "\"A1\" or \"Costs!A1\"")], "Reads a cell, on another sheet too", "its value"),
  RUN_PROGRAM: fn([p("program", "text", "a whole REXX program: lines, loops, routines"), opt("encoding", "text", "\"json\": the program is a JSON string")],
    "Runs a whole program with these same functions, as one step of undo, taken back whole if it fails", "{said: a stem of its SAY lines, result: its RESULT}"),
  CLEAR: fn([], "Empties the active sheet", "\"OK\""),
  // ---- looks, comments, editors
  SETFORMAT: fn([CELL, p("format", "text", "e.g. \"currency:USD\", \"percent:0.0%\", \"number:0.00\", \"bold;color:red;align:center\"")],
    "Sets a cell's number format and style", "the format"),
  GETFORMAT: fn([CELL], "Reads a cell's format", "the format, or \"\""),
  SETCOMMENT: fn([CELL, p("text", "text")], "Sets a cell's comment", "the comment"),
  GETCOMMENT: fn([CELL], "Reads a cell's comment", "the comment, or \"\""),
  AUTOFIT: fn([p("what", "text", "a column (\"B\", \"B:D\") or a row (3, \"3:5\")")], "AutoFit: columns as wide as their text, or rows as tall", "what it did"),
  SETSHEETCOLOR: fn([p("sheet", "text"), p("color", "text", "a colour's name or #rrggbb; none takes it off")], "A sheet's tab colour", "what it did"),
  HIDESHEET: fn([p("sheet", "text"), opt("off", "flag", "1 to show it again")], "Hides a sheet (never the last one shown)", "what it did"),
  UNHIDESHEET: fn([p("sheet", "text")], "Shows a hidden sheet", "what it did"),
  ISSHEETHIDDEN: fn([p("sheet", "text")], "Whether a sheet is hidden", "1 or 0"),
  COPYSHEET: fn([p("sheet", "text"), opt("to", "text", "the copy's name; \"Name copy\" when left out")], "Copies a sheet with its cells and looks, after it", "what it did"),
  SETDATESYSTEM: fn([p("system", "number", "1900 or 1904")], "The workbook's date system: serials count from 1900-01-00, or from 1904-01-01", "the system now"),
  GETDATESYSTEM: fn([], "The workbook's date system", "1900 or 1904"),
  SETZOOM: fn([p("zoom", "text", "a percentage, 10 to 400, or in or out a step")], "Zooms the sheet", "the zoom now"),
  GETZOOM: fn([], "The sheet's zoom", "a percentage"),
  SELECTRANGE: fn([p("range", "range", "a cell, a range, whole columns (\"B:B\"), whole rows (\"3:5\"), or \"all\"")], "Selects it on the page", "the range selected"),
  GETSELECTION: fn([], "The selection", "a range, e.g. B2:C4"),
  GOTOSPECIAL: fn([p("what", "text", "formulas, constants, blanks, errors, comments, validation or conditional"), opt("range", "range", "the used range when left out"), opt("type", "text", "numbers, text, logicals or errors, for formulas and constants")], "Go To Special: the cells of a kind", "a stem of cells, in sheet order"),
  GETLASTCELL: fn([], "Ctrl+End's cell: the last row used and the last column used", "a cell, e.g. D12"),
  LISTKEYS: fn([], "The keyboard map", "a stem of \"keys | what they do | command\" lines"),
  FLASHFILL: fn([p("range", "range", "one column, its first cells typed as examples")], "Flash Fill: the column's empty cells filled by the pattern the examples show, from the columns to its left", "what it did"),
  COPYFORMAT: fn([p("from", "range"), p("to", "range")], "Format Painter: the cells' look, number format and style onto a range (tiled), not their values or comments", "what it did"),
  SETALIGNMENT: fn([p("range", "range"), p("align", "text", "left, center, right or general")], "Aligns the cells' text", "OK"),
  SETFILLCOLOR: fn([p("range", "range"), p("color", "text", "a colour's name or #rrggbb; none takes it off")], "The cells' fill colour", "OK"),
  SETTEXTCOLOR: fn([p("range", "range"), p("color", "text", "a colour's name or #rrggbb; none takes it off")], "The cells' text colour", "OK"),
  ADDRULE: fn([p("range", "range"), p("kind", "text", "formula, greater, less, between, notBetween, equal, notEqual, greaterOrEqual, lessOrEqual, contains, notContains, beginsWith, endsWith, blanks, noBlanks, errors, noErrors, date, top, bottom, topPercent, bottomPercent, above, below, duplicates, unique, dataBar or iconSet (their value the colour or the icons, their style ADDDATABAR's or ADDICONSET's options)"),
    opt("value", "text", "the formula (\"=A1>B1\", for the range's top-left), the value compared, the period (today, yesterday, tomorrow, last7days, lastWeek, thisWeek, nextWeek, lastMonth, thisMonth, nextMonth) or the count"), opt("style", "text", "e.g. \"bold;color:red\"; Excel's light red fill by default"), opt("value2", "text", "between's other end")],
    "Adds a conditional-format rule after the sheet's others", "the rule, as Manage Rules lists it"),
  SETTABLETOTALS: fn([p("table", "text"), opt("functions", "text", "\"Qty=sum,Price=average\": each column's sum, average, count, counta, max, min, product, stdev, stdevp, var, varp, label or none (the first column Total, the last a sum by default); 0 takes the row off")], "A table's totals row, in the empty row under it", "what was done"),
  SETCALCULATEDCOLUMN: fn([p("table", "text"), p("column", "text", "a column's name"), p("formula", "text", "for the first data row, e.g. \"=[@Qty]*[@Price]\"")], "One formula down every data row of a table's column", "what was done"),
  SETTABLEBANDING: fn([p("table", "text"), opt("on", "flag", "1 (the default) or 0")], "Banded rows: every other data row shaded", "what was done"),
  TABLETORANGE: fn([p("table", "text")], "Convert to Range: the table gone, its cells kept, its structured references written as cells", "what was done"),
  INSERTSPARKLINE: fn([p("range", "range", "the numbers"), p("at", "cell", "where it goes"), opt("type", "text", "line (the default), column or winloss"), opt("options", "text", "Excel's Show: high;low;first;last, and axis (win/loss: a zero on the axis)")], "A sparkline at a cell, as a formula", "the cell and what it shows"),
  ADDDATABAR: fn([p("range", "range"), opt("color", "text", "#rrggbb; Excel's blue by default"), opt("options", "text", "name=value pairs split by ';': fill=gradient|solid, showValue=0 (the bar only), min=, max= (the shortest and longest bar's numbers)")], "Adds a data bar rule", "the rule"),
  ADDICONSET: fn([p("range", "range"), opt("icons", "text", "arrows, triangles or ratings"), opt("options", "text", "name=value pairs split by ';': thresholds=67,33 (high, middle), by=percent|percentile|number, reverse=1, showValue=0 (icons only)")], "Adds an icon set rule", "the rule"),
  LISTRULES: fn([], "The sheet's conditional-format rules, the first winning", "a stem of \"n. range: what\" lines"),
  DELETERULE: fn([p("index", "number", "from 1, as LISTRULES numbers them")], "Deletes a conditional-format rule", "what was deleted"),
  MOVERULE: fn([p("index", "number"), p("to", "number", "its new place, 1 first")], "Moves a rule up or down the list (Manage Rules' Move Up and Move Down)", "what was done"),
  STOPIFTRUE: fn([p("index", "number"), opt("on", "flag", "1 (the default) or 0 to take it off")], "Stop If True: the rules after this one leave the cells it formats alone", "the rule"),
  CLEARRULES: fn([p("range", "range")], "Clears the rules over the range", "what was cleared"),
  SETBORDER: fn([p("range", "range"), opt("edge", "text", "outline (the default), all, inside, horizontal, vertical, top, right, bottom, left, or none"), opt("style", "text", "thin (the default), medium, thick, double, dashed, dotted, hair, or none to take those lines off"), opt("color", "text", "a colour's name or #rrggbb; black by default")], "Borders on the range, as Excel's Borders menu", "what was done"),
  CLEARBORDERS: fn([p("range", "range")], "Takes every border off the range", "what was done"),
  FORMATVALUE: fn([p("value", "text", "a number, or text"), p("format", "text", "an Excel code (\"#,##0.00;[Red](#,##0.00)\") or a category: general, number, currency, accounting, date, time, percentage, fraction, scientific, text, elapsed")], "A value in a number format, as a cell in it shows it (and TEXT() gives it)", "the text"),
  SETNUMBERFORMAT: fn([p("range", "range"), p("format", "text", "number, currency, percent, accounting, fraction, scientific, date, time, text, general, or a pattern such as 0.00")], "The cells' number format", "OK"),
  SETWRAPTEXT: fn([p("range", "range"), p("on", "flag")], "Wraps text in a range's cells, or stops", "\"OK\""),
  GETWRAPTEXT: fn([CELL], "Whether a cell wraps its text", "1 or 0"),
  SETCELLEDITOR: fn([CELL, p("type", "text", "checkbox, dropdown or date"), opt("config", "json", "{\"options\": [...]} or {\"format\": ...}")],
    "Gives a cell an editor of its own", "the cell"),
  GETCELLEDITOR: fn([CELL], "Reads a cell's editor", "JSON {type, config}, or \"\""),
  REMOVECELLEDITOR: fn([CELL], "Takes a cell's editor away", "the cell"),
  // ---- references
  GETROW: fn([CELL], "A cell's row", "its number, from 1"),
  GETCOL: fn([CELL], "A cell's column", "its number, from 1"),
  GETCOLNAME: fn([CELL], "A cell's column", "its letter"),
  MAKECELLREF: fn([COLUMN, ROW], "A cell reference from a column and a row", "e.g. \"B3\""),
  // ---- rows and columns
  INSERTROW: fn([ROW], "Inserts a row there; what reads below it follows", "\"OK\""),
  DELETEROW: fn([ROW], "Deletes a row", "\"OK\""),
  INSERTCOLUMN: fn([COLUMN], "Inserts a column there", "\"OK\""),
  DELETECOLUMN: fn([COLUMN], "Deletes a column", "\"OK\""),
  MOVECOLUMNLEFT: fn([COLUMN], "Moves a column one place left", "\"\""),
  MOVECOLUMNRIGHT: fn([COLUMN], "Moves a column one place right", "\"\""),
  HIDEROW: fn([ROW], "Hides a row", "\"OK\""),
  UNHIDEROW: fn([ROW], "Shows a hidden row", "\"OK\""),
  HIDECOLUMN: fn([COLUMN], "Hides a column", "\"OK\""),
  UNHIDECOLUMN: fn([COLUMN], "Shows a hidden column", "\"OK\""),
  ISROWHIDDEN: fn([ROW], "Whether a row is hidden", "1 or 0"),
  ISCOLUMNHIDDEN: fn([COLUMN], "Whether a column is hidden", "1 or 0"),
  ISROWVISIBLE: fn([ROW], "Whether a row is shown", "\"1\" or \"0\""),
  SETROWHEIGHT: fn([ROW, p("height", "number", "pixels")], "Sets a row's height", "\"OK\""),
  GETROWHEIGHT: fn([ROW], "A row's height", "pixels, or 0 for the grid's own"),
  SETCOLWIDTH: fn([COLUMN, p("width", "number", "pixels")], "Sets a column's width", "\"OK\""),
  GETCOLWIDTH: fn([COLUMN], "A column's width", "pixels, or 0 for the grid's own"),
  FREEZEPANES: fn([p("rows", "number"), opt("columns", "number", "0 to 8")], "Freezes the top rows and the first columns", "\"OK\""),
  UNFREEZEPANES: fn([], "Unfreezes the rows and columns", "\"OK\""),
  GETFROZENPANES: fn([], "How many rows and columns are frozen", "{ROWS, COLUMNS}"),
  // ---- filling, sorting, finding
  FILLDOWN: fn([p("source", "range"), p("target", "range")], "Fills the target from the source above it", "\"OK\""),
  FILLUP: fn([p("source", "range"), p("target", "range")], "Fills the target from the source below it", "\"OK\""),
  FILLRIGHT: fn([p("source", "range"), p("target", "range")], "Fills the target from the source left of it", "\"OK\""),
  FILLLEFT: fn([p("source", "range"), p("target", "range")], "Fills the target from the source right of it", "\"OK\""),
  AUTOFILL: fn([p("source", "range"), p("target", "range", "beside the source, on any side")], "Fills the target from the source, whichever way it lies", "\"\""),
  SORTRANGE: fn([RANGE, p("by", "column"), opt("ascending", "flag", "1 or left out: ascending; 0: descending")], "Sorts a range's rows by a column", "\"OK\""),
  FIND: fn([p("text", "text"), opt("matchCase", "flag"), opt("wholeCell", "flag"), opt("inFormulas", "flag")], "Finds the cells holding some text", "a stem of cell references"),
  REPLACE: fn([p("text", "text"), p("with", "text"), opt("matchCase", "flag"), opt("wholeCell", "flag"), opt("inFormulas", "flag")],
    "Replaces some text everywhere it is found, as one step of undo", "how many cells changed"),
  APPLYROWFILTER: fn([COLUMN, p("criterion", "text", "as COUNTIF reads one: \">100\", \"=north\", \"apple\"")], "Filters the sheet's rows (the first row the headers) by a column", "\"\""),
  CLEARROWFILTER: fn([], "Shows every row again", "\"\""),
  GETFILTERCRITERIA: fn([], "The filter in force", "JSON {range, column, criterion}, or \"\""),
  // ---- merges
  MERGECELLS: fn([RANGE], "Merges a range's cells", "the merge's first cell"),
  UNMERGECELLS: fn([CELL], "Unmerges the merge a cell is in", "the merge's first cell"),
  GETMERGEDRANGE: fn([CELL], "The merge a cell is in", "its range, or \"\""),
  ISCELLMERGED: fn([CELL], "Whether a cell is in a merge", "\"1\" or \"0\""),
  // ---- sheets
  ADDSHEET: fn([NAME], "Adds a sheet", "its name"),
  DELETESHEET: fn([NAME], "Deletes a sheet; what read it shows #REF!", "\"\""),
  RENAMESHEET: fn([p("from", "text"), p("to", "text")], "Renames a sheet; formulas that name it follow", "the new name"),
  SETSHEETNAME: fn([p("to", "text")], "Renames the active sheet", "the new name"),
  GETSHEETNAME: fn([], "The active sheet's name", "its name"),
  SETACTIVESHEET: fn([NAME], "Makes a sheet the active one (what the other functions act on)", "its name"),
  GETACTIVESHEET: fn([], "The active sheet's name", "its name"),
  GETSHEETNAMES: fn([], "The sheets, in order", "a stem of names"),
  // ---- names
  DEFINENAMEDRANGE: fn([NAME, RANGE], "Names a range, for formulas to read by", "\"OK\""),
  DELETENAMEDRANGE: fn([NAME], "Deletes a name", "\"OK\""),
  GETNAMEDRANGE: fn([NAME], "The range a name stands for", "the range, or \"\""),
  GETALLNAMEDRANGES: fn([], "Every name", "a stem: n.NAME and n.RANGE"),
  // ---- validation
  SETCELLVALIDATION: fn([p("range", "range"), p("type", "text", "list, number or text"),
    { name: "limits", type: "text", optional: true, variadic: true, about: "list: the values, \"a,b,c\"; number: min, max; text: the shortest and longest length" }],
    "Sets the values a range's cells may take", "\"OK\""),
  CLEARCELLVALIDATION: fn([p("range", "range")], "Takes a range's validation away", "\"OK\""),
  VALIDATECELL: fn([CELL, p("value", "text")], "Whether a value would pass the cell's validation", "1 or 0"),
  // ---- tables
  DEFINETABLE: fn([NAME, RANGE, opt("headers", "flag", "1 or left out: the first row names the columns")], "Makes a range a table", "JSON of the table"),
  DELETETABLE: fn([NAME], "Deletes a table (its cells stay)", "\"\""),
  GETTABLE: fn([NAME], "Reads a table's definition", "JSON {name, range, headers, columns, ...}, or \"\""),
  LISTTABLES: fn([], "The tables", "a stem of names"),
  LIST_TABLES: fn([], "The tables", "a JSON array of names"),
  SET_TABLE_METADATA: fn([NAME, p("metadata", "json", "{range, columns: {name: letter}, hasHeader, types, descriptions}")], "Defines a table with its columns named", "a message"),
  GET_TABLE_METADATA: fn([NAME], "Reads a table's metadata", "JSON {range, columns, hasHeader, types, descriptions}, or \"\""),
  DELETE_TABLE_METADATA: fn([NAME], "Deletes a table", "a message"),
  SORTTABLEBYCOLUMN: fn([NAME, COLUMN], "Sorts a table by a column: ascending, then descending, then as it was", "how it is sorted now"),
  CONFIGURE_AUTO_ID: fn([opt("column", "column", "left out or \"\": auto IDs off"), opt("start", "number", "1 if left out"), opt("prefix", "text")],
    "Gives each inserted row the next ID in a column", "a message"),
  FIND_ROW_BY_ID: fn([p("id", "text")], "Finds the row with an auto ID", "its number, or \"\""),
  UPDATE_ROW_BY_ID: fn([p("id", "text"), p("updates", "stem", "of {column, value}")], "Sets cells in the row with an auto ID", "the row's number"),
  GET_NEXT_ID: fn([], "The next auto ID", "the ID"),
  // ---- pivots
  CREATEPIVOT: fn([p("id", "text"), p("config", "json", "{sourceRange, rowFields, colFields, valueField, aggFunction, outputCell, filters}; PIVOTCONFIG makes one")],
    "Makes a pivot table", "its id"),
  PIVOTCONFIG: fn([p("source", "range"), p("value", "text", "the column to total"), p("agg", "text", "SUM, COUNT, AVERAGE, MIN or MAX"), p("output", "cell"),
    opt("rows", "json", "a JSON array of columns"), opt("columns", "json", "a JSON array of columns")], "Makes a configuration for CREATEPIVOT", "JSON"),
  UPDATEPIVOT: fn([p("id", "text")], "Works a pivot table out again", "its id"),
  DELETEPIVOT: fn([p("id", "text")], "Deletes a pivot table", "its id"),
  GETPIVOT: fn([p("id", "text")], "Reads a pivot table's configuration", "JSON, or \"\""),
  LISTPIVOTS: fn([], "The pivot tables", "JSON"),
  // ---- the range functions
  SUM_RANGE: fn([RANGE], "Adds a range's numbers", "the sum"),
  AVERAGE_RANGE: fn([RANGE], "A range's numbers' mean", "the mean, or 0"),
  COUNT_RANGE: fn([RANGE], "Counts a range's cells that are not empty", "the count"),
  MIN_RANGE: fn([RANGE], "A range's smallest number", "it, or 0"),
  MAX_RANGE: fn([RANGE], "A range's largest number", "it, or 0"),
  MEDIAN_RANGE: fn([RANGE], "A range's numbers' median", "it, or 0"),
  STDEV_RANGE: fn([RANGE], "A range's numbers' sample standard deviation", "it, or 0"),
  STDEVP_RANGE: fn([RANGE], "A range's numbers' population standard deviation", "it, or 0"),
  VAR_RANGE: fn([RANGE], "A range's numbers' sample variance", "it, or 0"),
  VARP_RANGE: fn([RANGE], "A range's numbers' population variance", "it, or 0"),
  PRODUCT_RANGE: fn([RANGE], "Multiplies a range's numbers", "the product, or 0"),
  SUMIF_RANGE: fn([RANGE, p("condition", "text", "\">100\", \"<=5\", \"=3\", \"<>0\"")], "Adds the numbers that meet a condition", "the sum"),
  COUNTIF_RANGE: fn([RANGE, p("condition", "text", "\">100\", \"<=5\", \"=3\", \"<>0\"")], "Counts the numbers that meet a condition", "the count"),
  SPARKLINE_RANGE: fn([RANGE, opt("type", "text", "line (the default), column or winloss"), opt("options", "text", "high;low;first;last;axis")], "A range's numbers as a sparkline", "text of block characters"),
  // ---- queries: RANGE("A1:D9") |> WHERE("amount > 100") |> GROUP_BY("region") |> SUM("amount")
  RANGE: fn([p("ref", "range", "\"A1:D9\", \"Costs!A1:D9\" or a name")], "Starts a query over a range's rows", "a query to pipe on"),
  TABLE: fn([NAME], "Starts a query over a table's rows", "a query to pipe on"),
  WHERE: fn([QUERY, p("condition", "text", "REXX over the row's columns, e.g. \"amount > 100\"")], "Keeps the rows that meet a condition", "a query to pipe on"),
  GROUP_BY: fn([QUERY, FIELD], "Groups the rows by a column, for SUM, AVG or COUNT", "a query to pipe on"),
  PLUCK: fn([QUERY, FIELD], "One column of the rows", "a stem of its values"),
  SUM: fn([QUERY, FIELD], "Adds a column (per group, when grouped)", "the sum, or a stem of rows"),
  AVG: fn([QUERY, FIELD], "A column's mean (per group, when grouped)", "the mean, or a stem of rows"),
  COUNT: fn([QUERY], "Counts the rows (per group, when grouped)", "the count, or a stem of rows"),
  RESULT: fn([QUERY], "The query's rows", "a stem of rows"),
  // ---- the workbook
  SPREADSHEET_VERSION: fn([], "The app's version", "e.g. \"sae demo 18\""),
  SETFORMULALANGUAGE: fn([p("language", "text", "Excel or RexxJS")], "Sets the language formulas are written in", "the language now"),
  GETFORMULALANGUAGE: fn([], "The language formulas are written in", "Excel or RexxJS"),
  SETREFERENCESTYLE: fn([p("style", "text", "A1 or R1C1")], "Sets how the page shows Excel formulas and takes them typed (stored A1 either way)", "the style now"),
  GETREFERENCESTYLE: fn([], "How the page shows Excel formulas", "A1 or R1C1"),
  EVALUATE: fn([p("formula", "text", "in the formula language, without the =")], "Works a formula out against the sheet, in no cell", "its value"),
  RECALCULATE: fn([], "Works every formula out again", "how many there were"),
  SETCALCULATION: fn([p("mode", "text", "Automatic or Manual")], "Sets how formulas are worked out: on every change, or only when calculated (CALCULATE)", "the mode now"),
  GETCALCULATION: fn([], "How formulas are worked out", "Automatic or Manual"),
  COPY: fn([p("range", "range")], "Copies a range, for PASTE or PASTESPECIAL", "what it did"),
  CUT: fn([p("range", "range")], "Cuts a range: PASTE moves it", "what it did"),
  PASTE: fn([p("at", "cell")], "Pastes what was copied or cut, at a cell", "what it did"),
  PASTESPECIAL: fn([p("at", "cell"), p("what", "text", "all, formulas, values, formats or comments"), opt("operation", "text", "none, add, subtract, multiply or divide"), opt("skipBlanks", "flag"), opt("transpose", "flag")], "Paste Special: what was copied (COPY), at a cell", "what it did"),
  SELECTIONSTATS: fn([p("range", "range"), opt("figure", "text", "sum, average, count, numbers, min or max; none for the whole line")], "The status bar's figures for a range, as Excel's", "the line (\"Average: 2  Count: 3 ...\"), or the one figure"),
  SPLITWINDOW: fn([opt("on", "flag", "0 removes the split")], "Split: a second pane over the sheet, scrolling on its own", "1 split, 0 not"),
  ISSPLIT: fn([], "Whether the sheet is split in two panes", "1 or 0"),
  SHOWPANEL: fn([p("name", "text", "edit, looks, view, names, rules, validation, editors, data, find, audit, chart, file, script, lists, or none")], "Opens a panel under the grid, closing the one open", "the panel now, \"\" for none"),
  GETPANEL: fn([], "The panel open under the grid", "its name, \"\" for none"),
  CYCLEREFERENCE: fn([p("text", "text", "a formula as typed")], "F4: the reference at the end of a formula to its next form (A1, $A$1, A$1, $A1)", "the formula so"),
  COMPLETEFORMULA: fn([p("text", "text", "a formula as typed so far, with its =")], "The function names that complete it, and the call its end is in", "a stem of names, with .CALL the function and .ARG which argument"),
  LISTMENUS: fn([], "The menu bar's items, for finding a command by where it is on the menus", "a stem of \"Menu > Item | command.id | shortcut\" lines"),
  ERRORCHECK: fn([], "Error Checking over the active sheet: formulas that come to an error, are inconsistent with those beside them or omit a number next to their range, and numbers stored as text", "a stem of \"cell check message\" lines, check one of error, inconsistent, text-number, omits"),
  IGNOREERROR: fn([p("cell", "cell"), opt("off", "flag", "1 to check it again")], "Error Checking passes over the cell", "\"\""),
  CALCULATE: fn([], "Calculate Now: every sheet's formulas worked out", "how many there were"),
  CALCULATESHEET: fn([], "Calculate Sheet: the active sheet's formulas worked out", "how many there were"),
  STALECELLS: fn([], "The cells out of date under manual calculation", "a stem of cells, each with its sheet (Sheet1!B2)"),
  SETITERATION: fn([p("max", "number", "at most this many iterations (Excel's 100); 0 turns iteration off"), opt("change", "number", "until no cell moves by more than this (0.001)")], "Iterative calculation: a circular reference worked round instead of #CYCLE!", "the most iterations now, 0 for off"),
  GETITERATION: fn([], "Iterative calculation's settings", "\"max change\", e.g. \"100 0.001\", or 0 for off"),
  EXPORT: fn([], "The whole workbook as its file", "the file's JSON"),
  IMPORT: fn([p("file", "json", "a workbook file, as EXPORT gives one")], "Opens a workbook from its file's JSON (not undoable)", "\"OK\""),
  GETSETUPSCRIPT: fn([], "The setup script (REXX every formula reads)", "the script"),
  SETSETUPSCRIPT: fn([p("script", "text")], "Sets the setup script", "\"OK\""),
  EXECUTESETUPSCRIPT: fn([], "Runs the setup script, and works everything out again", "\"OK\""),
  UNDO: fn([], "Takes back the last change", "1, or 0 when there was none"),
  REDO: fn([], "Does again the last change taken back", "1, or 0 when there was none"),
  CANUNDO: fn([], "How many changes Undo can take back", "the count"),
  CANREDO: fn([], "How many changes Redo can do again", "the count"),
  LISTCOMMANDS: fn([opt("described", "flag", "1: each with its parameters, what it does and what it gives")], "The functions a command may call", "a stem of names, or of {name, signature, params, does, gives}"),
  BATCH_EXECUTE: fn([p("commands", "stem", "of {command, args}; or their JSON"), opt("stop", "flag", "1: the first that fails is the call's error (a control-bus request is then taken back whole)")],
    "Runs each command in turn, by name", "a JSON array of {index, command, success, result or error}"),
};

// "SETCELL(cell, value)": optional parameters in brackets, the rest as ...
export function signature(name: string, d: Described): string {
  return `${name}(${d.params.map((x) => (x.variadic ? `[${x.name}...]` : x.optional ? `[${x.name}]` : x.name)).join(", ")})`;
}
