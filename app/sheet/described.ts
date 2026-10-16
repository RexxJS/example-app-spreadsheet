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
  FREEZEPANES: fn([p("rows", "number"), opt("columns", "number", "0: only rows freeze")], "Freezes the top rows", "\"OK\""),
  UNFREEZEPANES: fn([], "Unfreezes the rows", "\"OK\""),
  GETFROZENPANES: fn([], "How many rows are frozen", "{ROWS, COLUMNS}"),
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
  SPARKLINE_RANGE: fn([RANGE], "A range's numbers as a sparkline", "text of block characters"),
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
