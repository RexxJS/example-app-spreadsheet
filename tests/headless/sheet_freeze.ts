// Frozen columns (3.8): the first columns frozen, Freeze Panes at a cell
// (the rows above and columns left of it), unfrozen; undone and saved with
// the file; FREEZEPANES and GETFROZENPANES from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const msg = async (id: string, p: Record<string, string>) => (await cmds.execute(id, p)).message;
eq("the first column frozen, then two", [await msg("view.freezeCols", { count: "1" }), wb.frozenCols(), await msg("view.freezeCols", { count: "2" }), wb.frozenCols()], ["froze 1 column", 1, "froze 2 columns", 2]);
eq("Freeze Panes at C4: three rows and two columns", [await msg("view.freezePanes", { cell: "C4" }), wb.frozenRows(), wb.frozenCols()], ["froze 3 rows and 2 columns", 3, 2]);
eq("saved with the file", (writeSheetFile(wb) as { view: { frozenRows: number; frozenCols?: number } }).view, { hiddenRows: [], hiddenCols: [], frozenRows: 3, frozenCols: 2, rowHeights: [], colWidths: [] } as never);
await cmds.run("edit.undo");
eq("undone", [wb.frozenRows(), wb.frozenCols()], [0, 2]);
eq("at A1 everything unfrozen; too many refused", [await msg("view.freezePanes", { cell: "A1" }), wb.frozenCols(), await msg("view.freezeCols", { count: "9" }), await msg("view.freezePanes", { cell: "K2" })],
  ["unfroze panes", 0, "9 is not a number of columns to freeze (0 to 8: the grid shows ten)", "K2: at most 8 columns freeze (the grid shows ten)"]);
const said: string[] = [];
eq("FREEZEPANES and GETFROZENPANES from REXX", [(await cmds.script('CALL FREEZEPANES 1, 3\nf = GETFROZENPANES()\nSAY f.ROWS f.COLUMNS\nCALL UNFREEZEPANES\nf = GETFROZENPANES()\nSAY f.ROWS f.COLUMNS', (l) => said.push(l))).message, said], ["ran", ["1 3", "0 0"]]);

// expect: ok the first column frozen, then two
// expect: ok Freeze Panes at C4: three rows and two columns
// expect: ok saved with the file
// expect: ok undone
// expect: ok at A1 everything unfrozen; too many refused
// expect: ok FREEZEPANES and GETFROZENPANES from REXX
