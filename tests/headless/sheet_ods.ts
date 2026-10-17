// ODS read (5.6a), on a file made by hand (lib/xlsx_fixtures.ts's ODS):
// its tables as sheets, each cell's cached value (strings with their
// spaces, floats, a percentage, booleans, a date, a duration, text like a
// number kept text), repeated rows and columns, a span as a merge, a
// formula as its value, said; OPENODS from REXX.
import { readOds } from "./sheet/ods.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { ODS } from "./lib/xlsx_fixtures.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("the report", readOds(fromBase64(ODS)).report, ["1 formula read as the value cached"]);
const wb = new Workbook();
const cmds = new Commands(wb);
const o = await cmds.run(`file.importOds name=Stock.ods base64=${ODS}`);
const v = (k: string) => show(wb.sheet.value(k));
eq("opened", [o.message, wb.pages().map((p) => p.name)], ["opened Stock: 16 cells, Excel formulas; 1 formula read as the value cached", ["Stock", "Two"]]);
eq("the values", [v("A1"), v("A2"), v("B2"), v("C2"), wb.textOf("D2"), v("A3"), v("A4"), wb.sheet.input("B3"), v("C4"), wb.textOf("A5"), wb.textOf("B5")],
  ["Item", "pens  red", "3", "6", "25%", "TRUE", "TRUE", "'007", "007", "2024-01-15", "36:00:00"]);
eq("a span as a merge", wb.page("Stock")!.merges.length, 1);
eq("not an ODS", (await cmds.run("file.importOds base64=aGVsbG8=")).message, "not a ZIP archive (an XLSX or ODS file is one)");
const said: string[] = [];
await cmds.script(`SAY OPENODS("${ODS}")`, (l) => said.push(l));
eq("OPENODS from REXX", said, ["opened Workbook: 16 cells, Excel formulas; 1 formula read as the value cached"]);

// expect: ok the report
// expect: ok opened
// expect: ok the values
// expect: ok a span as a merge
// expect: ok not an ODS
// expect: ok OPENODS from REXX
