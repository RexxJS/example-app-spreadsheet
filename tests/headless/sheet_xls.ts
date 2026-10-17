// XLS read (5.6b), on a file made by hand (lib/xls_fixtures.ts, mkxls.py):
// the compound file, the Workbook stream's records: two sheets (one
// hidden), shared strings (one across a CONTINUE, wide there), numbers
// (NUMBER, RK, MULRK), strings, booleans and errors, a date format, a
// merge; formulas written back as text from their tokens (SUM by
// ptgAttr, IF with a string result, another sheet's cell, an absolute
// reference, a unary minus, ROUND), an add-in function kept as its value
// and said; OPENXLS from REXX.
import { compoundStreams, formulaText, readXls } from "./sheet/xls.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { BOOK } from "./lib/xls_fixtures.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("the compound file's streams", [...compoundStreams(fromBase64(BOOK)).keys()], ["Workbook"]);
eq("tokens as text", [formulaText(new Uint8Array([0x1e, 1, 0, 0x1e, 2, 0, 0x03, 0x15]), () => null, () => null), formulaText(new Uint8Array([0x24, 0, 0, 0, 0xc0, 0x14]), () => null, () => null)], ["(1+2)", "A1%"]);
eq("the report", readXls(fromBase64(BOOK)).report, ["1 cell past Z100, the sheet's edge, left out", "1 formula kept as the value cached (an add-in or VBA function)"]);
const wb = new Workbook();
const cmds = new Commands(wb);
const o = await cmds.run(`file.importXls name=Old.xls base64=${BOOK}`);
const v = (k: string) => show(wb.sheet.value(k)), input = (k: string) => wb.sheet.input(k);
eq("opened: its sheets", [o.ok, wb.pages().map((p) => [p.name, p.hiddenSheet])], [true, [["Sales", false], ["Other", true]]]);
eq("strings, one across a CONTINUE; text like a number kept text", [v("A1"), v("A2"), v("A3"), input("A4"), v("A8")], ["Item", "pens", "ink", "'007", "Total"]);
eq("numbers: NUMBER, RK, MULRK (an integer, one over a hundred)", [v("B2"), v("B3"), v("B5"), v("C5")].slice(0, 3).concat([wb.textOf("B2")]), ["3.5", "12", "7", "3.50"]);
eq("booleans and errors; a date", [v("A6"), v("B6"), wb.textOf("A7")], ["TRUE", "#DIV/0!", "2024-01-15"]);
eq("formulas from their tokens", [input("C2"), v("C2"), input("C3"), v("C3"), input("C4"), v("C4"), input("C5"), v("C5")],
  ["=SUM(B2:B3)*2", "31", '=IF(B3>10,"big","small")&"!"', "big!", "=Other!A1+$B$2", "13.5", "=ROUND(-B2,0)", "-4"]);
eq("an add-in function: its value", input("C6"), "99");
eq("a merge", wb.page("Sales")!.merges.length, 1);
eq("not an XLS", (await cmds.run("file.importXls base64=aGVsbG8=")).message, "not an XLS file (not a compound document)");
const said: string[] = [];
await cmds.script(`SAY OPENXLS("${BOOK}")`, (l) => said.push(l));
eq("OPENXLS from REXX", said[0].startsWith("opened Workbook: "), true);

// expect: ok the compound file's streams
// expect: ok tokens as text
// expect: ok the report
// expect: ok opened: its sheets
// expect: ok strings, one across a CONTINUE; text like a number kept text
// expect: ok numbers: NUMBER, RK, MULRK (an integer, one over a hundred)
// expect: ok booleans and errors; a date
// expect: ok formulas from their tokens
// expect: ok an add-in function: its value
// expect: ok a merge
// expect: ok not an XLS
// expect: ok OPENXLS from REXX
