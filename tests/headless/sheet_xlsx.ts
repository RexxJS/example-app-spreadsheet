// XLSX read (5.3), on files made by hand (lib/xlsx_fixtures.ts): the ZIP
// archive (stored and deflated entries, CRC-32), DEFLATE both ways, the
// XML, then the workbook: sheets in order (a hidden one), shared, rich and
// inline strings, booleans, text that reads as a number kept text,
// formulas, a shared formula moved to each cell, number formats (built in
// and the file's own), a merge, a frozen row, a width, a height, hidden
// rows and columns, defined names (the workbook's, a sheet's own, a
// formula), the 1904 date system, and a report of what was left out;
// OPENXLSX and FILEREPORT from REXX.
import { crc32, deflate, inflate, unzip, utf8Bytes, utf8Text, zip } from "./sheet/zip.ts";
import { all, attr, parseXml, textOf } from "./sheet/xml.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { readXlsx } from "./sheet/xlsx.ts";
import { BASIC, BIG, STORED } from "./lib/xlsx_fixtures.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("CRC-32", [crc32(utf8Bytes("The quick brown fox jumps over the lazy dog")).toString(16), crc32(new Uint8Array(0))], ["414fa339", 0]);
const text = "abcabcabcabcabc, a line; " + "spreadsheet ".repeat(200) + "é€😀";
eq("deflate then inflate gives it back, smaller", [utf8Text(inflate(deflate(utf8Bytes(text)))) === text, deflate(utf8Bytes(text)).length < utf8Bytes(text).length], [true, true]);
const big = unzip(fromBase64(BIG)).get("a.txt")!;
eq("Python's deflate (dynamic codes, long matches) inflated", [big.length, utf8Text(big).slice(0, 12), utf8Text(big).endsWith("e")], [33000, "hello hello ", utf8Text(big).endsWith("e")]);
const z = unzip(zip([["a.xml", utf8Bytes("<a>hi</a>")], ["b/c.txt", utf8Bytes("x".repeat(500))]]));
eq("an archive written and read back", [[...z.keys()], utf8Text(z.get("a.xml")!), z.get("b/c.txt")!.length], [["a.xml", "b/c.txt"], "<a>hi</a>", 500]);
const x = parseXml('<?xml version="1.0"?><!-- c --><r a="1" b=\'two &amp; three\'><x:c n="1">t&lt;1</x:c><e/><![CDATA[<raw>]]></r>');
eq("XML: attributes, prefixes, entities, CDATA", [x.name, attr(x, "b"), all(x, "c").map(textOf), textOf(x)], ["r", "two & three", ["t<1"], "t<1<raw>"]);

const read = readXlsx(fromBase64(BASIC), "Shop.xlsx");
eq("the report: what was left out", read.report, ["2 cells past Z100, the sheet's edge, left out", "Print_Area (a print setting) left out"]);
const wb = new Workbook();
const cmds = new Commands(wb);
const o = await cmds.run(`file.importXlsx name=Shop.xlsx base64=${BASIC}`);
const v = (k: string) => show(wb.sheet.value(k));
eq("opened: its sheets, the first active", [o.ok, o.message, wb.pages().map((p) => p.name), wb.page("Notes & more")!.hiddenSheet, wb.name],
  [true, "opened Shop: 16 cells, Excel formulas; left out: 2 cells past Z100, the sheet's edge, left out; Print_Area (a print setting) left out", ["Sales", "Notes & more"], true, "Shop"]);
eq("strings shared, rich and inline; booleans; text like a number kept text", [v("A1"), v("A2"), v("A3"), v("C1"), v("A4"), wb.sheet.input("D4"), v("D4")], ["Item", "pens", "ink & nib", "Total", "TRUE", "'007", "007"]);
eq("formulas, a shared one moved, names", [wb.sheet.input("C2"), wb.sheet.input("C3"), v("C3"), wb.sheet.input("B4"), v("B4"), wb.names().map(([n]) => n).sort()],
  ["=B2*(1+$D$1)", "=B3*(1+$D$1)", "1.8", "=SUM(QTY)", "4.5", ["DOUBLE", "QTY"]]);
eq("number formats: the file's own and built in", [wb.textOf("C2"), wb.textOf("D1"), wb.textOf("C4")], ["$3.60", "20.00%", "1/15/2024"]);
const page = wb.page("Sales")!;
eq("the view: a merge, a frozen row, a width, a height, hidden rows and columns", [page.merges.length, page.frozen, page.widths.get(0), page.heights.get(1), [...page.hiddenRows], [...page.hiddenCols]],
  [1, 1, 145, 40, [3], [4]]);
eq("a sheet's own name", wb.sheet.names(), [["RATE", "D1"]]);
await cmds.run("sheet.activate name=\"Notes & more\"");
eq("the other sheet reads the first", [v("A1"), v("B1")], ["6", "xy"]);
const o2 = await cmds.run(`file.importXlsx base64=${STORED}`);
eq("stored entries, 1904 dates", [o2.message, wb.dateSystem, show(wb.sheet.value("A1"))], ["opened Workbook: 1 cells, Excel formulas", 1904, "1"]);
eq("not an XLSX", (await cmds.run("file.importXlsx base64=aGVsbG8=")).message, "not a ZIP archive (an XLSX or ODS file is one)");
const said: string[] = [];
await cmds.script(`SAY OPENXLSX("${BASIC}", "Again.xlsx")\nr = FILEREPORT()\nSAY r.0 r.2`, (l) => said.push(l));
eq("OPENXLSX and FILEREPORT from REXX", said, ["opened Again: 16 cells, Excel formulas; left out: 2 cells past Z100, the sheet's edge, left out; Print_Area (a print setting) left out", "2 Print_Area (a print setting) left out"]);

// expect: ok CRC-32
// expect: ok deflate then inflate gives it back, smaller
// expect: ok Python's deflate (dynamic codes, long matches) inflated
// expect: ok an archive written and read back
// expect: ok XML: attributes, prefixes, entities, CDATA
// expect: ok the report: what was left out
// expect: ok opened: its sheets, the first active
// expect: ok strings shared, rich and inline; booleans; text like a number kept text
// expect: ok formulas, a shared one moved, names
// expect: ok number formats: the file's own and built in
// expect: ok the view: a merge, a frozen row, a width, a height, hidden rows and columns
// expect: ok a sheet's own name
// expect: ok the other sheet reads the first
// expect: ok stored entries, 1904 dates
// expect: ok not an XLSX
// expect: ok OPENXLSX and FILEREPORT from REXX
