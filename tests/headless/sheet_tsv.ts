// CSV and TSV (5.1): the delimiter detected (comma, tab, semicolon, bar)
// or given; a byte-order mark left out; bytes in UTF-8, UTF-16 and
// Windows-1252; export with any delimiter, a field holding it quoted, the
// cells as they show; IMPORTCSV and EXPORTCSV from REXX.
import { decode, detectDelimiter, fromBase64, fromCsv, toBase64, toCsv } from "./sheet/csv.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("the delimiter detected", [detectDelimiter("a,b,c\n1,2,3"), detectDelimiter("a\tb\n1\t2"), detectDelimiter("a;b\n1,5;2,5"), detectDelimiter("x|y\n1|2"), detectDelimiter('"a,b";c\n1;2'), detectDelimiter("one\ntwo")],
  [",", "\t", ";", "|", ";", ","]);
eq("a byte-order mark left out", fromCsv("﻿a,b\n1,2"), [["a", "b"], ["1", "2"]]);
const bytes = (...xs: number[]) => xs;
eq("encodings", [decode(bytes(0xef, 0xbb, 0xbf, 0x63, 0x61, 0x66, 0xc3, 0xa9)), decode(bytes(0xff, 0xfe, 0x68, 0x00, 0xe9, 0x00)), decode(bytes(0xfe, 0xff, 0x00, 0x68, 0x20, 0xac)),
  decode(bytes(0x63, 0x61, 0x66, 0xe9, 0x80)), decode(bytes(0x63, 0xe9), "latin1")], ["café", "hé", "h€", "café€", "cé"]);
eq("base64 both ways", [toBase64([104, 105, 33]), [...fromBase64("aGkh")], [...fromBase64(toBase64([1, 2, 3, 4, 5]))]], ["aGkh", [104, 105, 33], [1, 2, 3, 4, 5]]);
eq("written with any delimiter, quoting what holds it", [toCsv([["a,b", "c"]], "\t"), toCsv([["a\tb", "c;d"]], "\t"), toCsv([["1,5", "x"]], ";")], ["a,b\tc", '"a\tb"\tc;d', '1,5;x']);

const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
eq("TSV in, detected", [await message('file.importCsv at=A1 text="name\tqty\npens\t3"'), show(wb.sheet.value("B2"))], ["imported 2 rows at A1, split by tab: 4 cells set, 4 changed", "3"]);
const latin = toBase64([0x6e, 0x61, 0x6d, 0x65, 0x3b, 0x70, 0x72, 0x69, 0x63, 0x65, 0x0a, 0x63, 0x61, 0x66, 0xe9, 0x3b, 0x32, 0x0a]);
eq("bytes in Windows-1252, semicolons", [await message(`file.importCsv at=D1 base64=${latin}`), show(wb.sheet.value("D2")), show(wb.sheet.value("E2"))],
  ["imported 2 rows at D1, split by semicolon: 4 cells set, 4 changed", "café", "2"]);
await cmds.run("format.number range=B2 format=0.00");
eq("out as TSV, the cells as they show", (await cmds.run("file.exportCsv range=A1:B2 delimiter=tab")).output, "name\tqty\npens\t3.00");
eq("refusals", [await message("file.importCsv at=A1"), await message("file.importCsv at=A1 base64=AA== encoding=ebcdic"), await message("file.exportCsv delimiter=x")],
  ["Import CSV takes text= or base64=", "ebcdic is not an encoding (auto, utf-8, utf-16le, utf-16be, windows-1252, latin1)", "x is not a delimiter (comma, tab, semicolon, bar, space)"]);
const said: string[] = [];
await cmds.script('SAY IMPORTCSV("G1", "a|b")\nSAY IMPORTCSV("G2", "Y2Fm6Q==", "comma", "windows-1252")\nSAY EXPORTCSV("G1:H2", "semicolon")', (l) => said.push(l));
eq("IMPORTCSV and EXPORTCSV from REXX", said, ["imported 1 rows at G1, split by bar: 2 cells set, 2 changed", "imported 1 rows at G2: 1 cells set, 1 changed", "a;b\ncafé;"]);

// expect: ok the delimiter detected
// expect: ok a byte-order mark left out
// expect: ok encodings
// expect: ok base64 both ways
// expect: ok written with any delimiter, quoting what holds it
// expect: ok TSV in, detected
// expect: ok bytes in Windows-1252, semicolons
// expect: ok out as TSV, the cells as they show
// expect: ok refusals
// expect: ok IMPORTCSV and EXPORTCSV from REXX
