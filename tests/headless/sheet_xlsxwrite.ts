// XLSX write (5.4), and the round trip: a workbook written as XLSX, read
// back (xlsx.ts), equal: its sheets (a hidden one), inputs (formulas,
// numbers, text, text like a number, booleans), the values formulas came
// to as Excel would read them, number formats, merges, widths, heights,
// frozen panes, hidden rows and columns, names, the 1904 date system;
// RexxJS formulas written as their values, said; SAVEXLSX from REXX.
import { unzip, utf8Text } from "./sheet/zip.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const run = async (line: string) => { const o = await cmds.run(line); if (!o.ok) print(`refused: ${line}: ${o.message}`); };
await run("formulas.use language=Excel");
await run("sheet.rename name=Sheet1 to=Sales");
for (const [k, v] of [["A1", "Item"], ["B1", "Qty"], ["A2", "pens"], ["B2", "3"], ["A3", "'007"], ["B3", "1.5"], ["C2", "=B2*2"], ["C3", '="x"&A2'], ["D1", "=1>0"], ["D2", "=1/0"], ["E1", "45306"], ["F1", "a & <b>"]]) await run(`cell.set cell=${k} input="${v.replace(/"/g, '""')}"`);
await run("format.number range=E1 format=yyyy-mm-dd");
await run("format.number range=B2:B3 format=0.00");
await run("cells.merge range=A5:C5");
await run("columns.width range=A:A width=120");
await run("rows.height range=2:2 height=30");
await run("view.freezePanes cell=B2");
await run("rows.hide range=4:4");
await run("columns.hide range=G:G");
await run("names.define name=QTY range=Sales!B2:B3");
await run("names.define name=RATE range=D1 scope=sheet");
await run('names.define name=TWICE formula="=LAMBDA(x,x*2)"');
await run("sheet.add name=Other");
await run('cell.set cell=A1 input="=Sales!B2+SUM(QTY)"');
await run("sheet.hide name=Other");
await run("sheet.activate name=Sales");
await run("workbook.dateSystem system=1904");
const o = await cmds.run("file.exportXlsx");
const parts = unzip(fromBase64(String(o.output)));
eq("the archive's parts", [...parts.keys()], ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]);
const sheet1 = utf8Text(parts.get("xl/worksheets/sheet1.xml")!);
eq("cells as Excel writes them", [/<c r="C2"><f>B2\*2<\/f><v>6<\/v><\/c>/.test(sheet1), /<c r="C3" t="str"><f>&quot;x&quot;&amp;A2<\/f><v>xpens<\/v><\/c>/.test(sheet1), /<c r="D1" t="b"><f>1&gt;0<\/f><v>1<\/v><\/c>/.test(sheet1),
  /<c r="D2" t="e"><f>1\/0<\/f><v>#DIV\/0!<\/v><\/c>/.test(sheet1), /<pane xSplit="1" ySplit="1" topLeftCell="B2"/.test(sheet1)], [true, true, true, true, true]);
eq("names, absolute, a sheet's own by its number", /<definedName name="QTY">Sales!\$B\$2:\$B\$3<\/definedName><definedName name="TWICE">LAMBDA\(x,x\*2\)<\/definedName><definedName name="RATE" localSheetId="0">Sales!\$D\$1<\/definedName>/.test(utf8Text(parts.get("xl/workbook.xml")!)), true);
const wb2 = new Workbook(), cmds2 = new Commands(wb2);
const back = await cmds2.run(`file.importXlsx name=Round.xlsx base64=${o.output}`);
const v = (k: string) => show(wb2.sheet.value(k));
eq("read back: the sheets, the inputs", [back.ok, wb2.pages().map((p) => [p.name, p.hiddenSheet]), ["A1", "B2", "A3", "C2", "C3", "D1", "D2", "F1"].map((k) => wb2.sheet.input(k))],
  [true, [["Sales", false], ["Other", true]], ["Item", "3", "'007", "=B2*2", '="x"&A2', "=1>0", "=1/0", "a & <b>"]]);
eq("the values, the formats", [v("C3"), v("D2"), wb2.textOf("E1"), wb2.textOf("B3"), wb2.dateSystem], ["xpens", "#DIV/0!", "2028-01-16", "1.50", 1904]);
const p2 = wb2.page("Sales")!;
eq("the view", [p2.merges.map(([a, b]) => `${a.col},${a.row}:${b.col},${b.row}`), p2.widths.get(0), p2.heights.get(1), p2.frozen, p2.frozenCols, [...p2.hiddenRows], [...p2.hiddenCols]],
  [["0,4:2,4"], 120, 30, 1, 1, [3], [6]]);
eq("the names", [wb2.names(), wb2.sheet.names()], [[["QTY", "Sales!B2:B3"], ["TWICE", "=LAMBDA(x,x*2)"]], [["RATE", "D1"]]]);
await cmds2.run("sheet.activate name=Other");
eq("the other sheet reads the first by name", v("A1"), "7.5");
const wb3 = new Workbook(), cmds3 = new Commands(wb3);
await cmds3.run("formulas.use language=RexxJS");
await cmds3.run('cell.set cell=A1 input="=2+3"');
const r = await cmds3.run("file.exportXlsx");
const wb4 = new Workbook(), cmds4 = new Commands(wb4);
await cmds4.run(`file.importXlsx base64=${r.output}`);
eq("RexxJS formulas: their values, said", [r.message.endsWith("RexxJS formulas are not Excel's: their values written, not the formulas"), wb4.sheet.input("A1")], [true, "5"]);
const said: string[] = [];
await cmds.script('x = SAVEXLSX()\nIF LENGTH(x) > 1000 THEN SAY LEFT(x, 4) "long"', (l) => said.push(l));
eq("SAVEXLSX from REXX: base64 of a ZIP", said, ["UEsD long"]);

// expect: ok the archive's parts
// expect: ok cells as Excel writes them
// expect: ok names, absolute, a sheet's own by its number
// expect: ok read back: the sheets, the inputs
// expect: ok the values, the formats
// expect: ok the view
// expect: ok the names
// expect: ok the other sheet reads the first by name
// expect: ok RexxJS formulas: their values, said
// expect: ok SAVEXLSX from REXX: base64 of a ZIP
