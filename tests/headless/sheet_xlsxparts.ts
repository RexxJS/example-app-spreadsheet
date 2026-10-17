// XLSX's conditional formatting, tables, and what the app does not model
// (5.5b): cfRules as the app's rules (by priority, their dxf styles, Stop
// If True, scales, bars, icons, periods), one of a kind there is not
// reported; a table part as the sheet's table (totals row, its functions,
// style, stripes); charts, pictures, pivots, x14 sparklines reported; a
// theme and document properties carried over unchanged; all written back
// and read again equal.
import { dxfStyle, styleDxf } from "./sheet/xlsxparts.ts";
import { parseXml } from "./sheet/xml.ts";
import { readXlsx } from "./sheet/xlsx.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { unzip, utf8Text } from "./sheet/zip.ts";
import { PARTS } from "./lib/xlsx_fixtures.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("a dxf as a style and back", [dxfStyle(parseXml('<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')),
  styleDxf("background: #ffc7ce; color: accent1; font-weight: bold")],
  ["background: #ffc7ce; color: #9c0006; font-weight: bold", '<dxf><font><b/><color theme="4"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>']);
eq("the report", readXlsx(fromBase64(PARTS)).report, ["Rules!A2:A3: a dataBarX rule, left out", "1 chart left out", "1 picture left out", "1 pivot table left out", "1 sheet's sparkline group left out"]);
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run(`file.importXlsx name=Parts.xlsx base64=${PARTS}`);
const rules = (await cmds.run("format.rules")).output;
eq("its rules, by priority", rules, ["1. B2:B3: formula =$A2=\"ink\" (stop if true)", "2. B2:B3: cell value > 10", "3. A2:A3: text contains en", "4. A2:A3: colour scale #f8696b, #63be7b",
  "5. A2:A3: data bars #638ec6", "6. A2:A3: icon set arrows reversed", "7. A2:A3: a date in the last 7 days", "8. A2:A3: below average"]);
eq("drawn: the formula's rule stops the other on B3, the table's band on B2", [wb.styleOf("B3"), wb.styleOf("B2")], ["text-align: end; color: #4472c4", "text-align: end; background: #d9e1f2"]);
eq("its table: the totals row, its functions, style, stripes", (await cmds.run("table.get name=Stock")).output,
  { name: "Stock", range: "A1:B4", headers: true, columns: ["Item", "Qty"], types: {}, descriptions: {}, sort: null, totals: { Item: "label", Qty: "sum" }, banded: true, style: "Medium" });
eq("the totals row's formula works", show(wb.sheet.value("B4")), "15");
const out = await cmds.run("file.exportXlsx");
const parts = unzip(fromBase64(String(out.output)));
eq("carried over unchanged: the theme, the document properties", [utf8Text(parts.get("xl/theme/theme1.xml")!), utf8Text(parts.get("docProps/core.xml")!).includes("Paul"), parts.has("xl/charts/chart1.xml")],
  ['<a:theme xmlns:a="a" name="Mine"/>', true, false]);
const wb2 = new Workbook(), cmds2 = new Commands(wb2);
await cmds2.run(`file.importXlsx base64=${out.output}`);
eq("written and read again: the rules, the table", [(await cmds2.run("format.rules")).output, (await cmds2.run("table.get name=Stock")).output], [rules, (await cmds.run("table.get name=Stock")).output]);
const cmds3 = new Commands(new Workbook());
await cmds3.run("cell.set cell=A1 input=1");
await cmds3.run("insert.sparkline range=A1:A1 at=B1");
await cmds3.run("chart.insert kind=bar range=A1:A1");
eq("what it cannot write, said", (await cmds3.run("file.exportXlsx")).message.replace(/^[^;]*; /, ""),
  "1 chart not written (the app draws its own, not DrawingML); 1 sparkline written as their text (Excel has no SPARKLINE function)");

// expect: ok a dxf as a style and back
// expect: ok the report
// expect: ok its rules, by priority
// expect: ok drawn: the formula's rule stops the other on B3, the table's band on B2
// expect: ok its table: the totals row, its functions, style, stripes
// expect: ok the totals row's formula works
// expect: ok carried over unchanged: the theme, the document properties
// expect: ok written and read again: the rules, the table
// expect: ok what it cannot write, said
