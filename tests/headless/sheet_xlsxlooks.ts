// XLSX's looks, comments and validation (5.5a), read and written: a
// style's font (bold, italic, underline, strike, size, colour), fill,
// borders (by side, style and colour: rgb or a theme's, which stays a
// role), alignment and wrap; a comment's text; data validation (whole
// numbers between, a list, a custom formula; a list from a range reported,
// left out); written back and read again equal.
import { colourOf, readXlsx } from "./sheet/xlsx.ts";
import { parseXml } from "./sheet/xml.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { STYLED } from "./lib/xlsx_fixtures.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

eq("colours: rgb, a theme's (Excel's light-dark first), indexed", [colourOf(parseXml('<c rgb="FF12AB34"/>')), colourOf(parseXml('<c theme="1"/>')), colourOf(parseXml('<c theme="4"/>')), colourOf(parseXml('<c indexed="2"/>'))],
  ["#12ab34", "dark1", "accent1", "#ff0000"]);
eq("the report", readXlsx(fromBase64(STYLED)).report, ["Looks!D2: validation a list from $F$1:$F$3 (a range; the app's lists are values), left out"]);
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run(`file.importXlsx name=Looks.xlsx base64=${STYLED}`);
const n = (k: string) => wb.notes.get(k);
eq("a style's font, fill, alignment", n("A1").look, { bold: true, italic: true, underline: true, size: 19, color: "#ff0000", fill: "#ffff00", align: "center", wrap: true });
eq("strike, a theme's colour, borders by side and style", n("B1").look, { strike: true, color: "accent1", borders: { top: "medium accent2", bottom: "dashed #00ff00", left: "#000000" } });
eq("a number format and an alignment", [n("C1").look, wb.textOf("C1")], [{ align: "end", number: "#,##0.00" }, "1,234.50"]);
eq("drawn", [wb.styleOf("A1"), wb.styleOf("B1")],
  ["font-weight: bold; font-style: italic; text-decoration: underline; font-size: 19px; white-space: normal; text-align: center; background: #ffff00; color: #ff0000; --corner: #d00000",
    "text-align: end; text-decoration: line-through; color: #4472c4; border-top: 2px solid #ed7d31; border-bottom: 1px dashed #00ff00; border-left: 1px solid #000000"]);
eq("a comment's runs together", n("A1").comment, "me: the big one");
eq("validation", [n("A2").validation, n("B2").validation, n("C3").validation, n("D2").validation],
  [{ kind: "whole", min: 1, max: 10, mode: "warn", when: "always", prompt: "one to ten" }, { kind: "list", values: ["Yes", "No"], mode: "stop", when: "always" },
    { kind: "custom", formula: "C2>A2", mode: "stop", when: "always", origin: "C2" }, undefined]);
eq("and kept", [(await cmds.run("cell.set cell=A3 input=50")).message, (await cmds.run("cell.set cell=C3 input=1")).message], ["A3 = 50 (A3 must be a whole number from 1 to 10)", "C3 must make =C2>A2 true"]);
const out = await cmds.run("file.exportXlsx");
const wb2 = new Workbook(), cmds2 = new Commands(wb2);
await cmds2.run(`file.importXlsx base64=${out.output}`);
const n2 = (k: string) => wb2.notes.get(k);
eq("written and read again: the looks", [n2("A1").look, n2("B1").look, n2("C1").look], [n("A1").look, n("B1").look, n("C1").look]);
eq("the comment and the rules", [n2("A1").comment, n2("A2").validation, n2("B2").validation, n2("C3").validation], [n("A1").comment, n("A2").validation, n("B2").validation, n("C3").validation]);
const cmds3 = new Commands(new Workbook());
await cmds3.run("cell.set cell=B2 input=x");
await cmds3.run('cell.comment cell=B2 text="a note"');
await cmds3.run("format.border range=B2 edge=all style=thick color=accent6");
const out3 = await cmds3.run("file.exportXlsx");
const wb4 = new Workbook(), cmds4 = new Commands(wb4);
await cmds4.run(`file.importXlsx base64=${out3.output}`);
eq("a workbook's own: its comment and a theme's border", [wb4.notes.get("B2").comment, wb4.notes.get("B2").look?.borders?.left], ["a note", "thick accent6"]);

// expect: ok colours: rgb, a theme's (Excel's light-dark first), indexed
// expect: ok the report
// expect: ok a style's font, fill, alignment
// expect: ok strike, a theme's colour, borders by side and style
// expect: ok a number format and an alignment
// expect: ok drawn
// expect: ok a comment's runs together
// expect: ok validation
// expect: ok and kept
// expect: ok written and read again: the looks
// expect: ok the comment and the rules
// expect: ok a workbook's own: its comment and a theme's border
