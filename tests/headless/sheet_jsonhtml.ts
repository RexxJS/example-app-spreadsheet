// JSON and HTML (5.2): JSON rows (an array of objects, their keys the
// header row in the order first seen; an array of arrays; an object
// holding one array), an HTML page's table (th and td, tags and entities
// read, a colspan's cells), the cells out as an HTML table with their
// looks; from REXX.
import { fromHtml, fromJson, toHtml, unentity } from "./sheet/tabular.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

eq("JSON objects: their keys the header, in the order first seen", fromJson('[{"item":"pens","qty":3},{"qty":1,"item":"ink","note":null,"ok":true}]'),
  [["item", "qty", "note", "ok"], ["pens", "3", "", ""], ["ink", "1", "", "TRUE"]]);
eq("arrays, an object holding one, nested values as JSON", [fromJson("[[1,2],[\"a\",[3]]]"), fromJson('{"rows":[{"a":{"b":1}}]}')], [[["1", "2"], ["a", "[3]"]], [["a"], ['{"b":1}']]]);
eq("entities", unentity("a &amp; b &lt;c&gt; &#233; &#x20AC; &nbsp;&bogus;"), "a & b <c> é €  &bogus;".replace(" ", " "));
const page = `<html><body><p>intro</p><table id="t1"><tr><th>Name</th><th>Note</th></tr><tr><td><b>pens</b></td><td>red &amp; blue<br>fine</td></tr></table>
<table><thead><tr><th colspan="2">Wide</th><th>C</th></tr></thead><tbody><tr><td>1</td><td>2</td><td>3</td></tr></tbody></table></body></html>`;
eq("an HTML page's tables", [fromHtml(page), fromHtml(page, 2)], [[["Name", "Note"], ["pens", "red & blue\nfine"]], [["Wide", "", "C"], ["1", "2", "3"]]]);
eq("a table there is not", (() => { try { return fromHtml(page, 3); } catch (e) { return (e as Error).message; } })(), "the page has 2 tables, not 3");
eq("an HTML table out", toHtml([[{ text: "a<b", style: "font-weight: bold" }], [{ text: "3", style: "text-align: end; --fill: 0.5" }]], true),
  '<table>\n<tr><th style="font-weight: bold">a&lt;b</th></tr>\n<tr><td style="text-align: right">3</td></tr>\n</table>');

const wb = new Workbook();
const cmds = new Commands(wb);
const message = async (line: string) => (await cmds.run(line)).message;
eq("the commands", [await message(`file.importJson at=A1 text="${'[{"item":"pens","qty":3}]'.replace(/"/g, '""')}"`), show(wb.sheet.value("B2")),
  await message(`file.importHtml at=D1 table=2 text="${page.replace(/"/g, '""')}"`), show(wb.sheet.value("F2"))],
  ["imported 2 rows of JSON at A1: 4 cells set, 4 changed", "3", "imported 2 rows of HTML at D1: 6 cells set, 6 changed", "3"]);
await cmds.run("format.font range=A1 bold=1");
eq("exported", (await cmds.run("file.exportHtml range=A1:B2 header=1")).output,
  '<table>\n<tr><th style="font-weight: bold">item</th><th>qty</th></tr>\n<tr><td>pens</td><td style="text-align: right">3</td></tr>\n</table>');
eq("refusals", [await message("file.importJson at=A1 text={}"), await message("file.importJson at=A1 text=nope")], ["JSON rows are an array: of objects, or of arrays", "not JSON: unexpected token: 'nope'"]);
const said: string[] = [];
await cmds.script('SAY IMPORTJSON("H1", \'[[1,2]]\')\nSAY IMPORTHTML("H3", "<table><tr><td>x</td></tr></table>")\nSAY EXPORTHTML("H3")', (l) => said.push(l));
eq("from REXX", said, ["imported 1 rows of JSON at H1: 2 cells set, 2 changed", "imported 1 rows of HTML at H3: 1 cells set, 1 changed", "<table>\n<tr><td>x</td></tr>\n</table>"]);

// expect: ok JSON objects: their keys the header, in the order first seen
// expect: ok arrays, an object holding one, nested values as JSON
// expect: ok entities
// expect: ok an HTML page's tables
// expect: ok a table there is not
// expect: ok an HTML table out
// expect: ok the commands
// expect: ok exported
// expect: ok refusals
// expect: ok from REXX
