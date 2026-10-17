// Hyperlinks (6.11a): Insert Link over a range (its text in an empty cell,
// the Hyperlink style), a place in the workbook followed (its sheet
// activated, the place selected), a URL handed to the opener, HYPERLINK's
// shown name and its link followed, refusals, Remove Hyperlinks, undo, the
// tooltip, the sheet file and XLSX round trips, and the REXX functions.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const opened: string[] = [];
wb.opener = (url) => { opened.push(url); return url.includes("refused") ? "this app may not open it" : ""; };
await cmds.run("formulas.use language=Excel");
await cmds.run("sheet.add name=Costs");
await cmds.run("sheet.activate name=Sheet1");
await cmds.run("cell.set cell=B2 input=Kept");
const o = await cmds.run("cell.hyperlink range=A1:B2 target=Costs!C3 tip=\"the costs\"");
eq("Insert Link over a range", [o.ok, o.message, show(wb.sheet.value("A1")), show(wb.sheet.value("B2")), wb.notes.get("A2").hyperlink, wb.notes.get("B2").cellStyle, wb.styleOf("A1").includes("underline")],
  [true, "A1:B2 links to Costs!C3", "Costs!C3", "Kept", { target: "Costs!C3", tip: "the costs" }, "Hyperlink", true]);
const f = await cmds.run("cell.followHyperlink cell=A1");
eq("a place in the workbook followed", [f.message, f.output, wb.active, wb.selection], ["went to Costs!C3", "Costs!C3", "Costs", "C3"]);
await cmds.run("sheet.activate name=Sheet1");
await cmds.run("names.define name=Spot range=Costs!D4:D5");
await cmds.run('cell.hyperlink range=C1 target="#Spot" text=Here');
eq("a name, with Excel's #", [show(wb.sheet.value("C1")), (await cmds.run("cell.followHyperlink cell=C1")).output], ["Here", "Costs!D4:D5"]);
await cmds.run("sheet.activate name=Sheet1");
await cmds.run("cell.hyperlink range=D1 target=https://example.org/a text=Example");
const u = await cmds.run("cell.followHyperlink cell=D1");
eq("a URL handed to the opener", [u.message, opened, show(wb.sheet.value("D1"))], ["opened https://example.org/a", ["https://example.org/a"], "Example"]);
await cmds.run("cell.set cell=E1 input=https://example.org/b");
await cmds.run('cell.set cell=E2 input="=HYPERLINK(E1, ""Docs"")"');
await cmds.run('cell.set cell=E3 input="=HYPERLINK(""Costs!A1"")"');
eq("HYPERLINK: its name shown, its link followed", [show(wb.sheet.value("E2")), show(wb.sheet.value("E3")), (await cmds.run("cell.followHyperlink cell=E2")).output, (await cmds.run("cell.followHyperlink cell=E3")).output],
  ["Docs", "Costs!A1", "https://example.org/b", "Costs!A1"]);
await cmds.run("sheet.activate name=Sheet1");
const no = await cmds.run("cell.hyperlink range=F1 target=Nowhere!A1"), none = await cmds.run("cell.followHyperlink cell=F9");
await cmds.run("cell.hyperlink range=F2 target=https://refused.example/");
eq("refusals", [no.message, none.message, (await cmds.run("cell.followHyperlink cell=F2")).message, (await cmds.run("cell.hyperlink range=F3 target=NOT_A_NAME")).message],
  ["no sheet Nowhere for the link Nowhere!A1", "F9 has no link", "this app may not open it", "the link NOT_A_NAME is not a cell, a range or a name in the workbook"]);
eq("the tooltip", [wb.tooltipOf("A1"), wb.tooltipOf("D1")], ["the costs", "https://example.org/a\nOpen Hyperlink to follow it"]);
// Saved and opened again; written as XLSX and read back.
const again = new Workbook(), cmds2 = new Commands(again);
const saved = (await cmds.run("file.export")).output as string;
await cmds2.execute("file.import", { text: saved });
eq("saved with the file", [again.notes.get("A1").hyperlink, again.notes.get("D1").hyperlink], [{ target: "Costs!C3", tip: "the costs" }, { target: "https://example.org/a" }]);
const x = await cmds.run("file.exportXlsx");
const back = new Workbook(), cmds3 = new Commands(back);
await cmds3.run(`file.importXlsx name=L.xlsx base64=${x.output}`);
eq("an XLSX file's", [back.notes.get("A1").hyperlink, back.notes.get("D1").hyperlink, back.notes.get("C1").hyperlink], [{ target: "Costs!C3", tip: "the costs" }, { target: "https://example.org/a" }, { target: "Spot" }]);
const r = await cmds.run("cell.removeHyperlink range=A1:D2");
eq("Remove Hyperlinks, and undo", [r.message, wb.notes.get("A1").hyperlink, wb.notes.get("A1").cellStyle, (await cmds.run("edit.undo")).ok, wb.notes.get("A1").hyperlink?.target], ["6 links removed from A1:D2", undefined, undefined, true, "Costs!C3"]);
const said: string[] = [];
await cmds.script('SAY SETHYPERLINK("G1", "Costs!B2", "Go")\nSAY GETHYPERLINK("G1")\nSAY REMOVEHYPERLINK("G1")\nSAY "[" || GETHYPERLINK("G1") || "]"\nLET r = SETHYPERLINK("G1", "Costs!B2")\nSAY FOLLOWHYPERLINK("G1")', (l) => said.push(l));
eq("from REXX", said, ["G1:G1 links to Costs!B2", "Costs!B2", "1 link removed from G1:G1", "[]", "Costs!B2"]);

// expect: ok Insert Link over a range
// expect: ok a place in the workbook followed
// expect: ok a name, with Excel's #
// expect: ok a URL handed to the opener
// expect: ok HYPERLINK: its name shown, its link followed
// expect: ok refusals
// expect: ok the tooltip
// expect: ok saved with the file
// expect: ok an XLSX file's
// expect: ok Remove Hyperlinks, and undo
// expect: ok from REXX
