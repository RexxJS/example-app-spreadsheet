// PDF export (5.8a): the cells laid out on pages as Excel prints them
// (down, then over), each as it shows in Helvetica (bold, its colour,
// numbers right), fills, borders, gridlines and headings when asked, the
// foot's page numbers; a well-formed PDF (its cross-reference table's
// offsets each at an object); EXPORTPDF from REXX.
import { paginate, textWidth } from "./sheet/pdf.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

eq("Helvetica's widths", [textWidth("Hello", 10), textWidth("1234", 12), textWidth("é", 10)], [22.78, 26.688, 5.56]);
eq("pages: what fits, then the next", [paginate([10, 10, 10, 10], 25), paginate([30, 5], 25), paginate([], 10)], [[[0, 1], [2, 3]], [[0, 0], [1, 1]], []]);
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("cell.set cell=A1 input=Item");
await cmds.run("cell.set cell=B1 input=(price)");
await cmds.run("cell.set cell=B2 input=12.5");
await cmds.run("format.font range=A1 bold=1");
await cmds.run("format.fill range=A1 color=yellow");
await cmds.run("format.border range=B2 edge=bottom color=red");
const text = (o: { output?: unknown }) => Array.from(fromBase64(String(o.output)), (c) => String.fromCharCode(c)).join("");
const o = await cmds.run("file.exportPdf gridlines=1 headings=1");
const pdf = text(o);
eq("a page, its parts", [o.message, pdf.startsWith("%PDF-1.4"), pdf.trimEnd().endsWith("%%EOF"), (pdf.match(/\/Type \/Page /g) ?? []).length], ["Untitled.pdf: 1 page, " + pdf.length + " bytes", true, true, 1]);
eq("the cells: bold, a fill, a border, a number right, text escaped", [/\/F2 10 Tf 0\.000 0\.000 0\.000 rg [\d.]+ [\d.]+ Td \(Item\) Tj/.test(pdf), /1\.000 1\.000 0\.000 rg [\d.]+ [\d.]+ 60 18 re f/.test(pdf), /1\.000 0\.000 0\.000 RG/.test(pdf),
  /\(12\.5\) Tj/.test(pdf), /\(\\\(price\\\)\) Tj/.test(pdf), /\(A\) Tj/.test(pdf), /Page 1 of 1/.test(pdf)], [true, true, true, true, true, true, true]);
// Every offset in the cross-reference table is where its object starts.
const xref = Number(/startxref\n(\d+)/.exec(pdf)![1]), offsets = [...pdf.slice(xref).matchAll(/(\d{10}) 00000 n/g)].map((m) => Number(m[1]));
eq("its cross-reference table", offsets.every((off, i) => pdf.startsWith(`${i + 1} 0 obj`, off)) && offsets.length > 5, true);
for (let r = 1; r <= 60; r++) await cmds.run(`cell.set cell=A${r} input=${r}`);
eq("rows past a page go to the next; landscape takes fewer", [(await cmds.run("file.exportPdf")).message.split(",")[0], (await cmds.run("file.exportPdf landscape=1")).message.split(",")[0],
  (await cmds.run("file.exportPdf paper=letter range=A1:A10")).message.split(",")[0]], ["Untitled.pdf: 2 pages", "Untitled.pdf: 3 pages", "Untitled.pdf: 1 page"]);
eq("refusals", [(await cmds.run("file.exportPdf paper=a3")).message, (await cmds.run("file.exportPdf sheets=all range=A1")).message], ["a3 is not a paper size (a4, letter)", "a range is the active sheet's: not with sheets=all"]);
const said: string[] = [];
await cmds.script('x = EXPORTPDF("A1:A5", "paper=letter")\nSAY LEFT(x, 8)', (l) => said.push(l));
eq("EXPORTPDF from REXX: base64 of a PDF", said, ["JVBERi0x"]);

// expect: ok Helvetica's widths
// expect: ok pages: what fits, then the next
// expect: ok a page, its parts
// expect: ok the cells: bold, a fill, a border, a number right, text escaped
// expect: ok its cross-reference table
// expect: ok rows past a page go to the next; landscape takes fewer
// expect: ok refusals
// expect: ok EXPORTPDF from REXX: base64 of a PDF
