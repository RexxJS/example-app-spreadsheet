// Pictures, shapes and text boxes (6.11b): a picture's type and size read
// from its bytes; each inserted at a cell, named as Excel numbers them;
// moved, sized, recoloured, renamed, ordered, deleted; refusals; laid out
// for the page; moved with their cells; undo; the sheet file and XLSX
// round trips (a drawing part, its media); the REXX functions.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { pictureOf } from "./sheet/drawings.ts";
import { fromBase64 } from "./sheet/csv.ts";
import { eq } from "./lib/check.ts";

// Drawings compared part by part, whatever the order of their keys.
const canon = (ds: readonly object[]) => JSON.stringify(ds.map((d) => Object.fromEntries(Object.entries(d).sort(([x], [y]) => (x < y ? -1 : 1)))));
const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAEklEQVR4nGP4z8DwHwyBNBgAAEnICff5q7YNAAAAAElFTkSuQmCC";
const JPEG = "/9j/4AAQSkZJRgAAAAAAAAAAAAD/wAARCAADAAUDAAAAAAAAAAAA";
eq("a picture's type and size from its bytes", [pictureOf(fromBase64(PNG)), pictureOf(fromBase64(JPEG)), pictureOf(new Uint8Array([1, 2, 3]))],
  [{ mime: "image/png", w: 2, h: 2 }, { mime: "image/jpeg", w: 5, h: 3 }, null]);
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("view.select range=C3");
const a = await cmds.run(`insert.picture base64=${PNG} alt=Quad w=40 h=40`);
const b = await cmds.run("insert.shape shape=oval at=E2 text=Hello");
const c = await cmds.run('insert.textbox text="Two\\nlines" at=B8');
const d = await cmds.run("insert.shape shape=rightArrow at=A1 fill=accent2 line=red w=120 h=50");
eq("inserted, each named", [a.message, b.message, c.message, d.message],
  ["Picture 1: picture at C3, 40x40, alt \"Quad\"", "Oval 2: oval at E2, 96x64, \"Hello\"", "TextBox 3: textbox at B8, 160x60, \"Two / lines\"", "Right Arrow 4: rightArrow at A1, 120x50"]);
eq("refusals", [(await cmds.run("insert.picture base64=AAAA")).message, (await cmds.run("insert.shape shape=star")).message, (await cmds.run("insert.shape shape=oval w=0")).message,
  (await cmds.run("insert.shape shape=oval fill=nothing")).message, (await cmds.run("drawing.edit name=Nobody w=5")).message, (await cmds.run("insert.shape shape=oval name=\"Oval 2\"")).message],
  ["not a picture (PNG, JPEG, GIF or BMP)", "no shape star (rectangle, roundedRectangle, oval, triangle, rightArrow, line)", "0 is not a size (1 to 4000 pixels)",
    "nothing is not a colour for the fill", "no drawing Nobody on Sheet1", "a drawing is named Oval 2 already"]);
const views = wb.drawingViews();
eq("laid out for the page", [views.map((v) => [v.name, v.where, v.shapes.length]), views[1].shapes[0].kind, views[3].shapes[0].kind === "path" && (views[3].shapes[0] as { fill: string }).fill, views[2].shapes.filter((s) => s.kind === "text").map((s) => (s as { text: string }).text)],
  [[["Picture 1", "C3", 0], ["Oval 2", "E2", 2], ["TextBox 3", "B8", 3], ["Right Arrow 4", "A1", 1]], "path", "#ed7d31", ["Two", "lines"]]);
const e = await cmds.run("drawing.edit name=\"oval 2\" at=F5 w=50 h=30 fill=#00ff00 text=Hi to=Ball");
eq("moved, sized, recoloured, renamed, ordered", [e.message, (await cmds.run("drawing.order name=Ball to=back")).message, (await cmds.run("drawing.list")).output],
  ["Ball: oval at F5, 50x30, \"Hi\"", "Ball to the back", ["Ball: oval at F5, 50x30, \"Hi\"", "Picture 1: picture at C3, 40x40, alt \"Quad\"", "TextBox 3: textbox at B8, 160x60, \"Two / lines\"", "Right Arrow 4: rightArrow at A1, 120x50"]]);
await cmds.run("rows.insert at=2 count=2");
await cmds.run("columns.delete at=A");
eq("moved with their cells", wb.page("Sheet1")!.drawings.map((x) => `${x.name}@${x.at.col},${x.at.row}`), ["Ball@4,6", "Picture 1@1,4", "TextBox 3@0,9", "Right Arrow 4@0,0"]);
await cmds.run("edit.undo");
await cmds.run("edit.undo");
await cmds.run('drawing.delete name="TextBox 3"');
eq("deleted, and undone", [(await cmds.run("drawing.list")).output, (await cmds.run("edit.undo")).ok, wb.page("Sheet1")!.drawings.length], [["Ball: oval at F5, 50x30, \"Hi\"", "Picture 1: picture at C3, 40x40, alt \"Quad\"", "Right Arrow 4: rightArrow at A1, 120x50"], true, 4]);
// Saved and opened again; written as XLSX and read back.
const again = new Workbook(), cmds2 = new Commands(again);
await cmds2.execute("file.import", { text: (await cmds.run("file.export")).output as string });
eq("saved with the file", canon(again.page("Sheet1")!.drawings), canon(wb.page("Sheet1")!.drawings));
const x = await cmds.run("file.exportXlsx");
const back = new Workbook(), cmds3 = new Commands(back);
const r = await cmds3.run(`file.importXlsx name=D.xlsx base64=${x.output}`);
eq("an XLSX file's: a drawing part, its picture a media part", [r.ok, back.fileReport.filter((l) => /picture|shape|drawing/.test(l)), canon(back.page("Sheet1")!.drawings)], [true, [], canon(wb.page("Sheet1")!.drawings.map((d) => (d.kind === "picture" ? d : { ...d, fill: d.fill ?? (d.kind === "textbox" ? "light1" : "accent1"), line: d.line ?? (d.kind === "textbox" ? "#7f7f7f" : "#2f528f") })))]);
const said: string[] = [];
await cmds.script(`SAY INSERTSHAPE("triangle", "H1", "w=30;h=30;text=Up")\nSAY INSERTTEXTBOX("Note", "", "w=80")\nSAY INSERTPICTURE("${PNG}", "J1", "alt=Dot")\nSAY EDITDRAWING("TextBox 6", "w=90")\nSAY ARRANGEDRAWING("Picture 7", "back")\nSAY LISTDRAWINGS()\nSAY DELETEDRAWING("Isosceles Triangle 5")`, (l) => said.push(l));
eq("from REXX", said.slice(0, 5).concat(said.slice(-1)), ["Isosceles Triangle 5", "TextBox 6", "Picture 7", "TextBox 6: textbox at C3, 90x60, \"Note\"", "Picture 7 to the back", "deleted Isosceles Triangle 5"]);

// expect: ok a picture's type and size from its bytes
// expect: ok inserted, each named
// expect: ok refusals
// expect: ok laid out for the page
// expect: ok moved, sized, recoloured, renamed, ordered
// expect: ok moved with their cells
// expect: ok deleted, and undone
// expect: ok saved with the file
// expect: ok an XLSX file's: a drawing part, its picture a media part
// expect: ok from REXX
