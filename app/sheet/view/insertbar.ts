// The Insert panel: Insert Link (Ctrl+K: the panel opened, the link field
// focused) over the selection, to a URL or a place in the workbook
// (Sheet2!B3, a name), the text field shown in an empty cell; Remove link;
// Open link, the selected cell's followed (a place selected, a URL handed
// to the system: the installed app's shell.open). Then pictures (a file
// picked: documents), shapes and text boxes at the selected cell, and each
// drawing on the sheet drawn in vg, with its own buttons: Delete, to the
// front, moved to the selected cell.
import { fromBase64 } from "../csv.ts";
import type { Shape } from "../charts.ts";
import type { DrawingView } from "../workbook.ts";
import type { SheetView } from "./view.ts";

function draw(s: Shape): void {
  switch (s.kind) {
    case "rect": vg.rect(s.x, s.y, s.w, s.h, () => vg.fill(s.fill)); return;
    case "line": vg.line(s.x1, s.y1, s.x2, s.y2, () => vg.stroke(s.stroke, s.width)); return;
    case "polyline": vg.polyline([...s.points], () => { vg.fill("none"); vg.stroke(s.stroke, s.width); }); return;
    case "circle": vg.circle(s.cx, s.cy, s.r, () => vg.fill(s.fill)); return;
    case "path": vg.path(s.d, () => { vg.fill(s.fill); vg.stroke(s.stroke, 1); }); return;
    case "text": vg.text_anchored(s.x, s.y, s.size, s.anchor, s.text, () => vg.fill(s.fill)); return;
  }
}

export function insertBar(v: SheetView): void {
  const target = state(""), shown = state(""), tip = state(""), words = state("");
  ui.hstack(() => {
    const field = bind(ui.textfield("link to", () => {}), target);
    v.action("Link field", () => ui.focus(field));
    bind(ui.textfield("link text", () => {}), shown);
    bind(ui.textfield("link tip", () => {}), tip);
    v.btn("Link", () => v.run("cell.hyperlink", { range: v.range(), target: target.value, ...(shown.value ? { text: shown.value } : {}), ...(tip.value ? { tip: tip.value } : {}) }));
    v.btn("Remove link", () => v.run("cell.removeHyperlink", { range: v.range() }));
    v.btn("Open link", () => v.run("cell.followHyperlink", { cell: v.here() }));
  });
  ui.hstack(() => {
    v.btn("Picture", async () => {
      if (typeof documents === "undefined") return v.status.set("A picture from a file needs the installed app (documents)");
      const d = await documents.pick_open("Insert a picture: PNG, JPEG, GIF or BMP");
      if (!d) return;
      return v.run("insert.picture", { base64: await d.read_base64(), at: v.here(), alt: d.name });
    });
    const shape = (label: string, kind: string) => v.btn(label, () => v.run("insert.shape", { shape: kind, at: v.here(), ...(words.value ? { text: words.value } : {}) }));
    shape("Rectangle", "rectangle");
    shape("Rounded rectangle", "roundedRectangle");
    shape("Oval", "oval");
    shape("Triangle", "triangle");
    shape("Right arrow", "rightArrow");
    shape("Line", "line");
    bind(ui.textfield("drawing text", () => {}), words);
    v.btn("Text box", () => v.run("insert.textbox", { text: words.value || "Text", at: v.here() }));
  });
  each(v.drawings, (d: DrawingView) => `${d.name}|${d.where}|${d.w}x${d.h}|${d.fill ?? ""}|${d.line ?? ""}|${d.text ?? ""}|${d.picture?.length ?? 0}`, (d: DrawingView) => {
    ui.hstack(() => {
      ui.text(`${d.name}: ${d.kind === "shape" ? d.shape : d.kind} at ${d.where}, ${d.w}x${d.h}${d.alt ? ` (${d.alt})` : ""}`);
      ui.btn(`Delete ${d.name}`, () => v.run("drawing.delete", { name: d.name }));
      ui.btn(`${d.name} to front`, () => v.run("drawing.order", { name: d.name, to: "front" }));
      ui.btn(`Move ${d.name} here`, () => v.run("drawing.edit", { name: d.name, at: v.here() }));
    });
    vg.scene(`0 0 ${d.w} ${d.h}`, d.w, d.h, () => {
      if (d.kind === "picture" && d.picture) {
        vg.image(fromBase64(d.picture), () => { vg.box(0, 0, d.w, d.h); vg.fit("contain"); });
        return;
      }
      d.shapes.forEach(draw);
    });
  });
}
