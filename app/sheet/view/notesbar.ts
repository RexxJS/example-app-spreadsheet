// The selected cell's comment and number format (or style: "bold;
// color:red", "$#,##0.00", "percent:0.0%"); Return sets each.
// on_submit gives the field back.
import type { SheetView } from "./view.ts";

export function notesBar(v: SheetView): void {
  ui.hstack(() => {
    const comment = ui.on_submit(bind(ui.textfield("comment", () => {}), v.comment), () => v.run("cell.comment", { cell: v.here(), text: v.comment.value }));
    const format = ui.on_submit(bind(ui.textfield("format", () => {}), v.format), () => v.run("cell.format", { cell: v.here(), format: v.format.value }));
    // The menus' Comment and Number format go to the field.
    v.action("Comment", () => ui.focus(comment));
    v.action("Number format", () => ui.focus(format));
  });
}
