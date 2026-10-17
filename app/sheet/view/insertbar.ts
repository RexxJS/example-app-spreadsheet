// The Insert panel: Insert Link (Ctrl+K: the panel opened, the link field
// focused) over the selection, to a URL or a place in the workbook
// (Sheet2!B3, a name), the text field shown in an empty cell; Remove link;
// Open link, the selected cell's followed (a place selected, a URL handed
// to the system: the installed app's shell.open).
import type { SheetView } from "./view.ts";

export function insertBar(v: SheetView): void {
  const target = state(""), shown = state(""), tip = state("");
  ui.hstack(() => {
    const field = bind(ui.textfield("link to", () => {}), target);
    v.action("Link field", () => ui.focus(field));
    bind(ui.textfield("link text", () => {}), shown);
    bind(ui.textfield("link tip", () => {}), tip);
    v.btn("Link", () => v.run("cell.hyperlink", { range: v.range(), target: target.value, ...(shown.value ? { text: shown.value } : {}), ...(tip.value ? { tip: tip.value } : {}) }));
    v.btn("Remove link", () => v.run("cell.removeHyperlink", { range: v.range() }));
    v.btn("Open link", () => v.run("cell.followHyperlink", { cell: v.here() }));
  });
}
