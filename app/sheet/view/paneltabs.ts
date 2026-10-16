// The panels under the grid as a row of tabs, one open at a time (as
// Excel's task panes, or old Excel's toolbars shown on demand): a tab opens
// its panel, closing the one open; the open one's tab closes it.
import { PANELS } from "../menus.ts";
import type { SheetView } from "./view.ts";

export function panelTabs(v: SheetView): void {
  ui.hstack(() => {
    ui.text("Panels:");
    for (const p of PANELS) {
      ui.tooltip(ui.btn(p.label, () => v.run("view.panel", { name: v.wb.panel === p.name ? "none" : p.name })), `The ${p.label} panel: open, or closed when open`);
    }
  });
}
