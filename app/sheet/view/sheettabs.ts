// The workbook's sheets, as tabs: one button each (the active one named
// beside them), and a sheet's name to add one by, or to rename the active
// one to (what read it follows); a tab dragged among the others moves its
// sheet there (sae's ui.reorderable over the row each() makes), as do Move
// left and Move right for the active one; Delete sheet deletes it, as
// Excel's does. A formula reads another sheet's cells as Costs!B2 (Excel)
// or Costs_B2 (RexxJS).
import type { SheetView } from "./view.ts";

export function sheetTabs(v: SheetView): void {
  const name = state("");
  ui.hstack(() => {
    bind(ui.text(""), computed(() => `On ${v.active.value}:`, v.active));
    // A drop that changes the order: `to` is the tab's index after the
    // move, sheet.move's place counted from 1.
    ui.reorderable(
      each(v.sheets, (n: string) => n, (n: string) => ui.btn(n, () => v.run("sheet.activate", { name: n }))),
      (from: number, to: number) => v.run("sheet.move", { name: v.sheets.value[from], to: String(to + 1) }),
    );
    bind(ui.textfield("sheet", () => {}), name);
    v.btn("Add sheet", () => v.run("sheet.add", name.value.trim() ? { name: name.value } : {}));
    v.btn("Rename sheet", () => v.run("sheet.rename", { name: v.active.value, to: name.value }));
    v.btn("Move left", () => v.run("sheet.move", { name: v.active.value, way: "left" }));
    v.btn("Move right", () => v.run("sheet.move", { name: v.active.value, way: "right" }));
    v.btn("Delete sheet", () => v.run("sheet.delete", { name: v.active.value }));
  });
}
