// The workbook's sheets, as tabs: one button each (the active one named
// beside them), and a sheet's name to add one by, or to rename the active
// one to (what read it follows); a tab dragged among the others moves its
// sheet there (sae's ui.reorderable over the row each() makes), as do Move
// left and Move right for the active one; Delete sheet deletes it, as
// Excel's does. Copy sheet copies it with its cells and looks (named by the
// field, or "Name copy"); Hide sheet hides it, Unhide sheets shows them all;
// Tab color colours its tab by the field (a colour's name or #rrggbb). A formula reads another sheet's cells as Costs!B2 (Excel)
// or Costs_B2 (RexxJS).
import type { SheetView } from "./view.ts";

export function sheetTabs(v: SheetView): void {
  const name = state("");
  ui.hstack(() => {
    bind(ui.text(""), computed(() => `On ${v.active.value}:`, v.active));
    // A drop that changes the order: `to` is the tab's index after the
    // move, sheet.move's place counted from 1.
    // A tab coloured as its sheet's; the hidden sheets have none, so a
    // drop's place among the tabs is mapped to its place among the sheets.
    type Tab = { name: string; color: string | null };
    const rgb = (c: string) => [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16) / 255);
    ui.reorderable(
      each(v.tabs, (t: Tab) => `${t.name}|${t.color ?? ""}`, (t: Tab) => ui.button(t.name, () => {
        if (t.color) { const [r, g, b] = rgb(t.color); ui.bg_color(r, g, b, 1); }
        ui.onclick(() => v.run("sheet.activate", { name: t.name }));
      })),
      (from: number, to: number) => v.run("sheet.move", { name: v.tabs.value[from].name, to: String(v.sheets.value.indexOf(v.tabs.value[to].name) + 1) }),
    );
    bind(ui.textfield("sheet", () => {}), name);
    v.btn("Add sheet", () => v.run("sheet.add", name.value.trim() ? { name: name.value } : {}));
    v.btn("Rename sheet", () => v.run("sheet.rename", { name: v.active.value, to: name.value }));
    v.btn("Move left", () => v.run("sheet.move", { name: v.active.value, way: "left" }));
    v.btn("Move right", () => v.run("sheet.move", { name: v.active.value, way: "right" }));
    v.btn("Copy sheet", () => v.run("sheet.copy", name.value.trim() ? { name: v.active.value, to: name.value } : { name: v.active.value }));
    v.btn("Hide sheet", () => v.run("sheet.hide", { name: v.active.value }));
    v.btn("Unhide sheets", async () => { for (const p of v.wb.pages().filter((x) => x.hiddenSheet)) await v.run("sheet.hide", { name: p.name, off: "1" }); });
    v.btn("Tab color", () => v.run("sheet.color", { name: v.active.value, color: name.value.trim() || "none" }));
    v.btn("Delete sheet", () => v.run("sheet.delete", { name: v.active.value }));
  });
}
