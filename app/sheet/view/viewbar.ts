// Rows and columns out of view, over the selection, as Excel has them: Hide
// and Unhide rows and columns; a filter on the selected column of the
// selection's table (the criterion as COUNTIF writes one: >100, north,
// <>x), Clear filter; Freeze top row and Unfreeze; the selection's rows set
// to a height (empty or 0: the grid's).
import { addrOf, colName } from "../address.ts";
import type { SheetView } from "./view.ts";

export function viewBar(v: SheetView): void {
  const criterion = state(""), height = state("");
  ui.hstack(() => {
    v.btn("Hide rows", () => v.run("rows.hide", { range: v.range() }));
    v.btn("Unhide rows", () => v.run("rows.unhide", { range: v.range() }));
    v.btn("Hide columns", () => v.run("columns.hide", { range: v.range() }));
    v.btn("Unhide columns", () => v.run("columns.unhide", { range: v.range() }));
    bind(ui.textfield("filter by", () => {}), criterion);
    v.btn("Filter", () => v.run("data.filter", { range: v.range(), column: colName(addrOf(v.here())!.col), criterion: criterion.value }));
    v.btn("Clear filter", () => v.run("data.clearFilter"));
    v.btn("Freeze top row", () => v.run("view.freezeRows", { count: "1" }));
    v.btn("Split", () => v.run("view.split", v.wb.split ? { off: "1" } : {}));
    v.btn("Freeze first column", () => v.run("view.freezeCols", { count: "1" }));
    v.btn("Freeze panes", () => v.run("view.freezePanes", { cell: v.here() }));
    v.btn("Unfreeze", async () => { await v.run("view.freezeRows", { count: "0" }); await v.run("view.freezeCols", { count: "0" }); });
    bind(ui.textfield("row height", () => {}), height);
    v.btn("Zoom in", () => v.run("view.zoom", { by: "in" }));
    v.btn("Zoom out", () => v.run("view.zoom", { by: "out" }));
    v.btn("Zoom 100%", () => v.run("view.zoom", { percent: "100" }));
    v.btn("AutoFit columns", () => v.run("columns.autofit", { range: v.range() }));
    v.btn("AutoFit rows", () => v.run("rows.autofit", { range: v.range() }));
    v.btn("Row height", () => v.run("rows.height", { range: v.range(), height: height.value.trim() || "0" }));
  });
}
