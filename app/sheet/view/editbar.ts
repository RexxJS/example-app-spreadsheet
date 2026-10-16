// Undo and redo, sorting the selection by the cursor's column, rows or
// columns inserted before the selection (as many as it covers) or deleted;
// then Fill Series and AutoSum, Paste Special (what was copied, at the
// selection), and the Clears.
import { addrOf, colName } from "../address.ts";
import type { SheetView } from "./view.ts";

export function editBar(v: SheetView): void {
  // The selection's rows or columns: where they start, and how many.
  const span = (axis: "row" | "col") => {
    const s = v.selection() ?? [addrOf(v.here())!, addrOf(v.here())!];
    const lo = Math.min(s[0][axis], s[1][axis]);
    return { at: axis === "row" ? String(lo + 1) : colName(lo), count: String(Math.abs(s[0][axis] - s[1][axis]) + 1) };
  };
  const sort = (id: string) => {
    const s = v.selection();
    if (!s || s[0].row === s[1].row) {
      v.status.set("Select the rows to sort, the cursor in the column to sort them by");
      return;
    }
    v.run(id, { range: v.range(), by: colName(addrOf(v.here())!.col) });
  };
  ui.hstack(() => {
    v.btn("Undo", () => v.run("edit.undo"));
    v.btn("Redo", () => v.run("edit.redo"));
    bind(ui.text(""), v.history);
    v.btn("Sort A→Z", () => sort("data.sortAscending"));
    v.btn("Sort Z→A", () => sort("data.sortDescending"));
    v.btn("Insert row", () => v.run("rows.insert", span("row")));
    v.btn("Delete row", () => v.run("rows.delete", span("row")));
    v.btn("Insert column", () => v.run("columns.insert", span("col")));
    v.btn("Delete column", () => v.run("columns.delete", span("col")));
  });
  const at = () => v.range().split(":")[0];
  ui.hstack(() => {
    v.btn("Fill series", () => v.run("edit.fillSeries", { range: v.range() }));
    v.btn("AutoSum", () => v.run("formulas.autoSum", { cell: v.here() }));
    v.btn("Paste values", () => v.run("edit.pasteSpecial", { at: at(), what: "values" }));
    v.btn("Paste formats", () => v.run("edit.pasteSpecial", { at: at(), what: "formats" }));
    v.btn("Paste transposed", () => v.run("edit.pasteSpecial", { at: at(), what: "all", transpose: "true" }));
    v.btn("Paste comments", () => v.run("edit.pasteSpecial", { at: at(), what: "comments" }));
    v.btn("Paste adding", () => v.run("edit.pasteSpecial", { at: at(), what: "values", operation: "add" }));
    v.btn("Paste multiplying", () => v.run("edit.pasteSpecial", { at: at(), what: "values", operation: "multiply" }));
    v.btn("Paste skipping blanks", () => v.run("edit.pasteSpecial", { at: at(), what: "all", skipBlanks: "1" }));
    v.btn("Clear formats", () => v.run("range.clearFormats", { range: v.range() }));
    v.btn("Clear comments", () => v.run("range.clearComments", { range: v.range() }));
    v.btn("Clear cells", () => v.run("range.clearAll", { range: v.range() }));
  });
}
