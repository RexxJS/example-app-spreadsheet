// What-if and data tools over the selection, as Excel's Data tab has them:
// Goal Seek (the selected formula to the goal, by changing a cell), Data
// Table (the selection, with a row input, a column input, or both), Remove
// Duplicates, Text to Columns (split at the delimiter, a comma by default);
// The outline: Group and Ungroup rows (or columns), Hide and Show Detail,
// the level buttons, and Subtotal (at each change in the selection's first
// column, a sum of its last); Custom Sort (the sort keys field's, "B desc, C asc", the first row a
// header) and Advanced Filter (the criteria range field's) over the
// selection; and for the table the selected cell is in, Table Design's Total Row and
// Banded Rows (each on, or off when it has it) and Convert to Range.
import type { SheetView } from "./view.ts";
import { addrOf, colName } from "../address.ts";

// The table the selected cell is in, by name, or "" (the command refuses).
function tableHere(v: SheetView): string {
  const at = addrOf(v.here());
  const t = at && v.wb.sheet.tables().find(({ range: [a, b] }) => at.col >= Math.min(a.col, b.col) && at.col <= Math.max(a.col, b.col) && at.row >= Math.min(a.row, b.row) && at.row <= Math.max(a.row, b.row));
  return t ? t.name : "";
}

export function dataBar(v: SheetView): void {
  const goal = state(""), by = state(""), rowInput = state(""), columnInput = state(""), delimiter = state("");
  const given = (o: Record<string, string>) => Object.fromEntries(Object.entries(o).filter(([, x]) => x.trim() !== ""));
  ui.hstack(() => {
    bind(ui.textfield("goal", () => {}), goal);
    bind(ui.textfield("by changing", () => {}), by);
    v.btn("Goal seek", () => v.run("data.goalSeek", { set: v.here(), to: goal.value, by: by.value }));
    bind(ui.textfield("row input", () => {}), rowInput);
    bind(ui.textfield("column input", () => {}), columnInput);
    v.btn("Data table", () => v.run("data.table", { range: v.range(), ...given({ row: rowInput.value, column: columnInput.value }) }));
    v.btn("Remove duplicates", () => v.run("data.removeDuplicates", { range: v.range() }));
    bind(ui.textfield("delimiter", () => {}), delimiter);
    v.btn("Flash fill", () => v.run("data.flashFill", { range: v.range() }));
    v.btn("Text to columns", () => v.run("data.textToColumns", { range: v.range(), ...given({ delimiter: delimiter.value }) }));
  });
  const keys = state(""), criteria = state(""), sources = state("");
  ui.hstack(() => {
    bind(ui.textfield("sort keys", () => {}), keys);
    v.btn("Custom sort", () => v.run("data.sort", { range: v.range(), keys: keys.value, headers: "1" }));
    bind(ui.textfield("criteria range", () => {}), criteria);
    v.btn("Advanced filter", () => v.run("data.advancedFilter", { range: v.range(), criteria: criteria.value }));
    bind(ui.textfield("sources", () => {}), sources);
    v.btn("Consolidate", () => v.run("data.consolidate", { to: v.here(), sources: sources.value, labels: "both" }));
  });
  ui.hstack(() => {
    v.btn("Group rows", () => v.run("outline.group", { range: v.range(), axis: "rows" }));
    v.btn("Ungroup rows", () => v.run("outline.ungroup", { range: v.range(), axis: "rows" }));
    v.btn("Group columns", () => v.run("outline.group", { range: v.range(), axis: "columns" }));
    v.btn("Hide detail", () => v.run("outline.hideDetail", { range: v.range(), axis: "rows" }));
    v.btn("Show detail", () => v.run("outline.showDetail", { range: v.range(), axis: "rows" }));
    v.btn("Outline level 1", () => v.run("outline.level", { level: "1" }));
    v.btn("Outline level 2", () => v.run("outline.level", { level: "2" }));
    v.btn("All levels", () => v.run("outline.level", { level: "8" }));
    v.btn("Subtotal", () => {
      const s = v.selection();
      if (!s) return;
      return v.run("data.subtotal", { range: v.range(), by: colName(Math.min(s[0].col, s[1].col)), columns: colName(Math.max(s[0].col, s[1].col)) });
    });
  });
  ui.hstack(() => {
    v.btn("Total row", () => v.run("table.totals", { name: tableHere(v), ...(v.wb.sheet.table(tableHere(v))?.totals ? { on: "0" } : {}) }));
    v.btn("Banded rows", () => v.run("table.banding", { name: tableHere(v), rows: v.wb.sheet.table(tableHere(v))?.banded ? "0" : "1" }));
    v.btn("Convert to range", () => v.run("table.toRange", { name: tableHere(v) }));
  });
}
