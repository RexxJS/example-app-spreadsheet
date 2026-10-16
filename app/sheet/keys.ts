// The keyboard map, as Excel's: each key, what it does, and the command
// it runs (or the grid's own, or the page's). Pure data: the grid's
// on_key (view/grid.ts) follows it, a headless test checks each command
// is one there is, and LISTKEYS() gives it to a script.
export interface Key {
  readonly combo: string;
  readonly does: string;
  readonly id: string;      // a command, "grid" (the datagrid's own), or "page.*"
}

export const KEYS: readonly Key[] = [
  { combo: "Ctrl+Arrow", does: "to the edge of the data", id: "grid" },
  { combo: "Page Up/Down", does: "a screen up or down", id: "grid" },
  { combo: "Ctrl+Home", does: "to A1", id: "page.select" },
  { combo: "Ctrl+End", does: "to the last cell used", id: "sheet.lastCell" },
  { combo: "Shift+Space", does: "select the row", id: "page.select" },
  { combo: "Ctrl+Space", does: "select the column", id: "page.select" },
  { combo: "F2", does: "edit the cell", id: "grid" },
  { combo: "Delete", does: "clear the selection", id: "range.clear" },
  { combo: "Ctrl+C", does: "copy", id: "edit.copy" },
  { combo: "Ctrl+X", does: "cut", id: "edit.cut" },
  { combo: "Ctrl+V", does: "paste", id: "edit.paste" },
  { combo: "Ctrl+D", does: "fill down", id: "edit.fillDown" },
  { combo: "Ctrl+R", does: "fill right", id: "edit.fillRight" },
  { combo: "Ctrl+E", does: "Flash Fill", id: "data.flashFill" },
  { combo: "Ctrl+Z", does: "undo", id: "edit.undo" },
  { combo: "Ctrl+Y", does: "redo", id: "edit.redo" },
  { combo: "Ctrl+B", does: "bold", id: "format.font" },
  { combo: "Ctrl+I", does: "italic", id: "format.font" },
  { combo: "Ctrl+U", does: "underline", id: "format.font" },
  { combo: "Ctrl+Shift+$", does: "Currency format", id: "format.number" },
  { combo: "Ctrl+Shift+%", does: "Percent format", id: "format.number" },
  { combo: "Ctrl+F", does: "Find (its panel, its field)", id: "view.panel" },
  { combo: "Ctrl+H", does: "Replace (its panel)", id: "view.panel" },
  { combo: "Ctrl+S", does: "Save", id: "file.save" },
  { combo: "F4", does: "the reference's $s, in the formula bar", id: "formulas.cycleRef" },
  { combo: "F9", does: "Calculate Now", id: "formulas.recalculate" },
  { combo: "Shift+F9", does: "Calculate Sheet", id: "formulas.calculateSheet" },
];
