// The menu bar, as old Excel's: File, Edit, View, Insert, Format, Data,
// Formulas, Review, Script, each item one of the commands the panels run.
// An item is named as the panel's button that runs it is (Bold runs
// format.font, as the Bold button does), with its keyboard shortcut; the
// page runs that button's action (view/menubar.ts), so a menu does just
// what the panel does. Pure data: LISTMENUS() reads it, and a headless test
// checks every command is on a menu or said here not to need one.
export interface MenuItem {
  readonly label: string;   // the panel button's label, and the item's
  readonly id: string;      // the command it runs ("page.*": the page's own, a script run)
  readonly key?: string;    // its keyboard shortcut, shown beside it
}
export type MenuEntry = MenuItem | "-";
export interface Menu {
  readonly label: string;
  readonly items: readonly MenuEntry[];
}

const i = (label: string, id: string, key?: string): MenuItem => (key ? { label, id, key } : { label, id });

// The panels under the grid, one open at a time (view.panel), each with
// its menu item on View. Their names are what SHOWPANEL takes.
export const PANELS: readonly { readonly name: string; readonly label: string }[] = [
  { name: "edit", label: "Edit" }, { name: "looks", label: "Looks" }, { name: "view", label: "View" }, { name: "names", label: "Names" },
  { name: "rules", label: "Rules" }, { name: "validation", label: "Validation" }, { name: "editors", label: "Editors" },
  { name: "data", label: "Data" }, { name: "find", label: "Find & replace" }, { name: "audit", label: "Audit" },
  { name: "chart", label: "Charts" }, { name: "file", label: "File" }, { name: "script", label: "Script" }, { name: "lists", label: "Lists" },
];
export const panelItem = (label: string): string => `${label} panel`;

export const MENUS: readonly Menu[] = [
  { label: "File", items: [
    i("New", "file.new"), i("Open", "file.open"), i("Open file", "file.import"), i("Save", "file.save", "Ctrl+S"), i("Save file", "file.export"), i("Delete file", "file.delete"), "-",
    i("Import CSV", "file.importCsv"), i("Export CSV", "file.exportCsv"), i("Export CSV file", "file.exportCsv"), "-",
    i("Paste example", "example.paste"), i("Budget (RexxJS)", "file.open"), i("Test sheet (RexxJS)", "file.open"), i("Formatted (RexxJS)", "file.open"), "-",
    i("Clear all", "sheet.clear"),
  ] },
  { label: "Edit", items: [
    i("Undo", "edit.undo", "Ctrl+Z"), i("Redo", "edit.redo", "Ctrl+Y"), "-",
    i("Cut", "edit.cut", "Ctrl+X"), i("Copy", "edit.copy", "Ctrl+C"), i("Paste", "edit.paste", "Ctrl+V"),
    i("Paste values", "edit.pasteSpecial"), i("Paste formats", "edit.pasteSpecial"), i("Paste transposed", "edit.pasteSpecial"),
    i("Paste comments", "edit.pasteSpecial"), i("Paste adding", "edit.pasteSpecial"), i("Paste multiplying", "edit.pasteSpecial"), i("Paste skipping blanks", "edit.pasteSpecial"), "-",
    i("Fill Down", "edit.fillDown", "Ctrl+D"), i("Fill Right", "edit.fillRight", "Ctrl+R"), i("Fill Up", "edit.fillUp"), i("Fill Left", "edit.fillLeft"), i("Fill series", "edit.fillSeries"), "-",
    i("Clear contents", "range.clear", "Delete"), i("Clear cells", "range.clearAll"), i("Clear formats", "range.clearFormats"), i("Clear comments", "range.clearComments"), "-",
    i("Delete row", "rows.delete"), i("Delete column", "columns.delete"), i("Delete sheet", "sheet.delete"), "-",
    i("Find", "edit.find"), i("Replace all", "edit.replace"),
  ] },
  { label: "View", items: [
    ...PANELS.map((p) => i(panelItem(p.label), "view.panel")), "-",
    i("Show formulas", "formulas.showFormulas"), i("R1C1", "formulas.style"), "-",
    i("Hide rows", "rows.hide"), i("Unhide rows", "rows.unhide"), i("Hide columns", "columns.hide"), i("Unhide columns", "columns.unhide"), "-",
    i("Freeze top row", "view.freezeRows"), i("Unfreeze", "view.freezeRows"), i("Row height", "rows.height"), i("AutoFit columns", "columns.autofit"), i("AutoFit rows", "rows.autofit"),
  ] },
  { label: "Insert", items: [
    i("Insert row", "rows.insert"), i("Insert column", "columns.insert"), i("Add sheet", "sheet.add"), "-",
    i("Bar chart", "chart.insert"), i("Line chart", "chart.insert"), i("Pie chart", "chart.insert"), i("Sparkline", "insert.sparkline"), "-",
    i("Name the selection", "names.define"), i("Name the formula", "names.define"), i("Delete name", "names.delete"),
  ] },
  { label: "Format", items: [
    i("Bold", "format.font", "Ctrl+B"), i("Italic", "format.font", "Ctrl+I"), i("Underline", "format.font", "Ctrl+U"), i("Strikethrough", "format.font"), "-",
    i("Outline borders", "format.border"), i("All borders", "format.border"), i("No borders", "format.border"), "-",
    i("Align left", "format.align"), i("Align center", "format.align"), i("Align right", "format.align"),
    i("Comma style", "format.number"), i("Currency style", "format.number"), i("Percent style", "format.number"),
    i("Fill yellow", "format.fill"), i("Text red", "format.color"), "-",
    i("Format painter", "format.painter"), i("Wrap text", "format.wrap"), i("Merge & center", "cells.merge"), i("Merge", "cells.merge"), i("Unmerge", "cells.unmerge"), "-",
    i("Greater than", "format.highlight"), i("Less than", "format.highlight"), i("Equal to", "format.highlight"), i("Text contains", "format.highlight"),
    i("Top 10", "format.topBottom"), i("Above average", "format.aboveAverage"), i("Duplicates", "format.duplicates"),
    i("Color scale", "format.colorScale"), i("Icon set", "format.iconSet"), i("Data bars", "format.dataBar"), i("Clear rules", "format.clearRules"), "-",
    i("Rename sheet", "sheet.rename"), i("Move left", "sheet.move"), i("Move right", "sheet.move"),
  ] },
  { label: "Data", items: [
    i("Sort A→Z", "data.sortAscending"), i("Sort Z→A", "data.sortDescending"), i("Filter", "data.filter"), i("Clear filter", "data.clearFilter"), "-",
    i("Allow list", "validation.set"), i("Allow whole numbers", "validation.set"), i("Allow decimals", "validation.set"), i("Allow length", "validation.set"),
    i("Clear rule", "validation.clear"), i("Circle invalid", "validation.circle"), "-",
    i("Goal seek", "data.goalSeek"), i("Data table", "data.table"), i("Remove duplicates", "data.removeDuplicates"), i("Text to columns", "data.textToColumns"), i("Flash fill", "data.flashFill", "Ctrl+E"),
  ] },
  { label: "Formulas", items: [
    i("AutoSum", "formulas.autoSum"), "-",
    i("Excel formulas", "formulas.use"), i("RexxJS formulas", "formulas.use"), "-",
    i("Trace precedents", "formulas.tracePrecedents"), i("Trace dependents", "formulas.traceDependents"), i("Remove arrows", "formulas.removeArrows"),
    i("Evaluate formula", "formulas.evaluate"), i("Error checking", "formulas.errorCheck"), "-",
    i("Calculate now", "formulas.recalculate", "F9"), i("Calculate sheet", "formulas.calculateSheet", "Shift+F9"),
    i("Manual calculation", "formulas.calculation"), i("Iterative calculation", "formulas.iteration"),
  ] },
  { label: "Review", items: [
    i("Comment", "cell.comment"), i("Number format", "cell.format"), "-",
    i("Add watch", "formulas.watch"),
  ] },
  { label: "Script", items: [
    i("Run script", "page.script"), i("Run the setup script", "setup.run"), i("Commands", "page.commands"),
  ] },
];

// The commands on no menu, and why: what a script or the grid itself
// uses (setting and reading a cell), what a panel's own list runs on one
// of its rows (a watch, a chart, a finding), and the getters.
export const NOT_IN_MENUS: Readonly<Record<string, string>> = {
  "cell.set": "typing in a cell", "cell.get": "a getter", "cell.about": "a getter", "sheet.state": "a getter",
  "cell.editor": "a script's (SETCELLEDITOR)", "cell.toggle": "the cell's editor", "edit.pasteText": "a paste from another app",
  "ids.configure": "a script's (the Tauri app's auto IDs)", "ids.find": "a script's", "columns.move": "a column's header dragged",
  "columns.width": "a column's edge dragged", "names.list": "the names bar lists them", "formulas.unwatch": "a watch's own button",
  "format.rules": "a getter", "format.deleteRule": "a rule's own", "table.define": "a script's", "table.delete": "a script's",
  "table.get": "a getter", "table.sort": "a script's (the Tauri app's sort cycle)", "table.list": "a getter",
  "pivot.create": "a script's", "pivot.update": "a script's", "pivot.delete": "a script's", "pivot.get": "a getter", "pivot.list": "a getter",
  "chart.list": "a getter", "chart.delete": "a chart's own button", "formulas.calculate": "a script's (EVALUATE)",
  "formulas.ignoreError": "a finding's own button", "setup.get": "a getter", "setup.set": "a script's",
  "sheet.activate": "a sheet's tab", "file.list": "the file bar lists them", "menus.list": "a getter (LISTMENUS)", "range.stats": "the status bar shows it", "formulas.cycleRef": "F4 in the formula bar", "sheet.lastCell": "Ctrl+End", "keys.list": "a getter (LISTKEYS)", "formulas.complete": "the formula bar's hints",
};

// Each item as a line: "File > Save | file.save | Ctrl+S".
export function menuLines(): string[] {
  return MENUS.flatMap((m) => m.items.filter((e): e is MenuItem => e !== "-").map((e) => [`${m.label} > ${e.label}`, e.id, ...(e.key ? [e.key] : [])].join(" | ")));
}

// An item as the menu shows it: its label, then its shortcut.
export const shownLabel = (e: MenuItem): string => (e.key ? `${e.label} (${e.key})` : e.label);
