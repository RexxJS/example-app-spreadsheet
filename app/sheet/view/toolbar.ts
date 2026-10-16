// One toolbar row under the menus, as old Excel's Standard and Formatting
// toolbars had it (not a ribbon): Undo, Redo; Cut, Copy, Paste; Bold,
// Italic, Underline; Align Left, Center, Align Right; Comma, Currency and
// Percent Style; Borders; Fill Color (yellow) and Font Color (red), as
// Excel's buttons start; Sort Ascending, AutoFilter (the filter field's
// criterion); a chart of the selection. Each a short label with its full
// name and shortcut as its tooltip, and each running what the panels' and
// keys' do (their actions, by label, at the click: the panels come after).
import { toggleFont } from "./lookbar.ts";
import type { SheetView } from "./view.ts";

export const TOOLS: readonly (readonly [string, string, string])[] = [
  // [label, tooltip, the action it runs or the command]
  ["Undo", "Undo (Ctrl+Z)", "Undo"], ["Redo", "Redo (Ctrl+Y)", "Redo"],
  ["Cut", "Cut (Ctrl+X)", "Cut"], ["Copy", "Copy (Ctrl+C)", "Copy"], ["Paste", "Paste (Ctrl+V)", "Paste"],
  ["B", "Bold (Ctrl+B)", "Bold"], ["I", "Italic (Ctrl+I)", "Italic"], ["U", "Underline (Ctrl+U)", "Underline"],
  ["Left", "Align Left", "Align left"], ["Center", "Center", "Align center"], ["Right", "Align Right", "Align right"],
  [",", "Comma Style", "Comma style"], ["$", "Currency Style", "Currency style"], ["%", "Percent Style", "Percent style"],
  ["Borders", "Borders: an outline round the selection", "Outline borders"],
  ["Fill", "Fill Color: yellow", "Fill yellow"], ["Color", "Font Color: red", "Text red"],
  ["A→Z", "Sort Ascending, by the selected cell's column", "Sort A→Z"], ["AutoFilter", "AutoFilter: the selection's rows by the filter field's criterion", "Filter"],
  ["Chart", "Chart of the selection", "Bar chart"],
];

export function toolBar(v: SheetView): void {
  // The toolbar's own actions, for the menus too.
  v.action("Align left", () => v.run("format.align", { range: v.range(), align: "left" }));
  v.action("Align center", () => v.run("format.align", { range: v.range(), align: "center" }));
  v.action("Align right", () => v.run("format.align", { range: v.range(), align: "right" }));
  v.action("Comma style", () => v.run("format.number", { range: v.range(), format: "number" }));
  v.action("Currency style", () => v.run("format.number", { range: v.range(), format: "currency" }));
  v.action("Percent style", () => v.run("format.number", { range: v.range(), format: "percent" }));
  v.action("Fill yellow", () => v.run("format.fill", { range: v.range(), color: v.wb.notes.get(v.here()).look?.fill ? "none" : "yellow" }));
  v.action("Text red", () => v.run("format.color", { range: v.range(), color: v.wb.notes.get(v.here()).look?.color ? "none" : "red" }));
  v.action("Bold", () => toggleFont(v, "bold"));
  v.action("Italic", () => toggleFont(v, "italic"));
  v.action("Underline", () => toggleFont(v, "underline"));
  ui.hstack(() => {
    for (const [label, tip, action] of TOOLS) {
      ui.tooltip(ui.btn(label, () => {
        if (v.actions.has(action)) return v.perform(action);
        v.status.set(`${tip}: not here`);
      }), tip);
    }
  });
}
