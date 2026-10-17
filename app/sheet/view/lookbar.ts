// A cell's look, over the selection, as Excel's Home tab has it: Bold,
// Italic, Underline and Strikethrough (each on, or off when the selected
// cell has it), outline, all, inside, bottom, thick outline or no borders, Wrap text (its rows as tall as
// their text), and Merge & center, Merge and Unmerge (a merge keeps its
// top-left's value), and Format painter. Ctrl+B, Ctrl+I and Ctrl+U on the grid do the first three.
import type { SheetView } from "./view.ts";

export type FontPart = "bold" | "italic" | "underline" | "strike";

// The part toggled over the selection, by what the selected cell has.
export const toggleFont = (v: SheetView, part: FontPart) =>
  v.run("format.font", { range: v.range(), [part]: v.wb.notes.get(v.here()).look?.[part] ? "0" : "1" });

export function lookBar(v: SheetView): void {
  // Format Painter: the first click takes the selection's format, the
  // second paints it over the selection then.
  let painting: string | null = null;
  ui.hstack(() => {
    v.btn("Format painter", async () => {
      if (painting === null) {
        painting = v.range();
        v.status.set(`Format painter: ${painting}'s format taken; select where to paint it, then Format painter again`);
        print(`sheet: format painter took ${painting}`);
        return;
      }
      const from = painting;
      painting = null;
      await v.run("format.painter", { from, to: v.range() });
    });
    v.btn("Bold", () => toggleFont(v, "bold"));
    v.btn("Italic", () => toggleFont(v, "italic"));
    v.btn("Underline", () => toggleFont(v, "underline"));
    v.btn("Strikethrough", () => toggleFont(v, "strike"));
    v.btn("Outline borders", () => v.run("format.border", { range: v.range(), edge: "outline" }));
    v.btn("All borders", () => v.run("format.border", { range: v.range(), edge: "all" }));
    v.btn("Inside borders", () => v.run("format.border", { range: v.range(), edge: "inside" }));
    v.btn("Bottom border", () => v.run("format.border", { range: v.range(), edge: "bottom" }));
    v.btn("Thick outline", () => v.run("format.border", { range: v.range(), edge: "outline", style: "thick" }));
    v.btn("No borders", () => v.run("format.border", { range: v.range(), edge: "none" }));
    v.btn("Wrap text", () => v.run("format.wrap", v.wb.notes.get(v.here()).look?.wrap ? { range: v.range(), off: "1" } : { range: v.range() }));
    v.btn("Merge & center", () => v.run("cells.merge", { range: v.range(), center: "1" }));
    v.btn("Merge", () => v.run("cells.merge", { range: v.range() }));
    v.btn("Unmerge", () => v.run("cells.unmerge", { range: v.range() }));
  });
}
