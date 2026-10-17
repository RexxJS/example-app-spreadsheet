// Cell styles and table styles, as Excel's galleries have them: a cell
// style is a named look (formats.ts's Look), applied to cells by name and
// edited in one place, every cell of it drawn anew; a table style colours a
// table's header row, its banded rows or columns, its first and last
// columns and its totals row. The galleries here are this app's own looks
// under Excel's style names. Colours may be a theme's roles (accent1 ...:
// themes.ts), so a theme changed re-colours them.
import type { Look } from "./formats.ts";

export const CELL_STYLES: Readonly<Record<string, Look>> = {
  "Normal": {},
  "Good": { fill: "#c6efce", color: "#006100" },
  "Bad": { fill: "#ffc7ce", color: "#9c0006" },
  "Neutral": { fill: "#ffeb9c", color: "#9c5700" },
  "Title": { size: 18, bold: true, color: "dark2" },
  "Heading 1": { size: 15, bold: true, color: "dark2", borders: { bottom: "thick accent1" } },
  "Heading 2": { size: 13, bold: true, color: "dark2", borders: { bottom: "medium accent1" } },
  "Heading 3": { bold: true, color: "dark2", borders: { bottom: "accent1" } },
  "Total": { bold: true, borders: { top: "accent1", bottom: "double accent1" } },
  "Note": { fill: "#ffffcc", borders: { top: "#b2b2b2", right: "#b2b2b2", bottom: "#b2b2b2", left: "#b2b2b2" } },
  "Input": { fill: "#ffcc99", color: "#3f3f76", borders: { top: "#7f7f7f", right: "#7f7f7f", bottom: "#7f7f7f", left: "#7f7f7f" } },
  "Output": { fill: "#f2f2f2", color: "#3f3f3f", bold: true, borders: { top: "#3f3f3f", right: "#3f3f3f", bottom: "#3f3f3f", left: "#3f3f3f" } },
  "Calculation": { fill: "#f2f2f2", color: "#fa7d00", bold: true, borders: { top: "#7f7f7f", right: "#7f7f7f", bottom: "#7f7f7f", left: "#7f7f7f" } },
  "Warning Text": { color: "#ff0000" },
  "Explanatory Text": { italic: true, color: "#7f7f7f" },
  "Accent1": { fill: "accent1", color: "light1" },
  "Accent2": { fill: "accent2", color: "light1" },
  "Comma": { number: "#,##0.00" },
  "Currency": { number: "$#,##0.00" },
  "Percent": { number: "0%" },
  "Hyperlink": { color: "hyperlink", underline: true },
};

// A style's look under a cell's own: the cell's parts win, its borders
// side by side.
export function under(style: Look | undefined, own: Look | undefined): Look | undefined {
  if (!style) return own;
  if (!own) return style;
  return { ...style, ...own, ...(style.borders || own.borders ? { borders: { ...style.borders, ...own.borders } } : {}) };
}

// A table style: its header row's fill and text, its bands' fill, its
// lines' colour.
export interface TableStyle {
  readonly header: { readonly fill?: string; readonly color?: string };
  readonly band: string;
  readonly line: string;
}

export const TABLE_STYLES: Readonly<Record<string, TableStyle>> = {
  Light: { header: {}, band: "#d9e1f2", line: "accent1" },
  Medium: { header: { fill: "accent1", color: "light1" }, band: "#d9e1f2", line: "accent1" },
  Dark: { header: { fill: "dark1", color: "light1" }, band: "#d9d9d9", line: "dark1" },
  Green: { header: { fill: "accent6", color: "light1" }, band: "#e2efda", line: "accent6" },
  Orange: { header: { fill: "accent2", color: "light1" }, band: "#fce4d6", line: "accent2" },
  Grey: { header: { fill: "accent3", color: "light1" }, band: "#ededed", line: "accent3" },
};
export const TABLE_STYLE_NAMES = Object.keys(TABLE_STYLES);

// A style's name as given, matched to one any case (null: none).
export const styleNamed = (names: Iterable<string>, n: string): string | null => [...names].find((x) => x.toLowerCase() === n.trim().toLowerCase()) ?? null;
