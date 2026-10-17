// Drawings on a sheet, as Excel's Insert > Illustrations and Text have
// them: pictures (PNG, JPEG, GIF or BMP bytes, as base64), text boxes and
// basic shapes, each named ("Picture 1", "Oval 2", "TextBox 3"), anchored
// at a cell's top-left and of a size in pixels, moved with its cell as
// rows and columns come and go. A shape's outline is laid out here as
// charts.ts's shapes, for the page to draw in vg; a picture is drawn from
// its bytes.
import type { Addr } from "./address.ts";
import type { Shape } from "./charts.ts";

export type DrawingKind = "picture" | "textbox" | "shape";
// Excel's preset geometries, by the names its Shapes gallery uses.
export const SHAPES = ["rectangle", "roundedRectangle", "oval", "triangle", "rightArrow", "line"] as const;
export type ShapeKind = (typeof SHAPES)[number];

export interface Drawing {
  readonly kind: DrawingKind;
  readonly name: string;
  readonly at: Addr;
  readonly w: number;
  readonly h: number;
  readonly shape?: ShapeKind;
  readonly fill?: string;
  readonly line?: string;
  readonly text?: string;
  // A picture's bytes (base64), its type, and its alternative text.
  readonly picture?: string;
  readonly mime?: string;
  readonly alt?: string;
}

// Each shape's name as Excel numbers them ("Rounded Rectangle 2").
export const SHAPE_TITLES: Readonly<Record<ShapeKind, string>> = {
  rectangle: "Rectangle", roundedRectangle: "Rounded Rectangle", oval: "Oval", triangle: "Isosceles Triangle", rightArrow: "Right Arrow", line: "Straight Connector",
};
// And as DrawingML's preset geometry (prstGeom prst=).
export const PRESET: Readonly<Record<ShapeKind, string>> = {
  rectangle: "rect", roundedRectangle: "roundRect", oval: "ellipse", triangle: "triangle", rightArrow: "rightArrow", line: "line",
};

// A picture's type and natural size from its first bytes: PNG's IHDR,
// GIF's screen, BMP's header, JPEG's first frame; null when it is none.
export function pictureOf(b: Uint8Array): { mime: string; w: number; h: number } | null {
  const u16 = (i: number) => (b[i] << 8) | b[i + 1], u32 = (i: number) => ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3];
  const le16 = (i: number) => b[i] | (b[i + 1] << 8), le32 = (i: number) => (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24));
  if (b.length >= 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { mime: "image/png", w: u32(16), h: u32(20) };
  if (b.length >= 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { mime: "image/gif", w: le16(6), h: le16(8) };
  if (b.length >= 26 && b[0] === 0x42 && b[1] === 0x4d) return { mime: "image/bmp", w: Math.abs(le32(18)), h: Math.abs(le32(22)) };
  if (b.length >= 4 && b[0] === 0xff && b[1] === 0xd8) {
    for (let i = 2; i + 9 < b.length;) {
      if (b[i] !== 0xff) return null;
      const m = b[i + 1], len = u16(i + 2);
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { mime: "image/jpeg", w: u16(i + 7), h: u16(i + 5) };
      i += 2 + len;
    }
  }
  return null;
}

// A picture's file extension, by its type.
export const EXTENSION: Readonly<Record<string, string>> = { "image/png": "png", "image/jpeg": "jpeg", "image/gif": "gif", "image/bmp": "bmp" };

// A shape or text box laid out in its own w x h box: its outline, then
// its text, a line apiece (centred in a shape, from the top-left in a text
// box). The colours are as given (a theme's roles resolved by the caller).
export function drawingShapes(d: Drawing, fill: string, line: string, color: string): Shape[] {
  const { w, h } = d, out: Shape[] = [];
  const outline = (path: string) => out.push({ kind: "path", d: path, fill, stroke: line });
  switch (d.kind === "textbox" ? "rectangle" : d.shape ?? "rectangle") {
    case "rectangle": outline(`M0.5 0.5H${w - 0.5}V${h - 0.5}H0.5Z`); break;
    case "roundedRectangle": {
      const r = Math.min(w, h) / 6;
      outline(`M${r} 0.5H${w - r}A${r} ${r} 0 0 1 ${w - 0.5} ${r}V${h - r}A${r} ${r} 0 0 1 ${w - r} ${h - 0.5}H${r}A${r} ${r} 0 0 1 0.5 ${h - r}V${r}A${r} ${r} 0 0 1 ${r} 0.5Z`);
      break;
    }
    case "oval": outline(`M0.5 ${h / 2}A${w / 2 - 0.5} ${h / 2 - 0.5} 0 1 0 ${w - 0.5} ${h / 2}A${w / 2 - 0.5} ${h / 2 - 0.5} 0 1 0 0.5 ${h / 2}Z`); break;
    case "triangle": outline(`M${w / 2} 0.5L${w - 0.5} ${h - 0.5}H0.5Z`); break;
    case "rightArrow": {
      const head = Math.min(w / 2, h / 2), t = h / 4;
      outline(`M0.5 ${t}H${w - head}V0.5L${w - 0.5} ${h / 2}L${w - head} ${h - 0.5}V${h - t}H0.5Z`);
      break;
    }
    case "line": out.push({ kind: "line", x1: 0, y1: 0, x2: w, y2: h, stroke: line, width: 2 }); break;
  }
  const lines = (d.text ?? "").split("\n").filter((l, i, all) => l !== "" || i < all.length - 1);
  const size = 13, step = 17;
  if (d.kind === "textbox") lines.forEach((l, i) => out.push({ kind: "text", x: 6, y: 18 + i * step, size, anchor: "start", text: l, fill: color }));
  else lines.forEach((l, i) => out.push({ kind: "text", x: w / 2, y: h / 2 + 5 + (i - (lines.length - 1) / 2) * step, size, anchor: "middle", text: l, fill: color }));
  return out;
}
