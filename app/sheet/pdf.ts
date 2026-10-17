// PDF export: cells laid out on pages as Excel prints them (down, then
// over; A4 or Letter, portrait or landscape, its margins), each cell's text
// as it shows in Helvetica (bold, italic, its size and colour; numbers to
// the right as the grid has them), its fill, its borders, the gridlines
// and the row and column headings when asked, the sheet's name and the
// page number at the foot. A PDF 1.4 file of the standard fonts, written
// here (no library). Pure: cells in, bytes out.

// Helvetica's widths (thousandths of the size) for character codes 32 to
// 126, its AFM's; a character past them is as wide as an 'n'.
const W = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
  1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
  333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584];
export const textWidth = (s: string, size: number, bold = false): number =>
  [...s].reduce((w, ch) => { const c = ch.charCodeAt(0); return w + (c >= 32 && c <= 126 ? W[c - 32] : 556); }, 0) * size / 1000 * (bold ? 1.06 : 1);

export interface PdfCell {
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly size?: number;           // pt
  readonly color?: string;          // #rrggbb
  readonly fill?: string;
  readonly align?: "start" | "center" | "end";
  readonly borders?: { readonly top?: string; readonly right?: string; readonly bottom?: string; readonly left?: string };  // #rrggbb
}

export interface PdfSheet {
  readonly name: string;
  readonly cells: (row: number, col: number) => PdfCell | null;
  readonly rows: number;            // how many, from the range's top
  readonly cols: number;
  readonly firstRow: number;        // the sheet's row and column numbers of the range's top-left (0-based)
  readonly firstCol: number;
  readonly colWidth: (col: number) => number;    // pt
  readonly rowHeight: (row: number) => number;   // pt
}

export interface PageSetup {
  readonly paper: "a4" | "letter";
  readonly landscape: boolean;
  readonly margin: number;          // pt
  readonly gridlines: boolean;
  readonly headings: boolean;
}
export const PAPER: Readonly<Record<string, [number, number]>> = { a4: [595, 842], letter: [612, 792] };

const colName = (c: number): string => { let s = ""; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
// Text as a PDF string in WinAnsi: what Latin-1 has, else ?.
const pdfText = (s: string) => `(${[...s].map((ch) => { const c = ch.charCodeAt(0); return c > 255 ? "?" : ch === "(" || ch === ")" || ch === "\\" ? `\\${ch}` : c < 32 ? " " : c > 126 ? `\\${c.toString(8).padStart(3, "0")}` : ch; }).join("")})`;
const rgb = (c: string) => { const h = c.replace("#", ""); return [0, 2, 4].map((i) => (parseInt(h.slice(i, i + 2), 16) / 255).toFixed(3)).join(" "); };
const n = (x: number) => (Math.round(x * 100) / 100).toString();

// The pages: each a run of rows and of columns that fit, down then over.
export function paginate(sizes: readonly number[], room: number): [number, number][] {
  const out: [number, number][] = [];
  let start = 0, used = 0;
  sizes.forEach((s, i) => {
    if (used + s > room && i > start) { out.push([start, i - 1]); start = i; used = 0; }
    used += s;
  });
  if (sizes.length) out.push([start, sizes.length - 1]);
  return out;
}

export function writePdf(sheets: readonly PdfSheet[], setup: PageSetup): { bytes: Uint8Array; pages: number } {
  const [pw0, ph0] = PAPER[setup.paper], [pw, ph] = setup.landscape ? [ph0, pw0] : [pw0, ph0];
  const head = setup.headings ? 24 : 0, foot = 18;
  const streams: string[] = [];
  for (const sh of sheets) {
    const widths = Array.from({ length: sh.cols }, (_, c) => sh.colWidth(c)), heights = Array.from({ length: sh.rows }, (_, r) => sh.rowHeight(r));
    const rowPages = paginate(heights, ph - 2 * setup.margin - foot - (setup.headings ? 14 : 0)), colPages = paginate(widths, pw - 2 * setup.margin - head);
    for (const [c1, c2] of colPages) for (const [r1, r2] of rowPages) {
      const ops: string[] = [];
      const x0 = setup.margin + head, top = ph - setup.margin - (setup.headings ? 14 : 0);
      const xs = [x0], ys = [top];
      for (let c = c1; c <= c2; c++) xs.push(xs[xs.length - 1] + widths[c]);
      for (let r = r1; r <= r2; r++) ys.push(ys[ys.length - 1] - heights[r]);
      const line = (a: number, b: number, c: number, d: number, colour: string, w = 0.5) => ops.push(`${rgb(colour)} RG ${w} w ${n(a)} ${n(b)} m ${n(c)} ${n(d)} l S`);
      // Fills first, then the gridlines, then the borders and the text.
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
        const cell = sh.cells(r, c), x = xs[c - c1], y = ys[r - r1 + 1];
        if (cell?.fill) ops.push(`${rgb(cell.fill)} rg ${n(x)} ${n(y)} ${n(widths[c])} ${n(heights[r])} re f`);
      }
      if (setup.gridlines) {
        for (const x of xs) line(x, ys[0], x, ys[ys.length - 1], "#d0d0d0", 0.25);
        for (const y of ys) line(xs[0], y, xs[xs.length - 1], y, "#d0d0d0", 0.25);
      }
      if (setup.headings) {
        for (let c = c1; c <= c2; c++) { const s = colName(sh.firstCol + c), x = xs[c - c1] + (widths[c] - textWidth(s, 8)) / 2; ops.push(`BT /F1 8 Tf 0.4 0.4 0.4 rg ${n(x)} ${n(top + 4)} Td ${pdfText(s)} Tj ET`); }
        for (let r = r1; r <= r2; r++) { const s = String(sh.firstRow + r + 1), y = ys[r - r1 + 1] + (heights[r] - 8) / 2 + 1; ops.push(`BT /F1 8 Tf 0.4 0.4 0.4 rg ${n(x0 - 4 - textWidth(s, 8))} ${n(y)} Td ${pdfText(s)} Tj ET`); }
      }
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
        const cell = sh.cells(r, c);
        if (!cell) continue;
        const x = xs[c - c1], y = ys[r - r1 + 1], w = widths[c], h = heights[r], b = cell.borders ?? {};
        if (b.top) line(x, y + h, x + w, y + h, b.top);
        if (b.bottom) line(x, y, x + w, y, b.bottom);
        if (b.left) line(x, y, x, y + h, b.left);
        if (b.right) line(x + w, y, x + w, y + h, b.right);
        if (!cell.text) continue;
        const size = cell.size ?? 10, font = cell.bold && cell.italic ? "F4" : cell.bold ? "F2" : cell.italic ? "F3" : "F1";
        const tw = textWidth(cell.text, size, cell.bold), tx = cell.align === "end" ? x + w - 2 - tw : cell.align === "center" ? x + (w - tw) / 2 : x + 2;
        // Clipped to its cell, as the grid shows it.
        ops.push(`q ${n(x)} ${n(y)} ${n(w)} ${n(h)} re W n BT /${font} ${n(size)} Tf ${rgb(cell.color ?? "#000000")} rg ${n(tx)} ${n(y + (h - size) / 2 + size * 0.22)} Td ${pdfText(cell.text.split("\n")[0])} Tj ET Q`);
      }
      streams.push(ops.join("\n"));
      // The foot: the sheet's name, the page number (filled in below).
      streams[streams.length - 1] += `\n__FOOT__${sh.name}`;
    }
  }
  if (!streams.length) streams.push("__FOOT__");
  // The objects: catalog, pages, four fonts, then each page and its content.
  const objs: string[] = [];
  const total = streams.length, fontBase = 3;
  objs[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  const kidsIds = streams.map((_, i) => 7 + i * 2);
  objs[2] = `<< /Type /Pages /Kids [${kidsIds.map((k) => `${k} 0 R`).join(" ")}] /Count ${total} >>`;
  ["Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique"].forEach((f, i) => { objs[fontBase + i] = `<< /Type /Font /Subtype /Type1 /BaseFont /${f} /Encoding /WinAnsiEncoding >>`; });
  streams.forEach((st, i) => {
    const [body, name] = st.split("__FOOT__");
    const label = `${name ? `${name}  ·  ` : ""}Page ${i + 1} of ${total}`;
    const content = `${body}\nBT /F1 8 Tf 0.4 0.4 0.4 rg ${n((pw - textWidth(label, 8)) / 2)} ${n(setup.margin / 2)} Td ${pdfText(label)} Tj ET`;
    objs[7 + i * 2] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pw} ${ph}] /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R /F4 6 0 R >> >> /Contents ${8 + i * 2} 0 R >>`;
    objs[8 + i * 2] = `<< /Length ${content.length} >>\nstream\n${content}\nendstream`;
  });
  let pdf = "%PDF-1.4\n%âãÏÓ\n";
  const offsets: number[] = [];
  for (let i = 1; i < objs.length; i++) {
    offsets[i] = pdf.length;
    pdf += `${i} 0 obj\n${objs[i]}\nendobj\n`;
  }
  const xref = pdf.length;
  pdf += `xref\n0 ${objs.length}\n0000000000 65535 f \n${offsets.slice(1).map((o) => `${String(o).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size ${objs.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  // Every character is one byte (WinAnsi), so offsets above are byte offsets.
  const bytes = new Uint8Array(pdf.length);
  for (let i = 0; i < pdf.length; i++) bytes[i] = pdf.charCodeAt(i) & 255;
  return { bytes, pages: total };
}
