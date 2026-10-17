// Data in and out besides CSV: JSON (an array of objects, a record a row,
// their keys the header row in the order first seen; or an array of
// arrays), an HTML page's table (its rows and cells, th or td, a cell's
// text with its tags and entities read), and the cells out as an HTML
// table, each as it shows, with its look inline. Pure: text in, rows of
// inputs out, or rows of shown cells in and HTML out.

// An input for a JSON value: a string as typed (an apostrophe kept off a
// number-like one by the sheet's own reading), a number, true or false as
// TRUE or FALSE, null as nothing, anything else as its JSON.
const inputOf = (v: unknown): string =>
  v === null || v === undefined ? "" : typeof v === "string" ? v : typeof v === "number" ? String(v) : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : JSON.stringify(v);

export function fromJson(text: string): string[][] {
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new RangeError(`not JSON: ${(e as Error).message}`);
  }
  if (data && typeof data === "object" && !Array.isArray(data)) {
    // An object holding one array of records ({ "rows": [...] }) is its array.
    const arrays = Object.values(data as Record<string, unknown>).filter(Array.isArray);
    if (arrays.length === 1) data = arrays[0];
  }
  if (!Array.isArray(data)) throw new RangeError("JSON rows are an array: of objects, or of arrays");
  if (data.every((r) => Array.isArray(r))) return (data as unknown[][]).map((r) => r.map(inputOf));
  const keys: string[] = [];
  for (const r of data) if (r && typeof r === "object" && !Array.isArray(r)) for (const k of Object.keys(r)) if (!keys.includes(k)) keys.push(k);
  if (!keys.length) return (data as unknown[]).map((v) => [inputOf(v)]);
  return [keys, ...(data as unknown[]).map((r) => keys.map((k) => (r && typeof r === "object" && !Array.isArray(r) ? inputOf((r as Record<string, unknown>)[k]) : "")))];
}

const ENTITIES: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ", copy: "©", reg: "®", euro: "€", pound: "£", yen: "¥", cent: "¢", mdash: "—", ndash: "–", hellip: "…" };
export const unentity = (s: string): string =>
  s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => (e[0] === "#" ? String.fromCodePoint(e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : ENTITIES[e.toLowerCase()] ?? m));

// The n-th table of an HTML page (from 1) as rows of its cells' text: a
// br a line break, tags left out, runs of white space one, entities read;
// a colspan's cell followed by empty ones; a nested table's cells its own.
export function fromHtml(html: string, n = 1): string[][] {
  const tables: string[] = [];
  const re = /<table\b[^>]*>|<\/table\s*>/gi;
  let depth = 0, start = -1, m: RegExpExecArray | null;
  while ((m = re.exec(html))) {
    if (m[0][1] !== "/") { if (depth++ === 0) start = m.index; }
    else if (depth > 0 && --depth === 0) tables.push(html.slice(start, m.index + m[0].length));
  }
  if (n < 1 || n > tables.length) throw new RangeError(tables.length ? `the page has ${tables.length} table${tables.length === 1 ? "" : "s"}, not ${n}` : "the page has no table");
  const t = tables[n - 1].replace(/<!--[\s\S]*?-->/g, "");
  const inner = t.replace(/^<table\b[^>]*>/i, "").replace(/<\/table\s*>$/i, "");
  const rows: string[][] = [];
  for (const tr of inner.split(/<tr\b[^>]*>/i).slice(1)) {
    const row: string[] = [];
    const cell = /<(td|th)\b([^>]*)>([\s\S]*?)(?=<\/?(?:td|th|tr)\b|<\/tbody|<\/thead|<\/tfoot|$)/gi;
    let c: RegExpExecArray | null;
    while ((c = cell.exec(tr.replace(/<\/tr\s*>[\s\S]*$/i, "")))) {
      const text = unentity(c[3].replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").replace(/[ \t\r\f]+/g, " ").replace(/ ?\n ?/g, "\n").trim());
      row.push(text);
      const span = Number(/colspan\s*=\s*["']?(\d+)/i.exec(c[2])?.[1] ?? 1);
      for (let k = 1; k < Math.min(span, 50); k++) row.push("");
    }
    rows.push(row);
  }
  return rows;
}

const escaped = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\n/g, "<br>");

// Cells as an HTML table: each its text as shown and its style (the
// grid's AeCS is CSS's own words), the first row th when it is a header.
export function toHtml(rows: readonly (readonly { readonly text: string; readonly style: string }[])[], header: boolean): string {
  const css = (st: string) => st.split(";").map((d) => d.trim()).filter((d) => /^(background|color|font-weight|font-style|text-decoration|font-size|text-align|border(-top|-right|-bottom|-left)?|white-space)\s*:/.test(d) && !d.startsWith("--"))
    .map((d) => d.replace(/text-align:\s*start/, "text-align: left").replace(/text-align:\s*end/, "text-align: right")).join("; ");
  const cell = (tag: string, c: { text: string; style: string }) => `<${tag}${css(c.style) ? ` style="${escaped(css(c.style))}"` : ""}>${escaped(c.text)}</${tag}>`;
  return `<table>\n${rows.map((r, i) => `<tr>${r.map((c) => cell(header && i === 0 ? "th" : "td", c)).join("")}</tr>`).join("\n")}\n</table>`;
}
