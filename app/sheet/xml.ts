// A small XML reader and writer, for the parts of an XLSX or ODS file: an
// element's name (its prefix kept: "x:c"), attributes, children and text;
// entities and character references read, CDATA, comments and processing
// instructions passed over. Not a validating parser.
export interface El {
  readonly name: string;
  readonly attrs: Readonly<Record<string, string>>;
  readonly children: (El | string)[];
}

const ENT: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
export const unescapeXml = (s: string): string =>
  s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e: string) => (e[0] === "#" ? String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : ENT[e] ?? m));
export const escapeXml = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

export function parseXml(text: string): El {
  const root: El = { name: "#root", attrs: {}, children: [] };
  const stack: El[] = [root];
  let i = 0;
  while (i < text.length) {
    const lt = text.indexOf("<", i);
    if (lt < 0) { stack[stack.length - 1].children.push(unescapeXml(text.slice(i))); break; }
    if (lt > i) stack[stack.length - 1].children.push(unescapeXml(text.slice(i, lt)));
    if (text.startsWith("<!--", lt)) { i = text.indexOf("-->", lt) + 3; if (i < 3) break; continue; }
    if (text.startsWith("<![CDATA[", lt)) { const e = text.indexOf("]]>", lt); stack[stack.length - 1].children.push(text.slice(lt + 9, e < 0 ? text.length : e)); i = e < 0 ? text.length : e + 3; continue; }
    if (text[lt + 1] === "?" || text[lt + 1] === "!") { i = text.indexOf(">", lt) + 1; if (i <= 0) break; continue; }
    const gt = text.indexOf(">", lt);
    if (gt < 0) throw new RangeError("XML that ends inside a tag");
    let tag = text.slice(lt + 1, gt);
    i = gt + 1;
    if (tag[0] === "/") {
      const name = tag.slice(1).trim();
      while (stack.length > 1 && stack[stack.length - 1].name !== name) stack.pop();
      if (stack.length > 1) stack.pop();
      continue;
    }
    const self = tag.endsWith("/");
    if (self) tag = tag.slice(0, -1);
    const m = /^([^\s]+)/.exec(tag);
    const name = m ? m[1] : tag;
    const attrs: Record<string, string> = {};
    const re = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let a: RegExpExecArray | null;
    while ((a = re.exec(tag.slice(name.length)))) attrs[a[1]] = unescapeXml(a[3] ?? a[4] ?? "");
    const el: El = { name, attrs, children: [] };
    stack[stack.length - 1].children.push(el);
    if (!self) stack.push(el);
  }
  const top = root.children.find((c): c is El => typeof c !== "string");
  if (!top) throw new RangeError("XML with no element");
  return top;
}

// The name without its prefix: "x:c" is "c".
export const local = (n: string): string => n.slice(n.indexOf(":") + 1);
export const kids = (e: El, name?: string): El[] => e.children.filter((c): c is El => typeof c !== "string" && (name === undefined || local(c.name) === name));
export const kid = (e: El, name: string): El | undefined => kids(e, name)[0];
// Every descendant of a name, in document order.
export function all(e: El, name: string): El[] {
  const out: El[] = [];
  const walk = (x: El) => { for (const c of kids(x)) { if (local(c.name) === name) out.push(c); walk(c); } };
  walk(e);
  return out;
}
export const textOf = (e: El): string => e.children.map((c) => (typeof c === "string" ? c : textOf(c))).join("");
// An attribute by its local name.
export const attr = (e: El, n: string): string | undefined => e.attrs[n] ?? Object.entries(e.attrs).find(([k]) => local(k) === n)?.[1];
