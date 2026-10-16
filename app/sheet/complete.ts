// Help while a formula is typed: F4 cycles the reference at the caret
// through Excel's four forms (A1, $A$1, A$1, $A1, then A1 again), and the
// function names that complete the name being typed, with the call the
// caret is in and which argument of it. Pure: the formula bar and REXX
// use it alike.

// The reference ending at or holding `caret` (the text's end by default),
// cycled; the text unchanged when there is none there. A reference in a
// string is left alone.
export function cycleRef(text: string, caret = text.length): { text: string; caret: number } {
  const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])/g;
  let quoted = false, at = 0;
  const strings: [number, number][] = [];
  for (let i = 0; i < text.length; i++) {
    if (text[i] === '"') {
      if (!quoted) at = i; else strings.push([at, i]);
      quoted = !quoted;
    }
  }
  for (let m = re.exec(text); m; m = re.exec(text)) {
    const start = m.index, end = start + m[0].length;
    if (strings.some(([a, b]) => start > a && start < b)) continue;
    if (start > 0 && /[\w.]/.test(text[start - 1]) && text[start - 1] !== "$") continue;
    if (caret < start || caret > end) continue;
    const [col, row] = [m[2], m[4]], form = `${m[1] ? "$" : ""}A${m[3] ? "$" : ""}1`;
    const next = ({ A1: ["$", "$"], $A$1: ["", "$"], A$1: ["$", ""], $A1: ["", ""] } as Record<string, [string, string]>)[form];
    const ref = `${next[0]}${col}${next[1]}${row}`;
    return { text: text.slice(0, start) + ref + text.slice(end), caret: start + ref.length };
  }
  return { text, caret };
}

export interface Completion {
  readonly prefix: string;          // the name being typed ("" for none)
  readonly names: readonly string[]; // the functions it could be, in order
  readonly call: string | null;     // the function whose brackets the end is in
  readonly arg: number;             // which of its arguments, from 1
}

// What completes a formula typed so far (its end the caret), from a
// language's function names.
export function complete(text: string, functions: readonly string[], limit = 8): Completion {
  const body = text.startsWith("=") ? text.slice(1) : text;
  // Out of strings: what is typed, quotes and their text taken out.
  const bare = body.replace(/"[^"]*("|$)/g, (s) => (s.endsWith('"') && s.length > 1 ? '""' : '"'));
  const inString = (bare.match(/"/g) ?? []).length % 2 === 1;
  const m = inString || !text.startsWith("=") ? null : /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(bare);
  const prefix = m && !/\d/.test(m[1][0]) && !/^[A-Za-z]{1,3}\d+$/.test(m[1]) ? m[1].toUpperCase() : "";
  const names = prefix ? functions.filter((f) => f.toUpperCase().startsWith(prefix)).slice(0, limit) : [];
  // The innermost open call, and the commas at its own depth.
  const stack: { name: string; arg: number }[] = [];
  const toks = bare.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\(|\(|\)|,/g);
  for (const t of toks) {
    if (t[0] === ")") stack.pop();
    else if (t[0] === ",") { if (stack.length) stack[stack.length - 1].arg++; }
    else stack.push({ name: t[1] ? t[1].toUpperCase() : "", arg: 1 });
  }
  const top = [...stack].reverse().find((s) => s.name) ?? null;
  return { prefix, names, call: inString || !text.startsWith("=") ? null : top?.name ?? null, arg: top?.arg ?? 0 };
}

// A completion taken: the name being typed replaced by `name(`.
export const completed = (text: string, prefix: string, name: string): string =>
  `${text.slice(0, text.length - prefix.length)}${name}(`;
