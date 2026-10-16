// Flash Fill: the rest of a column filled by the pattern its first cells
// show, from the columns beside it. A program is a run of parts, each a
// literal or a word of a source cell (counted from its start or its end),
// whole or its first letter, as it is, in upper, lower or proper case. One
// is found from the first example (each place of it matched by the longest
// word that fits, else taken as a literal) and kept only if it gives every
// example; both countings are tried. Pure.

type Case = "as" | "upper" | "lower" | "proper";
type Part = { readonly lit: string } | { readonly src: number; readonly word: number; readonly fromEnd: boolean; readonly first: boolean; readonly case: Case };
export type Program = readonly Part[];

// A cell's words: runs of letters and digits.
const words = (s: string): string[] => s.match(/[\p{L}\p{N}]+/gu) ?? [];
const cased = (w: string, c: Case): string =>
  c === "upper" ? w.toUpperCase() : c === "lower" ? w.toLowerCase() : c === "proper" ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : w;

function wordOf(srcs: readonly string[], p: Exclude<Part, { lit: string }>): string | null {
  const ws = words(srcs[p.src] ?? "");
  const w = ws[p.fromEnd ? ws.length - 1 - p.word : p.word];
  if (w === undefined) return null;
  return cased(p.first ? w.charAt(0) : w, p.case);
}

export function run(program: Program, srcs: readonly string[]): string | null {
  let out = "";
  for (const p of program) {
    if ("lit" in p) { out += p.lit; continue; }
    const w = wordOf(srcs, p);
    if (w === null) return null;
    out += w;
  }
  return out;
}

// A program giving `want` from `srcs`, words counted from their start or
// their end.
function learn(srcs: readonly string[], want: string, fromEnd: boolean): Program {
  const parts: Part[] = [];
  const cands: Exclude<Part, { lit: string }>[] = [];
  srcs.forEach((s, src) => {
    const n = words(s).length;
    for (let word = 0; word < n; word++) {
      for (const c of ["as", "proper", "upper", "lower"] as Case[]) {
        for (const first of [false, true]) cands.push({ src, word: fromEnd ? n - 1 - word : word, fromEnd, first, case: c });
      }
    }
  });
  let i = 0;
  while (i < want.length) {
    let best: Exclude<Part, { lit: string }> | null = null, len = 0;
    for (const c of cands) {
      const w = wordOf(srcs, c)!;
      if (w.length > len && want.startsWith(w, i)) { best = c; len = w.length; }
    }
    if (best && len > 0) { parts.push(best); i += len; continue; }
    const last = parts[parts.length - 1];
    if (last && "lit" in last) parts[parts.length - 1] = { lit: last.lit + want[i] }; else parts.push({ lit: want[i] });
    i++;
  }
  return parts;
}

// The program the examples show: [sources, wanted] pairs, the first
// learnt from, all checked. Null when neither counting gives them all.
export function infer(examples: readonly (readonly [readonly string[], string])[]): Program | null {
  if (!examples.length) return null;
  for (const fromEnd of [false, true]) {
    const p = learn(examples[0][0], examples[0][1], fromEnd);
    if (examples.every(([s, w]) => run(p, s) === w)) return p;
  }
  return null;
}
