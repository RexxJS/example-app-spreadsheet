// What a cell carries besides its input, as the Tauri spreadsheet's cells
// do: a validation rule (validation.ts), a comment, a number format, a fixed
// style, and a conditional style (an
// expression in the sheet's language whose value is a style, worked out
// with the sheet: STYLE_IF(C5 < 0, RED_TEXT(), GREEN_TEXT()) in RexxJS
// formulas, IF(C5 < 0, "color: red", "") in Excel ones), and whether Error
// Checking ignores it.
import { addrOf, keyOf } from "./address.ts";
import type { Addr } from "./address.ts";
import type { Rule } from "./validation.ts";
import type { Look } from "./formats.ts";
import type { CellEditor } from "./editors.ts";

export interface Note {
  readonly validation?: Rule;
  readonly comment?: string;
  readonly format?: string;
  readonly style?: unknown;
  readonly styleExpression?: string;
  readonly look?: Look;
  readonly editor?: CellEditor;
  // Error Checking passes over the cell (Ignore Error).
  readonly ignoreError?: true;
  // Its cell style's name (styles.ts), under its own look.
  readonly cellStyle?: string;
  // A hyperlink: a URL (https:, mailto: ...) or a place in the workbook
  // (Sheet2!B3, a name; "#" before it as Excel writes one), and the tip
  // shown when the pointer rests on it.
  readonly hyperlink?: Hyperlink;
}

export interface Hyperlink {
  readonly target: string;
  readonly tip?: string;
}

export class CellNotes {
  readonly #notes = new Map<string, Note>();

  get(key: string): Note {
    return this.#notes.get(key) ?? {};
  }

  // Some parts of a cell's note; an empty part is none.
  set(key: string, parts: Note): void {
    const note: Record<string, unknown> = { ...this.get(key), ...parts };
    for (const [k, v] of Object.entries(note)) if (v === "" || v === null || v === undefined) delete note[k];
    if (Object.keys(note).length) this.#notes.set(key, note as Note); else this.#notes.delete(key);
  }

  // The cells with a note.
  keys(): string[] {
    return [...this.#notes.keys()];
  }

  // The cells with a conditional style, and its expression.
  conditional(): [string, string][] {
    return [...this.#notes].filter(([, n]) => n.styleExpression).map(([k, n]) => [k, n.styleExpression!]);
  }

  // Every note to where its cell went (rows or columns inserted or
  // deleted), its conditional style rewritten as its cell's formula is; a
  // note whose cell went is gone.
  remap(to: (a: Addr) => Addr | null, rewrite: (expression: string, from: Addr, to: Addr) => string = (e) => e): void {
    const notes = [...this.#notes];
    this.#notes.clear();
    for (const [key, note] of notes) {
      const from = addrOf(key)!, a = to(from);
      if (a) this.#notes.set(keyOf(a), note.styleExpression ? { ...note, styleExpression: rewrite(note.styleExpression, from, a) } : note);
    }
  }

  clear(): void {
    this.#notes.clear();
  }

  // Every note, to be put back (undo); a note is never changed in place.
  saved(): ReadonlyMap<string, Note> {
    return new Map(this.#notes);
  }

  restore(saved: ReadonlyMap<string, Note>): void {
    this.#notes.clear();
    for (const [k, n] of saved) this.#notes.set(k, n);
  }
}
