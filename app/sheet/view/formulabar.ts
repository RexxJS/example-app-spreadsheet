// The formula bar, as old Excel's: one row under the toolbar, the Name
// Box (the selected cell's address), fx, and what was typed into the cell,
// both bound two-way. The address is also the Name Box: Return on a named
// range's name selects it (going to its sheet, for a workbook's name of
// another one); Return on a cell moves to the formula. Return in the
// formula sets the cell. R1C1 switches how Excel formulas are shown and
// typed (formulas.style), and back. Under it, the function names that
// complete what is typed (complete.ts) and the call it is in; F4 cycles
// the last reference's $s.
import { addrOf } from "../address.ts";
import { complete, completed, cycleRef } from "../complete.ts";
import type { SheetView } from "./view.ts";

export function formulaBar(v: SheetView): void {
  ui.hstack(() => {
    const cell = ui.tooltip(bind(ui.textfield("cell", () => {}), v.address), "Name Box: a cell, or a name, to go to");
    ui.text("fx");
    const formula = ui.tooltip(bind(ui.textfield("formula", () => {}), v.input), "Formula bar: Return sets the cell");
    // The Name Box as wide as an address or a name, the formula taking the
    // rest of the row and growing with the window (sae's set_width and
    // fill_width): left to the stack, the two shared it and the formula
    // came out a few pixels wide.
    ui.set_width(cell, 110);
    ui.fill_width(formula);
    ui.on_submit(cell, async () => {
      const name = v.address.value.trim().toUpperCase(), named = v.wb.sheet.named(name);
      if (!named) return ui.focus(formula);
      if (named.sheet) await v.run("sheet.activate", { name: v.wb.pages().find((p) => p.name.toUpperCase() === named.sheet)!.name });
      v.select(named.range);
      const listed = [...v.wb.sheet.names(), ...v.wb.names()].find(([n]) => n === name)!;
      v.status.set(`${name} is ${listed[1]}`);
    });
    ui.on_submit(formula, () => {
      if (addrOf(v.address.value)) v.run("cell.set", { cell: v.here(), input: v.wb.enteredInput(v.here(), v.input.value.trim()) });
      else v.status.set(`${v.address.value}: not a cell on this sheet`);
    });
    v.btn("R1C1", async () => {
      await v.run("formulas.style", { style: v.wb.referenceStyle === "A1" ? "R1C1" : "A1" });
      v.follow();
    });
  });
  // As a formula is typed: the names that complete it, each a button that
  // takes it, and the call its end is in; F4 (the grid not focused) cycles
  // the reference at its end.
  const hints = computed(() => complete(v.input.value, v.wb.sheet.language.functions()), v.input);
  ui.hstack(() => {
    bind(ui.text(""), computed(() => (hints.value.call ? `${hints.value.call}( argument ${hints.value.arg}` : ""), hints));
    each(computed(() => hints.value.names.map((n) => ({ n, prefix: hints.value.prefix })), hints), (h: { n: string }) => h.n,
      (h: { n: string; prefix: string }) => ui.btn(`${h.n}(`, () => v.input.set(completed(v.input.value, h.prefix, h.n))));
  });
  // F4: the grid's on_key runs it (a page has one key handler).
  v.action("Cycle reference", () => {
    if (v.input.value.startsWith("=")) v.input.set(cycleRef(v.input.value).text);
  });
}
