// The formula bar, as old Excel's: one row under the toolbar, the Name
// Box (the selected cell's address), fx, and what was typed into the cell,
// both bound two-way. The address is also the Name Box: Return on a named
// range's name selects it (going to its sheet, for a workbook's name of
// another one); Return on a cell moves to the formula. Return in the
// formula sets the cell. R1C1 switches how Excel formulas are shown and
// typed (formulas.style), and back.
import { addrOf } from "../address.ts";
import type { SheetView } from "./view.ts";

export function formulaBar(v: SheetView): void {
  ui.hstack(() => {
    const cell = ui.tooltip(bind(ui.textfield("cell", () => {}), v.address), "Name Box: a cell, or a name, to go to");
    ui.text("fx");
    const formula = ui.tooltip(bind(ui.textfield("formula", () => {}), v.input), "Formula bar: Return sets the cell");
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
}
