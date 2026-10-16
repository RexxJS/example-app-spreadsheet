// Find and replace, with the Tauri app's three options. Find lists the
// matches and selects the next after the selected cell.
import { addrOf, order } from "../address.ts";
import type { SheetView } from "./view.ts";

export function findBar(v: SheetView): void {
  const text = state(""), by = state("");
  const options = { matchCase: false, wholeCell: false, inFormulas: false };
  const params = () => ({ text: text.value, matchCase: String(options.matchCase), wholeCell: String(options.wholeCell), inFormulas: String(options.inFormulas) });
  const find = async () => {
    const o = await v.run("edit.find", params());
    const keys = (o.output ?? []) as string[];
    const next = keys.find((k) => order(addrOf(k)!) > order(addrOf(v.here())!)) ?? keys[0];
    if (next) v.address.set(next);
    v.status.set(keys.length ? `${keys.length} found: ${keys.join(", ")}` : `"${text.value}" not found`);
  };
  ui.hstack(() => {
    const field = bind(ui.textfield("find", () => {}), text);
    // Ctrl+F and Ctrl+H: the panel opened, the field focused.
    v.action("Find field", () => ui.focus(field));
    bind(ui.textfield("replace with", () => {}), by);
    v.btn("Find", find);
    // Go To Special: the kind typed in the find field (blanks, formulas,
    // constants, errors, comments, validation, conditional), its cells
    // listed and the first selected.
    v.btn("Go To Special", async () => {
      const o = await v.run("edit.goToSpecial", { what: text.value.trim() || "blanks" });
      const keys = (o.output ?? []) as string[];
      if (keys.length) v.address.set(keys[0]);
      if (o.ok) v.status.set(o.message);
    });
    v.btn("Replace all", () => v.run("edit.replace", { ...params(), with: by.value }));
    ui.toggle("Match case", (on: boolean) => { options.matchCase = !!on; });
    ui.toggle("Whole cell", (on: boolean) => { options.wholeCell = !!on; });
    ui.toggle("In formulas", (on: boolean) => { options.inFormulas = !!on; });
  });
}
