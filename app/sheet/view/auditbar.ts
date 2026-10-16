// Formula auditing, as Excel's Formulas tab has it: trace what the selected
// cell reads (blue) or what reads it (orange), a level further each time;
// remove the arrows; Show Formulas (on, then off); Evaluate Formula's steps; and the Watch
// Window, the cells watched on any sheet as they change. And Calculation
// Options: manual calculation (on, then off), Calculate now (F9) and
// Calculate sheet (Shift+F9), iterative calculation (on, then off). Error
// checking lists what to check on the sheet, each with Go to and Ignore.
import { parseRange } from "../address.ts";
import type { Finding } from "../errorcheck.ts";
import type { Watch } from "../workbook.ts";
import type { SheetView } from "./view.ts";

export function auditBar(v: SheetView): void {
  const steps = state("");
  const findings = state<Finding[]>([]);
  const check = async () => {
    const o = await v.run("formulas.errorCheck");
    findings.set(o.ok ? o.output as Finding[] : []);
  };
  ui.hstack(() => {
    v.btn("Trace precedents", () => v.run("formulas.tracePrecedents", { cell: v.here() }));
    v.btn("Trace dependents", () => v.run("formulas.traceDependents", { cell: v.here() }));
    v.btn("Remove arrows", () => v.run("formulas.removeArrows"));
    v.btn("Show formulas", () => v.run("formulas.showFormulas", v.wb.showFormulas ? { off: "1" } : {}));
    v.btn("Evaluate formula", async () => {
      const o = await v.run("formulas.evaluate", { cell: v.here() });
      steps.set(o.ok ? o.message : "");
    });
    v.btn("Add watch", () => v.run("formulas.watch", { cell: v.here() }));
  });
  ui.hstack(() => {
    v.btn("Manual calculation", () => v.run("formulas.calculation", { mode: v.wb.calculation.manual ? "automatic" : "manual" }));
    v.btn("Calculate now", () => v.run("formulas.recalculate"));
    v.btn("Calculate sheet", () => v.run("formulas.calculateSheet"));
    v.btn("Iterative calculation", () => v.run("formulas.iteration", v.wb.calculation.iterate ? { off: "1" } : {}));
    v.btn("Error checking", check);
    v.btn("1904 date system", () => v.run("workbook.dateSystem", { system: v.wb.dateSystem === 1904 ? "1900" : "1904" }));
  });
  each(findings, (f: Finding) => `${f.cell} ${f.check}`, (f: Finding) => ui.hstack(() => {
    ui.text(`${f.cell} ${f.check}: ${f.message}`);
    ui.btn(`Go to ${f.cell}`, () => v.select(parseRange(f.cell)!));
    ui.btn(`Ignore ${f.cell}`, async () => {
      await v.run("formulas.ignoreError", { cell: f.cell });
      await check();
    });
  }));
  bind(ui.text_wrapped("", 720), steps);
  each(v.watches, (w: Watch) => `${w.sheet}!${w.cell}`, (w: Watch) => ui.hstack(() => {
    ui.text(`Watch ${w.sheet}!${w.cell}: ${w.input} = ${w.shown}`);
    ui.btn(`Unwatch ${w.sheet}!${w.cell}`, () => v.run("formulas.unwatch", { cell: `${w.sheet}!${w.cell}` }));
  }));
}
