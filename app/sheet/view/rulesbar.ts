// Conditional formatting over the selection, as Excel's menu has it:
// Highlight Cells (greater than, less than, equal to, text that contains,
// the value in the field), Top 10, Above average, Duplicates, a colour scale
// (Excel's red, yellow, green), an icon set (arrows), data bars, Clear rules;
// text that begins or ends with the field's, a formula's rule (the field's
// formula), blanks, errors, a date occurring (the field's period: today,
// thisWeek ...), and Manage Rules' Move up and Stop if true (the field's
// rule number); and the sheet's rules listed, the first one winning.
import type { SheetView } from "./view.ts";

export function rulesBar(v: SheetView): void {
  const value = state(""), listed = state("");
  const added = async (o: Promise<unknown>) => {
    await o;
    listed.set((await v.cmds.execute("format.rules")).message);
  };
  const highlight = (op: string) => () => added(v.run("format.highlight", { range: v.range(), op, value: value.value }));
  ui.hstack(() => {
    bind(ui.textfield("rule value", () => {}), value);
    v.btn("Greater than", highlight("greater"));
    v.btn("Less than", highlight("less"));
    v.btn("Equal to", highlight("equal"));
    v.btn("Text contains", highlight("contains"));
    v.btn("Top 10", () => added(v.run("format.topBottom", { range: v.range() })));
    v.btn("Above average", () => added(v.run("format.aboveAverage", { range: v.range() })));
    v.btn("Duplicates", () => added(v.run("format.duplicates", { range: v.range() })));
    v.btn("Color scale", () => added(v.run("format.colorScale", { range: v.range() })));
    v.btn("Icon set", () => added(v.run("format.iconSet", { range: v.range() })));
    v.btn("Data bars", () => added(v.run("format.dataBar", { range: v.range() })));
    v.btn("Clear rules", () => added(v.run("format.clearRules", { range: v.range() })));
  });
  ui.hstack(() => {
    v.btn("Begins with", highlight("beginsWith"));
    v.btn("Ends with", highlight("endsWith"));
    v.btn("Formula rule", () => added(v.run("format.formulaRule", { range: v.range(), formula: value.value })));
    v.btn("Blanks", () => added(v.run("format.blanks", { range: v.range() })));
    v.btn("Errors", () => added(v.run("format.errors", { range: v.range() })));
    v.btn("Date occurring", () => added(v.run("format.datePeriod", { range: v.range(), period: value.value || "today" })));
    v.btn("Move rule up", () => added(v.run("format.moveRule", { index: value.value, to: String(Math.max(1, Number(value.value) - 1)) })));
    v.btn("Stop if true", () => added(v.run("format.stopIfTrue", { index: value.value })));
  });
  bind(ui.text_wrapped("", 720), listed);
}
