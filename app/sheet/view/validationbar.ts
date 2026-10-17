// Data validation for the selection: what is allowed ("Yes,No" for a list,
// "1..10", ">=0" for numbers, "<=20" for a text length, "2024-01-01..
// 2024-12-31" for dates, "09:00..17:00" for times, "=A1>B1" for a custom
// formula, a pattern), stop or only warn, an input prompt, a rule cleared,
// and Circle invalid for what breaks one.
import { limitsOf } from "../validation.ts";
import type { SheetView } from "./view.ts";

// Two dates or times, "a..b" (or one, the earliest).
function spanOf(text: string): { min?: string; max?: string } {
  const [a, b] = text.split("..").map((x) => x.trim());
  return { ...(a ? { min: a } : {}), ...(b ? { max: b } : {}) };
}

export function validationBar(v: SheetView): void {
  const allowed = state(""), prompt = state("");
  let warn = false;
  const rule = (kind: string, extra: Record<string, string>) =>
    v.run("validation.set", { range: v.range(), kind, mode: warn ? "warn" : "stop", ...(prompt.value.trim() ? { prompt: prompt.value } : {}), ...extra });
  ui.hstack(() => {
    bind(ui.textfield("allowed", () => {}), allowed);
    v.btn("Allow list", () => rule("list", { values: allowed.value }));
    v.btn("Allow whole numbers", () => rule("whole", limitsOf(allowed.value)));
    v.btn("Allow decimals", () => rule("decimal", limitsOf(allowed.value)));
    v.btn("Allow length", () => rule("length", limitsOf(allowed.value)));
    v.btn("Allow dates", () => rule("date", spanOf(allowed.value)));
    v.btn("Allow times", () => rule("time", spanOf(allowed.value)));
    v.btn("Allow custom", () => rule("custom", { formula: allowed.value }));
    v.btn("Allow pattern", () => rule("pattern", { pattern: allowed.value }));
    ui.toggle("Warn only", (on: boolean) => { warn = !!on; });
    bind(ui.textfield("input prompt", () => {}), prompt);
    v.btn("Clear rule", () => v.run("validation.clear", { range: v.range() }));
    v.btn("Circle invalid", () => v.run("validation.circle"));
  });
}
