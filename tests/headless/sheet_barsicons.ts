// Data bar and icon set options (4.5): a solid fill or Excel's gradient,
// the bar or the icon only, the shortest and longest bar's numbers;
// thresholds as a percent, a percentile or numbers, reversed; saved with
// the file; ADDDATABAR, ADDICONSET and ADDRULE's dataBar from REXX.
import { applied } from "./sheet/rules.ts";
import type { CfRule } from "./sheet/rules.ts";
import { parseRange } from "./sheet/address.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import type { Value } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const at = (s: string) => parseRange(s)!;
const cells: Record<string, Value> = { A1: 10, A2: 20, A3: 30, A4: 40, A5: 100 };
const marks = (rule: CfRule) => [...applied([rule], (k) => cells[k] ?? null)].map(([k, m]) => `${k} ${m.icon ?? ""}${m.style ?? ""}${m.hide ? " hidden" : ""}`);
eq("a bar's share of the largest, or of min to max, clamped", [marks({ kind: "dataBar", range: at("A1:A2"), color: "#638ec6" }), marks({ kind: "dataBar", range: at("A1:A5"), color: "#638ec6", min: 20, max: 40, barOnly: true })],
  [["A1 --fill: 0.5; --fill-color: #638ec6", "A2 --fill: 1; --fill-color: #638ec6"],
    ["A1 --fill: 0; --fill-color: #638ec6 hidden", "A2 --fill: 0; --fill-color: #638ec6 hidden", "A3 --fill: 0.5; --fill-color: #638ec6 hidden", "A4 --fill: 1; --fill-color: #638ec6 hidden", "A5 --fill: 1; --fill-color: #638ec6 hidden"]]);
const icons = (rule: Partial<CfRule>) => marks({ kind: "iconSet", range: at("A1:A5"), icons: "arrows", ...rule } as CfRule).map((m) => m.split(" ")[1]).join("");
eq("thresholds: percent of the way, percentiles, numbers; reversed", [icons({}), icons({ thresholds: [80, 50], by: "percentile" }), icons({ thresholds: [40, 20], by: "number" }), icons({ reverse: true })],
  ["↓↓↓→↑", "↓↓→→↑", "↓→→↑↑", "↑↑↑→↓"]);
eq("icons only", marks({ kind: "iconSet", range: at("A5"), icons: "triangles", iconOnly: true }), ["A5 ▲ hidden"]);

const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
for (const [k, v] of [["B1", "1"], ["B2", "2"], ["B3", "3"], ["C1", "5"], ["C2", "50"], ["C3", "95"]]) await cmds.run(`cell.set cell=${k} input=${v}`);
eq("the commands, and the cells shown", [await message("format.dataBar range=B1:B3 fill=solid showValue=0 min=0 max=4"), wb.textOf("B2"), wb.styleOf("B2").includes("--fill: 0.5"),
  await message("format.iconSet range=C1:C3 icons=ratings thresholds=90,10 by=number reverse=1 showValue=0"), [wb.textOf("C1"), wb.textOf("C2"), wb.textOf("C3")]],
  ["B1:B3: data bars #638ec6 solid from 0 to 4, bar only", "", true, "C1:C3: icon set ratings reversed at 90, 10 (number), icons only", ["●●●", "●●○", "●○○"]]);
eq("refusals", [await message("format.dataBar range=B1 fill=striped"), await message("format.dataBar range=B1 min=5 max=1"), await message("format.iconSet range=C1 thresholds=10,90"),
  await message("format.iconSet range=C1 by=rank")],
  ["striped is not a fill (gradient, solid)", "the shortest bar's number (5) must be below the longest's (1)", "10,90 is not two thresholds, the higher first (67,33)", "rank is not a threshold type (percent, percentile, number)"]);
await cmds.run("file.save name=bars");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=bars");
eq("saved with the file", [(await cmds2.run("format.rules")).output, wb2.textOf("C3")],
  [["1. B1:B3: data bars #638ec6 solid from 0 to 4, bar only", "2. C1:C3: icon set ratings reversed at 90, 10 (number), icons only"], "●○○"]);
const said: string[] = [];
await cmds.script('SAY ADDDATABAR("D1:D3", "#ff0000", "fill=solid;max=10")\nSAY ADDICONSET("E1:E3", "triangles", "thresholds=80,50;by=percentile")\nSAY ADDRULE("F1:F3", "dataBar", "#00ff00", "showValue=0")', (l) => said.push(l));
eq("ADDDATABAR, ADDICONSET and ADDRULE's dataBar from REXX", said, ["D1:D3: data bars #ff0000 solid from automatic to 10", "E1:E3: icon set triangles at 80, 50 (percentile)", "F1:F3: data bars #00ff00, bar only"]);

// expect: ok a bar's share of the largest, or of min to max, clamped
// expect: ok thresholds: percent of the way, percentiles, numbers; reversed
// expect: ok icons only
// expect: ok the commands, and the cells shown
// expect: ok refusals
// expect: ok saved with the file
// expect: ok ADDDATABAR, ADDICONSET and ADDRULE's dataBar from REXX
