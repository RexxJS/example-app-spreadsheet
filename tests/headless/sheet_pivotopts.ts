// PivotTable options (6.5): tabular, outline and compact layouts,
// subtotals and grand totals, sorting by label or total, Show Values As,
// a date field grouped by year, quarter or month, a calculated field on
// its fields' sums, an outer group collapsed; saved and kept with the
// pivot; PIVOTOPTIONS from REXX.
import { calculate, pivoted } from "./sheet/pivot.ts";
import type { PivotFields } from "./sheet/pivot.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import type { Value } from "./sheet/values.ts";
import { eq } from "./lib/check.ts";

const sales: Value[][] = [["region", "product", "amount", "cost"], ["North", "pens", 10, 4], ["South", "ink", 4, 1], ["North", "ink", 6, 2], ["North", "pens", 5, 2], ["South", "pens", 5, 3]];
const f = (o: PivotFields["options"], extra: Partial<PivotFields> = {}): PivotFields => ({ rows: ["region", "product"], cols: [], value: "amount", agg: "SUM", filters: {}, options: o, ...extra });
eq("tabular (the Tauri app's), with subtotals and a grand total", pivoted(sales, f({ subtotals: true, grandTotals: true })),
  [["region", "product", "SUM of amount"], ["North", "pens", "15"], ["North", "ink", "6"], ["North Total", "", "21"], ["South", "ink", "4"], ["South", "pens", "5"], ["South Total", "", "9"], ["Grand Total", "", "30"]]);
eq("outline: an outer label once", pivoted(sales, f({ layout: "outline" })).map((r) => r[0]), ["region", "North", "", "South", ""]);
eq("compact: one column indented, the outer group its own row", pivoted(sales, f({ layout: "compact", subtotals: true })),
  [["Row Labels", "SUM of amount"], ["North", "21"], ["  pens", "15"], ["  ink", "6"], ["South", "9"], ["  ink", "4"], ["  pens", "5"]]);
eq("collapsed: its subtotal standing for it", pivoted(sales, f({ layout: "compact", collapsed: ["North"] })), [["Row Labels", "SUM of amount"], ["North", "21"], ["South", ""], ["  ink", "4"], ["  pens", "5"]]);
const one = (o: PivotFields["options"], cols: string[] = []) => pivoted(sales, f(o, { rows: ["region"], cols }));
eq("sorted by total, descending; by label", [one({ sort: "valueDesc" }).map((r) => r[0]), one({ sort: "desc" }).map((r) => r[0])], [["region", "North", "South"], ["region", "South", "North"]]);
eq("Show Values As: % of the grand total, of the row, a running total, a rank", [one({ showAs: "percentOfTotal" }).slice(1).map((r) => r[1]), one({ showAs: "percentOfRow" }, ["product"])[1], one({ showAs: "runningTotal" }).slice(1).map((r) => r[1]), one({ showAs: "rank" }).slice(1).map((r) => r[1])],
  [["70%", "30%"], ["North", "71.4285714286%", "28.5714285714%"], ["21", "30"], ["1", "2"]]);
eq("grand totals with columns: a row and a column", one({ grandTotals: true }, ["product"]), [["region", "pens", "ink", "Grand Total"], ["North", "15", "6", "21"], ["South", "5", "4", "9"], ["Grand Total", "20", "10", "30"]]);
eq("a calculated field on the fields' sums", [calculate("amount - cost", (x) => ({ amount: 10, cost: 4 } as Record<string, number>)[x]), calculate("(amount+2)*0.5", () => 2),
  pivoted(sales, { rows: ["region"], cols: [], value: "margin", agg: "SUM", filters: {}, options: { calculated: { margin: "amount / cost" } } })], [6, 2, [["region", "SUM of margin"], ["North", "2.625"], ["South", "2.25"]]]);
const dated: Value[][] = [["day", "n"], [45306, 1], [45337, 2], [45397, 3], [45672, 4]];
eq("a date field by year, quarter, month", ["year", "quarter", "month"].map((by) => pivoted(dated, { rows: ["day"], cols: [], value: "n", agg: "SUM", filters: {}, options: { dateGroup: { field: "day", by: by as "year" } } }).slice(1).map((r) => r.join(":"))),
  [["2024:6", "2025:4"], ["2024 Q1:3", "2024 Q2:3", "2025 Q1:4"], ["2024 Jan:1", "2024 Feb:2", "2024 Apr:3", "2025 Jan:4"]]);

const store = memoryStore();
const wb = new Workbook();
const cmds = new Commands(wb, store);
for (const [r, row] of sales.entries()) for (const [c, v] of row.entries()) await cmds.run(`cell.set cell=${"ABCD"[c]}${r + 1} input="${v}"`);
await cmds.run("pivot.create id=p source=A1:D6 value=amount agg=sum output=F1 rows=region");
eq("the command writes it again", [(await cmds.run("pivot.options id=p grandTotals=1 showAs=percentOfTotal")).message, show(wb.sheet.value("G4")), show(wb.sheet.value("G2"))],
  ["pivot p at F1:G4: 8 cells set, 8 changed", "100%", "70%"]);
eq("refusals", [(await cmds.run("pivot.options id=p layout=grid")).message, (await cmds.run("pivot.options id=p calculated=x")).message, (await cmds.run("pivot.options id=q layout=compact")).message],
  ["grid is not a layout (tabular, outline, compact)", "x is not Name=formula", "no pivot q"]);
await cmds.run("file.save name=pv");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=pv");
eq("saved with the pivot", ((await cmds2.run("pivot.get id=p")).output as { options: unknown }).options, { grandTotals: true, showAs: "percentOfTotal" });
const said: string[] = [];
await cmds.script('SAY PIVOTOPTIONS("p", "showAs=normal;calculated=Margin=amount-cost")', (l) => said.push(l));
await cmds.run("pivot.create id=p source=A1:D6 value=Margin agg=sum output=F1 rows=region");
eq("PIVOTOPTIONS from REXX; the calculated field the value then", [said, show(wb.sheet.value("G2"))], [["pivot p at F1:G4: 8 cells set, 8 changed"], "13"]);

// expect: ok tabular (the Tauri app's), with subtotals and a grand total
// expect: ok outline: an outer label once
// expect: ok compact: one column indented, the outer group its own row
// expect: ok collapsed: its subtotal standing for it
// expect: ok sorted by total, descending; by label
// expect: ok Show Values As: % of the grand total, of the row, a running total, a rank
// expect: ok grand totals with columns: a row and a column
// expect: ok a calculated field on the fields' sums
// expect: ok a date field by year, quarter, month
// expect: ok the command writes it again
// expect: ok refusals
// expect: ok saved with the pivot
// expect: ok PIVOTOPTIONS from REXX; the calculated field the value then
