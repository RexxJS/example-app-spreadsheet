// Panels on demand (3.0.4): one open at a time under the grid, by name or
// label, none to close; each has its item on the View menu; SHOWPANEL and
// GETPANEL from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { MENUS, PANELS, panelItem } from "./sheet/menus.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const msg = async (name: string) => (await cmds.execute("view.panel", { name })).message;
eq("none open at the start", wb.panel, null);
eq("one at a time, by name or label", [await msg("rules"), wb.panel, await msg("Find & replace"), wb.panel, await msg("none"), wb.panel],
  ["the rules panel", "rules", "the find panel", "find", "no panel", null]);
eq("a panel that is none", await msg("ribbon"), `ribbon is not a panel (${PANELS.map((p) => p.name).join(", ")}, none)`);
const view = MENUS.find((m) => m.label === "View")!.items.map((e) => (e === "-" ? "-" : e.label));
eq("each panel on the View menu", PANELS.filter((p) => !view.includes(panelItem(p.label))).map((p) => p.name), []);
eq("not a step to undo", wb.canUndo, false);
const said: string[] = [];
eq("SHOWPANEL and GETPANEL from REXX", [(await cmds.script('SAY SHOWPANEL("audit")\nSAY GETPANEL()\nSAY SHOWPANEL("none") "|" GETPANEL() "|"', (l) => said.push(l))).message, said],
  ["ran", ["audit", "audit", " |  |"]]);

// expect: ok none open at the start
// expect: ok one at a time, by name or label
// expect: ok a panel that is none
// expect: ok each panel on the View menu
// expect: ok not a step to undo
// expect: ok SHOWPANEL and GETPANEL from REXX
