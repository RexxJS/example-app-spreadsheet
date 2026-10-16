// The menu bar (3.0.1): every command is on a menu or said in menus.ts not
// to need one (a getter, a script's, a row's own button), each item runs a
// command there is, and LISTMENUS() gives them to a script.
import { MENUS, NOT_IN_MENUS, menuLines, shownLabel } from "./sheet/menus.ts";
import type { MenuItem } from "./sheet/menus.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const ids = cmds.catalogue().map((c) => c.id);
const items = MENUS.flatMap((m) => m.items.filter((e): e is MenuItem => e !== "-"));
const onMenus = new Set(items.map((e) => e.id));
eq("old Excel's menus, in its order", MENUS.map((m) => m.label), ["File", "Edit", "View", "Insert", "Format", "Data", "Formulas", "Review", "Script"]);
eq("every command is on a menu, or said not to need one", ids.filter((id) => !onMenus.has(id) && !(id in NOT_IN_MENUS)), []);
eq("and none is both", ids.filter((id) => onMenus.has(id) && id in NOT_IN_MENUS), []);
eq("each item runs a command there is (or the page's own: a script run)", items.filter((e) => !ids.includes(e.id) && !e.id.startsWith("page.")).map((e) => e.label), []);
eq("NOT_IN_MENUS names only commands there are", Object.keys(NOT_IN_MENUS).filter((id) => !ids.includes(id)), []);
eq("no item twice on a menu", MENUS.filter((m) => new Set(m.items.filter((e) => e !== "-").map((e) => (e as MenuItem).label)).size !== m.items.filter((e) => e !== "-").length).map((m) => m.label), []);
eq("shortcuts shown beside their items", [shownLabel(items.find((e) => e.label === "Copy")!), shownLabel(items.find((e) => e.label === "Calculate now")!), shownLabel(items.find((e) => e.label === "Merge")!)],
  ["Copy (Ctrl+C)", "Calculate now (F9)", "Merge"]);
eq("as lines", [menuLines()[0], menuLines().find((l) => l.startsWith("Edit > Copy"))], ["File > New | file.new", "Edit > Copy | edit.copy | Ctrl+C"]);
const said: string[] = [];
await cmds.script('m = LISTMENUS()\nSAY m.0\nDO i = 1 TO m.0\n  IF INDEX(m.i, "| formulas.recalculate") > 0 THEN SAY m.i\nEND', (l) => said.push(l));
eq("LISTMENUS() from REXX: a command found by where it is", said, [String(menuLines().length), "Formulas > Calculate now | formulas.recalculate | F9"]);

// expect: ok old Excel's menus, in its order
// expect: ok every command is on a menu, or said not to need one
// expect: ok and none is both
// expect: ok each item runs a command there is (or the page's own: a script run)
// expect: ok NOT_IN_MENUS names only commands there are
// expect: ok no item twice on a menu
// expect: ok shortcuts shown beside their items
// expect: ok as lines
// expect: ok LISTMENUS() from REXX: a command found by where it is
