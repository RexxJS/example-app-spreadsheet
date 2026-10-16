// The keyboard map (3.6): each key's command is one there is; Ctrl+End's
// last cell; LISTKEYS and GETLASTCELL from REXX.
import { KEYS } from "./sheet/keys.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

const cmds = new Commands(new Workbook());
const ids = cmds.catalogue().map((c) => c.id);
eq("each key runs a command there is, the grid's own, or the page's", KEYS.filter((k) => k.id !== "grid" && !k.id.startsWith("page.") && !ids.includes(k.id)).map((k) => k.combo), []);
eq("Excel's: navigation, editing, formatting, undo, find, save", ["Ctrl+Home", "Ctrl+End", "Ctrl+Space", "F2", "Ctrl+B", "Ctrl+Shift+$", "Ctrl+Z", "Ctrl+F", "Ctrl+S"].filter((c) => !KEYS.some((k) => k.combo === c)), []);
eq("an empty sheet's last cell is A1", (await cmds.execute("sheet.lastCell")).output, "A1");
await cmds.execute("cell.set", { cell: "C9", input: "1" });
await cmds.execute("cell.set", { cell: "F2", input: "1" });
eq("the last row used and the last column used", (await cmds.execute("sheet.lastCell")).output, "F9");
const said: string[] = [];
eq("GETLASTCELL and LISTKEYS from REXX", [(await cmds.script("SAY GETLASTCELL()\nk = LISTKEYS()\nSAY k.0 k.9", (l) => said.push(l))).message, said],
  ["ran", ["F9", `${KEYS.length} Ctrl+C | copy | edit.copy`]]);

// expect: ok each key runs a command there is, the grid's own, or the page's
// expect: ok Excel's: navigation, editing, formatting, undo, find, save
// expect: ok an empty sheet's last cell is A1
// expect: ok the last row used and the last column used
// expect: ok GETLASTCELL and LISTKEYS from REXX
