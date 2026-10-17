// Themes (4.7): colour roles (two darks, two lights, six accents, a
// hyperlink's) and major and minor fonts, named by looks, cell styles and
// table styles; a theme changed re-colours them; a role customised;
// undone; saved with the file; SETTHEME and the rest from REXX.
import { THEMES, resolved, setTheme } from "./sheet/themes.ts";
import { Commands, memoryStore } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";

eq("a role resolved by the theme; anything else as it is", [resolved("accent1"), resolved("#123456"), (setTheme(THEMES.Harbour), resolved("ACCENT1")), (setTheme(THEMES.Office), resolved("dark2"))],
  ["#4472c4", "#123456", "#0f6fc6", "#44546a"]);
const wb = new Workbook();
const store = memoryStore();
const cmds = new Commands(wb, store);
const message = async (line: string) => (await cmds.run(line)).message;
await cmds.run("cell.set cell=A1 input=x");
await cmds.run("format.fill range=A1 color=accent1");
await cmds.run("style.apply range=B1 name=Accent2");
eq("a look naming a role, kept as the role", [wb.notes.get("A1").look?.fill, wb.styleOf("A1"), wb.styleOf("B1")], ["accent1", "background: #4472c4", "background: #ed7d31; color: #ffffff"]);
eq("another theme re-colours them", [await message("theme.set name=harbour"), wb.styleOf("A1"), wb.styleOf("B1")],
  ["theme Harbour: accents #0f6fc6 #009dd9 #0bd0d9 #10cf9b #7cca62 #a5c249, headings Georgia, body Verdana", "background: #0f6fc6", "background: #009dd9; color: #ffffff"]);
eq("a role customised, the fonts too", [await message("theme.color role=accent1 color=#112233"), wb.styleOf("A1"), await message("theme.fonts major=Didot"), wb.theme.name],
  ["theme Custom: accent1 #112233", "background: #112233", "theme Custom: headings Didot, body Verdana", "Custom"]);
await cmds.run("edit.undo");
await cmds.run("edit.undo");
eq("undone", [wb.theme.name, wb.styleOf("A1")], ["Harbour", "background: #0f6fc6"]);
await cmds.run("theme.color role=accent6 color=#abcdef");
await cmds.run("file.save name=t");
await cmds.run("theme.set name=Office");
const wb2 = new Workbook(), cmds2 = new Commands(wb2, store);
await cmds2.run("file.open name=t");
eq("saved with the file", [wb2.theme.name, wb2.theme.colors.accent6, wb2.theme.colors.accent1, wb2.styleOf("A1")], ["Custom", "#abcdef", "#0f6fc6", "background: #0f6fc6"]);
await cmds2.run("file.new");
eq("a new workbook is Office", wb2.theme.name, "Office");
eq("refusals", [await message("theme.set name=Plaid"), await message("theme.color role=accent9 color=#000000"), await message("theme.color role=accent1 color=accent2")],
  ["Plaid is not a theme (Office, Greyscale, Harbour, Ember)", "accent9 is not a theme colour (dark1, light1, dark2, light2, accent1, accent2, accent3, accent4, accent5, accent6, hyperlink)", "accent2 is not a colour (#rrggbb, or a name)"]);
const said: string[] = [];
await cmds.script('SAY SETTHEME("Ember")\nSAY GETTHEME()\nSAY SETTHEMECOLOR("accent1", "red")\nSAY SETTHEMEFONTS("", "Optima")', (l) => said.push(l));
eq("from REXX", said, ["theme Ember: accents #c0504d #d26b2a #e8a33a #9b8357 #7a5c46 #a17d6c, headings Georgia, body Georgia", "Ember", "theme Custom: accent1 #ff0000", "theme Custom: headings Georgia, body Optima"]);

// expect: ok a role resolved by the theme; anything else as it is
// expect: ok a look naming a role, kept as the role
// expect: ok another theme re-colours them
// expect: ok a role customised, the fonts too
// expect: ok undone
// expect: ok saved with the file
// expect: ok a new workbook is Office
// expect: ok refusals
// expect: ok from REXX
