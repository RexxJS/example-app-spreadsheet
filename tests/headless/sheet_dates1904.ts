// The 1904 date system (1.6): the same serial, two dates; what the date
// functions give follows the system; saved with the file; a new workbook
// is 1900; SETDATESYSTEM and GETDATESYSTEM from REXX.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { writeSheetFile } from "./sheet/file.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const set = (cell: string, input: string) => cmds.execute("cell.set", { cell, input });
await cmds.run("formulas.use language=Excel");
await set("A1", "45306");
await cmds.run('cell.format cell=A1 format="yyyy-mm-dd"');
await set("B1", "=DATE(2024, 1, 15)");
await set("C1", "=YEAR(A1)");
await set("D1", '=TEXT(0, "yyyy-mm-dd")');
eq("1900: serial 45306 is 2024-01-15", [wb.textOf("A1"), show(wb.sheet.value("B1")), show(wb.sheet.value("C1")), show(wb.sheet.value("D1"))], ["2024-01-15", "45306", "2024", "1899-12-30"]);
eq("1904", (await cmds.execute("workbook.dateSystem", { system: "1904" })).message, "the 1904 date system");
eq("the same serial four years and a day on; DATE gives 1,462 less; 0 is 1904-01-01", [wb.textOf("A1"), show(wb.sheet.value("B1")), show(wb.sheet.value("C1")), show(wb.sheet.value("D1"))], ["2028-01-16", "43844", "2028", "1904-01-01"]);
eq("saved with the file", (writeSheetFile(wb) as { dateSystem?: number }).dateSystem, 1904);
const file = JSON.stringify(writeSheetFile(wb));
await cmds.execute("file.new");
eq("a new workbook is 1900", [wb.dateSystem, ((await cmds.execute("sheet.state")).output as { dateSystem: number }).dateSystem], [1900, 1900]);
await cmds.execute("file.import", { text: file });
eq("and a file read back is 1904 again", [wb.dateSystem, show(wb.sheet.value("B1"))], [1904, "43844"]);
eq("refused", (await cmds.execute("workbook.dateSystem", { system: "2000" })).message, "2000 is not a date system (1900, 1904)");
const said: string[] = [];
eq("SETDATESYSTEM and GETDATESYSTEM from REXX", [(await cmds.script('SAY GETDATESYSTEM()\nSAY SETDATESYSTEM(1900)', (l) => said.push(l))).message, said, show(wb.sheet.value("B1"))], ["ran", ["1904", "1900"], "45306"]);

// expect: ok 1900: serial 45306 is 2024-01-15
// expect: ok 1904
// expect: ok the same serial four years and a day on; DATE gives 1,462 less; 0 is 1904-01-01
// expect: ok saved with the file
// expect: ok a new workbook is 1900
// expect: ok and a file read back is 1904 again
// expect: ok refused
// expect: ok SETDATESYSTEM and GETDATESYSTEM from REXX
