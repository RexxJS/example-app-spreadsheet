// app/sheet/described.ts held to the functions themselves: every name
// LISTCOMMANDS() gives is described and nothing else is, each one's
// parameters are as many as it takes, and LISTCOMMANDS(1) gives them, over
// the control bus as a program driving the app asks for them.
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { CONTROL, controlFunctions } from "./sheet/control.ts";
import { DESCRIBED, signature } from "./sheet/described.ts";
import { rangeFunctions } from "./sheet/rexx.ts";
import { eq } from "./lib/check.ts";

const wb = new Workbook();
const cmds = new Commands(wb);
const asked = async (code: string) => {
  const o = await cmds.spreadsheet(code, () => {});
  if (!o.ok) throw new Error(o.message);
  return o.output as Record<string, any>;
};
const unstem = (s: Record<string, any>) => Array.from({ length: Number(s[0]) }, (_, i) => s[i + 1]);

const names: string[] = unstem(await asked("LISTCOMMANDS()"));
eq("LISTCOMMANDS() lists about 130 names", names.length >= 125, true);
eq("every name it lists is described", names.filter((n) => !DESCRIBED[n]), []);
eq("and nothing is described that it does not list", Object.keys(DESCRIBED).filter((n) => !names.includes(n)), []);

// The parameters a function takes before any rest parameter, by its length:
// a control function's first is `exec` (or `_`), which a script never gives.
const fixed = (name: string) => DESCRIBED[name].params.filter((x) => !x.variadic).length;
const takes: Record<string, number> = {};
for (const [name, f] of Object.entries(CONTROL)) takes[name] = Math.max(0, f.length - 1);
for (const [name, f] of Object.entries(rangeFunctions(() => ""))) takes[name] = f.length;
const own = controlFunctions(async () => ({ ok: true, message: "" }), {});
takes.LISTCOMMANDS = own.LISTCOMMANDS.length;
takes.BATCH_EXECUTE = own.BATCH_EXECUTE.length;
takes.SETCELL = 2;   // script.ts's (cell, input)
takes.GETCELL = 1;
takes.RUN_PROGRAM = 2;   // script.ts's (source, encoding)
eq("each one's parameters are as many as it takes", Object.keys(takes).filter((n) => DESCRIBED[n] && fixed(n) !== takes[n]).map((n) => `${n}: ${fixed(n)} described, takes ${takes[n]}`), []);
eq("only SETCELLVALIDATION takes the rest of its arguments", Object.keys(DESCRIBED).filter((n) => DESCRIBED[n].params.some((x) => x.variadic)), ["SETCELLVALIDATION"]);
const TYPES = ["cell", "range", "row", "column", "text", "number", "flag", "json", "stem", "query"];
eq("every parameter is of a kind described.ts names", Object.entries(DESCRIBED).flatMap(([n, d]) => d.params.filter((x) => !TYPES.includes(x.type)).map((x) => `${n}.${x.name}`)), []);
eq("and every function says what it does and gives", Object.entries(DESCRIBED).filter(([, d]) => !d.does || !d.gives).map(([n]) => n), []);
eq("optional parameters come last", Object.entries(DESCRIBED).filter(([, d]) => d.params.some((x, i) => x.optional && d.params.slice(i + 1).some((y) => !y.optional))).map(([n]) => n), []);

const described: Record<string, any>[] = unstem(await asked("LISTCOMMANDS(1)"));
eq("LISTCOMMANDS(1) describes each name LISTCOMMANDS() gives, in its order", described.map((d) => d.name), names);
const setcell = described.find((d) => d.name === "SETCELL")!;
eq("each with its signature, what it does and gives", [setcell.signature, setcell.does, setcell.gives], ["SETCELL(cell, value)", "Sets a cell, as typing it does", "the value set"]);
eq("its parameters a stem, typed", unstem(setcell.params).map((x: any) => [x.name, x.type, x.optional]), [["cell", "cell", 0], ["value", "text", 0]]);
const sort = described.find((d) => d.name === "SORTRANGE")!;
eq("an optional one in brackets, and said so", [sort.signature, unstem(sort.params)[2].optional, unstem(sort.params)[2].about],
  ["SORTRANGE(range, by, [ascending])", 1, "1 or left out: ascending; 0: descending"]);
eq("the rest of the arguments as ...", signature("SETCELLVALIDATION", DESCRIBED.SETCELLVALIDATION), "SETCELLVALIDATION(range, type, [limits...])");
eq("LISTCOMMANDS(0) is the names, as before", unstem(await asked("LISTCOMMANDS(0)")), names);
eq("asking changes nothing: no step of undo", wb.canUndo, false);

// expect: ok LISTCOMMANDS() lists about 130 names
// expect: ok every name it lists is described
// expect: ok and nothing is described that it does not list
// expect: ok each one's parameters are as many as it takes
// expect: ok only SETCELLVALIDATION takes the rest of its arguments
// expect: ok every parameter is of a kind described.ts names
// expect: ok and every function says what it does and gives
// expect: ok optional parameters come last
// expect: ok LISTCOMMANDS(1) describes each name LISTCOMMANDS() gives, in its order
// expect: ok each with its signature, what it does and gives
// expect: ok its parameters a stem, typed
// expect: ok an optional one in brackets, and said so
// expect: ok the rest of the arguments as ...
// expect: ok LISTCOMMANDS(0) is the names, as before
// expect: ok asking changes nothing: no step of undo
