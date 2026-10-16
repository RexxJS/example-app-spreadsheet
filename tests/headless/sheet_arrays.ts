// Dynamic arrays, as Excel 365 has them. A formula may come to an array: a
// range, arithmetic over ranges cell by cell (broadcast across a row or a
// column, #N/A where one runs out), IF over an array, SEQUENCE, FILTER,
// SORT, SORTBY, UNIQUE, TRANSPOSE, an XLOOKUP that returns a row, and MAP,
// REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY with a LAMBDA. In a sheet it spills
// into the cells below and to the right: they show its values and are read
// as any cell, A1# is all of it, and anything in the way (or the sheet's
// edge) makes it #SPILL! until it is cleared.
import { evaluate } from "./sheet/evaluate.ts";
import { parse } from "./sheet/parser.ts";
import { show } from "./sheet/format.ts";
import { CellError, isArray } from "./sheet/values.ts";
import type { Value } from "./sheet/values.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { eq } from "./lib/check.ts";
import { print } from "./sheet/audit.ts";

const cells = new Map<string, Value>([
  ["A1", 1], ["A2", 2], ["A3", 3], ["B1", 10], ["C1", 20],
  ["D1", "a"], ["D2", "b"], ["D3", "c"], ["D4", "d"], ["E1", 1], ["E2", 3], ["E3", 2], ["E4", 5],
  ["F1", "x"], ["F2", "Y"], ["F3", "y"], ["F4", "z"], ["F5", "x"], ["G1", 1], ["H1", 2], ["G2", 3], ["H2", 4],
]);
const env = { get: (k: string): Value => cells.get(k) ?? null };
const text = (v: Value) => (v instanceof CellError ? v.code : show(v));
const ev = (...srcs: string[]): string[] => srcs.map((src) => {
  const v = evaluate(parse(src, () => null), env);
  if (isArray(v)) return Array.from({ length: v.rows }, (_, i) => v.values.slice(i * v.cols, (i + 1) * v.cols).map(text).join(",")).join(";");
  return v instanceof CellError ? `${v.code} ${v.why}` : show(v);
});

eq("a range is an array; one of one cell its value", ev("A1:A3", "A1:A1"), ["1;2;3", "1"]);
eq("arithmetic cell by cell, broadcast, #N/A where one runs out", ev("A1:A3*2", "A1:A3+B1:C1", "A1:A2+A1:A3", "-A1:A2", 'A1:A2&"!"'),
  ["2;4;6", "11,21;12,22;13,23", "2;4;#N/A", "-1;-2", "1!;2!"]);
eq("IF over an array; a function of an array", ev('IF(A1:A3>1,"big","small")', "SUM(A1:A3*2)", "LET(xs, SEQUENCE(3), SUM(xs))"), ["small;big;big", "12", "6"]);
eq("SEQUENCE", ev("SEQUENCE(3)", "SEQUENCE(2,3,10,5)", "SEQUENCE(0)"), ["1;2;3", "10,15,20;25,30,35", "#CALC! SEQUENCE of no rows"]);
eq("FILTER", ev("FILTER(D1:E4,E1:E4>2)", 'FILTER(D1:D4,E1:E4>9,"none")', "FILTER(D1:D4,E1:E4>9)", "FILTER(D1:E4,E1:E2>2)"),
  ["b,3;d,5", "none", "#CALC! FILTER found nothing", "#VALUE! FILTER's include is not the array's height or width"]);
eq("SORT and SORTBY", ev("SORT(E1:E4)", "SORT(D1:E4,2,-1)", "SORTBY(D1:D4,E1:E4)", "SORTBY(D1:D4,E1:E4,-1)"), ["1;2;3;5", "d,5;b,3;c,2;a,1", "a;c;b;d", "d;b;c;a"]);
eq("UNIQUE, any case; or what is there exactly once", ev("UNIQUE(F1:F5)", "UNIQUE(F1:F5,FALSE,TRUE)", "UNIQUE(G1:H1,TRUE)"), ["x;Y;z", "z", "1,2"]);
eq("TRANSPOSE", ev("TRANSPOSE(A1:A3)", "TRANSPOSE(G1:H2)"), ["1,2,3", "1,3;2,4"]);
eq("XLOOKUP's return as a row", ev('XLOOKUP("b",D1:D4,D1:E4)'), ["b,3"]);
eq("MAP, REDUCE, SCAN", ev("MAP(A1:A3,LAMBDA(x,x*x))", "REDUCE(0,A1:A3,LAMBDA(a,x,a+x))", "SCAN(0,A1:A3,LAMBDA(a,x,a+x))", "MAP(A1:A2,B1:C1,LAMBDA(x,y,x+y))"), ["1;4;9", "6", "1;3;6", "#VALUE! MAP's arrays differ in size"]);
// 2.12: optional parameters and ISOMITTED; trim refs and TRIMRANGE;
// implicit intersection.
const at = (cell: string, ...srcs: string[]): string[] => srcs.map((src) => {
  const v = evaluate(parse(src, () => null), env, cell);
  return v instanceof CellError ? `${v.code} ${v.why}` : isArray(v) ? `${v.rows}x${v.cols}` : show(v);
});
eq("a LAMBDA's [optional] parameter: left out, or passed as nothing, ISOMITTED says so and it reads as empty", ev(
  "LET(f, LAMBDA(x, [y], IF(ISOMITTED(y), x, x*y)), f(5))", "LET(f, LAMBDA(x, [y], IF(ISOMITTED(y), x, x*y)), f(5, 2))", "LET(f, LAMBDA(x, [y], IF(ISOMITTED(y), x, x*y)), f(5, ))",
  "LAMBDA(x, [y], y + 1)(1)", "LAMBDA(x, ISOMITTED(x))(1)", "ISOMITTED(A1)", "LAMBDA(x, [y], [z], x + IF(ISOMITTED(z), 0, z))(1, 2)"), ["5", "10", "5", "1", "FALSE", "FALSE", "1"]);
eq("its errors: too few or too many, optional ones first, brackets elsewhere", ev("LAMBDA(x, [y], x)()", "LAMBDA(x, [y], x)(1, 2, 3)"), ["#VALUE! LAMBDA takes 1 to 2 values, given 0", "#VALUE! LAMBDA takes 1 to 2 values, given 3"]);
eq("and what does not parse", ["LAMBDA([x], y, x)", "LET([x], 1, x)", "[x] + 1"].map((src) => { try { parse(src, () => null); return "parsed"; } catch (e) { return String((e as Error).message); } }), [
  "LAMBDA's optional parameters come last: Y follows [X]", "LET's names take no brackets: only a LAMBDA's parameters are optional", "[x]: a name in brackets is a LAMBDA's optional parameter",
]);
eq("a trim ref drops empty rows at its end, its start, or both; TRIMRANGE the same", [...ev("ROWS(A1:.A9)", "ROWS(A1:A9)", "SUM(A1:.A9)", "COLUMNS(A1:.J1)", "ROWS(D1:.E9)", "ROWS(TRIMRANGE(A1:A9))", "ROWS(TRIMRANGE(A1:A9, 0))", "COLUMNS(TRIMRANGE(A1:J1, 3, 1))"),
  ...at("A1", "ROWS(Z1.:Z9)", "ROWS(Z1:Z9)")], ["3", "9", "6", "8", "4", "3", "9", "10", "#REF! Z1:Z9 trims to nothing", "9"]);
eq("TRIMRANGE's own errors", ev("TRIMRANGE(Z1:Z9)", "TRIMRANGE(A1:A9, 4)"), ["#REF! TRIMRANGE trims to nothing", "#VALUE! TRIMRANGE trims 0 to 3, not 4 and 3"]);
eq("printed and stepped with its dots", [print(parse("SUM(A1:.A9)", () => null)), print(parse("SUM(A1.:.A9)+@B1:B9", () => null)), print(parse("LAMBDA(x,[y],x)", () => null))], ["SUM(A1:.A9)", "SUM(A1.:.A9)+@B1:B9", "LAMBDA(X,[Y],X)"]);
eq("@: of a column the cell in the formula's row, of a row its column, of an array its first value", at("Z2", "@A1:A3", "@D1:D4&@E1:E4", "@B1:C1*0+@G1:H1", "@SEQUENCE(3)+1", "@A1", "@A5:A9", "@A1:B3"),
  ["2", "b3", "#VALUE! @B1:C1 does not meet the formula's row or column", "2", "1", "#VALUE! @A5:A9 does not meet the formula's row or column", "#VALUE! @A1:B3 does not meet the formula's row or column"]);
eq("and in the formula's column", at("B9", "@A1:C1", "@G1:H2"), ["10", "#VALUE! @G1:H2 does not meet the formula's row or column"]);
eq("BYROW, BYCOL, MAKEARRAY", ev("BYROW(G1:H2,LAMBDA(r,SUM(r)))", "BYCOL(G1:H2,LAMBDA(c,SUM(c)))", "MAKEARRAY(2,2,LAMBDA(i,j,i*10+j))", "MAP(A1:A2,5)"),
  ["3;7", "4,6", "11,12;21,22", "#VALUE! MAP takes a LAMBDA last"]);

const wb = new Workbook();
const cmds = new Commands(wb);
const shown = (...keys: string[]) => keys.map((k) => show(wb.sheet.value(k)));
const set = (cell: string, input: string) => cmds.run(`cell.set cell=${cell} input="${input.replace(/"/g, '""')}"`);
await set("A1", "=SEQUENCE(3)");
await set("B1", "=SUM(A1#)");
await set("B2", "=A3*10");
eq("a formula spills below; its cells are read as any; A1# is all of it", [shown("A1", "A2", "A3", "B1", "B2"), wb.sheet.input("A2")], [["1", "2", "3", "6", "30"], ""]);
await set("A1", "=SEQUENCE(4)");
eq("a bigger spill: what reads it follows", shown("A4", "B1", "B2"), ["4", "10", "30"]);
await set("C1", "=A4+1");
eq("something in the way: #SPILL!, the spill gone", [(await set("A3", "x")) && shown("A1", "A2", "B1", "C1"), String((wb.sheet.value("A1") as CellError).why)],
  [["#SPILL!", "", "#SPILL!", "1"], "A3 is in the way"]);
await set("A3", "");
eq("cleared: it spills again", shown("A1", "A3", "B1", "C1"), ["1", "3", "10", "5"]);
await cmds.run("edit.undo");
eq("undo puts the blocker back", shown("A1", "A3"), ["#SPILL!", "x"]);
await cmds.run("edit.redo");
await set("Z99", "=SEQUENCE(3)");
eq("the sheet's edge", [shown("Z99"), String((wb.sheet.value("Z99") as CellError).why)], [["#SPILL!"], "the result does not fit on the sheet"]);
await set("A1", "=SEQUENCE(2)");
eq("a smaller spill leaves what it left empty", shown("A3", "A4", "B1", "C1"), ["", "", "3", "1"]);
await cmds.run("edit.copy range=B1");
await cmds.run("edit.paste at=E2");
eq("A1# moves in a copy", wb.sheet.input("E2"), "=SUM(D2#)");
eq("a spilled cell is styled as what it shows", wb.styleOf("A2"), "text-align: end");
// In a sheet: a named LAMBDA with an optional parameter, calling itself; a
// trim ref that grows as its column fills; @ in a cell.
await cmds.run('names.define name=TWICES formula="=LAMBDA(n, [acc], IF(n = 0, IF(ISOMITTED(acc), 1, acc), TWICES(n - 1, IF(ISOMITTED(acc), 2, acc * 2))))"');
await set("K1", "=TWICES(10)");
await set("K2", "=TWICES(0)");
await set("L1", "4");
await set("L2", "5");
await set("M1", "=SUM(L1:.L50)");
await set("M2", "=ROWS(L1:.L50)");
eq("a named LAMBDA with an optional parameter calls itself; a trim ref counts what is there", shown("K1", "K2", "M1", "M2"), ["1024", "1", "9", "2"]);
await set("L3", "6");
await set("N3", "=@L1:L50 * 10");
eq("and grows as the column fills; @ takes its own row", shown("M1", "M2", "N3"), ["15", "3", "60"]);

// expect: ok a range is an array; one of one cell its value
// expect: ok arithmetic cell by cell, broadcast, #N/A where one runs out
// expect: ok IF over an array; a function of an array
// expect: ok SEQUENCE
// expect: ok FILTER
// expect: ok SORT and SORTBY
// expect: ok UNIQUE, any case; or what is there exactly once
// expect: ok TRANSPOSE
// expect: ok XLOOKUP's return as a row
// expect: ok MAP, REDUCE, SCAN
// expect: ok a LAMBDA's [optional] parameter: left out, or passed as nothing, ISOMITTED says so and it reads as empty
// expect: ok its errors: too few or too many, optional ones first, brackets elsewhere
// expect: ok and what does not parse
// expect: ok a trim ref drops empty rows at its end, its start, or both; TRIMRANGE the same
// expect: ok TRIMRANGE's own errors
// expect: ok printed and stepped with its dots
// expect: ok @: of a column the cell in the formula's row, of a row its column, of an array its first value
// expect: ok and in the formula's column
// expect: ok BYROW, BYCOL, MAKEARRAY
// expect: ok a formula spills below; its cells are read as any; A1# is all of it
// expect: ok a bigger spill: what reads it follows
// expect: ok something in the way: #SPILL!, the spill gone
// expect: ok cleared: it spills again
// expect: ok undo puts the blocker back
// expect: ok the sheet's edge
// expect: ok a smaller spill leaves what it left empty
// expect: ok A1# moves in a copy
// expect: ok a spilled cell is styled as what it shows
// expect: ok a named LAMBDA with an optional parameter calls itself; a trim ref counts what is there
// expect: ok and grows as the column fills; @ takes its own row
