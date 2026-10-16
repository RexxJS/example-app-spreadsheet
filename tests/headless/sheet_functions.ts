// app/sheet/functions.ts through formulas, against a fixed set of cells:
// every function in the library, its edge cases, and the errors it gives.
import { evaluate } from "./sheet/evaluate.ts";
import { criterion, names } from "./sheet/functions.ts";
import { CellError, isArray } from "./sheet/values.ts";
import type { Value } from "./sheet/values.ts";
import { parse } from "./sheet/parser.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

// A: numbers. B: text with a blank and a gap. C: numbers beside them.
// D: odds and ends. E: an ascending column. F1:H4: a lookup table.
const cells = new Map<string, Value>([
  ["A1", 3], ["A2", 1], ["A3", 4], ["A4", 1], ["A5", 5], ["A6", 9],
  ["B1", "apples"], ["B2", "pears"], ["B3", "apples"], ["B4", ""], ["B6", "Plums"],
  ["C1", 10], ["C2", 20], ["C3", 30], ["C4", 40], ["C5", 50], ["C6", 60],
  ["D1", true], ["D2", new CellError("#DIV/0!", "divide by zero")], ["D3", "  lots   of   space  "], ["D4", "hello world"], ["D5", "5"],
  ["E1", 10], ["E2", 20], ["E3", 30], ["E4", 40], ["E5", 50],
  ["F1", "id"], ["G1", "fruit"], ["H1", "price"],
  ["F2", 1], ["G2", "apple"], ["H2", 0.5],
  ["F3", 2], ["G3", "pear"], ["H3", 0.75],
  ["F4", 3], ["G4", "plum"], ["H4", 1.2],
  ["I1", 40], ["I2", 30], ["I3", 20], ["I4", 10],
  // A cell of another sheet, as the sheet's Env would answer it.
  ["COSTS!B2", 7],
]);
const env = { get: (k: string): Value => cells.get(k) ?? null };
const ev = (...srcs: string[]): string[] => srcs.map((src) => show(evaluate(parse(src), env)));
const why = (...srcs: string[]): string[] => srcs.map((src) => {
  const v = evaluate(parse(src), env);
  return v instanceof CellError ? `${v.code} ${v.why}` : `not an error: ${show(v)}`;
});

// An array as rows x cols and its values; a value as shown.
const grid = (...srcs: string[]): string[] => srcs.map((src) => {
  const v = evaluate(parse(src), env);
  return isArray(v) ? `${v.rows}x${v.cols}:${v.values.map(show).join(",")}` : show(v);
});

eq("the library", names().join(" "), [
  "ABS ACOS ACOSH ADDRESS AGGREGATE AND ARRAYTOTEXT ASIN ASINH ATAN ATAN2 ATANH AVERAGE AVERAGEA AVERAGEIF",
  "AVERAGEIFS CEILING CEILING.MATH CEILING.PRECISE CHAR CHOOSE CHOOSECOLS CHOOSEROWS CLEAN CODE COLUMN COLUMNS",
  "COMBIN CONCAT CONCATENATE COS COSH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS DATE DATEDIF DATEVALUE DAY DAYS",
  "DAYS360 DEGREES DOLLAR DROP EDATE EOMONTH ERROR.TYPE EVEN EXACT EXP EXPAND FACT FALSE FILTER FIND FIXED FLOOR",
  "FLOOR.MATH FLOOR.PRECISE FORMULATEXT FV GCD GROUPBY HLOOKUP HOUR HSTACK IF IFERROR IFNA IFS INDEX INDIRECT INFO",
  "INT IRR ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISOWEEKNUM ISREF ISTEXT",
  "LARGE LCM LEFT LEN LN LOG LOG10 LOOKUP LOWER MATCH MAX MAXA MAXIFS MEDIAN MID MIN MINA MINIFS MINUTE MOD MODE",
  "MONTH MROUND N NA NETWORKDAYS NETWORKDAYS.INTL NOT NOW NPER NPV NUMBERVALUE ODD OFFSET OR PERCENTILE",
  "PERCENTILE.EXC PERCENTILE.INC PERMUT PI PIVOTBY PMT POWER PRODUCT PROPER PV QUARTILE QUARTILE.EXC QUARTILE.INC",
  "QUOTIENT RADIANS RAND RANDARRAY RANDBETWEEN RANK RATE REGEXEXTRACT REGEXREPLACE REGEXTEST REPLACE REPT RIGHT",
  "ROUND ROUNDDOWN ROUNDUP ROW ROWS SEARCH SECOND SEQUENCE SHEET SHEETS SIGN SIN SINH SMALL SORT SORTBY SPARKLINE",
  "SQRT STDEV STDEVP SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SWITCH T TAKE TAN TANH TEXT TEXTAFTER",
  "TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE TOCOL TODAY TOROW TRANSPOSE TRIM TRUE TRUNC TYPE UNICHAR UNICODE",
  "UNIQUE UPPER VALUE VALUETOTEXT VAR VARP VLOOKUP VSTACK WEEKDAY WEEKNUM WORKDAY WORKDAY.INTL WRAPCOLS WRAPROWS",
  "XLOOKUP XMATCH XOR YEAR YEARFRAC",
].join(" "));

// Maths and statistics.
eq("PRODUCT", ev("PRODUCT(A1:A3)", "PRODUCT(B1:B2)"), ["12", "0"]);
eq("MEDIAN, odd and even", ev("MEDIAN(A1:A5)", "MEDIAN(A1:A6)"), ["3", "3.5"]);
eq("STDEV is the sample's", ev("ROUND(STDEV(C1:C6), 4)", "STDEV(2, 4, 4, 4, 5, 5, 7, 9)"), ["18.7083", "2.138089935"]);
eq("COUNTA counts what is there, COUNTBLANK what is not", ev("COUNTA(B1:B6)", "COUNTBLANK(B1:B6)"), ["5", "2"]);
eq("SIGN, INT", ev("SIGN(-3)", "SIGN(0)", "INT(2.7)", "INT(-2.5)"), ["-1", "0", "2", "-3"]);
eq("ROUND away from zero, with or without places", ev("ROUND(2.5)", "ROUND(-2.5)", "ROUND(2.675, 2)", "ROUND(1234, -2)"), ["3", "-3", "2.68", "1200"]);
eq("ROUNDUP and ROUNDDOWN", ev("ROUNDUP(2.01)", "ROUNDUP(-2.01)", "ROUNDDOWN(2.99)", "ROUNDDOWN(-2.99, 1)"), ["3", "-3", "2", "-2.9"]);
eq("MOD takes the divisor's sign", ev("MOD(7, 3)", "MOD(-7, 3)", "MOD(7, -3)", "MOD(5.5, 2)"), ["1", "2", "-2", "1.5"]);
eq("POWER, SQRT, PI", ev("POWER(2, 10)", "SQRT(A6)", "ROUND(PI(), 5)"), ["1024", "3", "3.14159"]);
eq("their errors", why("MOD(1, 0)", "SQRT(-1)", "MEDIAN(B1:B2)", "STDEV(1)"), [
  "#DIV/0! MOD by zero", "#NUM! the square root of a negative number", "#NUM! MEDIAN of no numbers", "#DIV/0! STDEV needs two numbers or more",
]);
eq("an error in a range flows out of an aggregate", why("SUM(D1:D2)", "MAX(D2, 1)"), ["#DIV/0! divide by zero", "#DIV/0! divide by zero"]);

// Conditional sums and counts.
eq("COUNTIF by text, ignoring case", ev(`COUNTIF(B1:B6, "apples")`, `COUNTIF(B1:B6, "PLUMS")`), ["2", "1"]);
eq("COUNTIF by comparison", ev(`COUNTIF(A1:A6, ">3")`, `COUNTIF(A1:A6, "<=1")`, `COUNTIF(A1:A6, "<>1")`, "COUNTIF(A1:A6, 1)"), ["3", "2", "4", "2"]);
eq("COUNTIF of blanks, and of what is not", ev(`COUNTIF(B1:B6, "")`, `COUNTIF(B1:B6, "<>")`), ["2", "4"]);
eq("a number criterion skips text, and text skips numbers", ev(`COUNTIF(B1:B6, ">0")`, `COUNTIF(A1:A6, "apples")`), ["0", "0"]);
eq("SUMIF over another range", ev(`SUMIF(B1:B6, "apples", C1:C6)`, `SUMIF(A1:A6, ">3")`), ["40", "18"]);
eq("AVERAGEIF", ev(`AVERAGEIF(B1:B6, "apples", C1:C6)`), ["20"]);
eq("AVERAGEIF of nothing", why(`AVERAGEIF(B1:B6, "kiwis", C1:C6)`), ["#DIV/0! AVERAGEIF matched no numbers"]);
eq("criterion on its own", [criterion(">=4")(4), criterion(">=4")(3), criterion("=pears")("Pears"), criterion("")(null), criterion(3)("3")], [true, false, true, true, false]);

// Logic and information.
eq("AND, OR, NOT", ev("AND(A1>1, A2=1)", "AND(A1:A6)", "OR(A2>1, FALSE)", "OR(D1)", "NOT(A1=3)"), ["TRUE", "TRUE", "FALSE", "TRUE", "FALSE"]);
eq("AND with nothing true or false to look at", why("AND(B1:B2)"), ["#VALUE! AND found no TRUE or FALSE"]);
eq("IFERROR catches, and passes the rest", ev(`IFERROR(1/0, "none")`, `IFERROR(D2, 0)`, `IFERROR(A1, "none")`), ["none", "0", "3"]);
eq("IS functions", ev("ISERROR(D2)", "ISERROR(A1)", "ISNUMBER(A1)", "ISNUMBER(D5)", "ISTEXT(B1)", "ISBLANK(B5)", "ISBLANK(B4)"), [
  "TRUE", "FALSE", "TRUE", "FALSE", "TRUE", "TRUE", "FALSE",
]);
eq("the rest of the family: ISERR, ISNA, ISLOGICAL, ISNONTEXT, ISEVEN, ISODD", ev("ISERR(D2)", "ISERR(NA())", "ISERR(A1)", "ISNA(NA())", "ISNA(D2)", "ISLOGICAL(D1)", "ISLOGICAL(1)", "ISNONTEXT(A1)", "ISNONTEXT(B1)", "ISNONTEXT(B5)", "ISNONTEXT(D2)", "ISEVEN(A3)", "ISEVEN(-2.5)", "ISEVEN(0)", "ISODD(A1)", "ISODD(2.9)", "ISODD(-3)"), [
  "TRUE", "FALSE", "FALSE", "TRUE", "FALSE", "TRUE", "FALSE", "TRUE", "FALSE", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "FALSE", "TRUE",
]);
eq("ISEVEN of words", why("ISEVEN(B1)", "ISODD(D2)"), ["#VALUE! 'apples' is not a number", "#DIV/0! divide by zero"]);
eq("N and T: a number's number, a text's text, else 0 or nothing", ev("N(A1)", "N(D1)", "N(FALSE)", "N(B1)", "N(B5)", "N(D5)", "T(B1)", "T(A1)", "T(D1)", "T(D5)", `LEN(T(B5))`), ["3", "1", "0", "0", "0", "0", "apples", "", "", "5", "0"]);
eq("N of an error is the error", why("N(D2)", "T(D2)"), ["#DIV/0! divide by zero", "#DIV/0! divide by zero"]);
eq("TYPE: number 1, text 2, logical 4, error 16, array 64", ev("TYPE(A1)", "TYPE(B1)", "TYPE(D1)", "TYPE(D2)", "TYPE(A1:A3)", "TYPE({1,2})", "TYPE(B5)", `TYPE("x")`), ["1", "2", "4", "16", "64", "64", "1", "2"]);
eq("ERROR.TYPE: Excel's numbers, the app's own past them", ev("ERROR.TYPE(D2)", "ERROR.TYPE(NA())", "ERROR.TYPE(1/0)", "ERROR.TYPE(#REF!)", "ERROR.TYPE(#VALUE!)", "ERROR.TYPE(#NAME?)", "ERROR.TYPE(#NUM!)", "ERROR.TYPE(#CALC!)"), ["2", "7", "2", "4", "3", "5", "6", "14"]);
eq("ERROR.TYPE of a value is #N/A", why("ERROR.TYPE(A1)", "ERROR.TYPE(B5)"), ["#N/A ERROR.TYPE of a value that is not an error", "#N/A ERROR.TYPE of a value that is not an error"]);
eq("XOR: an odd number TRUE; TRUE() and FALSE() as calls", ev("XOR(TRUE, FALSE)", "XOR(TRUE, TRUE)", "XOR(1, 1, 1)", "XOR(0, 0)", "XOR(A1:A6)", "XOR(D1, 0)", "TRUE()", "FALSE()", "IF(TRUE(), 1, 2)", "NOT(FALSE())"), ["TRUE", "FALSE", "TRUE", "FALSE", "FALSE", "TRUE", "TRUE", "FALSE", "1", "TRUE"]);
eq("XOR of no truths", why("XOR(B1:B2)", "XOR(D2)"), ["#VALUE! XOR found no TRUE or FALSE", "#DIV/0! divide by zero"]);
eq("INFO, what applies: numfile and recalc", [...ev(`INFO("numfile")`, `INFO("recalc")`, `INFO("NumFile")`), ...why(`INFO("osversion")`, `INFO(1, 2)`)], ["1", "Automatic", "1", "#VALUE! INFO of 'osversion': only numfile and recalc apply here", "#VALUE! INFO takes 1 values"]);
eq("CHOOSE", ev(`CHOOSE(2, "a", "b", "c")`, "CHOOSE(A2, C1, C2)"), ["b", "10"]);
eq("CHOOSE out of range", why(`CHOOSE(4, "a", "b")`), ["#VALUE! CHOOSE has no choice 4"]);

// Text.
eq("case", ev("UPPER(B1)", `LOWER("MiXeD")`, `PROPER("the o'neil BROTHERS")`), ["APPLES", "mixed", "The O'Neil Brothers"]);
eq("TRIM collapses spaces", ev("TRIM(D3)", "LEN(TRIM(D3))"), ["lots of space", "13"]);
eq("LEFT, RIGHT, MID", ev("LEFT(D4)", "LEFT(D4, 5)", "RIGHT(D4, 5)", "RIGHT(D4, 0)", "MID(D4, 7, 3)", "LEFT(D4, 99)"), ["h", "hello", "world", "", "wor", "hello world"]);
eq("FIND is case-sensitive, from a start", ev(`FIND("o", D4)`, `FIND("o", D4, 6)`, `FIND("w", D4)`), ["5", "8", "7"]);
eq("FIND not finding", why(`FIND("W", D4)`), ["#VALUE! 'W' is not in 'hello world'"]);
eq("SUBSTITUTE, REPT, EXACT", ev(`SUBSTITUTE(D4, "o", "0")`, `SUBSTITUTE(D4, "", "x")`, `REPT("ab", 3)`, `EXACT("a", "A")`, `EXACT(B1, B3)`), [
  "hell0 w0rld", "hello world", "ababab", "FALSE", "TRUE",
]);
eq("TEXTJOIN, skipping blanks or not", ev(`TEXTJOIN(", ", TRUE, B1:B6)`, `TEXTJOIN("-", FALSE, B1:B4)`), ["apples, pears, apples, Plums", "apples-pears-apples-"]);
eq("VALUE", ev("VALUE(D5)+1", `VALUE(" 2.5 ")`), ["6", "2.5"]);
eq("VALUE of words", why("VALUE(B1)"), ["#VALUE! 'apples' is not a number"]);

// Aggregates (2.6). The Env here hides nothing: a sheet's does (sheet_book).
eq("SUBTOTAL by function number, 1 to 11 and 101 to 111", ev("SUBTOTAL(9, A1:A6)", "SUBTOTAL(109, A1:A6)", "SUBTOTAL(1, C1:C3)", "SUBTOTAL(2, A1:B6)", "SUBTOTAL(3, B1:B6)", "SUBTOTAL(4, A1:A6, C1:C6)", "SUBTOTAL(5, E1:E5)", "SUBTOTAL(6, A1:A3)", "ROUND(SUBTOTAL(7, A1:A6), 4)", "SUBTOTAL(11, C1:C3)"),
  ["23", "23", "20", "6", "5", "60", "10", "12", "2.9944", "66.66666667"]);
eq("SUBTOTAL's errors", why("SUBTOTAL(12, A1:A6)", "SUBTOTAL(0, A1:A6)", "SUBTOTAL(9)", "SUBTOTAL(9, D2)"), ["#VALUE! SUBTOTAL's function number 12: 1 to 11, or 101 to 111", "#VALUE! SUBTOTAL's function number 0: 1 to 11, or 101 to 111", "#VALUE! SUBTOTAL takes a function number and a reference", "#DIV/0! divide by zero"]);
eq("AGGREGATE: the function, the options (2 skips errors), the k forms", ev("AGGREGATE(9, 6, A1:A6, D2)", "AGGREGATE(1, 6, A1:A6, D2)", "AGGREGATE(14, 6, A1:A6, 2)", "AGGREGATE(15, 4, A1:A6, 1)", "AGGREGATE(12, 0, A1:A6)", "AGGREGATE(13, 0, A1:A6)", "AGGREGATE(16, 6, A1:A6, 0.5)", "AGGREGATE(17, 6, A1:A6, 1)", "AGGREGATE(18, 6, A1:A6, 0.5)", "AGGREGATE(19, 6, A1:A6, 3)", "AGGREGATE(4, 0, {1, 5, 2})"),
  ["23", "3.833333333", "5", "1", "3.5", "1", "3.5", "1.5", "3.5", "6", "5"]);
eq("AGGREGATE's errors: an error kept, a bad number or options", why("AGGREGATE(9, 4, A1:A6, D2)", "AGGREGATE(20, 0, A1:A6)", "AGGREGATE(9, 8, A1:A6)", "AGGREGATE(14, 6, A1:A6, 9)"), ["#DIV/0! divide by zero", "#VALUE! AGGREGATE's function number 20: 1 to 19", "#VALUE! AGGREGATE's options are 0 to 7, not 8", "#NUM! LARGE's k is 9 of 6 numbers"]);
eq("MAXA, MINA, AVERAGEA: TRUE 1, text and FALSE 0", ev("MAXA(A1:A3, D1)", "MAXA(B1:B6)", "MINA(A1:A6, B1)", "MINA(D1, 2)", "AVERAGEA(A1:A3, D1)", "AVERAGEA(C1:C2, B1)", "AVERAGEA(B5)"), ["4", "0", "0", "1", "2.25", "10", "#DIV/0!"]);
eq("COUNTA over mixed cells: everything but the empty, the empty text included", ev("COUNTA(A1:D6)", "COUNTA(B1:B6)", `COUNTA(B5, "")`, "COUNTA(D1:D5)"), ["22", "5", "1", "5"]);
eq("the .INC and .EXC quantiles", ev("PERCENTILE.INC(A1:A6, 0.5)", "QUARTILE.INC(A1:A6, 3)", "PERCENTILE.EXC(A1:A6, 0.5)", "PERCENTILE.EXC(A1:A6, 0.25)", "QUARTILE.EXC(A1:A6, 1)", "QUARTILE.EXC(A1:A6, 3)"), ["3.5", "4.75", "3.5", "1", "1", "6"]);
eq("their ends", why("PERCENTILE.EXC(A1:A6, 0.1)", "PERCENTILE.EXC(A1:A6, 1)", "QUARTILE.EXC(A1:A6, 0)"), ["#NUM! the 0.1 point lies outside the numbers", "#NUM! PERCENTILE.EXC's k is between 0 and 1, not 1", "#NUM! QUARTILE.EXC's quart is 1, 2 or 3, not 0"]);
eq("GROUPBY: each group's values folded by a LAMBDA, or a function's bare name, sorted, with a total", grid("GROUPBY(B1:B3, A1:A3, LAMBDA(v, SUM(v)))", "GROUPBY(B1:B3, A1:A3, SUM)", "GROUPBY(B1:B3, A1:A3, COUNT, 0, 0)", "GROUPBY(B1:B3, A1:A3, MAX, 0, 1, -1)", "GROUPBY(G1:G4, H1:H4, SUM, 3)"),
  ["3x2:apples,7,pears,1,Total,8", "3x2:apples,7,pears,1,Total,8", "2x2:apples,2,pears,1", "3x2:pears,1,apples,4,Total,4", "5x2:fruit,price,apple,0.5,pear,0.75,plum,1.2,Total,2.45"]);
eq("GROUPBY over two fields, a sort by the value, and a filter", grid("GROUPBY(HSTACK(B1:B3, A1:A3), C1:C3, SUM, 0, 0)", "GROUPBY(B1:B3, C1:C3, SUM, 0, 0, -2)", "GROUPBY(B1:B3, C1:C3, SUM, 0, 1, 1, A1:A3 > 1)"),
  ["3x3:apples,3,10,apples,4,30,pears,1,20", "2x2:apples,40,pears,20", "2x2:apples,40,Total,40"]);
eq("PIVOTBY: rows by one field, columns by another, totals both ways", grid("PIVOTBY(B1:B3, A1:A3, C1:C3, SUM)", "PIVOTBY(B1:B3, A1:A3, C1:C3, SUM, 0, 0, 1, 0, -1)"),
  ["4x5:,1,3,4,Total,apples,,10,30,40,pears,20,,,20,Total,20,10,30,60", "3x4:,4,3,1,apples,30,10,,pears,,,20"]);
eq("GROUPBY's and PIVOTBY's errors", why("GROUPBY(B1:B3, A1:A2, SUM)", "GROUPBY(B1:B3, A1:A3, 1)", "GROUPBY(B1:B3, A1:A3, SUM, 4)", "PIVOTBY(B1:B3, A1:A3, C1:C3, SUM, 0, 1, 1, 1, 1, A1:A2)", "GROUPBY(B1:B3, A1:A3)"), [
  "#VALUE! GROUPBY's fields and values differ in height", "#VALUE! GROUPBY takes a LAMBDA last", "#VALUE! GROUPBY's field_headers is 0 to 3, not 4", "#VALUE! PIVOTBY's filter_array differs in height", "#VALUE! GROUPBY takes 3 to 7 values",
]);
eq("a function's bare name is a LAMBDA where one is taken", ev("SUM(MAP(A1:A3, SUM))", "SUM(BYROW(F2:F4, MAX))", "REDUCE(0, A1:A3, LAMBDA(a, v, a + v))"), ["8", "6", "8"]);
eq("and nowhere else: a bare name, or one in another function's arguments", [...why("SUM(A1:A3) + 0"), (() => { try { parse("SUM(MAX)"); return "parsed"; } catch (e) { return String((e as Error).message); } })()], ["not an error: 8", "unknown name 'MAX'"]);

// The dates batch (2.5): against values worked out by hand and by
// Python's calendar (ISO weeks by isocalendar, workdays by a loop).
eq("WEEKNUM: Sunday weeks (1), Monday weeks (2, 11), other starts, ISO (21)", ev(`WEEKNUM("2012-03-09")`, `WEEKNUM("2012-03-09", 2)`, `WEEKNUM("2021-01-01")`, `WEEKNUM("2021-01-03")`, `WEEKNUM("2021-01-02", 2)`, `WEEKNUM("2024-12-31")`, `WEEKNUM("2012-03-09", 16)`, `WEEKNUM("2021-01-01", 21)`, `WEEKNUM(DATE(2012, 3, 9), 11)`),
  ["10", "11", "1", "2", "1", "53", "10", "53", "11"]);
eq("ISOWEEKNUM: week 1 holds the first Thursday", ev(`ISOWEEKNUM("2012-03-09")`, `ISOWEEKNUM("2021-01-01")`, `ISOWEEKNUM("2021-01-04")`, `ISOWEEKNUM("2024-12-30")`, `ISOWEEKNUM("2010-01-03")`), ["10", "53", "1", "1", "53"]);
eq("WEEKNUM's type", why(`WEEKNUM("2012-03-09", 3)`), ["#NUM! WEEKNUM's type is 1, 2, 11 to 17 or 21, not 3"]);
eq("YEARFRAC on each basis", ev(`ROUND(YEARFRAC("2012-01-01", "2012-07-30"), 8)`, `ROUND(YEARFRAC("2012-01-01", "2012-07-30", 1), 8)`, `ROUND(YEARFRAC("2012-01-01", "2012-07-30", 2), 8)`, `ROUND(YEARFRAC("2012-01-01", "2012-07-30", 3), 8)`, `ROUND(YEARFRAC("2012-01-01", "2012-07-30", 4), 8)`, `ROUND(YEARFRAC("2011-06-30", "2013-03-15", 1), 8)`, `YEARFRAC("2012-07-30", "2012-01-01", 2)`),
  ["0.58055556", "0.57650273", "0.58611111", "0.57808219", "0.58055556", "1.7080292", "0.5861111111"]);
eq("YEARFRAC's basis", why(`YEARFRAC("2012-01-01", "2012-07-30", 5)`), ["#NUM! YEARFRAC's basis is 0 to 4, not 5"]);
eq("DAYS360, US and European", ev(`DAYS360("2011-01-30", "2011-02-01")`, `DAYS360("2011-01-01", "2011-12-31")`, `DAYS360("2011-01-01", "2011-12-31", TRUE)`, `DAYS360("2011-01-31", "2011-03-31")`, `DAYS360("2011-01-31", "2011-03-31", TRUE)`, `DAYS360("2012-02-29", "2012-03-31")`, `DAYS360("2011-02-28", "2012-02-29")`, `DAYS360("2011-03-31", "2011-01-31")`),
  ["1", "360", "359", "60", "60", "30", "360", "-60"]);
eq("TIMEVALUE: the fraction of a day, a date in front ignored", ev(`TIMEVALUE("2:24 AM")`, `TIMEVALUE("22:48:00")`, `TIMEVALUE("6:45 PM")`, `TIMEVALUE("12:00 am")`, `TIMEVALUE("12:00 pm")`, `TIMEVALUE("2011-07-18 12:00")`, `TIMEVALUE("1/15/2024 6:00:30 pm")`, `HOUR(TIMEVALUE("6:45 PM"))`),
  ["0.1", "0.95", "0.78125", "0", "0.5", "0.5", "0.7503472222", "18"]);
eq("TIMEVALUE of what is no time", why(`TIMEVALUE("noon")`, `TIMEVALUE("25:00")`, `TIMEVALUE("13:00 pm")`), ["#VALUE! 'noon' is not a time", "#VALUE! '25:00' is not a time", "#VALUE! '13:00 pm' is not a time"]);
eq("NETWORKDAYS.INTL: the weekend as a number or seven 1s and 0s, with holidays", ev(`NETWORKDAYS.INTL("2006-01-01", "2006-01-31")`, `NETWORKDAYS.INTL("2006-01-01", "2006-01-31", 7)`, `NETWORKDAYS.INTL("2006-01-01", "2006-01-31", 11)`, `NETWORKDAYS.INTL("2006-02-28", "2006-01-31", 1)`, `NETWORKDAYS.INTL("2006-01-01", "2006-01-31", "0010001", {"2006-01-02", "2006-01-16"})`, `NETWORKDAYS.INTL("2006-01-01", "2006-01-31", "0000011")`),
  ["22", "23", "26", "-21", "20", "22"]);
eq("WORKDAY.INTL: that many workdays on, the weekend as told", ev(`WORKDAY.INTL("2012-01-01", 30)`, `WORKDAY.INTL("2012-01-01", 90, 11)`, `WORKDAY.INTL("2012-01-01", 30, 17)`, `WORKDAY.INTL("2012-01-01", -5, 1, "2011-12-26")`, `WORKDAY.INTL("2012-01-01", 5, "1111110")`),
  ["40949", "41013", "40944", "40900", "40944"]);
eq("a weekend that is none", why(`NETWORKDAYS.INTL("2006-01-01", "2006-01-31", 8)`, `WORKDAY.INTL("2012-01-01", 1, "1111111")`, `WORKDAY.INTL("2012-01-01", 1, "01")`), [
  "#NUM! 8 is not a weekend: 1 to 7, 11 to 17, or seven 1s and 0s", "#NUM! 1111111 is not a weekend: 1 to 7, 11 to 17, or seven 1s and 0s", "#NUM! 01 is not a weekend: 1 to 7, 11 to 17, or seven 1s and 0s",
]);

// The text batch (2.4).
eq("CODE, CHAR, UNICODE, UNICHAR", ev(`CODE("A")`, `CODE("apple")`, "CHAR(65)", "CHAR(33)", `UNICODE("é")`, `UNICODE("😀")`, "UNICHAR(233)", "UNICHAR(128512)", `LEN(UNICHAR(128512))`), ["65", "97", "A", "!", "233", "128512", "é", "😀", "2"]);
eq("their errors", why(`CODE("")`, "CHAR(0)", "CHAR(256)", `UNICODE("")`, "UNICHAR(0)", "UNICHAR(55296)"), ["#VALUE! CODE of no text", "#VALUE! CHAR(0): 1 to 255", "#VALUE! CHAR(256): 1 to 255", "#VALUE! UNICODE of no text", "#VALUE! UNICHAR(0): not a character", "#VALUE! UNICHAR(55296): not a character"]);
eq("CLEAN takes the control characters out", ev(`CLEAN("a" & CHAR(9) & "b" & CHAR(10) & "c")`, `LEN(CLEAN(CHAR(1) & CHAR(31) & "x"))`), ["abc", "1"]);
eq("FIXED: places, commas or not, rounding left of the point", ev("FIXED(1234.567, 1)", "FIXED(1234.567, -1)", "FIXED(-1234.567, -1, TRUE)", "FIXED(44.332)", "FIXED(0.5, 0)", "FIXED(2.675, 2)", "FIXED(-0.001, 2)", "FIXED(1234567.891, 2, FALSE)"),
  ["1,234.6", "1,230", "-1230", "44.33", "1", "2.68", "0.00", "1,234,567.89"]);
eq("DOLLAR: a $ in front, a negative in parentheses", ev("DOLLAR(1234.567, 2)", "DOLLAR(-1234.567, -2)", "DOLLAR(-0.123, 4)", "DOLLAR(99)", "DOLLAR(-0.001)"), ["$1,234.57", "($1,200)", "($0.1230)", "$99.00", "$0.00"]);
eq("SEARCH: without regard to case, with wildcards, from a start", ev(`SEARCH("e", "Statements")`, `SEARCH("margin", "Profit Margin")`, `SEARCH("P?L", "Apple")`, `SEARCH("a*e", "banana cake")`, `SEARCH("~?", "what? why?")`, `SEARCH("a", "banana", 3)`, `SEARCH("", "abc")`),
  ["5", "8", "2", "2", "5", "4", "1"]);
eq("SEARCH's errors", why(`SEARCH("z", "apple")`, `SEARCH("a", "apple", 0)`, `SEARCH("a", "apple", 7)`), ["#VALUE! 'z' is not in 'apple'", "#VALUE! SEARCH from 0: outside the text", "#VALUE! SEARCH from 7: outside the text"]);
eq("REPLACE and CONCATENATE", ev(`REPLACE("abcdefghijk", 6, 5, "*")`, `REPLACE("2009", 3, 2, "10")`, `REPLACE("123456", 1, 3, "@")`, `REPLACE("abc", 5, 1, "X")`, `CONCATENATE("a", 1, TRUE, B5, "b")`, `CONCATENATE(B1, " & ", B2)`),
  ["abcde*k", "2010", "@456", "abcX", "a1TRUEb", "apples & pears"]);
eq("CONCATENATE over a range goes cell by cell; CONCAT joins it whole", grid(`CONCATENATE(A1:A3, "-")`, "CONCAT(A1:A3)"), ["3x1:3-,1-,4-", "314"]);
eq("REPLACE's errors", why(`REPLACE("abc", 0, 1, "x")`, `REPLACE("abc", 1, -1, "x")`), ["#VALUE! REPLACE from 0, 1 characters", "#VALUE! REPLACE from 1, -1 characters"]);
eq("NUMBERVALUE: the separators given, spaces ignored, a % divides", ev(`NUMBERVALUE("2.500,27", ",", ".")`, `NUMBERVALUE("3.5%")`, `NUMBERVALUE("1 234.5")`, `NUMBERVALUE("")`, `NUMBERVALUE("-1,234")`, `NUMBERVALUE("12%%")`, `NUMBERVALUE("1e3")`),
  ["2500.27", "0.035", "1234.5", "0", "-1234", "0.0012", "1000"]);
eq("NUMBERVALUE's errors", why(`NUMBERVALUE("abc")`, `NUMBERVALUE("1,2.3,4")`, `NUMBERVALUE("1.5", ".", ".")`), ["#VALUE! 'abc' is not a number", "#VALUE! '1,2.3,4' is not a number", "#VALUE! NUMBERVALUE's separators are the same"]);
eq("VALUETOTEXT and ARRAYTOTEXT, concise and strict", ev("VALUETOTEXT(A1)", "VALUETOTEXT(B1)", "VALUETOTEXT(B1, 1)", "VALUETOTEXT(D1, 1)", "VALUETOTEXT(D2)", "ARRAYTOTEXT(F1:H2)", "ARRAYTOTEXT(F1:H2, 1)", `ARRAYTOTEXT({1,"a";TRUE,#N/A}, 1)`, "ARRAYTOTEXT(A1:A3)"),
  ["3", "apples", "\"apples\"", "TRUE", "#DIV/0!", "id, fruit, price; 1, apple, 0.5", "{\"id\",\"fruit\",\"price\";1,\"apple\",0.5}", "{1,\"a\";TRUE,#N/A}", "3; 1; 4"]);
eq("ARRAYTOTEXT's format", why("ARRAYTOTEXT(A1:A2, 2)"), ["#VALUE! ARRAYTOTEXT's format is 0 or 1, not 2"]);
eq("REGEXTEST, case sensitive unless told", ev(`REGEXTEST("Apple 42", "\\d+")`, `REGEXTEST("Apple", "^a")`, `REGEXTEST("Apple", "^a", 1)`, `REGEXTEST(D4, "^hello\\s\\w+$")`, `REGEXTEST("café", "^\\p{L}+$")`), ["TRUE", "FALSE", "TRUE", "TRUE", "TRUE"]);
eq("REGEXEXTRACT: the first match, every match, or the groups", grid(`REGEXEXTRACT("a1b22c333", "\\d+")`, `REGEXEXTRACT("a1b22c333", "\\d+", 1)`, `REGEXEXTRACT("John Smith", "(\\w+) (\\w+)", 2)`, `REGEXEXTRACT("x=12", "(\\w)=(\\d+)", 2)`, `REGEXEXTRACT("one", "o(n)e", 2)`, `REGEXEXTRACT("ABC", "b", 0, 1)`),
  ["1", "3x1:1,22,333", "1x2:John,Smith", "1x2:x,12", "n", "B"]);
eq("REGEXREPLACE: all, or the nth, with $1 for a group", ev(`REGEXREPLACE("a1b22c333", "\\d+", "#")`, `REGEXREPLACE("a1b22c333", "\\d+", "#", 2)`, `REGEXREPLACE("John Smith", "(\\w+) (\\w+)", "$2, $1")`, `REGEXREPLACE("x-y-z", "-", "+", 3)`, `REGEXREPLACE("aaa", "a", "b", 2)`, `REGEXREPLACE("Hello", "L", "_", 0, 1)`),
  ["a#b#c#", "a1b#c333", "Smith, John", "x-y-z", "aba", "He__o"]);
eq("REGEX errors: a pattern that does not compile, no match, a bad mode", why(`REGEXTEST("a", "(")`, `REGEXEXTRACT("abc", "\\d")`, `REGEXEXTRACT("abc", "b", 3)`, `REGEXREPLACE("abc", "b", "x", -1)`),
  ["#VALUE! '(' is not a regular expression", "#N/A '\\d' matches nothing in 'abc'", "#VALUE! REGEXEXTRACT's mode is 0, 1 or 2, not 3", "#VALUE! REGEXREPLACE's occurrence is 0 or more, not -1"]);

// The maths batch (2.2).
eq("CEILING and FLOOR to a multiple, away from zero and toward it", ev("CEILING(2.5, 1)", "CEILING(-2.5, -2)", "CEILING(1.5, 0.1)", "CEILING(0.234, 0.01)", "FLOOR(3.7, 2)", "FLOOR(-2.5, -2)", "FLOOR(1.58, 0.1)", "FLOOR(0.234, 0.01)", "CEILING(5, 0)"),
  ["3", "-4", "1.5", "0.24", "2", "-2", "1.5", "0.23", "0"]);
eq("their errors", why("CEILING(-2.5, 2)", "FLOOR(2.5, -2)", "FLOOR(2, 0)"), ["#NUM! CEILING's number and significance differ in sign", "#NUM! FLOOR's number and significance differ in sign", "#DIV/0! FLOOR by a significance of 0"]);
eq("CEILING.MATH: up, or away from zero with a mode; FLOOR.MATH: down, or toward zero", ev("CEILING.MATH(24.3, 5)", "CEILING.MATH(6.7)", "CEILING.MATH(-8.1, 2)", "CEILING.MATH(-5.5, 2, -1)", "CEILING.MATH(-5.5, -2)", "FLOOR.MATH(24.3, 5)", "FLOOR.MATH(6.7)", "FLOOR.MATH(-8.1, 2)", "FLOOR.MATH(-5.5, 2, -1)"),
  ["25", "7", "-8", "-6", "-4", "20", "6", "-10", "-4"]);
eq("the .PRECISE pair: up and down, the significance's sign ignored", ev("CEILING.PRECISE(4.3)", "CEILING.PRECISE(-4.3)", "CEILING.PRECISE(4.3, 2)", "CEILING.PRECISE(-4.3, -2)", "FLOOR.PRECISE(-3.2, -1)", "FLOOR.PRECISE(3.2, 1)", "FLOOR.PRECISE(-3.2, 1)"),
  ["5", "-4", "6", "-4", "-4", "3", "-4"]);
eq("TRUNC, MROUND, EVEN and ODD", ev("TRUNC(8.9)", "TRUNC(-8.9)", "TRUNC(0.45, 1)", "TRUNC(2.675, 2)", "MROUND(10, 3)", "MROUND(-10, -3)", "MROUND(1.3, 0.2)", "MROUND(7, 0)", "EVEN(1.5)", "EVEN(3)", "EVEN(-1)", "EVEN(2)", "ODD(1.5)", "ODD(3)", "ODD(-1.5)", "ODD(0)"),
  ["8", "-8", "0.4", "2.67", "9", "-9", "1.4", "0", "2", "4", "-2", "2", "3", "3", "-3", "1"]);
eq("MROUND of different signs", why("MROUND(5, -2)"), ["#NUM! MROUND's number and multiple differ in sign"]);
eq("EXP, LN, LOG, LOG10", ev("EXP(1)", "EXP(0)", "LN(EXP(2))", "LN(86)", "LOG(8, 2)", "LOG(100)", "LOG(1000, 10)", "LOG10(86)", "LOG(2.7182818)"),
  ["2.718281828", "1", "2", "4.454347296", "3", "2", "3", "1.934498451", "0.4342944774"]);
eq("their errors", why("LN(0)", "LN(-1)", "LOG(8, 1)", "LOG(-8)", "LOG10(0)", "EXP(1000)"),
  ["#NUM! LN of 0: not positive", "#NUM! LN of -1: not positive", "#NUM! LOG to base 1", "#NUM! LOG of -8: not positive", "#NUM! LOG10 of 0: not positive", "#NUM! EXP overflows"]);
eq("QUOTIENT, GCD, LCM", ev("QUOTIENT(5, 2)", "QUOTIENT(4.5, 3.1)", "QUOTIENT(-10, 3)", "GCD(5, 2)", "GCD(24, 36)", "GCD(7, 1)", "GCD(5, 0)", "GCD(A1:A6)", "LCM(5, 2)", "LCM(24, 36)", "LCM(6, 4, 10)", "LCM(0, 5)"),
  ["2", "1", "-3", "1", "12", "1", "5", "1", "10", "72", "60", "0"]);
eq("their errors", why("QUOTIENT(1, 0)", "GCD(-2, 4)", "LCM(3, -1)"), ["#DIV/0! QUOTIENT by zero", "#NUM! GCD of a negative number, -2", "#NUM! LCM of a negative number, -1"]);
eq("FACT, COMBIN, PERMUT, SUMSQ", ev("FACT(5)", "FACT(1.9)", "FACT(0)", "COMBIN(8, 2)", "COMBIN(5, 5)", "COMBIN(49, 6)", "PERMUT(3, 2)", "PERMUT(100, 3)", "SUMSQ(3, 4)", "SUMSQ(A1:A3)"),
  ["120", "1", "1", "28", "1", "13983816", "6", "970200", "25", "26"]);
eq("their errors", why("FACT(-1)", "FACT(200)", "COMBIN(2, 3)", "PERMUT(-1, 0)"), ["#NUM! FACT of -1: negative", "#NUM! FACT of 200 overflows", "#NUM! COMBIN of 3 from 2", "#NUM! PERMUT of 0 from -1"]);
{
  // Random numbers: in range, whole when asked, a grid when asked, and
  // not the same twice (a draw of a million).
  const draws = Array.from({ length: 20 }, () => evaluate(parse("RAND()"), env) as number);
  const between = Array.from({ length: 20 }, () => evaluate(parse("RANDBETWEEN(-3, 3)"), env) as number);
  const grid = evaluate(parse("RANDARRAY(2, 3, 10, 20, TRUE)"), env);
  const pair = evaluate(parse("RANDARRAY(1, 2)"), env);
  eq("RAND, RANDBETWEEN and RANDARRAY draw in range", [
    draws.every((d) => d >= 0 && d < 1), between.every((n) => Number.isInteger(n) && n >= -3 && n <= 3),
    isArray(grid) && grid.rows === 2 && grid.cols === 3 && grid.values.every((v) => Number.isInteger(v) && (v as number) >= 10 && (v as number) <= 20),
    isArray(pair) && pair.values.every((v) => (v as number) >= 0 && (v as number) < 1),
    show(evaluate(parse("RANDBETWEEN(1, 1000000)"), env)) !== show(evaluate(parse("RANDBETWEEN(1, 1000000)"), env)),
  ], [true, true, true, true, true]);
  eq("their errors", why("RANDBETWEEN(3, 1)", "RANDARRAY(0)", "RANDARRAY(1, 1, 5, 2)"), ["#NUM! RANDBETWEEN 3 and 1: nothing between", "#VALUE! RANDARRAY of no rows or columns", "#VALUE! RANDARRAY from 5 to 2: nothing between"]);
}
eq("trigonometry in radians, with DEGREES and RADIANS", ev("SIN(PI()/2)", "COS(0)", "ROUND(TAN(PI()/4), 10)", "ASIN(1)", "ACOS(1)", "DEGREES(ATAN(1))", "DEGREES(ATAN2(1, 1))", "DEGREES(ATAN2(-1, -1))", "RADIANS(180)", "DEGREES(PI())"),
  ["1", "1", "1", "1.570796327", "0", "45", "45", "-135", "3.141592654", "180"]);
eq("hyperbolic", ev("SINH(0)", "COSH(0)", "ROUND(TANH(1), 6)", "ASINH(0)", "ACOSH(1)", "ATANH(0)", "ROUND(ASINH(SINH(2)), 10)"), ["0", "1", "0.761594", "0", "0", "0", "2"]);
eq("outside a domain", why("ASIN(2)", "ACOSH(0.5)", "ATANH(1)", "ATAN2(0, 0)"), ["#NUM! ASIN of 2: outside its domain", "#NUM! ACOSH of 0.5: outside its domain", "#NUM! ATANH of 1: outside its domain", "#DIV/0! ATAN2 of the origin"]);

// Lookup.
eq("ROWS and COLUMNS", ev("ROWS(F1:H4)", "COLUMNS(F1:H4)", "ROWS(A1)"), ["4", "3", "1"]);
eq("ROW and COLUMN of a reference, counted from 1", ev("ROW(C5)", "COLUMN(C5)", "ROW($A$1)", "COLUMN(Z1)"), ["5", "3", "1", "26"]);
eq("over a range: each row down, each column across", grid("ROW(A3:A5)", "COLUMN(A1:C1)", "ROW(A1:C1)", "COLUMN(F2:H4)"), ["3x1:3,4,5", "1x3:1,2,3", "1", "1x3:6,7,8"]);
eq("with nothing, the cell the formula is in", [show(evaluate(parse("ROW()"), env, "D7")), show(evaluate(parse("COLUMN()"), env, "D7")), show(evaluate(parse("ROW()+COLUMN()"), env, "B2"))], ["7", "4", "4"]);
eq("ROW's and COLUMN's errors", why("ROW()", "ROW(1)", "COLUMN(A1, B1)", "ROW(ZZZ1)"), [
  "#REF! ROW() is in no cell", "#VALUE! ROW takes a reference", "#VALUE! COLUMN takes a reference, or nothing", "#REF! ZZZ1 is not on the sheet",
]);
// The sheet is 26 columns by 100 rows (address.ts): past that is not a
// reference here, as everywhere in the app, until the sheet grows.
eq("past the sheet's edge, as any reference is", why("COLUMN(AA1)", "ROW(A101)"), ["#REF! AA1 is not on the sheet", "#REF! A101 is not on the sheet"]);
eq("ADDRESS: absolute by default, then the three other mixes", ev("ADDRESS(2, 3)", "ADDRESS(2, 3, 2)", "ADDRESS(2, 3, 3)", "ADDRESS(2, 3, 4)", "ADDRESS(1, 27)"), ["$C$2", "C$2", "$C2", "C2", "$AA$1"]);
eq("ADDRESS in R1C1, and with a sheet", ev("ADDRESS(2, 3, 1, FALSE)", "ADDRESS(2, 3, 2, FALSE)", "ADDRESS(2, 3, 4, FALSE)", `ADDRESS(1, 1, 1, TRUE, "Costs")`, `ADDRESS(1, 1, 4, TRUE, "My Costs")`),
  ["R2C3", "R2C[3]", "R[2]C[3]", "Costs!$A$1", "'My Costs'!A1"]);
eq("ADDRESS's errors", why("ADDRESS(0, 1)", "ADDRESS(1, 1, 5)"), ["#VALUE! ADDRESS of row 0, column 1: not on a sheet", "#VALUE! ADDRESS's abs is 1 to 4, not 5"]);
eq("LOOKUP along an ascending vector: the last not past the value, from the result vector or itself", ev("LOOKUP(25, E1:E5, A1:A5)", "LOOKUP(50, E1:E5)", "LOOKUP(999, E1:E5, A1:A5)"), ["1", "50", "5"]);
eq("LOOKUP in an array: down the first column when tall, along the first row when wide", ev("LOOKUP(2, F2:H4)", "LOOKUP(2.5, F2:H4)", `LOOKUP(25, {10,20,30;"a","b","c"})`), ["0.75", "0.75", "b"]);
eq("LOOKUP's errors", why("LOOKUP(5, E1:E5)", "LOOKUP(25, E1:E5, A1:A1)"), ["#N/A 5 is not there", "#N/A LOOKUP's result vector is shorter than its lookup vector"]);
// References made at run time.
eq("OFFSET: the cell that many rows down and columns across", ev("OFFSET(A1, 2, 0)", "OFFSET(A1, 0, 2)", "OFFSET(C3, -2, -2)", "ROW(OFFSET(A1, 4, 0))"), ["4", "10", "3", "5"]);
eq("over a range, as high and wide as it is, or as told", grid("OFFSET(A1:A2, 1, 0)", "OFFSET(A1, 0, 0, 3, 1)", "OFFSET(F1, 1, 1, 2, 2)", "OFFSET(A1:B2, 1, 0, , 1)", "SUM(OFFSET(A1, 0, 0, 3, 1))"),
  ["2x1:1,4", "3x1:3,1,4", "2x2:apple,0.5,pear,0.75", "2x1:1,4", "8"]);
eq("OFFSET's errors", why("OFFSET(A1, -1, 0)", "OFFSET(Z1, 0, 1)", "OFFSET(A1, 0, 0, 0, 1)", "OFFSET(1, 0, 0)", "OFFSET(A1, 1)"), [
  "#REF! OFFSET to row 0, column 1, 1 by 1, is off the sheet", "#REF! OFFSET to row 1, column 27, 1 by 1, is off the sheet",
  "#REF! OFFSET of 0 rows by 1 columns is no reference", "#VALUE! OFFSET takes a reference first", "#VALUE! OFFSET takes 3 to 5 values",
]);
eq("INDIRECT: the cell or range the text names", ev(`INDIRECT("A3")`, `INDIRECT("$c$2")`, `INDIRECT("A" & 1 + 4)`, `SUM(INDIRECT("C1:C3"))`, `COLUMN(INDIRECT("F1"))`), ["4", "20", "5", "60", "6"]);
eq("in R1C1 style, relative to the formula's cell", [
  ...ev(`INDIRECT("R3C1", FALSE)`),
  show(evaluate(parse(`INDIRECT("R[1]C[-1]", FALSE)`), env, "B2")), show(evaluate(parse(`INDIRECT("RC[1]", FALSE)`), env, "B2")), show(evaluate(parse(`INDIRECT("r2c", FALSE)`), env, "C5")),
], ["4", "4", "20", "20"]);
eq("on another sheet, quoted or not", ev(`INDIRECT("Costs!B2")`, `INDIRECT("'Costs'!B2")`), ["7", "7"]);
eq("INDIRECT's errors", why(`INDIRECT("hello")`, `INDIRECT("A101")`, `INDIRECT("R[1]C", FALSE)`, `INDIRECT("A1:B2:C3")`, `INDIRECT(D2)`), [
  "#REF! INDIRECT of 'hello': not a reference", "#REF! A101 is not on the sheet", "#REF! R[1]C is relative, and the formula is in no cell",
  "#REF! INDIRECT of 'A1:B2:C3': not a reference", "#DIV/0! divide by zero",
]);
{
  // The Env hears each cell a run-time reference came to (the sheet links
  // them as the formula's reads); a reference written in the formula is
  // its static read, not told.
  const heard: string[] = [];
  const hearing = { get: env.get, read: (k: string) => heard.push(k) };
  const values = [`SUM(OFFSET(A1, 1, 0, 2, 1))`, `INDIRECT("C1") + A1`, `ROW(OFFSET(A1, 3, 0))`].map((src) => show(evaluate(parse(src), hearing)));
  eq("the Env hears the cells a run-time reference reads", [values, heard], [["5", "13", "4"], ["A2", "A3", "C1"]]);
}
{
  // Asking about a reference: its cell's input and the workbook come from
  // the Env (a sheet's: sheet.ts); without them a formula is in a book of
  // one sheet, and no cell holds a formula.
  const inputs = new Map([["A1", "3"], ["C1", "=A1+7"], ["COSTS!B2", "=SUM(A1:A3)"]]);
  const told = { get: env.get, input: (k: string) => inputs.get(k) ?? "", book: () => ({ sheets: ["Data", "Prices", "Costs"], here: 1 }) };
  const at = (src: string, cell: string | null = null) => {
    const v = evaluate(parse(src), told, cell);
    return v instanceof CellError ? `${v.code} ${v.why}` : show(v);
  };
  eq("FORMULATEXT and ISFORMULA: the cell's input, when it is a formula", [at("FORMULATEXT(C1)"), at("FORMULATEXT(C1:D2)"), at("FORMULATEXT(Costs!B2)"), at("ISFORMULA(C1)"), at("ISFORMULA(A1)"), at("ISFORMULA(Z9)")],
    ["=A1+7", "=A1+7", "=SUM(A1:A3)", "TRUE", "FALSE", "FALSE"]);
  eq("their errors", [at("FORMULATEXT(A1)"), at("FORMULATEXT(B9)"), at("FORMULATEXT(1)"), at("ISFORMULA(\"C1\")"), at("FORMULATEXT(A101)")],
    ["#N/A A1 holds no formula", "#N/A B9 holds no formula", "#VALUE! FORMULATEXT takes a reference", "#VALUE! ISFORMULA takes a reference", "#REF! A101 is not on the sheet"]);
  eq("ISREF: a reference, however made; not a value, nor what is off the sheet", [at("ISREF(A1)"), at("ISREF(A1:B2)"), at("ISREF(OFFSET(A1, 1, 0))"), at("ISREF(INDIRECT(\"B\" & 2))"), at("ISREF(1)"), at("ISREF(\"A1\")"), at("ISREF(SUM(A1))"), at("ISREF(A101)")],
    ["TRUE", "TRUE", "TRUE", "TRUE", "FALSE", "FALSE", "FALSE", "FALSE"]);
  eq("SHEET: this one, a reference's, or by name; SHEETS: how many", [at("SHEET()"), at("SHEET(A1)"), at("SHEET(Costs!B2)"), at("SHEET(\"costs\")"), at("SHEET(\"Data\")"), at("SHEETS()"), at("SHEETS(Costs!A1:B2)")],
    ["2", "2", "3", "3", "1", "3", "1"]);
  eq("SHEET's and SHEETS's errors", [at("SHEET(\"Nope\")"), at("SHEET(7)"), at("SHEETS(7)"), at("SHEET(A1, B1)")],
    ["#N/A no sheet Nope", "#N/A 7 is not a sheet", "#VALUE! SHEETS takes a reference", "#VALUE! SHEET takes a reference, a name, or nothing"]);
  eq("a formula on its own is in a book of one sheet, where no cell holds a formula", ev("SHEET()", "SHEETS()", "ISFORMULA(A1)", "SHEET(A1)"), ["1", "1", "FALSE", "1"]);
}
eq("INDEX by row and column, or along a list", ev("INDEX(F1:H4, 3, 2)", "INDEX(A1:A6, 6)", "INDEX(F1:H1, 3)"), ["pear", "9", "price"]);
eq("INDEX outside", why("INDEX(F1:H4, 5, 1)"), ["#REF! INDEX 5, 1 is outside 4 x 3"]);
eq("MATCH exact, ignoring case", ev(`MATCH("PEAR", G1:G4, 0)`, "MATCH(3, F1:F4, 0)"), ["3", "4"]);
eq("MATCH approximate: the largest at most", ev("MATCH(35, E1:E5)", "MATCH(50, E1:E5, 1)", "MATCH(10, E1:E5)"), ["3", "5", "1"]);
eq("MATCH descending: the smallest at least", ev("MATCH(25, I1:I4, -1)", "MATCH(40, I1:I4, -1)", "MATCH(10, I1:I4, -1)"), ["2", "1", "4"]);
eq("MATCH not finding", why(`MATCH("kiwi", G1:G4, 0)`, "MATCH(5, E1:E5)"), ["#N/A kiwi is not there", "#N/A 5 is not there"]);
eq("VLOOKUP exact", ev("VLOOKUP(2, F2:H4, 2, FALSE)", "VLOOKUP(3, F2:H4, 3, 0)"), ["pear", "1.2"]);
eq("VLOOKUP approximate", ev("VLOOKUP(2.5, F2:H4, 2)", "VLOOKUP(9, F2:H4, 3, TRUE)"), ["pear", "1.2"]);
eq("VLOOKUP's errors", why("VLOOKUP(7, F2:H4, 2, FALSE)", "VLOOKUP(2, F2:H4, 4, FALSE)", "VLOOKUP(2, F2:H4, 0)", "VLOOKUP(0.5, F2:H4, 2)"), [
  "#N/A 7 is not there", "#REF! VLOOKUP column 4 of a table 3 wide", "#VALUE! VLOOKUP's column starts at 1", "#N/A 0.5 is not there",
]);
eq("a lookup in a formula", ev(`"A " & VLOOKUP(1, F2:H4, 2, FALSE) & " costs " & INDEX(H2:H4, MATCH("apple", G2:G4, 0))`), ["A apple costs 0.5"]);
// expect: ok the library
// expect: ok PRODUCT
// expect: ok MEDIAN, odd and even
// expect: ok STDEV is the sample's
// expect: ok COUNTA counts what is there, COUNTBLANK what is not
// expect: ok SIGN, INT
// expect: ok ROUND away from zero, with or without places
// expect: ok ROUNDUP and ROUNDDOWN
// expect: ok MOD takes the divisor's sign
// expect: ok POWER, SQRT, PI
// expect: ok their errors
// expect: ok an error in a range flows out of an aggregate
// expect: ok COUNTIF by text, ignoring case
// expect: ok COUNTIF by comparison
// expect: ok COUNTIF of blanks, and of what is not
// expect: ok a number criterion skips text, and text skips numbers
// expect: ok SUMIF over another range
// expect: ok AVERAGEIF
// expect: ok AVERAGEIF of nothing
// expect: ok criterion on its own
// expect: ok AND, OR, NOT
// expect: ok AND with nothing true or false to look at
// expect: ok IFERROR catches, and passes the rest
// expect: ok IS functions
// expect: ok the rest of the family: ISERR, ISNA, ISLOGICAL, ISNONTEXT, ISEVEN, ISODD
// expect: ok ISEVEN of words
// expect: ok N and T: a number's number, a text's text, else 0 or nothing
// expect: ok N of an error is the error
// expect: ok TYPE: number 1, text 2, logical 4, error 16, array 64
// expect: ok ERROR.TYPE: Excel's numbers, the app's own past them
// expect: ok ERROR.TYPE of a value is #N/A
// expect: ok XOR: an odd number TRUE; TRUE() and FALSE() as calls
// expect: ok XOR of no truths
// expect: ok INFO, what applies: numfile and recalc
// expect: ok CHOOSE
// expect: ok CHOOSE out of range
// expect: ok case
// expect: ok TRIM collapses spaces
// expect: ok LEFT, RIGHT, MID
// expect: ok FIND is case-sensitive, from a start
// expect: ok FIND not finding
// expect: ok SUBSTITUTE, REPT, EXACT
// expect: ok TEXTJOIN, skipping blanks or not
// expect: ok VALUE
// expect: ok VALUE of words
// expect: ok SUBTOTAL by function number, 1 to 11 and 101 to 111
// expect: ok SUBTOTAL's errors
// expect: ok AGGREGATE: the function, the options (2 skips errors), the k forms
// expect: ok AGGREGATE's errors: an error kept, a bad number or options
// expect: ok MAXA, MINA, AVERAGEA: TRUE 1, text and FALSE 0
// expect: ok COUNTA over mixed cells: everything but the empty, the empty text included
// expect: ok the .INC and .EXC quantiles
// expect: ok their ends
// expect: ok GROUPBY: each group's values folded by a LAMBDA, or a function's bare name, sorted, with a total
// expect: ok GROUPBY over two fields, a sort by the value, and a filter
// expect: ok PIVOTBY: rows by one field, columns by another, totals both ways
// expect: ok GROUPBY's and PIVOTBY's errors
// expect: ok a function's bare name is a LAMBDA where one is taken
// expect: ok and nowhere else: a bare name, or one in another function's arguments
// expect: ok WEEKNUM: Sunday weeks (1), Monday weeks (2, 11), other starts, ISO (21)
// expect: ok ISOWEEKNUM: week 1 holds the first Thursday
// expect: ok WEEKNUM's type
// expect: ok YEARFRAC on each basis
// expect: ok YEARFRAC's basis
// expect: ok DAYS360, US and European
// expect: ok TIMEVALUE: the fraction of a day, a date in front ignored
// expect: ok TIMEVALUE of what is no time
// expect: ok NETWORKDAYS.INTL: the weekend as a number or seven 1s and 0s, with holidays
// expect: ok WORKDAY.INTL: that many workdays on, the weekend as told
// expect: ok a weekend that is none
// expect: ok CODE, CHAR, UNICODE, UNICHAR
// expect: ok their errors
// expect: ok CLEAN takes the control characters out
// expect: ok FIXED: places, commas or not, rounding left of the point
// expect: ok DOLLAR: a $ in front, a negative in parentheses
// expect: ok SEARCH: without regard to case, with wildcards, from a start
// expect: ok SEARCH's errors
// expect: ok REPLACE and CONCATENATE
// expect: ok CONCATENATE over a range goes cell by cell; CONCAT joins it whole
// expect: ok REPLACE's errors
// expect: ok NUMBERVALUE: the separators given, spaces ignored, a % divides
// expect: ok NUMBERVALUE's errors
// expect: ok VALUETOTEXT and ARRAYTOTEXT, concise and strict
// expect: ok ARRAYTOTEXT's format
// expect: ok REGEXTEST, case sensitive unless told
// expect: ok REGEXEXTRACT: the first match, every match, or the groups
// expect: ok REGEXREPLACE: all, or the nth, with $1 for a group
// expect: ok REGEX errors: a pattern that does not compile, no match, a bad mode
// expect: ok CEILING and FLOOR to a multiple, away from zero and toward it
// expect: ok their errors
// expect: ok CEILING.MATH: up, or away from zero with a mode; FLOOR.MATH: down, or toward zero
// expect: ok the .PRECISE pair: up and down, the significance's sign ignored
// expect: ok TRUNC, MROUND, EVEN and ODD
// expect: ok MROUND of different signs
// expect: ok EXP, LN, LOG, LOG10
// expect: ok their errors
// expect: ok QUOTIENT, GCD, LCM
// expect: ok their errors
// expect: ok FACT, COMBIN, PERMUT, SUMSQ
// expect: ok their errors
// expect: ok RAND, RANDBETWEEN and RANDARRAY draw in range
// expect: ok their errors
// expect: ok trigonometry in radians, with DEGREES and RADIANS
// expect: ok hyperbolic
// expect: ok outside a domain
// expect: ok ROWS and COLUMNS
// expect: ok ROW and COLUMN of a reference, counted from 1
// expect: ok over a range: each row down, each column across
// expect: ok with nothing, the cell the formula is in
// expect: ok ROW's and COLUMN's errors
// expect: ok past the sheet's edge, as any reference is
// expect: ok ADDRESS: absolute by default, then the three other mixes
// expect: ok ADDRESS in R1C1, and with a sheet
// expect: ok ADDRESS's errors
// expect: ok LOOKUP along an ascending vector: the last not past the value, from the result vector or itself
// expect: ok LOOKUP in an array: down the first column when tall, along the first row when wide
// expect: ok LOOKUP's errors
// expect: ok OFFSET: the cell that many rows down and columns across
// expect: ok over a range, as high and wide as it is, or as told
// expect: ok OFFSET's errors
// expect: ok INDIRECT: the cell or range the text names
// expect: ok in R1C1 style, relative to the formula's cell
// expect: ok on another sheet, quoted or not
// expect: ok INDIRECT's errors
// expect: ok the Env hears the cells a run-time reference reads
// expect: ok FORMULATEXT and ISFORMULA: the cell's input, when it is a formula
// expect: ok their errors
// expect: ok ISREF: a reference, however made; not a value, nor what is off the sheet
// expect: ok SHEET: this one, a reference's, or by name; SHEETS: how many
// expect: ok SHEET's and SHEETS's errors
// expect: ok a formula on its own is in a book of one sheet, where no cell holds a formula
// expect: ok INDEX by row and column, or along a list
// expect: ok INDEX outside
// expect: ok MATCH exact, ignoring case
// expect: ok MATCH approximate: the largest at most
// expect: ok MATCH descending: the smallest at least
// expect: ok MATCH not finding
// expect: ok VLOOKUP exact
// expect: ok VLOOKUP approximate
// expect: ok VLOOKUP's errors
// expect: ok a lookup in a formula
