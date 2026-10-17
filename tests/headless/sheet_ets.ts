// FORECAST.ETS and the Forecast Sheet (6.10), on series whose forecast is
// known: a straight line goes on as the line; a season with a trend
// repeats, its period found; the confidence interval widening; the
// statistics; refusals; the Forecast Sheet's formulas and chart;
// FORECASTSHEET from REXX.
import { fitEts, forecastAt, seasonality } from "./sheet/ets.ts";
import { Commands } from "./sheet/commands.ts";
import { Workbook } from "./sheet/workbook.ts";
import { show } from "./sheet/format.ts";
import { eq } from "./lib/check.ts";

const line = Array.from({ length: 12 }, (_, i) => 3 + 2 * i);
const seasonal = Array.from({ length: 24 }, (_, i) => 100 + 2 * i + [10, -5, 0, -5][i % 4]);
const r = (v: number, k = 6) => Math.round(v * 10 ** k) / 10 ** k;
eq("seasonality: none for a line, 4 for a season of four", [seasonality(line), seasonality(seasonal)], [1, 4]);
const ml = fitEts(line), ms = fitEts(seasonal);
eq("a line goes on as the line", [1, 2, 5].map((h) => r(forecastAt(ml, h), 3)), [27, 29, 35]);
eq("a season repeats on its trend", [1, 2, 3, 4].map((h) => r(forecastAt(ms, h), 3)), [158, 145, 152, 149]);
const wb = new Workbook();
const cmds = new Commands(wb);
await cmds.run("formulas.use language=Excel");
for (let i = 0; i < 24; i++) { await cmds.run(`cell.set cell=A${i + 1} input=${i + 1}`); await cmds.run(`cell.set cell=B${i + 1} input=${seasonal[i]}`); }
const ev = async (f: string) => { await cmds.run(`cell.set cell=D1 input="${f.replace(/"/g, '""')}"`); return show(wb.sheet.value("D1")); };
eq("FORECAST.ETS, its seasonality, its statistics", [await ev("=FORECAST.ETS(25,B1:B24,A1:A24)"), await ev("=FORECAST.ETS.SEASONALITY(B1:B24,A1:A24)"), await ev("=FORECAST.ETS(25,B1:B24,A1:A24,1)") !== "158",
  await ev("=FORECAST.ETS.STAT(B1:B24,A1:A24,8)"), await ev("=ROUND(FORECAST.ETS.STAT(B1:B24,A1:A24,7),6)")], ["158", "4", true, "1", "0"]);
eq("the interval widens with the steps", [Number(await ev("=FORECAST.ETS.CONFINT(30,B1:B24,A1:A24,0.95,1)")) > Number(await ev("=FORECAST.ETS.CONFINT(26,B1:B24,A1:A24,0.95,1)"))], [true]);
eq("refusals", [await ev("=FORECAST.ETS(10,B1:B24,A1:A24)"), await ev("=FORECAST.ETS(25,B1:B3,A1:A3,4)"), await ev("=FORECAST.ETS.STAT(B1:B24,A1:A24,9)")], ["#NUM!", "#NUM!", "#NUM!"]);
await cmds.run("cell.set cell=D1 input=");
const o = await cmds.run("data.forecastSheet timeline=A1:A24 values=B1:B24 end=28 seasonality=4");
eq("the Forecast Sheet: formulas, values, a chart", [o.message, wb.active, wb.sheet.input("C26"), show(wb.sheet.value("C26")), show(wb.sheet.value("A29")), wb.charts().map((c) => [c.kind, c.range])],
  ["forecast to 28 on Forecast: 4 points: 72 cells set, 72 changed", "Forecast", "=FORECAST.ETS(A26,$B$2:$B$25,$A$2:$A$25,4)", "158", "28", [["line", "A1:E29"]]]);
await cmds.run("sheet.activate name=Sheet1");
const said: string[] = [];
await cmds.script('SAY FORECASTSHEET("A1:A24", "B1:B24", 26, "to=Ahead")', (l) => said.push(l));
eq("FORECASTSHEET from REXX", said, ["forecast to 26 on Ahead: 2 points: 64 cells set, 64 changed"]);

// expect: ok seasonality: none for a line, 4 for a season of four
// expect: ok a line goes on as the line
// expect: ok a season repeats on its trend
// expect: ok FORECAST.ETS, its seasonality, its statistics
// expect: ok the interval widens with the steps
// expect: ok refusals
// expect: ok the Forecast Sheet: formulas, values, a chart
// expect: ok FORECASTSHEET from REXX
