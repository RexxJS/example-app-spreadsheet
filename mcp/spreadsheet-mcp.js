#!/usr/bin/env node
/*!
 * spreadsheet-mcp | MIT License (../LICENSE-PROTOCOL-MIT)
 */
/**
 * An MCP server for the sae spreadsheet (../app/): an AI client (Claude, or
 * any MCP client) drives a running spreadsheet through its control bus.
 *
 *   claude mcp add spreadsheet -- node /path/to/mcp/spreadsheet-mcp.js
 *
 * MCP over stdio: one JSON-RPC message a line in, one a line out. Each tool
 * call is one control-bus request. Which tools a client is offered is set
 * when the server starts, by --tools=json|rexx|all (or SPREADSHEET_MCP_TOOLS):
 *
 *   json (the default)  arguments as JSON, the server writing the app's
 *                       commands from them: nothing for a model to escape
 *     spreadsheet_write_cells    rows of values from a cell
 *     spreadsheet_call           one function, or several as one step,
 *                                its arguments checked before it is sent
 *     spreadsheet_read_range     a range's values and formulas, as rows
 *     spreadsheet_list_commands  the functions, each with its parameters
 *   rexx                for those who would write the app's REXX themselves
 *     spreadsheet_run            a command, a string of REXX the app runs
 *     spreadsheet_read_range, spreadsheet_list_commands
 *   all                 all five
 *
 * The bus is sae's: the app started with `--control-bus`, a port chosen at
 * random and a key fresh for every launch, in a courier file only this user
 * can read, and every message sealed (sae's docs/control-bus.md). This
 * server does none of that itself: it runs sae's `sae-bus` for each call,
 * as rexx/spreadsheet-address.js does.
 *
 *   SAE_BUS            the sae-bus to run (default: ../sae/target/sae-bus,
 *                      beside this repository's sae link)
 *   SPREADSHEET_APP    the app's name, which finds its courier file
 *                      (default "Spreadsheet", app/app.json's name)
 *   SAE_APPDATA_DIR    the app's data folder, when the app was started
 *                      with one (sae-bus reads it)
 */
'use strict';

const { execFile } = require('child_process');
const path = require('path');
const readline = require('readline');

const PROTOCOL_VERSIONS = ['2025-06-18', '2025-03-26', '2024-11-05'];
const SERVER = { name: 'spreadsheet', version: '1.1.0' };
// sae-bus waits up to 10 s for the app's answer; a little more here, so its
// own reason is what comes back.
const BUS_DEADLINE_MS = 12000;
// The most cells one spreadsheet_write_cells writes: each is a command, and
// all of them must be done within the bus's 10 s.
const MOST_CELLS = 2000;
const SURFACES = ['json', 'rexx', 'all'];

const LIMITS =
  'Each call is one request to the running app and one step of its undo (the person takes it back with one Undo); ' +
  'one that fails is taken back whole. ' +
  'Request and reply only: you are not told when the person edits the sheet, so read again before relying on what you read earlier. ' +
  'One call at a time, each answered within 10 seconds. The app must be running with its control bus on (sae --app app --control-bus).';

// ---- the tools -------------------------------------------------------------

const WRITE_CELLS = {
  name: 'spreadsheet_write_cells',
  title: 'Write rows of values',
  description:
    'Write a block of values into the active sheet: `rows` is a list of rows, each a list of values, written from `cell` ' +
    'rightwards and down. A value is text, a number, true/false, null (empties the cell), or a formula starting with "=" ' +
    '(e.g. "=SUM(B2:B9)"). Nothing needs escaping. At most ' + MOST_CELLS + ' cells a call. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: {
      cell: { type: 'string', description: 'The top-left cell, e.g. "A1"' },
      rows: {
        type: 'array',
        description: 'e.g. [["Date", "Item", "Amount"], ["2026-10-01", "Taxi", 23.5]]',
        items: { type: 'array', items: { type: ['string', 'number', 'boolean', 'null'] } },
      },
    },
    required: ['cell', 'rows'],
    additionalProperties: false,
  },
};

const CALL = {
  name: 'spreadsheet_call',
  title: 'Call a spreadsheet function',
  description:
    'Call one of the spreadsheet\'s functions by name, its arguments as a JSON list: ' +
    '{"function": "SORTRANGE", "args": ["A1:C10", "B", 0]}. Or several in turn as one step of undo, all or nothing: ' +
    '{"calls": [{"function": "INSERTROW", "args": [2]}, {"function": "SETCELL", "args": ["A2", "=SUM(B3:B9)"]}]}. ' +
    'spreadsheet_list_commands lists the functions and what each takes; the arguments are checked against that before anything is ' +
    'sent. Common ones: GETCELL(cell), SETFORMAT(cell, format) with "currency:USD", "percent:0.0%", "number:0.00" or ' +
    '"bold;color:red"; SORTRANGE(range, by, [ascending]); SUM_RANGE(range); INSERTROW(row); ADDSHEET(name); ' +
    'SETACTIVESHEET(name); GETSHEETNAMES(); UNDO(). A list argument is a JSON list; a JSON one may be an object. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: {
      function: { type: 'string', description: 'A function\'s name, e.g. "SORTRANGE"' },
      args: { type: 'array', description: 'Its arguments, in order', items: {} },
      calls: {
        type: 'array',
        description: 'Instead of function and args: several calls, done in turn as one step',
        items: {
          type: 'object',
          properties: { function: { type: 'string' }, args: { type: 'array', items: {} } },
          required: ['function'],
          additionalProperties: false,
        },
      },
    },
    additionalProperties: false,
  },
};

const RUN = {
  name: 'spreadsheet_run',
  title: 'Run a spreadsheet command (REXX)',
  description:
    'Run one command in the spreadsheet: a REXX expression calling its control functions, as text. ' +
    'Values go inside the command as REXX strings in double quotes, and a quote inside one is written twice ("say ""hi"""). ' +
    'RexxJS reads \\n in a string as a new line, \\t as a tab and \\r as a return, and a backslash before one of those ' +
    'cannot be written in a string at all ("C:\\temp" arrives with a tab): put text with backslashes in through ' +
    'spreadsheet_write_cells or spreadsheet_call (--tools=all), which escape nothing. Common ones:\n' +
    '  SETCELL("A1", "100")          a value, or a formula: SETCELL("A3", "=A1+A2")\n' +
    '  GETCELL("A3")                 what a cell shows\n' +
    '  GETEXPRESSION("A3")           its formula, without the =\n' +
    '  CLEARCELL("A1:B2")            empty a range\n' +
    '  SETFORMAT("D2", "currency:USD")   or "percent:0.0%", "number:0.00", "bold;color:red;align:center"\n' +
    '  SETCOMMENT("A1", "checked")\n' +
    '  INSERTROW(3) / DELETEROW(3) / INSERTCOLUMN("B") / DELETECOLUMN("B")\n' +
    '  ADDSHEET("Costs") / SETACTIVESHEET("Costs") / GETSHEETNAMES() / GETACTIVESHEET()\n' +
    '  SORTRANGE("A1:C10", "B", 0)   by column B, descending (1 or left out: ascending)\n' +
    '  RESULT(RANGE("A1:C9") |> WHERE("amount > 100"))   a query over a range\'s rows\n' +
    '  UNDO() / REDO()\n' +
    '  EXPORT()                      the whole workbook, as its file\'s JSON\n' +
    'Several can go in one command, separated by ";": they are one step of undo, and the answer is the FIRST one\'s ' +
    '(read with a call of its own).\n' +
    'spreadsheet_list_commands lists every function with its parameters. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: { command: { type: 'string', description: 'The REXX to run, e.g. SETCELL("A1", "100")' } },
    required: ['command'],
    additionalProperties: false,
  },
};

const READ_RANGE = {
  name: 'spreadsheet_read_range',
  title: 'Read a range',
  description:
    'Read a rectangle of the active sheet, e.g. "A1:C5": each row\'s values as shown, and the formulas behind ' +
    'any that have one. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: { range: { type: 'string', description: 'A1-style, e.g. "A1:C5" (or one cell, "B2")' } },
    required: ['range'],
    additionalProperties: false,
  },
};

const listCommands = (via) => ({
  name: 'spreadsheet_list_commands',
  title: 'List the spreadsheet functions',
  description:
    'Every function ' + via + ' may call (about 130), each as its signature, what it does and what it gives: ' +
    'SORTRANGE(range, by, [ascending]): optional parameters in brackets. Give a name for one function in full, with what ' +
    'each parameter takes. ' + LIMITS,
  inputSchema: {
    type: 'object',
    properties: { name: { type: 'string', description: 'One function, e.g. "SETCELLVALIDATION"; left out: all of them' } },
    additionalProperties: false,
  },
});

function toolsFor(surface) {
  if (surface === 'rexx') return [RUN, READ_RANGE, listCommands('spreadsheet_run')];
  if (surface === 'all') return [WRITE_CELLS, CALL, READ_RANGE, listCommands('spreadsheet_call (or spreadsheet_run)'), RUN];
  return [WRITE_CELLS, CALL, READ_RANGE, listCommands('spreadsheet_call')];
}

// --tools=X, or SPREADSHEET_MCP_TOOLS; json when neither.
function surfaceOf(argv, env) {
  const flag = argv.find((a) => a.startsWith('--tools='));
  const asked = (flag ? flag.slice('--tools='.length) : env.SPREADSHEET_MCP_TOOLS || 'json').trim().toLowerCase();
  if (!SURFACES.includes(asked)) throw new Error('--tools is json, rexx or all, not ' + JSON.stringify(asked));
  return asked;
}

// ---- the bus ---------------------------------------------------------------

// A string as a REXX literal: double quotes, any inside written twice.
function rexxString(s) {
  return '"' + String(s).replace(/"/g, '""') + '"';
}

// One request over the bus: { ok, result } or { ok: false, error }.
function bus(command) {
  const saeBus = process.env.SAE_BUS || path.join(__dirname, '..', 'sae', 'target', 'sae-bus');
  const app = process.env.SPREADSHEET_APP || 'Spreadsheet';
  return new Promise((resolve) => {
    execFile(saeBus, [app, command], { encoding: 'utf8', timeout: BUS_DEADLINE_MS, maxBuffer: 16 * 1024 * 1024 }, (error, stdout, stderr) => {
      if (error) {
        const why = (stderr || '').trim() || (error.killed ? 'no answer within ' + BUS_DEADLINE_MS / 1000 + ' s' : error.message);
        resolve({ ok: false, error: 'no spreadsheet to ask: ' + why });
        return;
      }
      let answer;
      try {
        answer = JSON.parse(stdout);
      } catch (e) {
        resolve({ ok: false, error: 'sae-bus answered with something other than JSON: ' + stdout.slice(0, 200) });
        return;
      }
      if (answer.success) resolve({ ok: true, result: answer.result === undefined ? null : answer.result });
      else resolve({ ok: false, error: answer.error || 'the command failed (error ' + (answer.errorCode || 1) + ')' });
    });
  });
}

// A REXX stem ({0: count, 1: ..., n: ...}) as an array; anything else as is.
function unstem(v) {
  if (!isStem(v)) return v;
  const out = [];
  for (let i = 1; i <= Number(v['0']); i++) out.push(v[String(i)]);
  return out;
}

function isStem(v) {
  if (v === null || typeof v !== 'object' || Array.isArray(v) || !('0' in v)) return false;
  const n = Number(v['0']);
  if (!Number.isInteger(n) || n < 0) return false;
  for (let i = 1; i <= n; i++) if (!(String(i) in v)) return false;
  return true;
}

// Stems as arrays, all the way down.
function plain(v) {
  if (Array.isArray(v)) return v.map(plain);
  if (isStem(v)) return unstem(v).map(plain);
  if (v !== null && typeof v === 'object') return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, plain(x)]));
  return v;
}

// ---- ranges and cells ------------------------------------------------------

const RANGE = /^\s*([A-Za-z]{1,3})([1-9][0-9]*)\s*(?::\s*([A-Za-z]{1,3})([1-9][0-9]*)\s*)?$/;
const REF = /^([A-Za-z]+)([0-9]+)$/;

function colNumber(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function colLetters(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

// "A1:C5" as its corners, or null when it is not a range.
function parseRange(text) {
  const m = RANGE.exec(String(text));
  if (!m) return null;
  const c1 = colNumber(m[1]), r1 = Number(m[2]);
  const c2 = m[3] ? colNumber(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
  return { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: Math.min(c1, c2), right: Math.max(c1, c2) };
}

// GETCELLS's cells laid out as the range's rows, and the formulas by cell.
function asRows(span, cells) {
  const width = span.right - span.left + 1;
  const rows = [];
  for (let r = span.top; r <= span.bottom; r++) rows.push(new Array(width).fill(''));
  const formulas = {};
  for (const c of cells) {
    if (!c || typeof c !== 'object') continue;
    const m = REF.exec(String(c.ref));
    if (!m) continue;
    const r = Number(m[2]) - span.top, k = colNumber(m[1]) - span.left;
    if (r < 0 || r >= rows.length || k < 0 || k >= width) continue;
    rows[r][k] = c.value === undefined || c.value === null ? '' : c.value;
    if (c.expression) formulas[String(c.ref).toUpperCase()] = '=' + c.expression;
  }
  return { rows, formulas };
}

// ---- the functions, described by the app (LISTCOMMANDS(1)) -----------------

// Asked for once a server: the app's functions do not change while it runs.
let catalogue = null;

async function functions() {
  if (catalogue) return { ok: true, list: catalogue };
  const a = await bus('LISTCOMMANDS(1)');
  if (!a.ok) return a;
  const listed = unstem(a.result);
  const list = (Array.isArray(listed) ? listed : [])
    .filter((c) => c && typeof c === 'object')
    .map((c) => ({ ...c, params: (unstem(c.params) || []).map((x) => ({ ...x, optional: x.optional === 1, variadic: x.variadic === 1 })) }))
    .sort((x, y) => String(x.name).localeCompare(String(y.name)));
  catalogue = list;
  return { ok: true, list };
}

const isNumberish = (v) => (typeof v === 'number' && Number.isFinite(v)) || (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v)));

// One argument as the app takes it, or { error }. `where` names it.
function argument(param, v, where) {
  const bad = (wanted) => ({ error: where + ' (' + param.name + ') must be ' + wanted + ', not ' + JSON.stringify(v) });
  switch (param.type) {
    case 'row':
      return isNumberish(v) && Number(v) >= 1 && Number.isInteger(Number(v)) ? { value: Number(v) } : bad('a row number from 1');
    case 'number':
      return isNumberish(v) ? { value: Number(v) } : bad('a number');
    case 'flag':
      if (v === true || v === 1 || v === '1') return { value: 1 };
      if (v === false || v === 0 || v === '0') return { value: 0 };
      return bad('1 or 0 (or true or false)');
    case 'column':
      if (typeof v === 'string' && /^[A-Za-z]{1,3}$/.test(v.trim())) return { value: v.trim().toUpperCase() };
      return isNumberish(v) && Number(v) >= 1 && Number.isInteger(Number(v)) ? { value: Number(v) } : bad('a column, "B" or 2');
    case 'stem':
      return Array.isArray(v) ? { value: v } : bad('a list');
    case 'json':
      return typeof v === 'string' ? { value: v } : v !== null && typeof v === 'object' ? { value: JSON.stringify(v) } : bad('JSON, as text or an object');
    case 'cell':
    case 'range':
    case 'text':
      return typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean' ? { value: String(v) } : bad('text');
    default:
      return { value: v };
  }
}

// A call checked against the function's description: { command, args } as
// BATCH_EXECUTE takes it, or { error }.
function checked(list, call, which) {
  const name = String(call && call.function || '').trim().toUpperCase();
  const where = which ? 'call ' + which + ', ' : '';
  if (!name) return { error: where + 'no function named; spreadsheet_list_commands lists them' };
  const fn = list.find((c) => c.name === name);
  if (!fn) return { error: where + 'no function ' + JSON.stringify(call.function) + '; spreadsheet_list_commands lists them' };
  if (fn.params.some((x) => x.type === 'query')) {
    return { error: where + name + ' works on a query (RANGE(...) |> ' + name + '(...)): run it with spreadsheet_run, with --tools=rexx or all' };
  }
  const args = call.args === undefined ? [] : call.args;
  if (!Array.isArray(args)) return { error: where + name + '\'s args must be a list' };
  const fixed = fn.params.filter((x) => !x.variadic);
  const rest = fn.params.find((x) => x.variadic);
  const needed = fixed.filter((x) => !x.optional).length;
  if (args.length < needed || (!rest && args.length > fixed.length)) {
    const takes = rest ? 'at least ' + needed : needed === fixed.length ? String(needed) : needed + ' to ' + fixed.length;
    return { error: where + fn.signature + ' takes ' + takes + ' argument' + (takes === '1' ? '' : 's') + ', not ' + args.length };
  }
  const out = [];
  for (let i = 0; i < args.length; i++) {
    const param = i < fixed.length ? fixed[i] : rest;
    const a = argument(param, args[i], where + name + '\'s argument ' + (i + 1));
    if (a.error) return { error: a.error };
    out.push(a.value);
  }
  return { command: name, args: out };
}

// Calls as one request, BATCH_EXECUTE's, stopping at the first that fails
// (so the request is taken back whole): each one's result, or { error }.
async function batch(calls) {
  const a = await bus(batchCommand(calls));
  if (!a.ok) return { error: a.error };
  let results;
  try {
    results = typeof a.result === 'string' ? JSON.parse(a.result) : a.result;
  } catch (e) {
    return { error: 'the app answered with something other than its results: ' + String(a.result).slice(0, 200) };
  }
  return { results: (Array.isArray(results) ? results : []).map((r) => plain(r.result === undefined ? null : r.result)) };
}

// The REXX that carries calls to the app with nothing in it to escape. A
// value's text never appears in a REXX string: RexxJS reads \n, \t and \r
// in its strings, in single quotes as in double, before it reads \\, so
// "C:\temp" arrives with a tab in it and no escaping gets it there whole
// (RexxJS's TODO.md). The calls go as JSON in base64,
// which RexxJS's BASE64_DECODE gives back; that reads bytes as Latin-1, so
// the JSON is first made ASCII (\uXXXX for the rest, which JSON.parse in
// the app turns back).
function batchCommand(calls) {
  const ascii = JSON.stringify(calls).replace(/[\u007f-\uffff]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));
  return "BATCH_EXECUTE(BASE64_DECODE('" + Buffer.from(ascii, 'ascii').toString('base64') + "'), 1)";
}

// ---- answers ---------------------------------------------------------------

function text(value) {
  return typeof value === 'string' ? value : JSON.stringify(value, null, 2);
}

function ok(value, structured) {
  const r = { content: [{ type: 'text', text: text(value) }], isError: false };
  if (structured !== undefined) r.structuredContent = structured;
  return r;
}

function failed(message) {
  return { content: [{ type: 'text', text: message }], isError: true };
}

async function callTool(name, args, surface = 'all') {
  args = args || {};
  if (!toolsFor(surface).some((t) => t.name === name)) return null;
  if (name === 'spreadsheet_run') {
    if (typeof args.command !== 'string' || args.command.trim() === '') return failed('spreadsheet_run needs a command, e.g. SETCELL("A1", "100")');
    const a = await bus(args.command);
    if (!a.ok) return failed(a.error);
    const result = plain(a.result);
    return ok(result === null ? '(done; no result)' : result, { result });
  }
  if (name === 'spreadsheet_call') {
    const many = args.calls !== undefined;
    if (many === (args.function !== undefined)) return failed('spreadsheet_call takes a function (with its args) or calls, one of the two');
    if (many && args.args !== undefined) return failed('spreadsheet_call: args go inside each of the calls');
    const calls = many ? args.calls : [{ function: args.function, args: args.args }];
    if (!Array.isArray(calls) || calls.length === 0) return failed('spreadsheet_call: calls must be a list of at least one');
    const f = await functions();
    if (!f.ok) return failed(f.error);
    const sent = [];
    for (let i = 0; i < calls.length; i++) {
      const c = checked(f.list, calls[i], many ? i + 1 : 0);
      if (c.error) return failed(c.error + (many ? ' (nothing was sent)' : ''));
      sent.push(c);
    }
    const b = await batch(sent);
    if (b.error) return failed(b.error);
    if (!many) {
      const result = b.results[0] === undefined ? null : b.results[0];
      return ok(result === null || result === '' ? '(done; no result)' : result, { result });
    }
    return ok(b.results, { results: b.results });
  }
  if (name === 'spreadsheet_write_cells') {
    const m = REF.exec(String(args.cell || '').trim());
    if (!m) return failed('spreadsheet_write_cells needs a cell to start from, such as "A1", not ' + JSON.stringify(args.cell));
    if (!Array.isArray(args.rows) || args.rows.some((r) => !Array.isArray(r))) return failed('spreadsheet_write_cells: rows must be a list of lists');
    const left = colNumber(m[1]), top = Number(m[2]);
    const sets = [];
    let width = 0;
    for (let r = 0; r < args.rows.length; r++) {
      width = Math.max(width, args.rows[r].length);
      for (let k = 0; k < args.rows[r].length; k++) {
        const v = args.rows[r][k];
        if (v !== null && !['string', 'number', 'boolean'].includes(typeof v)) {
          return failed('spreadsheet_write_cells: row ' + (r + 1) + ', value ' + (k + 1) + ' is ' + JSON.stringify(v) + ': text, a number, true/false or null');
        }
        sets.push({ command: 'SETCELL', args: [colLetters(left + k) + (top + r), v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)] });
      }
    }
    if (sets.length === 0) return failed('spreadsheet_write_cells: no values to write');
    if (sets.length > MOST_CELLS) return failed('spreadsheet_write_cells writes at most ' + MOST_CELLS + ' cells a call, not ' + sets.length + ': split the rows');
    const b = await batch(sets);
    if (b.error) return failed(b.error);
    const range = colLetters(left) + top + ':' + colLetters(left + width - 1) + (top + args.rows.length - 1);
    return ok('wrote ' + sets.length + ' cell' + (sets.length === 1 ? '' : 's') + ' in ' + range, { range, cells: sets.length });
  }
  if (name === 'spreadsheet_read_range') {
    const span = parseRange(args.range);
    if (!span) return failed('spreadsheet_read_range needs a range such as "A1:C5", not ' + JSON.stringify(args.range));
    const range = String(args.range).replace(/\s+/g, '').toUpperCase();
    const a = await bus('GETCELLS(' + rexxString(range.includes(':') ? range : range + ':' + range) + ')');
    if (!a.ok) return failed(a.error);
    const cells = unstem(a.result);
    const laid = asRows(span, Array.isArray(cells) ? cells : []);
    const structured = { range, rows: laid.rows, formulas: laid.formulas };
    return ok(structured, structured);
  }
  if (name === 'spreadsheet_list_commands') {
    const f = await functions();
    if (!f.ok) return failed(f.error);
    const commands = f.list;
    if (args.name !== undefined) {
      const one = commands.find((c) => c.name === String(args.name).trim().toUpperCase());
      if (!one) return failed('no function ' + JSON.stringify(args.name) + '; spreadsheet_list_commands with no name lists them');
      const lines = [one.signature, one.does + '; gives ' + one.gives + '.'];
      for (const x of one.params) lines.push('  ' + x.name + ': ' + x.type + (x.optional ? ', optional' : '') + (x.variadic ? ', any number' : '') + (x.about ? ' (' + x.about + ')' : ''));
      return ok(lines.join('\n'), one);
    }
    return ok(commands.map((c) => c.signature + ': ' + c.does + '; gives ' + c.gives).join('\n'), { commands });
  }
  return null;
}

// ---- JSON-RPC over stdio -------------------------------------------------

function send(message) {
  process.stdout.write(JSON.stringify(message) + '\n');
}

function reply(id, result) {
  send({ jsonrpc: '2.0', id, result });
}

function error(id, code, message) {
  send({ jsonrpc: '2.0', id, error: { code, message } });
}

async function handle(msg, surface) {
  const isRequest = msg.id !== undefined && msg.id !== null;
  switch (msg.method) {
    case 'initialize': {
      const asked = msg.params && msg.params.protocolVersion;
      reply(msg.id, {
        protocolVersion: PROTOCOL_VERSIONS.includes(asked) ? asked : PROTOCOL_VERSIONS[0],
        capabilities: { tools: { listChanged: false } },
        serverInfo: SERVER,
        instructions:
          'Drives a running sae spreadsheet over its control bus. Read before you write, keep each call small, ' +
          'and tell the person what you changed: each call is one Undo for them.',
      });
      return;
    }
    case 'ping':
      if (isRequest) reply(msg.id, {});
      return;
    case 'tools/list':
      reply(msg.id, { tools: toolsFor(surface) });
      return;
    case 'tools/call': {
      const p = msg.params || {};
      const result = await callTool(p.name, p.arguments, surface);
      if (result === null) error(msg.id, -32602, 'no tool named ' + JSON.stringify(p.name));
      else reply(msg.id, result);
      return;
    }
    default:
      // Notifications (initialized, cancelled) need no answer; an unknown
      // request does.
      if (isRequest) error(msg.id, -32601, 'no method ' + JSON.stringify(msg.method));
  }
}

if (require.main === module) {
  let surface;
  try {
    surface = surfaceOf(process.argv.slice(2), process.env);
  } catch (e) {
    process.stderr.write('spreadsheet-mcp: ' + e.message + '\n');
    process.exit(2);
  }
  // Requests are answered in order: the bus takes one command at a time.
  let queue = Promise.resolve();
  const lines = readline.createInterface({ input: process.stdin });
  lines.on('line', (line) => {
    if (line.trim() === '') return;
    let msg;
    try {
      msg = JSON.parse(line);
    } catch (e) {
      // In its turn, like any answer.
      queue = queue.then(() => error(null, -32700, 'not JSON: ' + e.message));
      return;
    }
    queue = queue.then(() => handle(msg, surface)).catch((e) => {
      if (msg.id !== undefined && msg.id !== null) error(msg.id, -32603, String(e && e.message || e));
    });
  });
  lines.on('close', () => queue.then(() => process.exit(0)));
}

module.exports = { toolsFor, surfaceOf, rexxString, batchCommand, unstem, plain, parseRange, asRows, checked, callTool, handle };
