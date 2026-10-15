// mcp/spreadsheet-mcp.js spoken to as an MCP client does, over stdio, with
// tests/mcp/fake-sae-bus.js in place of sae's sae-bus: the handshake, the
// tools each --tools setting offers, each tool's request on the bus and its
// answer, the arguments checked before anything is sent, and what a client
// is told when a call fails or there is no app to ask.
//   node --test tests/mcp/
// The app itself is driven through this server by tests/specs/spec_app_sheet_mcp.ae.
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { spawn, spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const SERVER = path.join(__dirname, '..', '..', 'mcp', 'spreadsheet-mcp.js');
const FAKE = path.join(__dirname, 'fake-sae-bus.js');

// A server with the fake bus; `talk(messages)` sends them one a line and
// gives back the replies to those with ids, in order, and the bus's calls.
async function talk(messages, env = {}, argv = []) {
  const log = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-mcp-')), 'bus.log');
  const child = spawn(process.execPath, [SERVER, ...argv], {
    env: { ...process.env, SAE_BUS: FAKE, FAKE_BUS_LOG: log, SPREADSHEET_MCP_TOOLS: '', ...env },
    stdio: ['pipe', 'pipe', 'inherit'],
  });
  let out = '';
  child.stdout.on('data', (d) => (out += d));
  for (const m of messages) child.stdin.write((typeof m === 'string' ? m : JSON.stringify(m)) + '\n');
  child.stdin.end();
  const code = await new Promise((resolve) => child.on('close', resolve));
  const replies = out.split('\n').filter((l) => l.trim() !== '').map((l) => JSON.parse(l));
  const calls = fs.existsSync(log) ? fs.readFileSync(log, 'utf8').split('\n').filter((l) => l).map((l) => JSON.parse(l)) : [];
  return { code, replies, calls };
}

const INIT = { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '0' } } };
const INITIALIZED = { jsonrpc: '2.0', method: 'notifications/initialized' };
const LIST = { jsonrpc: '2.0', id: 2, method: 'tools/list' };
const call = (id, name, args) => ({ jsonrpc: '2.0', id, method: 'tools/call', params: { name, arguments: args } });
const REXX = ['--tools=rexx'];
// The calls a BATCH_EXECUTE request carried, as the app reads them.
// Nothing but base64 in it, so nothing a value holds is read as REXX.
const batched = (command) => JSON.parse(Buffer.from(/^BATCH_EXECUTE\(BASE64_DECODE\('([A-Za-z0-9+/=]*)'\), 1\)$/.exec(command)[1], 'base64').toString('latin1'));

// ---- the handshake, and the tools each setting offers

test('the handshake: the version asked for, tools offered, and the notification left unanswered', async () => {
  const { code, replies } = await talk([INIT, INITIALIZED, { jsonrpc: '2.0', id: 2, method: 'ping' }]);
  assert.equal(code, 0);
  assert.equal(replies.length, 2, 'two requests, two replies: none to the notification');
  assert.equal(replies[0].result.protocolVersion, '2025-06-18');
  assert.deepEqual(replies[0].result.capabilities, { tools: { listChanged: false } });
  assert.equal(replies[0].result.serverInfo.name, 'spreadsheet');
  assert.deepEqual(replies[1], { jsonrpc: '2.0', id: 2, result: {} });
});

test('a protocol version it does not know is answered with the newest it does', async () => {
  const { replies } = await talk([{ ...INIT, params: { ...INIT.params, protocolVersion: '1999-01-01' } }]);
  assert.equal(replies[0].result.protocolVersion, '2025-06-18');
});

const names = (r) => r.result.tools.map((t) => t.name);

test('json, the default: the four tools that take JSON, and no REXX to write', async () => {
  const { replies } = await talk([INIT, LIST]);
  assert.deepEqual(names(replies[1]), ['spreadsheet_write_cells', 'spreadsheet_call', 'spreadsheet_read_range', 'spreadsheet_list_commands']);
  for (const t of replies[1].result.tools) {
    assert.equal(t.inputSchema.type, 'object', t.name);
    assert.match(t.description, /one step of its undo/, t.name);
    assert.match(t.description, /not told when the person edits/, t.name);
  }
  assert.match(replies[1].result.tools[3].description, /Every function spreadsheet_call may call/);
});

test('--tools=rexx: spreadsheet_run, with REXX\'s escaping said, and the two that read', async () => {
  const { replies } = await talk([INIT, LIST], {}, REXX);
  assert.deepEqual(names(replies[1]), ['spreadsheet_run', 'spreadsheet_read_range', 'spreadsheet_list_commands']);
  assert.match(replies[1].result.tools[0].description, /a quote inside one is written twice/);
  assert.match(replies[1].result.tools[2].description, /Every function spreadsheet_run may call/);
});

test('--tools=all, or SPREADSHEET_MCP_TOOLS=all: all five', async () => {
  const flag = await talk([INIT, LIST], {}, ['--tools=all']);
  const env = await talk([INIT, LIST], { SPREADSHEET_MCP_TOOLS: 'ALL' });
  for (const { replies } of [flag, env]) {
    assert.deepEqual(names(replies[1]), ['spreadsheet_write_cells', 'spreadsheet_call', 'spreadsheet_read_range', 'spreadsheet_list_commands', 'spreadsheet_run']);
  }
});

test('the flag wins over the environment', async () => {
  const { replies } = await talk([INIT, LIST], { SPREADSHEET_MCP_TOOLS: 'all' }, REXX);
  assert.equal(names(replies[1]).length, 3);
});

test('a setting it does not know: said on stderr, exit 2, before any client is answered', () => {
  const r = spawnSync(process.execPath, [SERVER, '--tools=everything'], { input: JSON.stringify(INIT) + '\n', encoding: 'utf8' });
  assert.equal(r.status, 2);
  assert.equal(r.stdout, '');
  assert.match(r.stderr, /--tools is json, rexx or all, not "everything"/);
});

test('a tool the setting does not offer is no tool at all', async () => {
  const json = await talk([INIT, call(2, 'spreadsheet_run', { command: 'GETCELL("A1")' })]);
  assert.equal(json.replies[1].error.code, -32602);
  assert.deepEqual(json.calls, []);
  const rexx = await talk([INIT, call(2, 'spreadsheet_call', { function: 'GETCELL', args: ['A1'] })], {}, REXX);
  assert.equal(rexx.replies[1].error.code, -32602);
});

// ---- spreadsheet_call

test('spreadsheet_call: the function checked against LISTCOMMANDS(1), then sent as BATCH_EXECUTE that stops at a failure', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_call', { function: 'sortrange', args: ['A1:C10', 'b', false] })]);
  assert.deepEqual(calls.map((c) => c[1].slice(0, 29)), ['LISTCOMMANDS(1)', 'BATCH_EXECUTE(BASE64_DECODE(\'']);
  assert.deepEqual(batched(calls[1][1]), [{ command: 'SORTRANGE', args: ['A1:C10', 'B', 0] }], 'its name and column in capitals, false as 0');
  assert.equal(replies[1].result.isError, false);
  assert.deepEqual(replies[1].result.structuredContent, { result: 'OK' });
});

test('spreadsheet_call: text with quotes, backslashes, new lines and more than ASCII goes in as it is, nothing for the model to escape', async () => {
  const value = 'Lunch "client" meeting\\n, C:\\temp\nsecond line "" \'single\' ünï € 😀';
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_call', { function: 'SETCELL', args: ['B2', value] })]);
  assert.deepEqual(batched(calls[1][1]), [{ command: 'SETCELL', args: ['B2', value] }]);
  assert.equal(replies[1].result.structuredContent.result, value);
});

test('spreadsheet_call: a list for a stem, an object for JSON, and a stem given back as a list', async () => {
  const { replies, calls } = await talk([INIT,
    call(2, 'spreadsheet_call', { function: 'SETCELLS', args: ['A1:A3', [1, 'two', 3]] }),
    call(3, 'spreadsheet_call', { function: 'SETCELLEDITOR', args: ['C1', 'dropdown', { options: ['a', 'b'] }] }),
    call(4, 'spreadsheet_call', { function: 'GETSHEETNAMES' })]);
  assert.deepEqual(batched(calls[1][1])[0].args, ['A1:A3', [1, 'two', 3]]);
  assert.deepEqual(batched(calls[2][1])[0].args, ['C1', 'dropdown', '{"options":["a","b"]}']);
  assert.deepEqual(replies[3].result.structuredContent, { result: ['Sheet1', 'Costs'] });
  assert.equal(calls.filter((c) => c[1] === 'LISTCOMMANDS(1)').length, 1, 'the functions asked for once');
});

test('spreadsheet_call with calls: in turn, as one request, each one\'s result', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_call', { calls: [
    { function: 'INSERTROW', args: ['2'] }, { function: 'SETCELL', args: ['A2', '=SUM(B3:B9)'] }, { function: 'GETCELL', args: ['A2'] }] })]);
  assert.deepEqual(batched(calls[1][1]), [
    { command: 'INSERTROW', args: [2] }, { command: 'SETCELL', args: ['A2', '=SUM(B3:B9)'] }, { command: 'GETCELL', args: ['A2'] }]);
  assert.deepEqual(replies[1].result.structuredContent, { results: ['OK', '=SUM(B3:B9)', 'shown'] });
});

test('spreadsheet_call: one that fails in the app is a tool error saying which, and that nothing was kept', async () => {
  const { replies } = await talk([INIT, call(2, 'spreadsheet_call', { calls: [{ function: 'SETCELL', args: ['A1', 'x'] }, { function: 'SORTRANGE', args: ['A1:B2', 'Q'] }] })]);
  assert.equal(replies[1].result.isError, true);
  assert.equal(replies[1].result.content[0].text, 'command 2 (SORTRANGE): Q is not in A1:B2 (nothing it did was kept)');
});

const refused = async (args) => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_call', args)]);
  assert.equal(replies[1].result.isError, true, JSON.stringify(args));
  assert.deepEqual(calls.filter((c) => c[1].startsWith('BATCH')), [], 'nothing sent');
  return replies[1].result.content[0].text;
};

test('spreadsheet_call: arguments checked before anything is sent, each refusal naming what is wrong', async () => {
  assert.equal(await refused({ function: 'NOPE' }), 'no function "NOPE"; spreadsheet_list_commands lists them');
  assert.equal(await refused({ function: 'SETCELL', args: ['A1'] }), 'SETCELL(cell, value) takes 2 arguments, not 1');
  assert.equal(await refused({ function: 'SORTRANGE', args: ['A1:B2'] }), 'SORTRANGE(range, by, [ascending]) takes 2 to 3 arguments, not 1');
  assert.equal(await refused({ function: 'INSERTROW', args: [0] }), 'INSERTROW\'s argument 1 (row) must be a row number from 1, not 0');
  assert.equal(await refused({ function: 'SORTRANGE', args: ['A1:B2', 'B', 'yes'] }), 'SORTRANGE\'s argument 3 (ascending) must be 1 or 0 (or true or false), not "yes"');
  assert.equal(await refused({ function: 'SORTRANGE', args: ['A1:B2', 'B2'] }), 'SORTRANGE\'s argument 2 (by) must be a column, "B" or 2, not "B2"');
  assert.equal(await refused({ function: 'SETCELLS', args: ['A1:A2', '1,2'] }), 'SETCELLS\'s argument 2 (values) must be a list, not "1,2"');
  assert.equal(await refused({ function: 'SETCELL', args: ['A1', { x: 1 }] }), 'SETCELL\'s argument 2 (value) must be text, not {"x":1}');
  assert.match(await refused({ function: 'WHERE', args: ['x', 'amount > 1'] }), /WHERE works on a query .* run it with spreadsheet_run, with --tools=rexx or all/);
  assert.equal(await refused({ function: 'GETCELL', args: ['A1'], calls: [] }), 'spreadsheet_call takes a function (with its args) or calls, one of the two');
  assert.equal(await refused({ calls: [] }), 'spreadsheet_call: calls must be a list of at least one');
  assert.equal(await refused({ calls: [{ function: 'GETCELL', args: ['A1'] }, { function: 'SETCELL', args: [] }] }),
    'call 2, SETCELL(cell, value) takes 2 arguments, not 0 (nothing was sent)');
});

test('spreadsheet_call: the rest of the arguments, for a function that takes them', async () => {
  const { calls } = await talk([INIT, call(2, 'spreadsheet_call', { function: 'SETCELLVALIDATION', args: ['B2:B9', 'number', 0, 100] })]);
  assert.deepEqual(batched(calls[1][1])[0].args, ['B2:B9', 'number', '0', '100']);
  assert.equal(await refused({ function: 'SETCELLVALIDATION', args: ['B2:B9'] }), 'SETCELLVALIDATION(range, type, [limits...]) takes at least 2 arguments, not 1');
});

// ---- spreadsheet_write_cells

test('spreadsheet_write_cells: rows from a cell, one SETCELL a value, as one request', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_write_cells', { cell: 'b2', rows: [['Date', 'Item', 'Amount'], ['2026-10-01', 'Taxi "airport"', 23.5], ['', null, '=SUM(D3:D3)', true]] })]);
  assert.equal(calls.length, 1, 'no LISTCOMMANDS needed');
  assert.deepEqual(batched(calls[0][1]), [
    { command: 'SETCELL', args: ['B2', 'Date'] }, { command: 'SETCELL', args: ['C2', 'Item'] }, { command: 'SETCELL', args: ['D2', 'Amount'] },
    { command: 'SETCELL', args: ['B3', '2026-10-01'] }, { command: 'SETCELL', args: ['C3', 'Taxi "airport"'] }, { command: 'SETCELL', args: ['D3', '23.5'] },
    { command: 'SETCELL', args: ['B4', ''] }, { command: 'SETCELL', args: ['C4', ''] }, { command: 'SETCELL', args: ['D4', '=SUM(D3:D3)'] }, { command: 'SETCELL', args: ['E4', 'TRUE'] },
  ]);
  assert.deepEqual(replies[1].result.structuredContent, { range: 'B2:E4', cells: 10 });
  assert.equal(replies[1].result.content[0].text, 'wrote 10 cells in B2:E4');
});

test('spreadsheet_write_cells: past column Z, the letters go on (Z, AA, AB)', async () => {
  const { calls } = await talk([INIT, call(2, 'spreadsheet_write_cells', { cell: 'Z1', rows: [[1, 2, 3]] })]);
  assert.deepEqual(batched(calls[0][1]).map((c) => c.args[0]), ['Z1', 'AA1', 'AB1']);
});

test('spreadsheet_write_cells: refused before anything is sent', async () => {
  const cases = [
    [{ cell: 'A1:B2', rows: [[1]] }, /needs a cell to start from, such as "A1"/],
    [{ cell: 'A1', rows: [1, 2] }, /rows must be a list of lists/],
    [{ cell: 'A1', rows: [[{ a: 1 }]] }, /row 1, value 1 is \{"a":1\}/],
    [{ cell: 'A1', rows: [[]] }, /no values to write/],
    [{ cell: 'A1', rows: Array.from({ length: 41 }, () => new Array(50).fill(1)) }, /at most 2000 cells a call, not 2050: split the rows/],
  ];
  for (const [args, why] of cases) {
    const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_write_cells', args)]);
    assert.equal(replies[1].result.isError, true);
    assert.match(replies[1].result.content[0].text, why);
    assert.deepEqual(calls, []);
  }
});

// ---- spreadsheet_run (--tools=rexx)

test('spreadsheet_run: the command goes to the bus as it is, for the app named', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_run', { command: 'SETCELL("A1", "100")' })], { SPREADSHEET_APP: 'Sheet Two' }, REXX);
  assert.deepEqual(calls, [['Sheet Two', 'SETCELL("A1", "100")']]);
  assert.equal(replies[1].result.content[0].text, '100');
  assert.deepEqual(replies[1].result.structuredContent, { result: '100' });
});

test('spreadsheet_run: a stem comes back as a list, nothing as "done"', async () => {
  const { replies } = await talk([INIT, call(2, 'spreadsheet_run', { command: 'GETSHEETNAMES()' }), call(3, 'spreadsheet_run', { command: 'UNDO()' })], {}, REXX);
  assert.deepEqual(replies[1].result.structuredContent, { result: ['Sheet1', 'Costs'] });
  assert.equal(replies[2].result.content[0].text, '(done; no result)');
});

test('spreadsheet_run: a refused command is a tool error with the app\'s reason, not a protocol error', async () => {
  const { replies } = await talk([INIT, call(2, 'spreadsheet_run', { command: 'BOGUS("A1")' })], {}, REXX);
  assert.equal(replies[1].error, undefined);
  assert.equal(replies[1].result.isError, true);
  assert.equal(replies[1].result.content[0].text, 'Function BOGUS is not available');
});

test('spreadsheet_run: an empty command is refused before the bus', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_run', { command: '  ' })], {}, REXX);
  assert.equal(replies[1].result.isError, true);
  assert.deepEqual(calls, []);
});

// ---- spreadsheet_read_range and spreadsheet_list_commands

test('spreadsheet_read_range: GETCELLS asked for, its stem laid out as rows, formulas beside', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_read_range', { range: ' a1 : b2 ' })]);
  assert.deepEqual(calls.map((c) => c[1]), ['GETCELLS("A1:B2")']);
  assert.deepEqual(replies[1].result.structuredContent, { range: 'A1:B2', rows: [['2', 'x'], [5, '']], formulas: { A2: '=A1+3' } });
  assert.deepEqual(JSON.parse(replies[1].result.content[0].text), replies[1].result.structuredContent);
});

test('spreadsheet_read_range: one cell is a range of one', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_read_range', { range: 'C3' })]);
  assert.deepEqual(calls.map((c) => c[1]), ['GETCELLS("C3:C3")']);
  assert.deepEqual(replies[1].result.structuredContent.rows, [['one']]);
});

test('spreadsheet_read_range: what is not a range never reaches the app (no REXX can ride in on it)', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_read_range', { range: 'A1") ; DELETESHEET("Costs' })]);
  assert.equal(replies[1].result.isError, true);
  assert.match(replies[1].result.content[0].text, /needs a range such as "A1:C5"/);
  assert.deepEqual(calls, []);
});

test('spreadsheet_list_commands: each function a line of signature, does and gives, sorted', async () => {
  const { replies, calls } = await talk([INIT, call(2, 'spreadsheet_list_commands', {})]);
  assert.deepEqual(calls.map((c) => c[1]), ['LISTCOMMANDS(1)']);
  const lines = replies[1].result.content[0].text.split('\n');
  assert.equal(lines[0], 'GETCELL(cell): Reads a cell; gives what it shows');
  assert.ok(lines.includes('SORTRANGE(range, by, [ascending]): Sorts a range\'s rows by a column; gives "OK"'));
  const sort = replies[1].result.structuredContent.commands.find((c) => c.name === 'SORTRANGE');
  assert.deepEqual(sort.params.map((x) => [x.name, x.type, x.optional]), [['range', 'range', false], ['by', 'column', false], ['ascending', 'flag', true]]);
});

test('spreadsheet_list_commands with a name: that one in full, each parameter said', async () => {
  const { replies } = await talk([INIT, call(2, 'spreadsheet_list_commands', { name: 'sortrange' })]);
  assert.equal(replies[1].result.content[0].text, [
    'SORTRANGE(range, by, [ascending])',
    'Sorts a range\'s rows by a column; gives "OK".',
    '  range: range',
    '  by: column',
    '  ascending: flag, optional (1 or left out: ascending; 0: descending)',
  ].join('\n'));
  assert.equal(replies[1].result.structuredContent.name, 'SORTRANGE');
});

test('spreadsheet_list_commands with a name it does not have: a tool error', async () => {
  const { replies } = await talk([INIT, call(2, 'spreadsheet_list_commands', { name: 'NOPE' })]);
  assert.equal(replies[1].result.isError, true);
  assert.match(replies[1].result.content[0].text, /no function "NOPE"/);
});

// ---- failures and the protocol

test('no app to ask: a tool error carrying sae-bus\'s reason', async () => {
  const { replies } = await talk([INIT, call(2, 'spreadsheet_list_commands', {})], { FAKE_BUS_NO_APP: '1' });
  assert.equal(replies[1].result.isError, true);
  assert.equal(replies[1].result.content[0].text, 'no spreadsheet to ask: sae-bus: no courier file at /nonexistent/.control-bus');
});

test('an answer that is not JSON is said so, not thrown', async () => {
  const { replies } = await talk([INIT, call(2, 'spreadsheet_read_range', { range: 'A1' })], { FAKE_BUS_GARBLED: '1' });
  assert.equal(replies[1].result.isError, true);
  assert.match(replies[1].result.content[0].text, /something other than JSON: not json/);
});

test('protocol errors: an unknown tool, an unknown method, a line that is not JSON', async () => {
  const { replies } = await talk([INIT, call(2, 'spreadsheet_nope', {}), { jsonrpc: '2.0', id: 3, method: 'resources/list' }, '{not json']);
  assert.equal(replies[1].error.code, -32602);
  assert.equal(replies[2].error.code, -32601);
  assert.equal(replies[3].error.code, -32700);
  assert.equal(replies[3].id, null);
});

test('requests are answered in the order they came, one bus call at a time', async () => {
  const msgs = [INIT];
  for (let i = 0; i < 6; i++) msgs.push(call(10 + i, 'spreadsheet_run', { command: `SETCELL("A${i + 1}", "${i}")` }));
  const { replies, calls } = await talk(msgs, {}, REXX);
  assert.deepEqual(replies.slice(1).map((r) => r.id), [10, 11, 12, 13, 14, 15]);
  assert.deepEqual(calls.map((c) => c[1]), msgs.slice(1).map((m) => m.params.arguments.command));
});
