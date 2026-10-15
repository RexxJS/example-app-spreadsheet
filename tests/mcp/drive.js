#!/usr/bin/env node
// One MCP tool call to mcp/spreadsheet-mcp.js, as a client makes it:
//   node tests/mcp/drive.js <tool> '<arguments as JSON>'
//   node tests/mcp/drive.js <tool> @tests/mcp/args/<file>.json
// starts the server, does the handshake, calls the tool, and prints one line
// of JSON, {isError, result}: the structured result, or the text when there
// is none. For tests/specs/spec_app_sheet_mcp.ae, which checks the grid.
'use strict';
const { spawn } = require('child_process');
const path = require('path');
const readline = require('readline');

const [tool, given = '{}'] = process.argv.slice(2);
const args = given.startsWith('@') ? require('fs').readFileSync(given.slice(1), 'utf8') : given;
const server = spawn(process.execPath, [path.join(__dirname, '..', '..', 'mcp', 'spreadsheet-mcp.js')], { stdio: ['pipe', 'pipe', 'inherit'] });
const send = (m) => server.stdin.write(JSON.stringify(m) + '\n');

readline.createInterface({ input: server.stdout }).on('line', (line) => {
  const m = JSON.parse(line);
  if (m.id === 1) {
    send({ jsonrpc: '2.0', method: 'notifications/initialized' });
    send({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: tool, arguments: JSON.parse(args) } });
  } else if (m.id === 2) {
    const r = m.result;
    const out = r ? { isError: r.isError, result: r.structuredContent !== undefined ? r.structuredContent : r.content[0].text } : { error: m.error };
    process.stdout.write(JSON.stringify(out) + '\n');
    server.stdin.end();
  }
});
server.on('close', (code) => process.exit(code));
send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'spec', version: '0' } } });
