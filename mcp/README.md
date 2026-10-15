# An MCP server for the spreadsheet

`spreadsheet-mcp.js` lets an AI client (Claude, or any MCP client) drive a
running spreadsheet. It speaks MCP over stdio and sends each tool call to
the app over its control bus, through sae's `sae-bus`, as
`rexx/spreadsheet-address.js` does for REXX scripts. It needs only node,
with no packages.

## Using it

Start the app with its control bus on:

```bash
../sae/target/build/bin/sae --app app --control-bus
```

Then add the server to the client. For Claude Code:

```bash
claude mcp add spreadsheet -- node /path/to/spraedsheet/mcp/spreadsheet-mcp.js
```

Or in a client's JSON configuration:

```json
{ "mcpServers": { "spreadsheet": { "command": "node", "args": ["/path/to/spraedsheet/mcp/spreadsheet-mcp.js"] } } }
```

`SAE_BUS` names another `sae-bus` (the default is `../sae/target/sae-bus`),
`SPREADSHEET_APP` another app name (default `Spreadsheet`), and
`SAE_APPDATA_DIR` the app's data folder when it was started with one.

## Which tools: `--tools=json`, `rexx` or `all`

Set when the server starts, by `--tools=...` after the script or
`SPREADSHEET_MCP_TOOLS` (the flag wins):

| Setting | Tools | For |
|---|---|---|
| `json` (the default) | `spreadsheet_write_cells`, `spreadsheet_call`, `spreadsheet_run_program`, `spreadsheet_read_range`, `spreadsheet_list_commands` | any model, any client: arguments are JSON, and nothing needs escaping |
| `rexx` | `spreadsheet_run`, `spreadsheet_run_program`, `spreadsheet_read_range`, `spreadsheet_list_commands` | those who would write the app's REXX commands themselves |
| `all` | all six | both |

```bash
claude mcp add spreadsheet -- node /path/to/mcp/spreadsheet-mcp.js --tools=all
```

## The tools

| Tool | What it does |
|---|---|
| `spreadsheet_write_cells` | Writes rows of values from a cell: `{"cell": "A1", "rows": [["Date", "Item", "Amount"], ["2026-10-01", "Taxi", 23.5]]}`. Text, numbers, true/false, null (empties a cell) or a formula (`"=SUM(C2:C4)"`); up to 2000 cells a call. |
| `spreadsheet_call` | Calls one of the app's functions by name, its arguments a JSON list: `{"function": "SORTRANGE", "args": ["A2:C9", "C", 0]}`. Or several in turn, as one step: `{"calls": [{"function": ..., "args": [...]}, ...]}`. The arguments are checked against the function's description (how many, and of what kind) before anything is sent, and a refusal names the argument. |
| `spreadsheet_run_program` | Runs a whole REXX program in the app, next to the data: loops, conditions, routines, `SAY`, calling the same functions as `spreadsheet_call`. One step of undo, all or nothing; its `SAY` lines and `RESULT` come back. A loop over hundreds of cells is one call, not hundreds. Sent as it is (base64 JSON to the app's `RUN_PROGRAM`), so nothing in it needs escaping. |
| `spreadsheet_read_range` | Reads a range (`"A1:C5"`) as rows of shown values, with the formulas behind them. |
| `spreadsheet_list_commands` | Lists the app's functions (about 130), each as its signature (`SORTRANGE(range, by, [ascending])`), what it does and what it gives, from the app's `LISTCOMMANDS(1)` (`app/sheet/described.ts`); with a `name`, one in full, each parameter said. |
| `spreadsheet_run` | Runs one command of REXX, e.g. `SETCELL("A1", "100"); SETCELL("A2", "=A1*2")`; queries too (`RESULT(RANGE("A1:C9") \|> WHERE("amount > 100"))`), which `spreadsheet_call` does not take. |

There is one generic `spreadsheet_call`, not a tool for each of the app's
130 functions: a client carries every tool's description into each turn,
and models choose better among a few. `spreadsheet_list_commands` is how a
client finds the function it wants, and a function added to the app needs
nothing here.

## Nothing to escape

`spreadsheet_write_cells`, `spreadsheet_call` and `spreadsheet_run_program`
never put a value's text in a REXX string. They send the calls (or the
program) to the app's `BATCH_EXECUTE` (or `RUN_PROGRAM`) as JSON,
base64-encoded (`BATCH_EXECUTE(BASE64_DECODE('...'), 1)`), so quotes,
backslashes, new lines and any language's letters arrive as they were
sent. That matters: RexxJS reads `\n`, `\t` and `\r` in a string literal
as escapes, and a backslash before one of them cannot be written in a
string a host function is given (`"C:\temp"` arrives with a tab: RexxJS's
`TODO.md`). In `spreadsheet_run`, a quote in a string is written twice
(`"say ""hi"""`), and text with backslashes is better put in through the
JSON tools.

## What a client should know

The tools' descriptions say it too:

- **Each call is one step of undo** in the app: whatever one call does, the
  person takes it back with one Undo. `spreadsheet_call` with several
  calls, a whole `spreadsheet_run_program`, and `spreadsheet_run` with
  several statements, are one step.
- **All or nothing.** A call that fails part way is taken back whole, and
  the error says so: "(nothing it did was kept)".
- **Request and reply only.** The client is not told when the person edits
  the sheet, so it should read again before relying on an earlier read.
- **A program is sandboxed and bounded.** It runs in the page's own
  JavaScript engine inside the sae process, with the app's functions and
  nothing else: no files, no shell, no network, no `REQUIRE`. sae stops
  one stretch of work past 5 s, and the bus answers for a request still
  running at 10 s; either way the app takes the request back whole and
  the next request is one Undo as before.
- **One call at a time**, each answered within 10 seconds, on this machine
  only (the bus is loopback, off unless asked for, and its key is fresh
  each launch and readable only by this user: sae's `docs/control-bus.md`).

## Tests

```bash
./run-tests.sh mcp                  # the protocol, with a stand-in for sae-bus
./run-tests.sh spec_app_sheet_mcp   # the app driven through the server
```

`tests/mcp/server.test.js` runs on `node --test` against
`tests/mcp/fake-sae-bus.js`; `tests/specs/spec_app_sheet_mcp.ae` starts the
app with `--control-bus`, makes each call through `tests/mcp/drive.js`
(its larger arguments in `tests/mcp/args/`), and checks the grid.
