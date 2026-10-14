# REXX scripts that drive the spreadsheet

Scripts that control a running spreadsheet from another program, ARexx
style: `ADDRESS SPREADSHEET`, then `SETCELL`, `GETCELL` and the rest of the
control protocol (MIT; see [../LICENSING.md](../LICENSING.md)).

| Script | What it does |
|---|---|
| `test-connection.rexx` | Checks the spreadsheet answers, and says what to do when it does not |
| `test-spreadsheet-address.rexx` | The ARexx-style demonstration: sets cells, formulas and formats, reads them back |
| `spreadsheet-functions-test.rexx` | The range functions (SUM_RANGE, MEDIAN_RANGE, SUMIF_RANGE ...), checked with RexxJS's expectations |
| `spreadsheet-control-bus-test.rexx` | The control commands: cells, metadata, references, ranges, sheets, setup scripts, LISTCOMMANDS |
| `spreadsheet-row-column-ops.rexx` | Rows and columns inserted and deleted, and bad ones refused |
| `spreadsheet-address.js` | The ADDRESS library: `ADDRESS SPREADSHEET` for the sae app |

## Running them

From the repository's root, with sae built (including `target/sae-bus`)
and RexxJS checked out beside this repository (the `RexxJS` link):

```bash
../sae/target/build/bin/sae --app app --control-bus
```

then, in another terminal:

```bash
RexxJS/core/rexx rexx/test-spreadsheet-address.rexx
```

`./run-tests.sh spec_app_sheet_rexx` runs all five against the app and
checks what they say and what the grid then shows.

## How a command reaches the app

```rexx
REQUIRE "./spreadsheet-address.js"
ADDRESS SPREADSHEET
'SETCELL("A1", "100")'
'GETCELL("A1")'
SAY RC RESULT
```

The app's control bus is sae's (`../sae/docs/control-bus.md`): off unless
the app is started with `--control-bus`; a port chosen at random; a key
fresh for every launch, in a courier file only this user can read; every
request and answer sealed with it. A script cannot know the port or key
in advance, and an unsealed request is refused.

So `spreadsheet-address.js` does none of that itself. For each command it
runs sae's `sae-bus Spreadsheet <command>`, which finds the courier file,
seals the command, sends it, and prints the opened answer. RC is 0 and
RESULT what the command came to when it worked. Otherwise RC is 1 and
ERRORTEXT says why. RC is 2 when there was no app to ask (not running, or
the bus off), and ERRORTEXT then has sae-bus's reason. `SAE_BUS`,
`SPREADSHEET_APP` and `SAE_APPDATA_DIR` point it elsewhere (the library's
header says how).

## Where they differ from the Tauri app's

These scripts were written for the Tauri app (`../tauri/`), whose bus
was plain HTTP at `http://localhost:2410` with a fixed token. What changed
to drive the sae app:

- **Connecting.** `ADDRESS "http://localhost:2410/api/spreadsheet" AUTH
  "dev-token-12345" AS SPREADSHEET` became `REQUIRE
  "./spreadsheet-address.js"` and `ADDRESS SPREADSHEET`.
  `test-connection.rexx` used to POST JSON itself with `HTTP_POST`; it now
  asks `SPREADSHEET_VERSION()` the same way.
- **Formula language.** The Tauri app's sheets were RexxJS formulas only.
  The sae app also has Excel's, and starts with it. So each script first
  sends `SETFORMULALANGUAGE("RexxJS")`, a control function the sae app adds
  (with `GETFORMULALANGUAGE()`).
- **Values in the command.** A command runs in the app, which cannot see
  the script's variables. `'SETCELLS("E1:E3", values)'` became
  `'SETCELLS("E1:E3", "[100, 200, 300]")'`.
- **Numbers.** The Tauri app gave every value as a string; the sae app
  gives a number as a number. REXX makes no difference between them, but
  RexxJS's `should be string` does, so those checks became `should equal`.
- **LISTCOMMANDS** lists 130 names, not the Tauri app's 26 (all 26 among
  them).
- **`spreadsheet-row-column-ops.rexx`** only printed what each test should
  do; it now runs those cases against the sheet.
- The Tauri app's `TestImportExport` subroutine was never called, and
  still is not: it hands the app a variable of the script's
  (`IMPORT(exported)`).

## RexxJS bugs met on the way, fixed there

Found driving these scripts, and fixed in RexxJS (2026-10-14, `e987768`
and `6ec1ab7`), which is canonical for REXX:

- `rexx` (the CLI) exited 0 after `EXIT 1`; it now exits with EXIT's code.
  `spec_app_sheet_rexx` checks both the status and each script's own
  "passed" line.
- Inside a string, a doubled quote of its own kind stayed as two quotes;
  it is now one, as in classic REXX (`'it''s'` is it's).
- `SAY "(RC" RC || ")"` printed its source text; each side of `||` is now
  its terms joined by blanks.

`vendor/rexxjs/` was re-vendored from that commit, so REXX inside the app
reads strings the same way.
