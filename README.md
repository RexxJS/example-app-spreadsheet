# RexxSheet: a spreadsheet you can script in REXX

A spreadsheet with two formula languages, Excel's and RexxJS's, and every
action a command that a REXX script can drive. This repository has three
parts:

| Folder | What it is | License |
|---|---|---|
| [`app/`](app/) | **The spreadsheet, current.** A [sae](https://github.com/aether-lang-dev/sae-it-aint-so) app in sae's TypeScript dialect: it runs as a page in the sae browser or as an installed app, natively on aether-ui (GTK4, AppKit, Win32). | MIT |
| [`rexx/`](rexx/) | REXX scripts that drive a running spreadsheet from outside, over its control bus. | MIT |
| [`tauri/`](tauri/) | **The original, legacy.** The React + Tauri app the sae one grew from. New features go into `app/`; this stays as it was, for reference. | AGPL (see [LICENSING.md](LICENSING.md)) |

## The sae app

`app/` began in sae as its demo 18 and moved here on 2026-10-14; its
history up to then is in sae's git log (to sae commit `61e4adb`). What it
does and what is still to do: [docs/spreadsheet_demo_done_and_todo.md](docs/spreadsheet_demo_done_and_todo.md).
The bugs it found in the Tauri app: [docs/tauri-spreadsheet-upstream-notes.md](docs/tauri-spreadsheet-upstream-notes.md).

```
app/
  app.json          the installed app's manifest: its start page and grants
                    (documents, the control bus)
  sheet.ts          the page: panels declared over the commands
  sheet/            the spreadsheet with no ui: workbook, sheets, both formula
                    languages, every command (commands.ts), RexxJS (rexxjs.js,
                    generated)
  sheet/view/       the ui: each panel reads the workbook and runs commands
tests/
  headless/         the model and commands, no window: sae's lowerer harness
                    on sae's QuickJS (sheet -> ../../app/sheet)
  specs/            the app driven in sae-driver: grid, keys, clipboard, files,
                    the control bus
  lib/griddriver.ae reads and clicks the datagrid for the specs
vendor/rexxjs/      RexxJS's interpreter and parser, vendored (MIT, its LICENSE)
tools/              bundle-rexxjs.mjs: re-vendors RexxJS and writes app/sheet/rexxjs.js
docs/               the plan, and notes for the Tauri app
sae -> ../sae       the sae checkout it builds and tests against
```

### Running it

It needs sae checked out beside this repository (the `sae` symlink points
at `../sae`) and built, as sae's AGENTS.md "Build and test" says. Then:

```bash
../sae/target/build/bin/sae --app app
```

`--control-bus` after it switches on the control bus, so other programs on
this machine can drive the sheet (sae's `docs/control-bus.md`).

### Testing it

```bash
./run-tests.sh
```

That runs the headless tests and the three specs, one line of PASS or FAIL
each, with logs under `target/run-tests/`. `./run-tests.sh headless` runs
only the headless tests; `./run-tests.sh spec_sheet` only that spec.
`SAE_DIR=...` picks another sae checkout. On Linux with no display it runs
under `xvfb-run`.

### Re-vendoring RexxJS

```bash
node tools/bundle-rexxjs.mjs ../RexxJS
```

copies RexxJS's interpreter and parser into `vendor/rexxjs/` (recording the
commit in `UPSTREAM`) and regenerates `app/sheet/rexxjs.js`. RexxJS is the
canonical source of the REXX language; changes to it go there first.

## REXX scripts

See [rexx/README.md](rexx/README.md).

## The Tauri app

See [tauri/README.md](tauri/README.md). It was written to live inside the
RexxJS repository (its scripts reach `../../core`), and does not build
from here as it is.
