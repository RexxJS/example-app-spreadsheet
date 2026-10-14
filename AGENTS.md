# Notes for agents

Start with [README.md](README.md): what the three folders are, and how to
run and test the sae app. Then [docs/spreadsheet_demo_done_and_todo.md](docs/spreadsheet_demo_done_and_todo.md),
the plan: what is done, what is next, and how the code is laid out.

- **`app/` is where the work is.** `tauri/` is legacy and kept as it was;
  read it for what the original did, don't develop it. `rexx/` holds the
  control-protocol scripts.
- **It runs on sae, beside this checkout** (`sae -> ../sae`). A feature
  that needs something new of sae (a page service, a grant, an aether-ui
  widget) is built in sae, with sae's own test of it, and used here after;
  `run-tests.sh` runs sae's harnesses on this repo's tests. When sae
  changes a page API this app uses, run `./run-tests.sh` here before
  pushing sae.
- **RexxJS is canonical for REXX.** `vendor/rexxjs/` is a copy; fix RexxJS
  itself (`../RexxJS`, with its tests), then re-vendor with
  `node tools/bundle-rexxjs.mjs ../RexxJS`. No sae or app code goes into
  RexxJS.
- **Tests:** `./run-tests.sh`, all green before a push. Every command is
  tested headless (`tests/headless/`, one file per area, `// expect:` lines
  as verdicts); what only the ui shows is tested in `tests/specs/`.
- **Commits** go straight to `main`. Docs-only commits carry
  `[skip actions]`.
