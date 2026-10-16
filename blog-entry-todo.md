# To do: a new take on spreadsheets, from Paul's 2012 blog entry

Paul's [It's time for a new take on spreadsheets](https://paulhammant.com/2012/08/10/its-time-for-a-new-take-on-spreadsheets/)
(2012-08-10), reread on 2026-10-16. Its argument:
- the grid mixes model and view;
- a cell's styling is hidden behind menus;
- `.xlsx` is zipped XML nobody can read or diff;
- exports lose the formulas.

It proposed:
- a terse model with formulas in it, apart from the view;
- text kept in source control, so a diff says who changed what;
- an edit mode over a sheet that is otherwise read.

This file is the plan that follows from it, kept apart from `TODO.md`,
whose steps are GridCraft parity. Each step here follows the same
rules (`docs/gridcraft_parity.md`): one commit with its tests and the
control function it gains, described, so scripts, the bus and the MCP
server reach it. A step done is struck through, with its date.

## What happened since 2012

- **Terse lambda formulas.** Excel added `LAMBDA`, `LET`, `MAP` and
  `REDUCE` in 2020. This app has them too, optional parameters
  included (`TODO.md` 2.12).
- **Collaboration.** Google Sheets won it by syncing live, not through
  source control. "Who changed what" arrived as version history inside
  the product, not as git diffs.
- **The file format.** `.xlsx` is still zipped XML, and no business
  user diffs a spreadsheet. That critique stands untouched.

## Why it reads differently now

There is a new kind of user, the AI model. The model/view split, a
terse readable source form and an audit trail of changes are what a
model needs, more than a person does.

- **Models read the model, not the grid.** The MCP server already works
  this way: a model drives the workbook through commands, never the
  screen.
- **Reviewing an AI's edits is a diff problem.** "What did the AI change
  in my sheet?" is the question a diff answers. One MCP request is
  already one undo step. The next move is one request, one reviewable
  diff.
- **The Excel Depression link is sharper.** Mistakes like
  Reinhart-Rogoff's get easier to make when an AI writes the formulas.
  A readable form, diffs and tests kept in the sheet are a credible
  answer.

## The steps

| # | Step | Tests prove | REXX |
|---|---|---|---|
| B.1 | A canonical text format that diffs well: one cell per line, in a stable order (sheet, then row, then column); formulas as typed; values only where there is no formula; a look listed only where it differs from the default; names, tables, rules and validation in their own sections, sorted. Read and written beside the JSON and `.rexxsheet` formats, through sae's `documents`. | headless: round-trips every sample workbook exactly; writing twice gives the same bytes; changing one cell changes one line | `EXPORTTEXT()`, `IMPORTTEXT(text)` |
| B.2 | A workbook as a REXX program: written as `SETCELL`, `SETFORMULA`, look and name calls that rebuild it when run, as chaess's recorder writes a game as the script that replays it. A source form that is both readable and runnable. | headless: write it, run it in a new workbook, B.1's text of the two is identical | `WORKBOOKSCRIPT()` |
| B.3 | Everything non-default about a cell: its value, formula, number format, every look that is not the default, rules that apply to it, validation, note, name, table, and what it reads and what reads it. A panel, and a REXX function. | headless: a styled sample cell lists exactly its non-defaults; spec: the panel | `CELLINFO("B3")` gives a stem |
| B.4 | Git-backed saving, optional: a workbook saved as B.1's text in a git repository, each save a commit with a message (the command's label, or the MCP request's), and a diff view between any two versions, cell by cell (old and new value or formula), before a change is kept. | headless: two saves, the diff names the changed cells; spec: the diff view | `SAVEVERSION(message)`, `VERSIONS()`, `DIFFVERSIONS(a, b)` |
| B.5 | One MCP request, one reviewable diff: the server can hold a request's changes for review (the diff of B.4 against the state before it), and keep or take them back whole. | tests/mcp: a request held, its diff read, taken back; then kept | `spreadsheet_review` in the MCP server |
| B.6 | Assertions in the sheet: checks kept in the workbook (a formula that must be TRUE, with a message), run on every recalculation, shown in the status bar and a panel when one fails, and saved with the file. | headless: a total that stops adding up fails its check, by name; spec: the status bar says so | `ADDCHECK(name, formula, message)`, `CHECKS()` gives the failing ones |
| B.7 | A read view with an edit mode: the workbook shown read-only (no editor, no selection chrome) until edit mode is switched on, as the 2012 entry's "hotkey into edit mode" had it. | spec: typing in read mode changes nothing; edit mode on, it does | `SETEDITMODE(1)` |

The order is a proposal. B.1 comes first because B.2, B.4 and B.5
build on it. B.3 and B.6 stand alone and could go earlier.
