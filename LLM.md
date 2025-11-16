# LLM Primer: RexxJS Spreadsheet Application

## Deployment Modes (3 Corners)
1. **Tauri Desktop** - Native app (Mac/Windows/Linux) with Rust backend
2. **Standalone Web** - Static HTML/JS served from any web server
3. **iframe/RexxJS-Controlled** - Embedded in RexxJS Director/Worker architecture

All modes share the same React codebase; control interfaces differ.

## Architecture Stack
- **Frontend**: React 19 + Vite
- **Desktop Backend**: Rust + Tauri 2.9 (file I/O, HTTP control bus)
- **Formula Engine**: RexxJS interpreter (REXX language, not Excel)
- **Integration**: Variable resolver for lazy cell references, postMessage/HTTP APIs

## Key Source Files
```
src/
├── spreadsheet-model.js              # Pure JS core (DOM-free, testable)
├── spreadsheet-rexx-adapter.js       # RexxJS integration (variable resolver)
├── SpreadsheetApp.jsx                # React UI
├── spreadsheet-control-functions.js  # Control protocol (MIT licensed)
└── spreadsheet-import-export.js      # CSV/JSON/TOML/YAML

spreadsheet-controlbus.js             # postMessage handler (iframe mode)
src-tauri/src/lib.rs                  # HTTP API (Tauri mode, port 2410)
```

## Development
```bash
npm run dev          # Web mode (localhost:5173)
npm run tauri:dev    # Desktop mode
npm test             # Jest + Playwright
```

## Critical Concepts

### 1. RexxJS Formula Language
- **REXX syntax**, not Excel: `=SQRT(A1*A1 + B1*B1)`
- **Pipelines**: `=A1:A10 |> SUM_RANGE('A1:A10') |> ROUND(2)`
- **Variable resolver**: Cell refs (A1, Sheet2_A1) resolved on-demand, not pre-injected
- **Setup scripts**: Page-level REXX code for globals (TAX_RATE, REQUIRE'd functions)
- **Async evaluation**: All formulas async (affects calc order)

### 2. Control Bus (ARexx-inspired, MIT Protocol)
**Purpose**: External automation from REXX scripts or other apps

**Modes**:
- **Web/iframe**: `postMessage` with `type: 'spreadsheet-control'`
- **Desktop**: HTTP long-polling on `http://localhost:2410/command`
- **Commands**: 60+ (SETCELL, GETCELL, MERGECELLS, CREATEPIVOT, etc.)

**Usage in iframe**:
```js
parent.postMessage({
  type: 'spreadsheet-control',
  command: 'setCell',
  params: {cellRef: 'A1', content: '=B1+B2'},
  requestId: 123
}, '*');
```

**RexxJS Director/Worker Pattern**:
- RexxJS runs spreadsheet in worker iframe
- Director iframe sends commands via postMessage
- Enables pause/resume, progress monitoring, distributed workflows
- APPLICATION ADDRESSING for secure iframe integration

### 3. Model-View Separation
- `spreadsheet-model.js`: Zero React/DOM deps (pure business logic)
- Multi-sheet state (Map of sheets), dependency graph, undo/redo stack
- Testable via Jest without browser
- UI triggers CustomEvents (`spreadsheet-update`) for re-renders

### 4. Multi-Sheet + Advanced Features
- **Sheets**: Tabbed UI, cross-sheet refs (`Sheet2_A1` in REXX)
- **Merged cells**: Range metadata, top-left anchor
- **Conditional formatting**: RexxJS expressions evaluated per cell
- **Data validation**: List, number, text patterns
- **Pivot tables**: JSON config, dynamic aggregation
- **Named ranges**: Symbolic refs (SalesData → A1:B10)
- **PyOdide**: Optional NumPy/SciPy (PY_LINREGRESS, PY_FFT, PY_SOLVE)

### 5. File Format
- **Extensions**: `.json`, `.rexxsheet` (preferred for file associations)
- **Imports**: CSV, JSON, TOML, YAML (converted to internal format)
- **Structure**: `{version, sheets: [{name, cells, setupScript}], ...}`

## Testing
- **Unit**: `tests/*.spec.js` (Jest, model logic)
- **E2E**: `tests/*-web.spec.js` (Playwright)
- **Control Bus**: `tests/*.rexx` (REXX automation scripts)

## Pitfalls
1. **REXX != Excel**: Different syntax (concat, indexing, functions)
2. **Cell refs**: Use variable resolver (no pre-injection), Sheet2_A1 not Sheet2.A1
3. **Async**: All evals async, circular deps detected
4. **RexxJS path**: Must be at `../../core/src/repl/dist/rexxjs.bundle.js` for builds

## RexxJS Integration (from reference docs)
- **Autonomous Web Mode**: Standalone browser execution
- **Controlled Web Mode**: Director/Worker iframes + postMessage
- **Variable Resolver**: Lazy resolution without caching (ideal for spreadsheets)
- **Control Bus**: Checkpoints for pause/resume/abort
- **ADDRESS handlers**: SQLite, System, APPLICATION (iframe integration)
- **Excel-compat functions**: VLOOKUP, stats, financial (via REQUIRE)

## License
- App code: AGPL-3.0
- Control protocol (spreadsheet-control-functions.js): MIT (encourages ecosystem)
