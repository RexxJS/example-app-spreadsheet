# LLM Primer: RexxJS Spreadsheet Application

## Quick Overview
Hybrid web/desktop spreadsheet app using React + Tauri. Uses **RexxJS (REXX interpreter)** for formulas instead of traditional Excel syntax.

## Architecture
- **Frontend**: React 19 + Vite (src/)
- **Desktop Backend**: Rust + Tauri 2.9 (src-tauri/)
- **Formula Engine**: RexxJS (external dependency, loaded from CDN or bundled)
- **Deployment**: Static web app OR native desktop app (same codebase)

## Key Source Files
```
src/
├── spreadsheet-model.js              # Core logic (pure JS, DOM-free, testable)
├── spreadsheet-rexx-adapter.js       # RexxJS integration layer
├── SpreadsheetApp.jsx                # React UI components
├── spreadsheet-control-functions.js  # Control Bus commands (automation API)
└── spreadsheet-import-export.js      # CSV/JSON/TOML/YAML support

src-tauri/src/
└── lib.rs                            # HTTP Control Bus API, file handling
```

## Development Commands
```bash
npm run dev              # Web mode (localhost:5173)
npm run tauri:dev        # Desktop mode
npm test                 # Jest + Playwright tests
npm run build            # Production web build
npm run tauri:build      # Desktop app build
```

## Critical Concepts

### 1. RexxJS Formula Language
- Formulas use REXX syntax, not Excel: `=SQRT(A1*A1 + B1*B1)`
- Function pipelines supported: `=A1:A10 |> SUM() |> ROUND(2)`
- Setup Scripts for globals/imports
- Custom functions in `public/lib/spreadsheet-functions.js`

### 2. Model-View Separation
- `spreadsheet-model.js` has ZERO React/DOM dependencies
- All business logic testable with Jest (no browser needed)
- UI components in `SpreadsheetApp.jsx` call model methods

### 3. Control Bus (ARexx-inspired)
- **Web mode**: iframe postMessage
- **Desktop mode**: HTTP API (port 2410, long-polling)
- Allows external REXX scripts to automate spreadsheet
- 26 commands: SET_CELL, GET_CELL, LOAD_FILE, etc.

### 4. Multi-Sheet Architecture
- Tabbed interface, independent cell data per sheet
- Sheet state in `spreadsheet-model.js` (sheets array)

### 5. Advanced Features
- Conditional formatting (RexxJS expressions)
- Data validation, pivot tables, merged cells
- Named ranges, charts
- Optional PyOdide for NumPy/SciPy (Python, not JS polyfills)

## Testing Strategy
- **Unit tests**: `tests/*.spec.js` (Jest, model logic)
- **E2E tests**: `tests/*-web.spec.js` (Playwright, browser UI)
- **Integration**: `tests/*.rexx` (REXX scripts via Control Bus)

## Common Pitfalls
1. **RexxJS dependency**: Must exist at `../../core/src/repl/dist/rexxjs.bundle.js` for production builds
2. **Cargo.lock**: Should be committed (application, not library)
3. **Formula syntax**: REXX, not Excel (different string concat, array indexing)
4. **Async formulas**: RexxJS formulas are async (affects calculation order)

## Where to Start
1. Read `README.md` for user-facing features
2. Explore `src/spreadsheet-model.js` for core logic
3. Check `tests/*.spec.js` for usage examples
4. See `docs/FEATURES.md` for visual feature guide

## License
- Application: AGPL-3.0
- Control Protocol: MIT
