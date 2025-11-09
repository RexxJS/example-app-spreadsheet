# RexxJS Spreadsheet POC

A proof-of-concept spreadsheet powered by RexxJS expressions, built with React.

**Deployment Modes:**
- **Web**: Static HTML/JS that can be served from any web server
- **Desktop**: Native application for Mac, Windows, and Linux via Tauri

## Quick Start

### Development Mode (Tauri)

**Launch with sample data:**
```bash
./rexxsheet-dev
```

**Load a specific spreadsheet:**
```bash
./rexxsheet-dev test-sheet.json
./rexxsheet-dev sample-budget.json
./rexxsheet-dev /path/to/your-spreadsheet.json
```

**Show help:**
```bash
./rexxsheet-dev --help
```

### Production Binary (Static Build)

**Quick start with pre-built binary:**
```bash
./rexxsheet-static
./rexxsheet-static spreadsheet.json
./rexxsheet-static --control-bus  # Enable HTTP control bus
```

The static binary (`./rexxsheet-static`) is a symlink to `src-tauri/target/release/rexxsheet` and includes all dependencies.

See **[BUILDING-BINARY.md](BUILDING-BINARY.md)** for instructions on building and distributing a standalone binary.

### Sample Files

- **`test-sheet.json`** - Simple test spreadsheet with basic formulas
  - Numbers, text, SUM_RANGE, UPPER functions
  - Good for quick testing and learning
  - Matches the default programmatic sample data

- **`sample-budget.json`** - Budget spreadsheet with tax calculations
  - Demonstrates setupScript with global variables
  - Cell dependencies and complex formulas
  - Uses SUM_RANGE for totals

### File Format

See **[FILE-LOADING.md](FILE-LOADING.md)** for complete file format documentation and loading options.

## Features

- **Cells with Values or Expressions**: Enter literal values or formulas starting with `=`
- **A1-Style Cell References**: Reference other cells using standard notation (A1, B2, etc.)
- **RexxJS Expression Evaluation**: Full RexxJS language support in formulas
- **Function Pipelines**: Use the `|>` operator for data transformation chains
- **Range Functions**: Built-in functions for working with cell ranges
- **Dependency Tracking**: Automatic recalculation when referenced cells change
- **Cell Comments & Formats**: Attach metadata to cells for documentation
- **Named Variables**: Define constants in Setup Script (e.g., `LET TAX_RATE = 0.07`)
- **Enhanced Info Panel**: Shows cell details, dependencies, type, comments
- **View Mode Hotkeys**: Press V/E/F/N to toggle between different views
- **Basic Styling**: Clean, modern UI with visual feedback
- **📁 Load from File/URL**: Load spreadsheet data from JSON files (web mode via hash parameter)
- **🔌 Control Bus**: Remote control via ARexx-inspired cross-application scripting (web mode: iframe postMessage, Tauri mode: HTTP API)
- **🔢 NumPy Integration**: Real Python NumPy via PyOdide for 100% accurate scientific computing
- **Row/Column Operations**: Insert and delete rows/columns with automatic cell shifting

## Deployment

### Web Deployment

The spreadsheet can run as a web application using Vite's dev server.

1. **Build the RexxJS bundle** (if not already built):
   ```bash
   cd core/src/repl
   npm install
   npm run build
   # This creates core/src/repl/dist/rexxjs.bundle.js
   ```

2. **Copy RexxJS bundle to public directory**:
   ```bash
   cd examples/spreadsheet-poc
   cp ../../core/src/repl/dist/rexxjs.bundle.js public/
   ```

3. **Start Vite dev server**:
   ```bash
   npm run dev:vite
   # Opens on http://localhost:5173/
   ```

4. **Load a spreadsheet** (optional):
   ```
   # Load sample budget
   http://localhost:5173/#load=sample-budget.json

   # Load test sheet
   http://localhost:5173/#load=test-sheet.json
   ```

**Alternative: Static HTTP Server**

You can also serve the built files via any HTTP server:
```bash
# Build for production
npm run build

# Serve the dist directory
npx http-server dist -p 8082 -c-1
```

### Desktop Deployment (Tauri)

The spreadsheet can also run as a native desktop application on Mac, Windows, and Linux.

1. **Build the RexxJS bundle** (if not already built):
   ```bash
   cd core/src/repl
   npm install
   npm run build
   ```

2. **Navigate to spreadsheet directory**:
   ```bash
   cd examples/spreadsheet-poc
   npm install
   ```

3. **Run in development mode**:
   ```bash
   # Using the convenience script (recommended)
   ./rexxsheet-dev

   # Or with a specific file
   ./rexxsheet-dev sample-budget.json

   # Or via npm directly
   npm run tauri:dev
   ```

4. **Build desktop application**:
   ```bash
   npm run tauri:build
   ```

   This creates platform-specific installers in `src-tauri/target/release/bundle/`:
   - **Mac**: `.dmg` and `.app`
   - **Windows**: `.msi` and `.exe`
   - **Linux**: `.deb`, `.appimage`

### Choosing Between Web and Desktop

**Use Web Deployment When:**
- You want to share via URL
- No installation required
- Cross-platform without builds
- Easy to deploy to static hosting

**Use Desktop Deployment When:**
- You want a standalone application
- Better integration with OS (file system, etc.)
- Offline usage is important
- Professional application feel

## Usage

### Loading Spreadsheet Data

#### Web Mode - Load from URL

Use the hash parameter to load spreadsheet data from a JSON file:

```bash
# Load from relative path
http://localhost:8082/examples/spreadsheet-poc/index.html#load=sample-budget.json

# Load from absolute path
http://localhost:8082/examples/spreadsheet-poc/index.html#load=/data/mysheet.json

# Load from HTTP(S) URL
http://localhost:8082/examples/spreadsheet-poc/index.html#load=https://example.com/sheet.json
```

**Spreadsheet JSON Format:**
```json
{
  "name": "My Spreadsheet",
  "version": "1.0",
  "setupScript": "LET TAX_RATE = 0.08",
  "cells": {
    "A1": { "value": "Item", "expression": null, "format": null, "comment": "Column header" },
    "B2": { "value": "1200", "expression": null, "format": "$0.00", "comment": null },
    "C2": { "value": "96", "expression": "B2 * TAX_RATE", "format": null, "comment": null }
  },
  "metadata": {
    "rows": 100,
    "cols": 26,
    "created": "2025-11-05T18:00:00.000Z"
  }
}
```

See `sample-budget.json` for a complete example.

#### Tauri Mode - Load from Filesystem

In Tauri desktop mode, pass the file path as a command-line argument:

```bash
# Launch with file path (development)
./rexxsheet-dev /path/to/mysheet.json

# Or via npm directly
npm run tauri:dev:file /path/to/mysheet.json

# Without file argument (loads sample data)
npm run tauri:dev
```

### Control Bus - Remote Scripting

The spreadsheet supports remote control via an ARexx-inspired control bus, allowing other applications or Rexx scripts to manipulate the spreadsheet programmatically.

#### Web Mode - iframe Communication

Open the control bus demo to see it in action:

```
http://localhost:8082/examples/spreadsheet-poc/spreadsheet-controlbus-demo.html
```

**Example Rexx Script (controlling spreadsheet from another iframe):**

```rexx
-- Control the spreadsheet remotely
ADDRESS spreadsheet

-- Set cell values
"setCell" ref="A1" content="100"
"setCell" ref="A2" content="200"
"setCell" ref="A3" content="=A1+A2"

-- Get results
LET result = "getCellValue" ref="A3"
SAY "Sum: " || result.value  -- Output: Sum: 300

-- Export data
LET data = "export" name="MySheet"
SAY "Exported " || data.name
```

**Available Control Bus Commands:**

**All Control Bus Commands (26 total):**

**Basic Cell Operations:**
1. `SETCELL(cellRef, content)` - Set cell value or formula
2. `GETCELL(cellRef)` - Get cell value
3. `GETEXPRESSION(cellRef)` - Get cell formula/expression
4. `CLEARCELL(cellRef)` - Clear cell content

**Cell Metadata:**
5. `SETFORMAT(cellRef, format)` - Set cell formatting
6. `GETFORMAT(cellRef)` - Get cell format
7. `SETCOMMENT(cellRef, comment)` - Set cell comment
8. `GETCOMMENT(cellRef)` - Get cell comment

**Cell References:**
9. `GETROW(cellRef)` - Get row number from cell reference
10. `GETCOL(cellRef)` - Get column number from cell reference
11. `GETCOLNAME(cellRef)` - Get column letter from cell reference
12. `MAKECELLREF(col, row)` - Create cell reference from column and row

**Range Operations:**
13. `GETCELLS(rangeRef)` - Get multiple cells by range (e.g., "A1:B5")
14. `SETCELLS(rangeRef, values)` - Set multiple cells at once

**Spreadsheet Management:**
15. `CLEAR()` - Clear all cells in spreadsheet
16. `EXPORT()` - Export spreadsheet to JSON
17. `IMPORT(jsonData)` - Import spreadsheet from JSON
18. `GETSHEETNAME()` - Get spreadsheet name
19. `SETSHEETNAME(name)` - Set spreadsheet name

**Expression Evaluation:**
20. `EVALUATE(expression)` - Evaluate RexxJS expression in spreadsheet context
21. `RECALCULATE()` - Force recalculation of all formulas

**Setup Script Management:**
22. `GETSETUPSCRIPT()` - Get setup script
23. `SETSETUPSCRIPT(script)` - Set setup script
24. `EXECUTESETUPSCRIPT()` - Execute setup script

**Utility:**
25. `LISTCOMMANDS()` - Get list of available commands
26. `SPREADSHEET_VERSION()` - Get version

**Testing:** See `examples/spreadsheet-poc/tests/spreadsheet-control-bus-test.rexx` for comprehensive tests of all commands.

#### Tauri Mode - HTTP API (ADDRESS Remote)

The Tauri spreadsheet can be controlled via HTTP using RexxJS's built-in ADDRESS facility:

```bash
# Terminal 1: Start spreadsheet with control bus
./rexxsheet-dev --control-bus

# Terminal 2: Control it with REXX
../../core/rexx test-spreadsheet-address.rexx
```

**Example Control Script:**
```rexx
-- Register the remote spreadsheet endpoint (automatically switches to it)
ADDRESS "http://localhost:2410/api/spreadsheet" AUTH "dev-token-12345" AS SPREADSHEET

-- Send commands (already in SPREADSHEET context)
'SETCELL("A1", "100")'
'SETCELL("A2", "200")'
'SETCELL("A3", "=A1+A2")'

IF RC = 0 THEN
  SAY "Spreadsheet updated successfully"
```

**Key Features:**
- Built-in to RexxJS core (no external libraries needed)
- Bearer token authentication for security (`Authorization: Bearer <token>`)
- Classic ARexx-style inter-process communication
- Works from command-line, automation scripts, or CI/CD pipelines
- Automatic ADDRESS context switching after registration

See **[TESTING-CONTROL-BUS.md](TESTING-CONTROL-BUS.md)** for complete documentation.

This design is inspired by **ARexx** from the Amiga, which allowed applications like DPaint, PageStream, and Directory Opus to be controlled by external scripts, enabling powerful workflow automation and inter-application communication.

### Keyboard Shortcuts (Hotkeys)

The spreadsheet supports view mode switching via keyboard:

- **V** - Values view: Show only literal values (formulas show blank)
- **E** - Expressions view: Show only formulas (value cells hidden)
- **F** - Formats view: Show format strings only
- **N** - Normal view: Standard spreadsheet view (default)

Press these keys while the grid is active (not editing a cell) to switch views.

**Use Cases:**
- Press **V** to audit which cells are hard-coded values
- Press **E** to review all formulas at once
- Press **F** to see formatting specifications
- Press **N** to return to normal operation

### Named Variables

Define constants and reusable values in the **⚙️ Setup** script:

```rexx
/* Setup Script */
LET TAX_RATE = 0.07
LET SHIPPING_COST = 15.00
LET DAYS_IN_YEAR = 365
```

Then use in any cell:
```rexx
=Revenue * TAX_RATE
=Subtotal + SHIPPING_COST
=Hours / DAYS_IN_YEAR
```

### Cell Comments & Formats

Cells can have metadata attached:
- **Comments**: Documentation notes (shown as 💬 icon in cell)
- **Formats**: Display formatting specs (shown as green border)

Access these via the Info Panel when a cell is selected, or hover over cells to see tooltips.

### Sheet Name via Hash Parameter

The sheet name is specified via the URL hash:
- `#Sheet1` - Default sheet
- `#Budget2024` - Custom sheet name
- If no hash is provided, defaults to `Sheet1`

## Formula Examples

### Basic Arithmetic
```
=A1 + B1
=A2 * 2
=(A1 + A2) / 2
```

### Cell References
```
=A1 + A2 + A3
=B1 * C1
=SUM_RANGE("A1:A5")
```

### RexxJS String Functions
```
=UPPER("hello")
=SUBSTR("Hello World", 1, 5)
=LENGTH("test")
=REVERSE("stressed")
```

### Function Pipelines
```
="hello" |> UPPER |> LENGTH
=A1 |> UPPER |> SUBSTR(_, 1, 3)
="  trim me  " |> STRIP |> LENGTH
```

### Range Functions
```
=SUM_RANGE("A1:A5")
=AVERAGE_RANGE("B1:B10")
=MAX_RANGE("C1:C20")
=MIN_RANGE("D1:D15")
=COUNT_RANGE("E1:E10")
```

### Complex Expressions
```
=UPPER(A1) || " " || LOWER(B1)
=(A1 + A2) |> MULTIPLY(_, 2)
=LENGTH(A1) > 5
```

## Architecture

### Components

1. **spreadsheet-model.js**
   - Pure JavaScript spreadsheet model
   - Cell storage with A1-style addressing
   - Dependency tracking and recalculation
   - Testable with Jest (no DOM dependencies)

2. **spreadsheet-rexx-adapter.js**
   - Integration layer between SpreadsheetModel and RexxJS
   - Provides cell reference functions (A1(), B2(), etc.)
   - Range functions (SUM_RANGE, AVERAGE_RANGE, etc.)
   - Expression evaluation context
   - Setup script execution

3. **spreadsheet-loader.js**
   - Load/save spreadsheet data from JSON files
   - Supports both Tauri filesystem and web URLs
   - Import/export model data
   - Validate spreadsheet format

4. **SpreadsheetApp.jsx** (React components)
   - `App` - Main application with state management
   - `Grid` - Spreadsheet grid with row/column headers
   - `Cell` - Individual cell with edit/view modes
   - `FormulaBar` - Cell formula/value editor
   - `InfoPanel` - Cell details and help
   - `SettingsModal` - Setup script editor

5. **main.jsx** (Vite entry point)
   - Waits for RexxJS to load
   - Initializes React app
   - Handles Tauri file events

6. **index.html**
   - Loads RexxJS bundle
   - Mounts React via Vite
   - Entry point for both web and Tauri modes

7. **lib.rs** (Tauri backend)
   - CLI argument handling
   - File path event emission
   - Filesystem plugin initialization

## Testing

### Jest Tests (Model)
```bash
cd core
npm test -- --testPathPattern="spreadsheet"
```

### Playwright Tests (UI)
```bash
PLAYWRIGHT_HTML_OPEN=never npx playwright test examples/spreadsheet-poc/tests/
```

## Limitations & Future Enhancements

### Current Limitations
- No save functionality (load is supported via JSON files)
- No undo/redo
- No paste (copy is supported with Ctrl+C)
- Basic styling only
- Limited to 100 rows × 26 columns (configurable in code)
- No backend integration

### Implemented Features
- ✅ Load from JSON files (web mode: hash parameter, Tauri mode: CLI argument)
- ✅ Cell comments and formats
- ✅ Keyboard navigation (arrow keys, Tab, Enter)
- ✅ Cell selection ranges (click and drag, or Shift+Click)
- ✅ Copy selection to clipboard (Ctrl+C)
- ✅ View mode hotkeys (V/E/F for values/expressions/formats)
- ✅ Named variables via Setup Script
- ✅ Control Bus for remote scripting (web mode: iframe postMessage, Tauri mode: HTTP API on port 2410)
- ✅ Static binary build for production deployment (`./rexxsheet-static`)

### Potential Enhancements
- Save/export to JSON file (complement the load feature)
- Undo/redo functionality
- Paste from clipboard
- Add more Excel-like functions (IF, VLOOKUP via extras/functions/excel)
- Visual cell formatting (colors, fonts, borders)
- Multi-sheet support (tabs)
- Import/export CSV format
- Collaborative editing
- Charts and visualizations

## RexxJS Functions Available

The spreadsheet has access to three levels of functions:

### 1. Built-in RexxJS Functions (from `core/src/*`)

Always available without any setup:

**String Functions:**
- `UPPER`, `LOWER`, `LENGTH`, `SUBSTR`, `STRIP`, `REVERSE`
- `INDEX`, `POS`, `LASTPOS`, `WORDPOS`
- `LEFT`, `RIGHT`, `CENTER`
- `TRANSLATE`, `COPIES`, `SPACE`

**Math Functions:**
- `ABS`, `CEIL`, `FLOOR`, `ROUND`, `SQRT`, `POW`
- `MIN`, `MAX`, `SUM`, `AVERAGE`
- `RANDOM`, `SIGN`, `TRUNC`

**JSON Functions:**
- `JSON_PARSE`, `JSON_STRINGIFY`
- `JSON_GET`, `JSON_SET`

**Array Functions:**
- `MAP`, `FILTER`, `REDUCE`
- `SORT`, `REVERSE`, `JOIN`
- `SLICE`, `CONCAT`

**Date/Time Functions:**
- `DATE`, `TIME`, `FORMAT_DATE`
- `PARSE_DATE`, `ADD_DAYS`, `DIFF_DAYS`

### 2. Custom Spreadsheet Functions (from `examples/spreadsheet-poc/*`)

Spreadsheet-specific functions for working with cell ranges:

- `SUM_RANGE(rangeRef)` - Sum cells in range (e.g., `SUM_RANGE("A1:A5")`)
- `AVERAGE_RANGE(rangeRef)` - Average of cells in range
- `COUNT_RANGE(rangeRef)` - Count non-empty cells
- `MIN_RANGE(rangeRef)` - Minimum value in range
- `MAX_RANGE(rangeRef)` - Maximum value in range
- `CELL(ref)` - Get cell value by reference
- `ROW(ref)` - Get row number of cell
- `COLUMN(ref)` - Get column number of cell

### 3. Extra Function Libraries (from `extras/functions/*`)

Extended function libraries can be loaded via the **Setup Script** feature.

#### Loading Extra Libraries via Setup Button

Click the **⚙️ Setup** button in the header to open the setup editor. Add REQUIRE statements for the libraries you want to use:

**Example Setup Script:**
```rexx
// Load Excel-like functions (VLOOKUP, SUMIF, etc.)
REQUIRE "cwd:../../extras/functions/excel/src/excel-functions.js"

// Load R-inspired statistical functions
REQUIRE "cwd:../../extras/functions/r-inspired/src/r-statistics-functions.js"

// Load GraphViz visualization
REQUIRE "cwd:../../extras/functions/graphviz/src/graphviz-functions.js"
```

Click **Save & Execute** to load the libraries. Now all functions are available in all cells:

**Excel Functions:**
```rexx
=VLOOKUP(A1, B1:D10, 2, 0)
=SUMIF(A1:A10, ">5", B1:B10)
```

**R-Inspired Statistical Functions:**
```rexx
=MEAN([1,2,3,4,5])
=MEDIAN([1,2,3,4,5])
=SD([1,2,3,4,5])
```

**GraphViz Visualization:**
```rexx
=DIGRAPH("A -> B -> C")
```

**Benefits of Setup Script Approach:**
- ✅ Load libraries once, use everywhere
- ✅ Clean formulas (no REQUIRE in each cell)
- ✅ Easy to see what libraries are loaded
- ✅ Setup script is saved with spreadsheet data

**Note:** Extra libraries require proper path resolution. Use `cwd:` prefix for relative paths from the HTML file location.

### 4. NumPy Functions via PyOdide (Real Python)

The spreadsheet can integrate **real Python NumPy** via PyOdide, providing 100% accurate scientific computing with authentic NumPy algorithms.

#### Features

- **100% NumPy Compatible**: Runs real Python NumPy code (not JavaScript approximations)
- **Cached After First Load**: PyOdide initializes once (~10 seconds), then functions are fast (~1ms)
- **Full Linear Algebra**: Eigenvalues, matrix operations, determinants with perfect accuracy
- **25+ Functions**: Array creation, math, statistics, linear algebra, array manipulation

#### Installation

1. Load PyOdide and numpy-via-pyoide in your HTML:
```html
<!-- Load PyOdide -->
<script src="https://cdn.jsdelivr.net/pyodide/v0.26.1/full/pyodide.js"></script>

<!-- Load PyOdide ADDRESS handler -->
<script src="../../extras/addresses/pyodide/src/pyodide-address.js"></script>

<!-- Load NumPy via PyOdide -->
<script src="../../extras/functions/numpy-via-pyoide/numpy.js"></script>
```

2. In your spreadsheet initialization code, call:
```javascript
await adapter.installNumpyFunctions();
```

#### Available Functions

All numpy functions are available with the `NP_` prefix:

**Array Creation:**
- `NP_ZEROS(shape)` - Create array of zeros
- `NP_ONES(shape)` - Create array of ones
- `NP_EYE(n)` - Create identity matrix
- `NP_ARANGE(start, stop, step)` - Create range array
- `NP_LINSPACE(start, stop, num)` - Create evenly spaced array

**Math Functions:**
- `NP_SIN(x)`, `NP_COS(x)`, `NP_EXP(x)`
- `NP_LOG(x)`, `NP_SQRT(x)`

**Statistics:**
- `NP_MEAN(arr)`, `NP_MEDIAN(arr)`
- `NP_STD(arr)`, `NP_SUM(arr)`

**Linear Algebra:**
- `NP_DOT(a, b)` - Dot product
- `NP_MATMUL(a, b)` - Matrix multiplication
- `NP_DET(matrix)` - Determinant
- `NP_INV(matrix)` - Matrix inverse
- `NP_EIGVALS(matrix)` - Eigenvalues
- `NP_EIG(matrix)` - Eigenvalues and eigenvectors

**Array Manipulation:**
- `NP_RESHAPE(arr, shape)` - Reshape array
- `NP_TRANSPOSE(arr)` - Transpose matrix
- `NP_FLATTEN(arr)` - Flatten to 1D

#### Example Formulas

```rexx
// Square roots of array
=NP_SQRT([4, 9, 16, 25])
// Returns: [2, 3, 4, 5]

// Evenly spaced numbers
=NP_LINSPACE(0, 10, 5)
// Returns: [0, 2.5, 5, 7.5, 10]

// Sine of angles in radians
=NP_SIN([0, 1.57, 3.14])
// Returns: [0, 0.9999996829318346, 0.0015926529164868282]

// Mean of 2D array
=NP_MEAN([[1,2,3], [4,5,6]])
// Returns: 3.5

// Dot product
=NP_DOT([1,2,3], [4,5,6])
// Returns: 32

// 3x3 identity matrix
=NP_EYE(3)
// Returns: [[1,0,0], [0,1,0], [0,0,1]]

// Eigenvalues of matrix
=NP_EIGVALS([[4,1], [1,3]])
// Returns: [5, 2]

// Matrix determinant
=NP_DET([[1,2], [3,4]])
// Returns: -2.0
```

#### Performance Notes

- **First load**: ~10 seconds (loading PyOdide and NumPy)
- **Subsequent calls**: ~1ms per function (PyOdide is cached)
- **Accuracy**: 100% identical to `pip install numpy`
- **Bundle size**: ~15MB (PyOdide + NumPy)

#### Example Page

See `numpy-example.html` for a complete example with NumPy integration.

## Development

### File Structure
```
examples/spreadsheet-poc/
├── README.md                      # This file
├── FILE-LOADING.md               # File format and loading documentation
├── BUILDING-BINARY.md            # Production binary build guide
├── rexxsheet-dev                 # Development launch script (executable)
├── package.json                  # Node.js/npm configuration
├── vite.config.js                # Vite build configuration
├── index.html                    # HTML entry point (Vite-served)
├── spreadsheet-styles.css        # Stylesheet
├── test-sheet.json               # Sample test spreadsheet
├── sample-budget.json            # Sample budget spreadsheet
├── spreadsheet-loader.js         # Load/save spreadsheet files (legacy location)
├── src/                          # Source files (Vite entry)
│   ├── main.jsx                  # Vite entry point
│   ├── SpreadsheetApp.jsx        # React components
│   ├── spreadsheet-model.js      # Core spreadsheet logic
│   ├── spreadsheet-rexx-adapter.js  # RexxJS integration layer
│   └── spreadsheet-loader.js     # Load/save spreadsheet files
├── src-tauri/                    # Tauri (desktop app) backend
│   ├── Cargo.toml                # Rust dependencies
│   ├── tauri.conf.json           # Tauri configuration
│   └── src/
│       ├── main.rs               # Rust entry point
│       └── lib.rs                # CLI argument handling
├── public/                       # Static assets served by Vite
│   ├── rexxjs.bundle.js          # RexxJS interpreter bundle
│   └── lib/
│       └── spreadsheet-functions.js  # Custom spreadsheet functions
└── tests/                        # Test files
    ├── spreadsheet-model.test.js    # Jest tests
    ├── spreadsheet-loader.spec.js   # Loader tests
    └── web/                         # Playwright browser tests
        ├── spreadsheet-ui.spec.js
        ├── spreadsheet-loader-web.spec.js
        └── spreadsheet-controlbus-web.spec.js
```

### Code Organization

- **No spreadsheet logic in core/**: All spreadsheet-specific code is in `examples/`
- **Reusable RexxJS libraries**: Could be extracted to `extras/` if needed
- **REQUIRE system**: Libraries can be loaded via RexxJS `REQUIRE` statements

## License

**RexxSheet uses dual licensing:**

- **Application**: GNU Affero General Public License v3 (AGPL)
- **Control Protocol/Grammar**: MIT License

This means:
- ✅ The spreadsheet UI and application code is AGPL (share modifications)
- ✅ The wire protocol for controlling spreadsheets is MIT (use freely)
- ✅ You can build proprietary clients using the MIT protocol
- ✅ Alternative spreadsheet implementations can use the MIT protocol
- ⚠️ Modifications to the application itself must remain open source (AGPL)

**See [LICENSING.md](LICENSING.md) for complete details.**

Quick summary:
- Want to control the spreadsheet? → **MIT** (do whatever you want)
- Want to modify the app? → **AGPL** (share your changes)
