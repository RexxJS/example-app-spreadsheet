# RexxJS Spreadsheet - Visual Feature Guide

This guide provides a visual walkthrough of the RexxJS Spreadsheet features with screenshots demonstrating key functionality.

## Table of Contents

1. [Getting Started](#getting-started)
2. [Basic Cell Editing](#basic-cell-editing)
3. [Formulas and Cell References](#formulas-and-cell-references)
4. [RexxJS Functions](#rexxjs-functions)
5. [Range Functions](#range-functions)
6. [Number Formatting](#number-formatting)
7. [Cell Styling](#cell-styling)
8. [User Interface](#user-interface)
9. [Complete Examples](#complete-examples)

---

## Getting Started

### Initial Spreadsheet Load

When you first open the RexxJS Spreadsheet, you'll see a clean, Excel-like interface with sample data pre-loaded to help you get started.

![Initial spreadsheet with sample data](screenshots/overview-01-initial-load.png)

**Key Interface Elements:**
- **Header**: Application title and toolbar buttons
- **Formula Bar**: Shows the currently selected cell reference and its content/formula
- **Grid**: Main spreadsheet grid with column headers (A, B, C...) and row numbers
- **Info Panel**: Displays detailed information about the selected cell

---

## Basic Cell Editing

### Selecting a Cell

Click on any cell to select it. The selected cell is highlighted, and the formula bar shows the cell reference.

![Cell selected](screenshots/basic-editing-01-cell-selected.png)

### Entering Values

Double-click a cell or press Enter to start editing. Type your value and press Enter to confirm.

![Entering a value](screenshots/basic-editing-02-entering-value.png)

![Value entered and displayed](screenshots/basic-editing-03-value-entered.png)

**Tips:**
- Press **Enter** to confirm and move down
- Press **Tab** to confirm and move right
- Press **Escape** to cancel editing
- Use **arrow keys** for navigation

---

## Formulas and Cell References

### Basic Arithmetic Formulas

Formulas start with `=` and can reference other cells using A1-style notation.

#### Setting Up Values

![Values ready for formula](screenshots/formulas-01-values-setup.png)

#### Entering a Formula

Type a formula like `=A1 + A2` to add two cells together.

![Formula being entered](screenshots/formulas-02-formula-entry.png)

#### Viewing the Result

The formula automatically calculates and displays the result.

![Formula result displayed](screenshots/formulas-03-formula-result.png)

### Cell Dependencies and Automatic Updates

When a cell references another cell, it automatically updates when the referenced cell changes.

#### Initial Dependency

Cell A2 contains `=A1 * 2`, showing the value 20 when A1 is 10.

![Initial dependency](screenshots/formulas-04-dependency-initial.png)

#### Automatic Update

When A1 changes to 25, A2 automatically recalculates to 50.

![Dependency updated](screenshots/formulas-05-dependency-updated.png)

**How It Works:**
- The spreadsheet tracks dependencies between cells
- When you change a cell value, all dependent cells automatically recalculate
- Supports complex dependency chains (A1 → A2 → A3 → ...)

---

## RexxJS Functions

The spreadsheet includes powerful RexxJS functions for string manipulation, math, dates, and more.

### String Functions

RexxJS provides many built-in string functions:

![RexxJS string functions](screenshots/rexxjs-functions-01-string-functions.png)

**Examples shown:**
- `=UPPER("hello world")` → HELLO WORLD
- `=LENGTH("RexxJS")` → 6
- `=SUBSTR("Spreadsheet", 1, 6)` → Spread

**Available String Functions:**
- `UPPER(text)` - Convert to uppercase
- `LOWER(text)` - Convert to lowercase
- `LENGTH(text)` - Get string length
- `SUBSTR(text, start, length)` - Extract substring
- `REVERSE(text)` - Reverse string
- `STRIP(text)` - Remove leading/trailing whitespace
- And many more...

### Function Pipelines

RexxJS supports the powerful `|>` pipeline operator for chaining operations.

#### Entering a Pipeline

![Pipeline being entered](screenshots/rexxjs-functions-02-pipeline-entry.png)

#### Pipeline Result

![Pipeline result](screenshots/rexxjs-functions-03-pipeline-result.png)

**How Pipelines Work:**
```rexx
="hello" |> UPPER |> LENGTH
```

1. Start with "hello"
2. Pass to UPPER → "HELLO"
3. Pass to LENGTH → 5

**Benefits:**
- Cleaner, more readable formulas
- Left-to-right data flow
- Easy to add/remove transformation steps

---

## Range Functions

Range functions operate on multiple cells at once using range notation like `A1:A5`.

### Setting Up a Range

![Values setup for range function](screenshots/range-functions-01-values-setup.png)

### Using SUM_RANGE

The `SUM_RANGE` function adds all values in a range.

![SUM_RANGE result](screenshots/range-functions-02-sum-result.png)

**Available Range Functions:**
- `SUM_RANGE("A1:A5")` - Sum all values in range
- `AVERAGE_RANGE("A1:A5")` - Calculate average
- `MIN_RANGE("A1:A5")` - Find minimum value
- `MAX_RANGE("A1:A5")` - Find maximum value
- `COUNT_RANGE("A1:A5")` - Count non-empty cells

**Range Syntax:**
- `"A1:A5"` - Single column range
- `"A1:C1"` - Single row range
- `"A1:C5"` - Rectangular range

---

## Number Formatting

### Unformatted Numbers

By default, numbers are displayed as entered.

![Unformatted number](screenshots/formatting-01-unformatted-number.png)

### Accessing Format Options

Right-click a cell to open the context menu.

![Context menu](screenshots/formatting-02-context-menu.png)

### Number Format Menu

The Number Format submenu provides various formatting options.

![Number format submenu](screenshots/formatting-03-number-format-submenu.png)

### Currency Formatting

Apply currency formatting to display values with dollar signs and thousand separators.

![Currency formatting applied](screenshots/formatting-04-currency-applied.png)

**Result:** 1234.56 → $1,234.56

### Percentage Formatting

Convert decimals to percentages.

![Percentage formatting applied](screenshots/formatting-05-percentage-applied.png)

**Result:** 0.856 → 85.6%

**Available Number Formats:**
- **Currency (USD)**: $#,##0.00
- **Percentage**: 0.0%
- **Number (2 decimals)**: #,##0.00
- **Integer**: #,##0
- **Date (ISO)**: yyyy-mm-dd
- **Date (US)**: mm/dd/yyyy

---

## Cell Styling

### Text Styling - Bold and Color

Apply bold formatting and text colors to emphasize important cells.

![Bold red text styling](screenshots/styling-01-bold-red-text.png)

**How to Apply:**
1. Right-click the cell
2. Hover over "Format"
3. Select "Bold"
4. Repeat for color: "Format" → "Text: Red"

### Background Colors

Use background colors to organize and highlight data.

![Background color styling](screenshots/styling-02-background-colors.png)

**Available Colors:**
- Red, Green, Blue, Yellow, Orange (backgrounds)
- Red, Green, Blue, Orange (text colors)

### Text Alignment

Control how text is aligned within cells.

![Text alignment options](screenshots/styling-03-text-alignment.png)

**Alignment Options:**
- **Left** (default)
- **Center**
- **Right**

**Use Cases:**
- Left: Text labels
- Center: Headers, titles
- Right: Numbers (especially in columns)

---

## User Interface

### Formula Bar - Empty Cell

When you select an empty cell, the formula bar shows the cell reference.

![Formula bar with empty cell](screenshots/interface-01-formula-bar-empty.png)

### Formula Bar - Editing

You can edit cell content directly in the formula bar.

![Formula bar while editing](screenshots/interface-02-formula-bar-editing.png)

### Formula Bar - Showing Formula

When a cell with a formula is selected, the formula bar displays the formula while the cell shows the result.

![Formula bar showing formula](screenshots/interface-03-formula-bar-result.png)

**Formula Bar Features:**
- Shows cell reference (e.g., "A1")
- Displays raw formula for formula cells
- Allows direct editing without double-clicking the cell
- Supports all keyboard shortcuts (Ctrl+C, Ctrl+V, etc.)

### Info Panel

The Info Panel shows detailed information about the selected cell.

![Info panel with cell details](screenshots/interface-04-info-panel.png)

**Information Displayed:**
- **Cell Reference**: The A1-style cell address
- **Value**: The calculated or literal value
- **Formula**: The formula (if any)
- **Type**: Value type (number, string, formula)
- **Dependencies**: Which cells this cell depends on
- **Format**: Applied number format
- **Style**: Applied visual styles

---

## Complete Examples

### Budget Spreadsheet

A complete example showing headers, values, formulas, and formatting working together.

![Budget spreadsheet example](screenshots/examples-01-budget-spreadsheet.png)

**Features Demonstrated:**
- **Headers**: Bold formatting for column titles
- **Currency Formatting**: All amounts displayed with $ and commas
- **Range Formula**: `=SUM_RANGE("B2:B4")` calculates the total
- **Practical Layout**: Organized like a real budget

**Try It Yourself:**

1. Create headers: "Item" and "Amount"
2. List expense items in column A
3. Enter amounts in column B
4. Add a "Total" row
5. Use `=SUM_RANGE("B2:B4")` to calculate the sum
6. Format headers as bold
7. Format amounts as currency

---

## Keyboard Shortcuts

**Navigation:**
- **Arrow Keys**: Move between cells
- **Tab**: Move right
- **Shift+Tab**: Move left
- **Enter**: Move down
- **Shift+Enter**: Move up

**Editing:**
- **F2 or Double-Click**: Start editing
- **Escape**: Cancel editing
- **Ctrl+C**: Copy cell content
- **Ctrl+V**: Paste (limited support)

**View Modes:**
- **V**: Values view (show only literal values)
- **E**: Expressions view (show only formulas)
- **F**: Formats view (show format strings)
- **N**: Normal view (default)

---

## Advanced Features

### Conditional Formatting

Use RexxJS expressions to dynamically style cells based on their values.

**Example:**
```rexx
// In cell styleExpression property:
STYLE_IF(A1 < 0, RED_TEXT(), GREEN_TEXT())
```

See [README.md](../README.md#cell-formatting--styling) for more details.

### Named Variables

Define constants in the Setup Script for reuse across formulas.

**Setup Script:**
```rexx
LET TAX_RATE = 0.07
LET SHIPPING = 15.00
```

**In Cells:**
```rexx
=Subtotal * TAX_RATE
=Total + SHIPPING
```

### Control Bus

Control the spreadsheet programmatically via:
- **Web Mode**: iframe postMessage API
- **Desktop Mode**: HTTP REST API

See [TESTING-CONTROL-BUS.md](../TESTING-CONTROL-BUS.md) for complete documentation.

---

## Advanced Query & Data Features

RexxJS Spreadsheet includes powerful features for working with data like a database, including query chaining, table metadata, context-aware validation, and auto-increment IDs.

### Query Chaining with RANGE()

Build complex data queries using RexxJS's pipe operator (`|>`), similar to SQL or LINQ.

**Basic Filtering:**
```rexx
// Filter rows where column C > 1000
=RANGE('A1:D100') |> WHERE('column_C > 1000') |> RESULT()
```

**Extract Single Column:**
```rexx
// Get all values from column B
=RANGE('A1:D100') |> PLUCK('B')

// Or using column names if headers exist
=RANGE('SalesData') |> PLUCK('ProductName')
```

**Group and Aggregate:**
```rexx
// Sum amounts by region
=RANGE('A1:D100') |> GROUP_BY('B') |> SUM('D')

// Multiple operations - pipe operator chains left-to-right
=RANGE('Sales!A:D') |> WHERE('Amount > 1000') |> GROUP_BY('Region') |> AVG('Amount')
```

**Available Pipeline Functions:**
- `WHERE(condition)` - Filter rows (supports column_X, colX, or column names)
- `PLUCK(column)` - Extract single column
- `GROUP_BY(column)` - Group rows for aggregation
- `SUM(column)` - Sum values (works with GROUP_BY)
- `AVG(column)` - Average values (works with GROUP_BY)
- `COUNT()` - Count rows (works with GROUP_BY)
- `RESULT()` - Return final data array

**Note:** Backward-compatible method chaining (`.WHERE().GROUP_BY()`) still works.

### Table Metadata & SQL-like Queries

Define named ranges as "tables" with column metadata for more readable queries.

**Define Table Metadata (via Control Bus):**
```javascript
// JavaScript example using Control Bus
CALL SET_TABLE_METADATA("SalesData", '{
  "range": "A1:D100",
  "columns": {
    "id": "A",
    "region": "B",
    "product": "C",
    "amount": "D"
  },
  "hasHeader": true,
  "types": {
    "id": "number",
    "amount": "number"
  },
  "descriptions": {
    "region": "Sales region (North, South, East, West)",
    "amount": "Sale amount in dollars"
  }
}')
```

**Query Using Column Names:**
```rexx
// Much more readable than column letters!
=TABLE('SalesData') |> WHERE('region == "West"') |> GROUP_BY('product') |> SUM('amount')

// vs the old way:
=RANGE('A1:D100') |> WHERE('column_B == "West"') |> GROUP_BY('C') |> SUM('D')
```

**Benefits:**
- **Readability**: Use meaningful names ('region') instead of letters ('B')
- **Type Safety**: Optional type information for validation
- **Documentation**: Built-in descriptions for each column
- **Refactoring**: Change column order without breaking formulas

**Control Bus Commands:**
- `SET_TABLE_METADATA(name, jsonMetadata)` - Define table schema
- `GET_TABLE_METADATA(name)` - Retrieve metadata as JSON
- `DELETE_TABLE_METADATA(name)` - Remove table definition
- `LIST_TABLES()` - Get array of all table names

### Context-Aware Validation

Apply different validation rules for creating vs. updating cells.

**Setup Validation:**
```javascript
// Via model API
model.setCellValidation('A4', {
  type: 'contextual',
  onCreate: 'UNIQUE("A1:A100", value)',  // Only check when first populated
  onUpdate: 'value > PREVIOUS("A4")',    // Only check when editing
  always: 'value != ""'                   // Always check
});
```

**Use Cases:**
```javascript
// Ensure unique customer IDs on creation
onCreate: 'UNIQUE("A:A", value)'

// Require values to increase (audit trail)
onUpdate: 'value > PREVIOUS("B5")'

// Allow empty on update but not creation
onCreate: 'value != ""'
onUpdate: 'true'  // Always valid on update
```

**Helper Functions:**
- `UNIQUE(range, value)` - Check if value is unique in range
- `PREVIOUS(cellRef)` - Get previous value of cell (for comparisons)

### Auto-Increment Row IDs

Automatically generate sequential IDs for new rows, like database primary keys.

**Configure Auto-ID (via Control Bus):**
```javascript
// Enable auto-ID on column A, starting at 1
CALL CONFIGURE_AUTO_ID("A", 1, "")

// With custom prefix
CALL CONFIGURE_AUTO_ID("A", 1000, "ORDER-")
// Generates: ORDER-1000, ORDER-1001, ORDER-1002, ...
```

**Insert Row (Auto-ID Applied):**
```javascript
// When you insert a row, the ID is automatically set
model.insertRow(1);  // Cell A1 gets next ID

// IDs increment automatically
model.insertRow(2);  // A2 gets next ID
model.insertRow(3);  // A3 gets next ID
```

**Find and Update by ID:**
```javascript
// Find row by ID value
rowNum = FIND_ROW_BY_ID("ORDER-1005")

// Update specific row by ID
CALL UPDATE_ROW_BY_ID("ORDER-1005", '[
  {"column": "B", "value": "Updated Name"},
  {"column": "C", "value": "New Status"}
]')

// Get next ID that will be assigned
nextId = GET_NEXT_ID()
```

**Per-Sheet Configuration:**
Each sheet can have its own ID sequence with different prefixes.

### Batch Operations

Execute multiple operations in a single call for better performance.

**Batch Set Cells:**
```javascript
// Update multiple cells at once
CALL BATCH_SET_CELLS('[
  {"address": "A1", "value": "100"},
  {"address": "B2", "value": "200"},
  {"address": "C3", "value": "=A1+B2"}
]')

// Returns: {"total": 3, "success": 3, "errors": 0}
```

**Batch Execute Commands:**
```javascript
// Execute multiple Control Bus commands
CALL BATCH_EXECUTE('[
  {"command": "SETCELL", "args": ["A1", "10"]},
  {"command": "SETCELL", "args": ["A2", "20"]},
  {"command": "INSERTROW", "args": ["3"]}
]')

// Returns detailed results for each command
```

**Benefits:**
- **Performance**: Reduces HTTP round-trips in desktop mode
- **Atomicity**: All operations processed together
- **Error Handling**: Continues on error, reports which operations failed
- **Automation**: Ideal for bulk data imports and scripts

### Complete Example: Sales Tracking

Here's a complete example using all the advanced features:

**1. Define Table Schema:**
```javascript
CALL SET_TABLE_METADATA("Sales", '{
  "range": "A1:E1000",
  "columns": {
    "order_id": "A",
    "customer": "B",
    "region": "C",
    "product": "D",
    "amount": "E"
  },
  "hasHeader": true,
  "types": {
    "order_id": "string",
    "amount": "number"
  }
}')
```

**2. Configure Auto-IDs:**
```javascript
CALL CONFIGURE_AUTO_ID("A", 1000, "ORD-")
```

**3. Add Validation:**
```javascript
// Order IDs must be unique
model.setCellValidation('A2:A1000', {
  type: 'contextual',
  onCreate: 'UNIQUE("A:A", value)'
});

// Amounts must be positive
model.setCellValidation('E2:E1000', {
  type: 'formula',
  formula: 'value > 0'
});
```

**4. Query Data:**
```rexx
// Total sales by region
=TABLE('Sales') |> GROUP_BY('region') |> SUM('amount')

// High-value West coast orders
=TABLE('Sales') |> WHERE('region == "West" && amount > 5000') |> RESULT()

// Average order value by product
=TABLE('Sales') |> GROUP_BY('product') |> AVG('amount')
```

**5. Bulk Import Data:**
```javascript
CALL BATCH_SET_CELLS('[
  {"address": "B2", "value": "Acme Corp"},
  {"address": "C2", "value": "West"},
  {"address": "D2", "value": "Widget"},
  {"address": "E2", "value": "1500"}
]')
// A2 automatically gets "ORD-1000"
```

---

## Next Steps

- **Try the Examples**: Load `sample-budget.json` or `example-formatted.json`
- **Read the README**: [README.md](../README.md) for complete feature list
- **Explore Functions**: See the full list of RexxJS functions in the README
- **Build Custom Sheets**: Create your own spreadsheets with formulas and formatting

---

## Generating These Screenshots

These screenshots are automatically generated from tests. To regenerate them:

```bash
# Run tests with screenshot mode enabled
TAKE_SCREENSHOT=1 npm run test:web -- spreadsheet-screenshots-web.spec.js

# Screenshots are saved to docs/screenshots/
# Manifest is saved to docs/screenshots/manifest.json
```

The screenshot generation is controlled by the `TAKE_SCREENSHOT` environment variable. When set to `1`, the test suite captures screenshots at key moments and saves them with descriptive names for documentation purposes.

---

*This documentation is generated from the test file: `tests/spreadsheet-screenshots-web.spec.js`*
