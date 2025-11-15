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
