# Excel/Google Sheets Features in RexxSheet

This document describes the Excel and Google Sheets-like features available in RexxSheet, including both UI interactions and programmatic control via the Control Bus (RexxJS commands).

## Table of Contents

- [Cell Formatting](#cell-formatting)
- [Text Wrapping](#text-wrapping)
- [Freeze Panes](#freeze-panes)
- [Sort Data](#sort-data)
- [Find and Replace](#find-and-replace)
- [Data Validation](#data-validation)
- [Fill Down/Right (Autofill)](#fill-downright-autofill)
- [Hide/Unhide Rows and Columns](#hideunhide-rows-and-columns)
- [Named Ranges](#named-ranges)
- [Row/Column Operations](#rowcolumn-operations)
- [Undo/Redo](#undoredo)

---

## Cell Formatting

### Number Formatting

Format cell values using various number formats.

**Available Formats:**
- Currency (USD, EUR, GBP, JPY)
- Percentage
- Decimals (0, 0.0, 0.00, 0.000)
- Scientific notation
- Custom formats

**UI:**
- Right-click cell → Number Format → Select format

**Control Bus:**
```rexx
-- Set currency format
CALL SETFORMAT("A1", "currency:USD")

-- Set percentage
CALL SETFORMAT("B1", "number:0.00%")

-- Get format
format = GETFORMAT("A1")
```

### Visual Styling

Apply colors, fonts, and alignment to cells.

**Available Styles:**
- **Text:** Bold, Italic, Underline
- **Colors:** Text colors (red, blue, green, orange)
- **Backgrounds:** Yellow, light blue, light green, etc.
- **Alignment:** Left, Center, Right, Justify
- **Font Size:** Small (14px), Medium (16px), Large (20px)

**UI:**
- Right-click cell → Format → Select style
- Right-click cell → Align → Select alignment

**Control Bus:**
```rexx
-- Set bold
CALL SETFORMAT("A1", "bold")

-- Set text color
CALL SETFORMAT("A1", "color:red")

-- Set background color
CALL SETFORMAT("A1", "background:yellow")

-- Set alignment
CALL SETFORMAT("A1", "align:center")

-- Combine multiple styles (semicolon-separated)
CALL SETFORMAT("A1", "bold;color:red;align:center")

-- Clear format
CALL SETFORMAT("A1", "")
```

---

## Text Wrapping

Wrap text within cells to display long content across multiple lines.

**UI:**
- Right-click cell → Wrap Text (or Unwrap Text)
- The menu item shows current state: ↩️ if not wrapped, ⬜ if wrapped

**Control Bus:**
```rexx
-- Enable text wrapping
CALL SETWRAPTEXT("A1", 1)

-- Disable text wrapping
CALL SETWRAPTEXT("A1", 0)

-- Check wrap status
wrapped = GETWRAPTEXT("A1")  /* Returns 1 if wrapped, 0 if not */
```

**CSS Behavior:**
When text wrapping is enabled, cells apply:
- `white-space: pre-wrap`
- `word-wrap: break-word`
- `overflow-wrap: break-word`

---

## Freeze Panes

Lock rows and columns in place while scrolling through the spreadsheet.

**UI:**
Coming soon - View menu with freeze options

**Control Bus:**
```rexx
-- Freeze top 2 rows and left 3 columns
CALL FREEZEPANES(2, 3)

-- Freeze only top row
CALL FREEZEPANES(1, 0)

-- Unfreeze all
CALL UNFREEZEPANES()

-- Get frozen pane settings
frozen = GETFROZENPANES()
SAY "Frozen rows:" frozen.ROWS
SAY "Frozen columns:" frozen.COLUMNS
```

**Use Cases:**
- Keep header row visible while scrolling data
- Keep ID column visible while scrolling right
- Navigate large spreadsheets more easily

---

## Sort Data

Sort a range of cells by a specific column.

**UI:**
Coming soon - Context menu sort option

**Control Bus:**
```rexx
-- Sort range A1:C10 by column A, ascending
CALL SORTRANGE("A1:C10", "A", 1)

-- Sort by column B, descending
CALL SORTRANGE("A1:C10", "B", 0)

-- Sort by column number (1=A, 2=B, etc.)
CALL SORTRANGE("A1:C10", 1, 1)
```

**Features:**
- Numeric sorting (for number values)
- Alphabetic sorting (for text values)
- Maintains row integrity (all columns move together)
- Preserves formulas and formatting

**Example:**
```rexx
-- Sort a sales report by revenue (column C), highest first
CALL SORTRANGE("A2:C100", "C", 0)
```

---

## Find and Replace

Search for values in cells and optionally replace them.

**UI:**
Coming soon - Find dialog (Ctrl+F)

**Control Bus:**

### Find

```rexx
-- Basic find
results = FIND("hello")

-- Find with options
results = FIND("hello", matchCase, matchEntireCell, searchFormulas)
--        FIND(search, 1/0,       1/0,             1/0)

-- Results are returned as a REXX stem array
SAY "Found" results.0 "matches"
DO i = 1 TO results.0
  SAY "Match" i ":" results.i
END
```

**Find Options:**
- `matchCase` (0/1): Case-sensitive search
- `matchEntireCell` (0/1): Match whole cell vs. partial match
- `searchFormulas` (0/1): Search in formulas instead of values

### Replace

```rexx
-- Basic replace
count = REPLACE("old", "new")

-- Replace with options
count = REPLACE("old", "new", matchCase, matchEntireCell, searchFormulas)

SAY "Replaced" count "occurrences"
```

**Example:**
```rexx
-- Find all cells containing "Total"
results = FIND("Total", 0, 0, 0)

-- Replace "Q1" with "Q2" in all formulas
count = REPLACE("Q1", "Q2", 0, 0, 1)
```

---

## Data Validation

Restrict cell input to specific values or ranges.

**Validation Types:**
1. **List** - Dropdown of allowed values
2. **Number** - Numeric range (min/max)
3. **Text** - Length constraints and pattern matching
4. **Custom** - Custom formula validation

**UI:**
Coming soon - Data validation dialog

**Control Bus:**

### List Validation (Dropdown)

```rexx
-- Create dropdown with options
CALL SETCELLVALIDATION("A1", "list", "Red,Green,Blue")

-- User must select from: Red, Green, or Blue
```

### Number Validation

```rexx
-- Allow numbers between 0 and 100
CALL SETCELLVALIDATION("B1", "number", 0, 100)

-- Min only (no max)
CALL SETCELLVALIDATION("B2", "number", 0)
```

### Text Validation

```rexx
-- Text length between 5 and 10 characters
CALL SETCELLVALIDATION("C1", "text", 5, 10)

-- Text matching pattern
CALL SETCELLVALIDATION("D1", "text", 0, 0, "^[A-Z]{3}-[0-9]{4}$")
```

### Clear Validation

```rexx
CALL CLEARCELLVALIDATION("A1")
```

### Validate a Value

```rexx
-- Check if a value is valid for a cell
valid = VALIDATECELL("A1", "Purple")  /* Returns 0 (invalid) */
valid = VALIDATECELL("A1", "Red")     /* Returns 1 (valid) */
```

**Use Cases:**
- Dropdown lists for categories (Status: Pending, Complete, In Progress)
- Restrict age to 0-120
- Enforce product code format
- Limit text length for comments

---

## Fill Down/Right (Autofill)

Copy cell content and formulas to adjacent cells.

**UI:**
Coming soon - Drag handle for autofill

**Control Bus:**

### Fill Down

```rexx
-- Copy A1 down to A2:A10
CALL FILLDOWN("A1", "A2:A10")

-- Copy A1:B1 down to A2:B10 (multiple columns)
CALL FILLDOWN("A1:B1", "A2:B10")
```

### Fill Right

```rexx
-- Copy A1 right to B1:E1
CALL FILLRIGHT("A1", "B1:E1")

-- Copy A1:A2 right to B1:E2 (multiple rows)
CALL FILLRIGHT("A1:A2", "B1:E2")
```

**Features:**
- Automatic formula adjustment (relative references)
- Preserves formatting
- Handles both values and formulas

**Example:**
```rexx
-- Set up headers
CALL SETCELL("A1", "Jan")
CALL SETCELL("B1", "Feb")
CALL SETCELL("C1", "Mar")

-- Set up formulas
CALL SETCELL("A2", "100")
CALL SETCELL("A3", "=A2 * 1.1")

-- Fill formula down
CALL FILLDOWN("A3", "A4:A10")
-- A4 becomes =A3*1.1, A5 becomes =A4*1.1, etc.
```

---

## Hide/Unhide Rows and Columns

Temporarily hide rows or columns without deleting them.

**UI:**
Coming soon - Context menu options

**Control Bus:**

### Hide/Unhide Rows

```rexx
-- Hide row 5
CALL HIDEROW(5)

-- Unhide row 5
CALL UNHIDEROW(5)

-- Check if row is hidden
hidden = ISROWHIDDEN(5)  /* Returns 1 if hidden, 0 if visible */
```

### Hide/Unhide Columns

```rexx
-- Hide column B
CALL HIDECOLUMN("B")
CALL HIDECOLUMN(2)  /* Same as above */

-- Unhide column B
CALL UNHIDECOLUMN("B")

-- Check if column is hidden
hidden = ISCOLUMNHIDDEN("B")
```

**Use Cases:**
- Hide helper columns with intermediate calculations
- Hide rows with sensitive data
- Simplify view for presentations

---

## Named Ranges

Define named ranges for easier formula references.

**UI:**
Coming soon - Named range manager

**Control Bus:**

```rexx
-- Define a named range
CALL DEFINENAMEDRANGE("SalesData", "A2:A100")
CALL DEFINENAMEDRANGE("TaxRate", "B1")

-- Use in formulas
CALL SETCELL("C1", "=SUM_RANGE(SalesData)")
CALL SETCELL("D1", "=A1 * TaxRate")

-- Get a named range
range = GETNAMEDRANGE("SalesData")  /* Returns "A2:A100" */

-- Delete a named range
CALL DELETENAMEDRANGE("SalesData")

-- Get all named ranges
ranges = GETALLNAMEDRANGES()
SAY "Range count:" ranges.0
DO i = 1 TO ranges.0
  SAY ranges.i.NAME "=" ranges.i.RANGE
END
```

**Use Cases:**
- Reference data ranges by meaningful names
- Make formulas more readable
- Centralize range definitions

---

## Row/Column Operations

Insert or delete entire rows and columns.

**UI:**
- Right-click cell → Row → Insert Row Above/Below, Delete This Row
- Right-click cell → Column → Insert Column Left/Right, Delete This Column

**Control Bus:**

```rexx
-- Insert row at position 5
CALL INSERTROW(5)

-- Delete row 5
CALL DELETEROW(5)

-- Insert column at position B (column 2)
CALL INSERTCOLUMN("B")
CALL INSERTCOLUMN(2)  /* Same as above */

-- Delete column B
CALL DELETECOLUMN("B")
```

**Features:**
- Automatic formula adjustment (cell references update)
- Preserves formatting
- Invalid references become #REF!

---

## Undo/Redo

Undo and redo changes to the spreadsheet.

**UI:**
- Ctrl+Z to undo
- Ctrl+Y or Ctrl+Shift+Z to redo

**Control Bus:**

```rexx
-- Undo last action
success = UNDO()  /* Returns 1 if undo performed, 0 if nothing to undo */

-- Redo last undone action
success = REDO()  /* Returns 1 if redo performed, 0 if nothing to redo */

-- Check if undo/redo available
canUndo = CANUNDO()  /* Returns 1 if undo available */
canRedo = CANREDO()  /* Returns 1 if redo available */
```

**Features:**
- 100-level undo history
- Tracks all cell changes, row/column operations, etc.
- Redo stack cleared when new action performed

---

## Complete Example

Here's a complete example demonstrating multiple features:

```rexx
/* Budget Spreadsheet with Excel-like features */

-- Setup headers
CALL SETCELL("A1", "Item")
CALL SETCELL("B1", "Amount")
CALL SETCELL("C1", "Tax")
CALL SETCELL("D1", "Total")

-- Format headers
CALL SETFORMAT("A1:D1", "bold;background:lightblue;align:center")
CALL SETWRAPTEXT("A1", 1)  /* Enable wrapping for long headers */

-- Freeze header row
CALL FREEZEPANES(1, 0)

-- Add data
CALL SETCELL("A2", "Office Supplies")
CALL SETCELL("B2", "1250")
CALL SETCELL("C2", "=B2 * 0.08")
CALL SETCELL("D2", "=B2 + C2")

-- Format currency columns
CALL SETFORMAT("B2", "currency:USD")
CALL SETFORMAT("C2", "currency:USD")
CALL SETFORMAT("D2", "currency:USD")

-- Fill down formulas
CALL FILLDOWN("C2:D2", "C3:D10")

-- Add data validation for amounts
CALL SETCELLVALIDATION("B2:B10", "number", 0, 10000)

-- Define named range
CALL DEFINENAMEDRANGE("Expenses", "D2:D10")

-- Add total row
CALL SETCELL("D11", "=SUM_RANGE(Expenses)")
CALL SETFORMAT("D11", "bold;currency:USD;background:yellow")

-- Sort by total (column D), highest first
CALL SORTRANGE("A2:D10", "D", 0)

SAY "Budget spreadsheet created successfully!"
```

---

## Comparison with Excel/Google Sheets

| Feature | Excel/Sheets | RexxSheet | Notes |
|---------|-------------|-----------|-------|
| Cell Formatting | ✅ | ✅ | Full support |
| Text Wrapping | ✅ | ✅ | |
| Freeze Panes | ✅ | ✅ | |
| Sort | ✅ | ✅ | Single column sort |
| Find/Replace | ✅ | ✅ | |
| Data Validation | ✅ | ✅ | List, number, text |
| Autofill | ✅ | ⚠️ | FILLDOWN/FILLRIGHT only |
| Hide Rows/Cols | ✅ | ✅ | |
| Named Ranges | ✅ | ✅ | |
| Undo/Redo | ✅ | ✅ | 100-level history |
| Merge Cells | ✅ | ❌ | Not yet implemented |
| Conditional Formatting | ✅ | ✅ | Via RexxJS expressions |
| Charts | ✅ | ✅ | Via Chart.js integration |
| Pivot Tables | ✅ | ❌ | Not yet implemented |
| Multi-sheet | ✅ | ❌ | Not yet implemented |

---

## Future Enhancements

Planned features to increase Excel/Google Sheets compatibility:

1. **Merge Cells** - Combine multiple cells into one
2. **Smart Autofill** - Detect patterns (1,2,3... or Mon,Tue,Wed...)
3. **Filter** - Filter rows based on column values
4. **Conditional Formatting UI** - Visual builder for conditional formats
5. **Data Validation UI** - Dialog for setting validation rules
6. **Multi-sheet Support** - Multiple tabs/worksheets
7. **Sparklines** - Mini charts in cells
8. **Cell Borders** - Advanced border controls
9. **Protect Cells** - Lock cells from editing
10. **CSV Import/Export** - Load and save CSV files

---

## See Also

- [README.md](README.md) - Main documentation
- [FEATURES.md](docs/FEATURES.md) - Visual feature guide with screenshots
- [TESTING-CONTROL-BUS.md](TESTING-CONTROL-BUS.md) - Control Bus documentation
- [FILE-LOADING.md](FILE-LOADING.md) - File format and loading
