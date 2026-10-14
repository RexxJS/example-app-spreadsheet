# Spreadsheet Column/Row Resizing & Enhanced Copy/Paste Features

This document describes the newly implemented features for the spreadsheet POC.

## 1. Interactive Column & Row Resizing

### Visual Resize Handles
- **Column headers**: Blue highlight on hover, draggable handles on the right edge
- **Row headers**: Blue highlight on hover, draggable handles on the bottom edge
- **Mouse-based resizing**: Click and drag to resize columns (horizontal) or rows (vertical)

### Constraints
- **Columns**: Minimum width of 20px
- **Rows**: Minimum height of 15px

### Persistent Sizing
- Column widths and row heights are stored in the model
- Automatically exported/imported with spreadsheet JSON

## 2. REXX Control Functions for Sizing

Four new spreadsheet control functions are available for programmatic control:

### SETCOLWIDTH(col, width)
Set column width in pixels. Accepts column number or letter.
```rexx
LET result = SETCOLWIDTH("A", 150)  -- Set column A to 150px
LET result = SETCOLWIDTH(1, 150)    -- Same as above
```

### GETCOLWIDTH(col)
Get column width in pixels.
```rexx
LET width = GETCOLWIDTH("A")  -- Returns current width of column A
```

### SETROWHEIGHT(row, height)
Set row height in pixels.
```rexx
LET result = SETROWHEIGHT(1, 50)  -- Set row 1 to 50px height
```

### GETROWHEIGHT(row)
Get row height in pixels.
```rexx
LET height = GETROWHEIGHT(1)  -- Returns current height of row 1
```

## 3. Smart Formula Copy/Paste with Relative References

### Formula Adjustment
When copying formulas and pasting elsewhere, cell references adjust relatively:

**Example:**
- Copy `=A1+B2` from cell A3
- Paste to cell C5
- Result: `=C3+D4` (2 columns right, 2 rows down)

### How It Works
- **Offset tracking**: Clipboard stores source position (sourceCol, sourceRow)
- **Reference adjustment**: `adjustCellReferences()` function parses formulas and updates cell refs by offset
- **Case-insensitive**: Cell references like `a1` work the same as `A1`

## 4. Enhanced Keyboard Shortcuts

### New Clipboard Operations
- **Ctrl+C**: Copy selected cell (formula or value)
- **Ctrl+X**: Cut selected cell (clears source after paste)
- **Ctrl+V**: Paste with relative reference adjustment

### Existing View Shortcuts
- **V**: View values only
- **E**: View expressions only
- **F**: View formats only
- **N**: Normal view (default)

### Priority Handling
Clipboard operations (Ctrl+C/X/V) are checked FIRST before plain key handlers, preventing conflicts like 'v' triggering view mode when doing Ctrl+V.

## 5. UI/UX Improvements

### Dynamic Cell Dimensions
- Cells now render with actual column width and row height from model
- Column headers and row headers use the same dimensions
- Resize handles are visible on hover

### Help Panel Updates
- Added Ctrl+X and Ctrl+V to keyboard shortcuts documentation
- Documented all clipboard features

## 6. Data Persistence

Column widths and row heights are now part of the spreadsheet JSON format:

```json
{
  "setupScript": "...",
  "cells": {
    "A1": "value",
    "B2": "=A1*2"
  },
  "columnWidths": {
    "1": 150,
    "2": 200
  },
  "rowHeights": {
    "1": 50,
    "3": 40
  }
}
```

### Backward Compatibility
- Old format (without columnWidths/rowHeights) still loads correctly
- New format exports sizing data only if columns/rows have been resized
- Default sizes (100px width, 32px height) are not stored

## Implementation Details

### Files Modified

1. **spreadsheet-model.js**
   - Added `columnWidths` and `rowHeights` properties
   - Added `getColumnWidth()`, `setColumnWidth()`, `getRowHeight()`, `setRowHeight()` methods
   - Updated `toJSON()` and `fromJSON()` to persist sizing data

2. **spreadsheet-rexx-adapter.js**
   - Added `SETCOLWIDTH()`, `GETCOLWIDTH()`, `SETROWHEIGHT()`, `GETROWHEIGHT()` functions
   - Functions support both column letters and numbers

3. **spreadsheet-app.jsx**
   - Updated `Cell` component to use dynamic width/height
   - Updated `ColumnHeader` with resize handle and mouse drag logic
   - Updated `RowHeader` with resize handle and mouse drag logic
   - Updated `Grid` to pass widths, heights, and resize handlers
   - Added clipboard state to `App` component
   - Implemented `handleCopy()`, `handleCut()`, `handlePaste()` with relative reference adjustment
   - Updated keyboard event handler to support Ctrl+C/X/V
   - Updated help panel with new shortcuts

4. **spreadsheet-styles.css**
   - Added `.resize-handle` styles for column and row handles
   - Added hover effects for column/row headers
   - Handle visibility and cursor changes

## Testing

The implementation includes:
- Minimum size constraints (20px for columns, 15px for rows)
- Case-insensitive cell reference parsing
- Relative reference adjustment for formulas
- Clipboard state management (copy vs cut)
- Data persistence compatibility

All core JavaScript files pass syntax validation.
