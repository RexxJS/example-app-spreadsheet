# RexxJS Spreadsheet Enhancement Plan
## Object-Spreadsheet back and forth

**Status**: Planning
**Date**: 2025-11-16

---

## Overview

This document outlines object to s/sheet enhancements that align with RexxJS's strengths

---

## Priority 1: Query Chaining for Named Ranges

### Concept

Extend named ranges to support method chaining for data manipulation  using RexxJS pipeline syntax.

### Implementation

**New Functions** (add to `public/lib/spreadsheet-functions.js`):
```rexx
/* Filter rows in a range */
RANGE('A1:D100') |> WHERE('column_C > 1000') |> RESULT()

/* Aggregate operations */
RANGE('Sales!A:D') |> WHERE('region="West"') |> SUM('amount') |> RESULT()

/* Extract column */
RANGE('Data!A1:Z100') |> PLUCK('B') |> RESULT()

/* Group and aggregate */
RANGE('A:D') |> GROUP_BY('region') |> SUM('sales') |> RESULT()
```

**Benefits**:

- Leverages existing `|>` pipeline operator
- More readable than nested formulas
- Natural fit for REXX syntax

**Files to Modify**:
- `public/lib/spreadsheet-functions.js` - Add WHERE, PLUCK, GROUP_BY functions
- `src/spreadsheet-rexx-adapter.js` - Ensure range objects support chaining
- `tests/formula-evaluation.spec.js` - Add tests for chaining

**Estimated Effort**: 2-3 days

---

## Priority 2: Batch Control Bus Commands

### Concept

Reduce Control Bus round-trips by adding batch-aware commands, especially important for HTTP API in desktop mode.

### Implementation

**New Commands** (add to `src/spreadsheet-control-functions.js`):
```javascript
// Existing: SET_CELL (single cell)
// New: BATCH_SET_CELLS (multiple cells in one call)
BATCH_SET_CELLS: {
  handler: (model, args) => {
    const updates = args.updates; // [{address: 'A1', value: '42'}, ...]
    updates.forEach(({address, value}) => {
      model.setCellValue(address, value);
    });
  }
}

// Intelligent upsert (?)
CREATE_OR_UPDATE_ROW: {
  handler: (model, args) => {
    const {keyColumn, keyValue, rowData} = args;
    // Find row by key, update if exists, insert if not
  }
}
```

**Rust HTTP API Changes** (`src-tauri/src/lib.rs`):
```rust
// Add batch endpoint
"/control/batch" => {
  // Accept array of commands, execute in sequence
  // Return consolidated results
}
```

**Benefits**:
- Desktop mode: Reduce HTTP latency (currently 26 separate endpoints)
- Web mode: Fewer postMessage calls
- Better performance for automation scripts

**Files to Modify**:
- `src/spreadsheet-control-functions.js` - Add BATCH_* commands
- `src-tauri/src/lib.rs` - Add batch endpoint
- `tests/control-bus.spec.js` - Test batch operations

**Estimated Effort**: 3-4 days

---

## Priority 3: Context-Aware Validation

### Concept
Extend data validation to support different rules for create vs. update operations.

### Implementation

**Enhanced Validation Syntax**:
```rexx
/* Current validation (simple expression) */
=A1 > 0 & A1 < 100

/* New: Context-aware */
VALIDATE_ON_CREATE: =UNIQUE(A:A, A1)     /* Only when cell is first populated */
VALIDATE_ON_UPDATE: =A1 > PREVIOUS(A1)   /* Only when editing existing cell */
VALIDATE_ALWAYS: =A1 != ''               /* Both contexts */
```

**Model Changes** (`src/spreadsheet-model.js`):
```javascript
class SpreadsheetModel {
  validateCell(address, value, context) {
    // context: 'create' | 'update' | 'always'
    const validation = this.getValidation(address);

    if (validation.onCreate && context === 'create') {
      // Run create-specific validation
    }
    if (validation.onUpdate && context === 'update') {
      // Run update-specific validation
    }
    // Always run 'always' validations
  }
}
```

**Benefits**:
- Enforce UNIQUE constraints only on new entries
- Historical comparisons (require values to increase)
- Audit-friendly (different rules for initial vs. subsequent edits)

**Files to Modify**:
- `src/spreadsheet-model.js` - Add context parameter to validation
- `src/SpreadsheetApp.jsx` - Pass context when validating
- `tests/data-validation.spec.js` - Test context-aware rules

**Estimated Effort**: 2 days

---

## Priority 4: Auto-Increment Row IDs

### Concept
Optional auto-ID column for tracking rows, similar to database primary keys.

### Implementation

**Sheet-Level Setting**:
```javascript
// In sheet metadata
{
  name: "Customers",
  autoIdColumn: "#",  // Column for auto-IDs (null = disabled)
  nextId: 1001,       // Next ID to assign
}
```

**UI Enhancement** (`src/SpreadsheetApp.jsx`):
- Sheet settings dialog: "Enable Auto-ID Column"
- Dropdown to select column (A, B, C, ... or # for first column)

**Model Logic** (`src/spreadsheet-model.js`):
```javascript
insertRow(rowIndex) {
  if (this.currentSheet.autoIdColumn) {
    const idCol = this.currentSheet.autoIdColumn;
    const newId = this.currentSheet.nextId++;
    this.setCellValue(`${idCol}${rowIndex}`, newId);
  }
  // ... rest of insert logic
}
```

**Use Cases**:
- Track merged cell ownership
- Audit trails (correlate with edit history)
- Relational lookups: `=VLOOKUP(1005, A:D, 2)` instead of cell references
- Control Bus: `UPDATE_ROW_BY_ID` command

**Files to Modify**:
- `src/spreadsheet-model.js` - Add auto-ID logic
- `src/SpreadsheetApp.jsx` - Add settings UI
- `src/spreadsheet-control-functions.js` - Add *_BY_ID commands

**Estimated Effort**: 3 days

---

## Priority 5: Named Range Query Builder

### Concept
Define named ranges as "tables" with column metadata, enabling SQL-like queries.

### Implementation

**Named Range Metadata**:
```javascript
namedRanges: {
  "SalesData": {
    range: "A1:D100",
    columns: {
      id: "A",
      region: "B",
      product: "C",
      amount: "D"
    },
    hasHeader: true
  }
}
```

**RexxJS Query Syntax**:
```rexx
/* Use column names instead of letters */
=TABLE('SalesData') |> WHERE('region = "West" & amount > 1000') |> RESULT()

/* Aggregations */
=TABLE('SalesData') |> GROUP_BY('region') |> SUM('amount') |> RESULT()

/* Joins (advanced) */
=TABLE('Sales') |> JOIN(TABLE('Products'), 'product_id') |> RESULT()
```

**Benefits**:
- Readable formulas (column names vs. letters)
- Schema validation (catch typos in column references)
- Foundation for relational features

**Files to Modify**:
- `src/spreadsheet-model.js` - Add table metadata to named ranges
- `public/lib/spreadsheet-functions.js` - Add TABLE() function
- `src/SpreadsheetApp.jsx` - UI for defining table schemas

**Estimated Effort**: 4-5 days

---

## Implementation Roadmap

### Phase 1: Foundation (1 week)
- [ ] Query chaining for ranges (Priority 1)
- [ ] Context-aware validation (Priority 3)
- [ ] Test coverage for new features

### Phase 2: Performance (1 week)
- [ ] Batch Control Bus commands (Priority 2)
- [ ] Benchmark improvements in desktop mode
- [ ] Update Control Bus documentation

### Phase 3: Advanced Features (1.5 weeks)
- [ ] Auto-increment row IDs (Priority 4)
- [ ] Named range query builder (Priority 5)
- [ ] Integration tests via REXX scripts

### Phase 4: Documentation & Polish
- [ ] Update README.md with new features
- [ ] Add examples to docs/FEATURES.md
- [ ] Tutorial: "Building a CRM with RexxJS Spreadsheet"

---

## Design Principles

**Alignment with Existing Architecture**:
1. RexxJS-First: All features use REXX syntax, not JavaScript DSLs
2. Model Purity: Keep `spreadsheet-model.js` DOM-free
3. Control Bus Compatible: Every feature should be automatable
4. Test-Driven: Unit tests before implementation
5. SoC, IoC as usual

---

## Open Questions

1. **Query Chaining Syntax**: Should `RESULT()` be required, or auto-evaluate?
   - Pro `RESULT()`: Explicit, allows lazy evaluation
   - Con: Extra boilerplate

2. **Auto-ID Performance**: Should IDs be stored in metadata instead of cells?
   - Pro: No visible column clutter
   - Con: Breaks "what you see is what you get" principle

3. **Validation Context Detection**: How to distinguish create vs. update?
   - Option A: Check if cell was previously empty
   - Option B: Explicit undo/redo tracking
   - Option C: User declares context in formula

4. **Batch Command Atomicity**: Should batch operations be transactional?
   - If one command fails, roll back all?
   - Or continue with best-effort?

---

## Success Metrics

- [ ] Query chaining reduces formula complexity by 30%+ (measured by character count)
- [ ] Batch commands improve desktop automation speed by 50%+ (time for 100-cell updates)
- [ ] Context-aware validation enables new use cases (collect 3+ user testimonials)
- [ ] Zero breaking changes to existing spreadsheets
- [ ] Test coverage remains >80%

---

## References

- **REXX Pipelines**: IBM CMS Pipelines documentation
- **SQLite JSON Functions**: Inspiration for query syntax
- **Excel Power Query**: Comparison for query chaining UX
