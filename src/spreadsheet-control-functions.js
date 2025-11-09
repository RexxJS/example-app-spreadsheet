/**
 * Spreadsheet Control Protocol/Grammar - MIT Licensed
 *
 * Copyright (c) 2024-2025 Paul Hammant and Contributors
 * Licensed under the MIT License (see LICENSE-PROTOCOL-MIT)
 *
 * This file defines the wire protocol/grammar for controlling spreadsheets
 * via REXX ADDRESS commands. The protocol is intentionally MIT licensed to
 * encourage ecosystem growth and interoperability.
 *
 * REXX functions for remote control of the spreadsheet.
 * These are used both by:
 * - Remote ADDRESS commands (via HTTP control bus)
 * - Cell expressions (e.g., =SETCELL("B1", GETCELL("A1")))
 *
 * This ensures a single source of truth for spreadsheet manipulation.
 *
 * Anyone can implement compatible clients in any language or create
 * alternative spreadsheet implementations using this protocol.
 */

/**
 * Create spreadsheet control functions bound to a specific model and adapter
 * @param {SpreadsheetModel} model - The spreadsheet model
 * @param {SpreadsheetRexxAdapter} adapter - The REXX adapter
 * @returns {Object} Functions to register with interpreter
 */
export function createSpreadsheetControlFunctions(model, adapter) {
  return {
    /**
     * SETCELL - Set cell content (value or formula)
     * Usage: CALL SETCELL("A1", "100")
     *        CALL SETCELL("A3", "=A1+A2")
     *        result = SETCELL("B1", "Hello")
     */
    SETCELL: async function(cellRef, content) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('SETCELL requires cell reference as first argument (e.g., "A1")');
      }
      if (content === undefined || content === null) {
        throw new Error('SETCELL requires content as second argument');
      }

      // Convert content to string
      const contentStr = String(content);

      // Set the cell
      await model.setCell(cellRef, contentStr, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      // Return the set value (REXX convention)
      return contentStr;
    },

    /**
     * GETCELL - Get cell value
     * Usage: value = GETCELL("A1")
     *        SAY "Cell A1 is:" GETCELL("A1")
     */
    GETCELL: function(cellRef) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('GETCELL requires cell reference as argument (e.g., "A1")');
      }

      const cell = model.getCell(cellRef);
      return cell.value || '';
    },

    /**
     * GETEXPRESSION - Get cell formula/expression
     * Usage: formula = GETEXPRESSION("A3")
     */
    GETEXPRESSION: function(cellRef) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('GETEXPRESSION requires cell reference as argument (e.g., "A1")');
      }

      const cell = model.getCell(cellRef);
      return cell.expression || '';
    },

    /**
     * CLEARCELL - Clear cell content
     * Usage: CALL CLEARCELL("A1")
     */
    CLEARCELL: async function(cellRef) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('CLEARCELL requires cell reference as argument (e.g., "A1")');
      }

      await model.setCell(cellRef, '', adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return '';
    },

    /**
     * SPREADSHEET_VERSION - Get spreadsheet control functions version
     * Usage: SAY SPREADSHEET_VERSION()
     */
    SPREADSHEET_VERSION: function() {
      return 'var_missing-2024-11-06';
    },

    /**
     * SETFORMAT - Set cell format
     * Usage: CALL SETFORMAT("A1", "bold")
     *        CALL SETFORMAT("A1", "bold;italic;color:red")
     */
    SETFORMAT: async function(cellRef, format) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('SETFORMAT requires cell reference as first argument (e.g., "A1")');
      }
      if (format === undefined || format === null) {
        format = '';
      }

      const formatStr = String(format);
      model.setCellMetadata(cellRef, { format: formatStr });

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return formatStr;
    },

    /**
     * GETFORMAT - Get cell format
     * Usage: format = GETFORMAT("A1")
     */
    GETFORMAT: function(cellRef) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('GETFORMAT requires cell reference as argument (e.g., "A1")');
      }

      const cell = model.getCell(cellRef);
      return cell.format || '';
    },

    /**
     * SETCOMMENT - Set cell comment
     * Usage: CALL SETCOMMENT("A1", "This is a note")
     */
    SETCOMMENT: async function(cellRef, comment) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('SETCOMMENT requires cell reference as first argument (e.g., "A1")');
      }
      if (comment === undefined || comment === null) {
        comment = '';
      }

      const commentStr = String(comment);
      model.setCellMetadata(cellRef, { comment: commentStr });

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return commentStr;
    },

    /**
     * GETCOMMENT - Get cell comment
     * Usage: comment = GETCOMMENT("A1")
     */
    GETCOMMENT: function(cellRef) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('GETCOMMENT requires cell reference as argument (e.g., "A1")');
      }

      const cell = model.getCell(cellRef);
      return cell.comment || '';
    },

    /**
     * GETROW - Get row number from cell reference
     * Usage: row = GETROW("A5")  -> 5
     */
    GETROW: function(cellRef) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('GETROW requires cell reference as argument (e.g., "A1")');
      }

      const parsed = SpreadsheetModel.parseCellRef(cellRef);
      return parsed.row;
    },

    /**
     * GETCOL - Get column number from cell reference
     * Usage: col = GETCOL("C1")  -> 3
     */
    GETCOL: function(cellRef) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('GETCOL requires cell reference as argument (e.g., "A1")');
      }

      const parsed = SpreadsheetModel.parseCellRef(cellRef);
      return SpreadsheetModel.colLetterToNumber(parsed.col);
    },

    /**
     * GETCOLNAME - Get column letter from cell reference
     * Usage: col = GETCOLNAME("C1")  -> "C"
     */
    GETCOLNAME: function(cellRef) {
      if (!cellRef || typeof cellRef !== 'string') {
        throw new Error('GETCOLNAME requires cell reference as argument (e.g., "A1")');
      }

      const parsed = SpreadsheetModel.parseCellRef(cellRef);
      return parsed.col;
    },

    /**
     * MAKECELLREF - Create cell reference from column and row
     * Usage: ref = MAKECELLREF(3, 5)  -> "C5"
     *        ref = MAKECELLREF("C", 5)  -> "C5"
     */
    MAKECELLREF: function(col, row) {
      if (col === undefined || col === null) {
        throw new Error('MAKECELLREF requires column as first argument (number or letter)');
      }
      if (row === undefined || row === null) {
        throw new Error('MAKECELLREF requires row as second argument (number)');
      }

      return SpreadsheetModel.formatCellRef(col, Number(row));
    },

    /**
     * GETCELLS - Get multiple cells by range
     * Usage: cells = GETCELLS("A1:B5")
     * Returns: REXX stem array with cell values
     */
    GETCELLS: function(rangeRef) {
      if (!rangeRef || typeof rangeRef !== 'string') {
        throw new Error('GETCELLS requires range reference as argument (e.g., "A1:B5")');
      }

      // Parse range reference (e.g., "A1:B5")
      const match = rangeRef.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
      if (!match) {
        throw new Error('Invalid range reference format. Expected format: "A1:B5"');
      }

      const start = SpreadsheetModel.parseCellRef(match[1]);
      const end = SpreadsheetModel.parseCellRef(match[2]);

      const startCol = SpreadsheetModel.colLetterToNumber(start.col);
      const endCol = SpreadsheetModel.colLetterToNumber(end.col);
      const startRow = start.row;
      const endRow = end.row;

      // Build REXX stem array
      const result = { 0: 0 }; // Count will be updated
      let index = 1;

      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          const ref = SpreadsheetModel.formatCellRef(col, row);
          const cell = model.getCell(ref);
          result[index] = {
            ref: ref,
            value: cell.value || '',
            expression: cell.expression || '',
            comment: cell.comment || '',
            format: cell.format || ''
          };
          index++;
        }
      }

      result[0] = index - 1; // Set count
      return result;
    },

    /**
     * SETCELLS - Set multiple cells at once
     * Usage: CALL SETCELLS("A1:A3", ["100", "200", "300"])
     */
    SETCELLS: async function(rangeRef, values) {
      if (!rangeRef || typeof rangeRef !== 'string') {
        throw new Error('SETCELLS requires range reference as first argument (e.g., "A1:A3")');
      }
      if (!values) {
        throw new Error('SETCELLS requires values array as second argument');
      }

      // Parse range reference
      const match = rangeRef.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
      if (!match) {
        throw new Error('Invalid range reference format. Expected format: "A1:B5"');
      }

      const start = SpreadsheetModel.parseCellRef(match[1]);
      const end = SpreadsheetModel.parseCellRef(match[2]);

      const startCol = SpreadsheetModel.colLetterToNumber(start.col);
      const endCol = SpreadsheetModel.colLetterToNumber(end.col);
      const startRow = start.row;
      const endRow = end.row;

      // Convert values to array if needed (handle REXX stem arrays)
      let valuesArray = [];
      if (Array.isArray(values)) {
        valuesArray = values;
      } else if (typeof values === 'object' && values[0] !== undefined) {
        // REXX stem array format: {0: count, 1: val1, 2: val2, ...}
        const count = values[0];
        for (let i = 1; i <= count; i++) {
          valuesArray.push(values[i]);
        }
      } else {
        throw new Error('SETCELLS values must be an array or REXX stem array');
      }

      // Set cells
      let valueIndex = 0;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          if (valueIndex < valuesArray.length) {
            const ref = SpreadsheetModel.formatCellRef(col, row);
            await model.setCell(ref, String(valuesArray[valueIndex]), adapter);
            valueIndex++;
          }
        }
      }

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return valueIndex; // Return count of cells set
    },

    /**
     * CLEAR - Clear all cells in spreadsheet
     * Usage: CALL CLEAR()
     */
    CLEAR: async function() {
      model.cells.clear();
      model.dependents.clear();
      model.evaluationInProgress.clear();

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * EXPORT - Export spreadsheet to JSON
     * Usage: data = EXPORT()
     * Returns: JSON string of spreadsheet data
     */
    EXPORT: function() {
      const data = model.toJSON();
      return JSON.stringify(data, null, 2);
    },

    /**
     * IMPORT - Import spreadsheet from JSON
     * Usage: CALL IMPORT(jsonData)
     */
    IMPORT: async function(jsonData) {
      if (!jsonData || typeof jsonData !== 'string') {
        throw new Error('IMPORT requires JSON string as argument');
      }

      try {
        const data = JSON.parse(jsonData);
        model.fromJSON(data, adapter);

        // Trigger UI update
        if (typeof window !== 'undefined') {
          window.dispatchEvent(new CustomEvent('spreadsheet-update'));
        }

        return 'OK';
      } catch (error) {
        throw new Error(`IMPORT failed: ${error.message}`);
      }
    },

    /**
     * GETSHEETNAME - Get spreadsheet name
     * Usage: name = GETSHEETNAME()
     */
    GETSHEETNAME: function() {
      return model.name || 'Sheet1';
    },

    /**
     * SETSHEETNAME - Set spreadsheet name
     * Usage: CALL SETSHEETNAME("Budget2024")
     */
    SETSHEETNAME: function(name) {
      if (!name || typeof name !== 'string') {
        throw new Error('SETSHEETNAME requires name as argument');
      }

      model.name = name;

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return name;
    },

    /**
     * EVALUATE - Evaluate RexxJS expression in spreadsheet context
     * Usage: result = EVALUATE("A1 + A2")
     */
    EVALUATE: async function(expression) {
      if (!expression || typeof expression !== 'string') {
        throw new Error('EVALUATE requires expression as argument');
      }

      try {
        const result = await adapter.evaluate(expression, model);
        return result;
      } catch (error) {
        throw new Error(`EVALUATE failed: ${error.message}`);
      }
    },

    /**
     * RECALCULATE - Force recalculation of all formulas
     * Usage: CALL RECALCULATE()
     */
    RECALCULATE: async function() {
      // Get all cells with formulas
      const cellsWithFormulas = [];
      for (const [ref, cell] of model.cells.entries()) {
        if (cell.expression) {
          cellsWithFormulas.push(ref);
        }
      }

      // Recalculate each cell
      for (const ref of cellsWithFormulas) {
        await model.evaluateCell(ref, adapter);
      }

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return cellsWithFormulas.length; // Return count of cells recalculated
    },

    /**
     * GETSETUPSCRIPT - Get setup script
     * Usage: script = GETSETUPSCRIPT()
     */
    GETSETUPSCRIPT: function() {
      return model.getSetupScript() || '';
    },

    /**
     * SETSETUPSCRIPT - Set setup script
     * Usage: CALL SETSETUPSCRIPT("LET TAX_RATE = 0.07")
     */
    SETSETUPSCRIPT: function(script) {
      if (script === undefined || script === null) {
        script = '';
      }

      model.setSetupScript(String(script));

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * EXECUTESETUPSCRIPT - Execute setup script
     * Usage: CALL EXECUTESETUPSCRIPT()
     */
    EXECUTESETUPSCRIPT: async function() {
      const script = model.getSetupScript();
      if (!script) {
        return 'OK'; // Nothing to execute
      }

      try {
        await adapter.executeSetupScript(script);
        return 'OK';
      } catch (error) {
        throw new Error(`EXECUTESETUPSCRIPT failed: ${error.message}`);
      }
    },

    /**
     * INSERTROW - Insert a row at the specified position
     * Usage: CALL INSERTROW(2)  -- Insert before row 2
     */
    INSERTROW: async function(rowNum) {
      if (!rowNum || typeof rowNum !== 'number' && typeof rowNum !== 'string') {
        throw new Error('INSERTROW requires row number as argument (e.g., 2)');
      }

      const rowNumber = parseInt(rowNum, 10);
      if (isNaN(rowNumber)) {
        throw new Error('INSERTROW requires a valid row number');
      }

      model.insertRow(rowNumber, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * DELETEROW - Delete a row at the specified position
     * Usage: CALL DELETEROW(2)  -- Delete row 2
     */
    DELETEROW: async function(rowNum) {
      if (!rowNum || typeof rowNum !== 'number' && typeof rowNum !== 'string') {
        throw new Error('DELETEROW requires row number as argument (e.g., 2)');
      }

      const rowNumber = parseInt(rowNum, 10);
      if (isNaN(rowNumber)) {
        throw new Error('DELETEROW requires a valid row number');
      }

      model.deleteRow(rowNumber, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * INSERTCOLUMN - Insert a column at the specified position
     * Usage: CALL INSERTCOLUMN(2)  -- Insert before column B
     *        CALL INSERTCOLUMN("B")  -- Insert before column B
     */
    INSERTCOLUMN: async function(colNum) {
      if (!colNum) {
        throw new Error('INSERTCOLUMN requires column number or letter as argument (e.g., 2 or "B")');
      }

      model.insertColumn(colNum, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * DELETECOLUMN - Delete a column at the specified position
     * Usage: CALL DELETECOLUMN(2)  -- Delete column B
     *        CALL DELETECOLUMN("B")  -- Delete column B
     */
    DELETECOLUMN: async function(colNum) {
      if (!colNum) {
        throw new Error('DELETECOLUMN requires column number or letter as argument (e.g., 2 or "B")');
      }

      model.deleteColumn(colNum, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * LISTCOMMANDS - Get list of available commands
     * Usage: commands = LISTCOMMANDS()
     * Returns: REXX stem array with command names
     */
    LISTCOMMANDS: function() {
      const commands = [
        'SETCELL', 'GETCELL', 'GETEXPRESSION', 'CLEARCELL',
        'SPREADSHEET_VERSION', 'SETFORMAT', 'GETFORMAT',
        'SETCOMMENT', 'GETCOMMENT', 'GETROW', 'GETCOL',
        'GETCOLNAME', 'MAKECELLREF', 'GETCELLS', 'SETCELLS',
        'CLEAR', 'EXPORT', 'IMPORT', 'GETSHEETNAME', 'SETSHEETNAME',
        'EVALUATE', 'RECALCULATE', 'GETSETUPSCRIPT', 'SETSETUPSCRIPT',
        'EXECUTESETUPSCRIPT', 'INSERTROW', 'DELETEROW', 'INSERTCOLUMN', 'DELETECOLUMN',
        'LISTCOMMANDS'
      ];

      // Return as REXX stem array
      const result = { 0: commands.length };
      commands.forEach((cmd, index) => {
        result[index + 1] = cmd;
      });

      return result;
    }
  };
}

/**
 * Function metadata for RexxJS parameter handling
 */
export const functionMetadata = {
  SETCELL: {
    name: 'SETCELL',
    params: ['cellRef', 'content'],
    description: 'Set cell content (value or formula)',
    examples: [
      'CALL SETCELL("A1", "100")',
      'CALL SETCELL("A3", "=A1+A2")',
      'result = SETCELL("B1", "Hello")'
    ]
  },
  GETCELL: {
    name: 'GETCELL',
    params: ['cellRef'],
    description: 'Get cell value',
    examples: [
      'value = GETCELL("A1")',
      'SAY "Cell A1 is:" GETCELL("A1")'
    ]
  },
  GETEXPRESSION: {
    name: 'GETEXPRESSION',
    params: ['cellRef'],
    description: 'Get cell formula/expression',
    examples: [
      'formula = GETEXPRESSION("A3")'
    ]
  },
  CLEARCELL: {
    name: 'CLEARCELL',
    params: ['cellRef'],
    description: 'Clear cell content',
    examples: [
      'CALL CLEARCELL("A1")'
    ]
  },
  SPREADSHEET_VERSION: {
    name: 'SPREADSHEET_VERSION',
    params: [],
    description: 'Get spreadsheet control functions version',
    examples: [
      'SAY SPREADSHEET_VERSION()'
    ]
  },
  SETFORMAT: {
    name: 'SETFORMAT',
    params: ['cellRef', 'format'],
    description: 'Set cell format (bold, italic, color, etc.)',
    examples: [
      'CALL SETFORMAT("A1", "bold")',
      'CALL SETFORMAT("A1", "bold;italic;color:red")'
    ]
  },
  GETFORMAT: {
    name: 'GETFORMAT',
    params: ['cellRef'],
    description: 'Get cell format',
    examples: [
      'format = GETFORMAT("A1")'
    ]
  },
  SETCOMMENT: {
    name: 'SETCOMMENT',
    params: ['cellRef', 'comment'],
    description: 'Set cell comment/note',
    examples: [
      'CALL SETCOMMENT("A1", "Important note")'
    ]
  },
  GETCOMMENT: {
    name: 'GETCOMMENT',
    params: ['cellRef'],
    description: 'Get cell comment',
    examples: [
      'comment = GETCOMMENT("A1")'
    ]
  },
  GETROW: {
    name: 'GETROW',
    params: ['cellRef'],
    description: 'Get row number from cell reference',
    examples: [
      'row = GETROW("A5")  -- returns 5'
    ]
  },
  GETCOL: {
    name: 'GETCOL',
    params: ['cellRef'],
    description: 'Get column number from cell reference',
    examples: [
      'col = GETCOL("C1")  -- returns 3'
    ]
  },
  GETCOLNAME: {
    name: 'GETCOLNAME',
    params: ['cellRef'],
    description: 'Get column letter from cell reference',
    examples: [
      'col = GETCOLNAME("C1")  -- returns "C"'
    ]
  },
  MAKECELLREF: {
    name: 'MAKECELLREF',
    params: ['col', 'row'],
    description: 'Create cell reference from column and row',
    examples: [
      'ref = MAKECELLREF(3, 5)  -- returns "C5"',
      'ref = MAKECELLREF("C", 5)  -- returns "C5"'
    ]
  },
  GETCELLS: {
    name: 'GETCELLS',
    params: ['rangeRef'],
    description: 'Get multiple cells by range',
    examples: [
      'cells = GETCELLS("A1:B5")',
      'count = cells.0  -- number of cells'
    ]
  },
  SETCELLS: {
    name: 'SETCELLS',
    params: ['rangeRef', 'values'],
    description: 'Set multiple cells at once',
    examples: [
      'CALL SETCELLS("A1:A3", ["100", "200", "300"])'
    ]
  },
  CLEAR: {
    name: 'CLEAR',
    params: [],
    description: 'Clear all cells in spreadsheet',
    examples: [
      'CALL CLEAR()'
    ]
  },
  EXPORT: {
    name: 'EXPORT',
    params: [],
    description: 'Export spreadsheet to JSON',
    examples: [
      'data = EXPORT()',
      'SAY data'
    ]
  },
  IMPORT: {
    name: 'IMPORT',
    params: ['jsonData'],
    description: 'Import spreadsheet from JSON',
    examples: [
      'CALL IMPORT(jsonString)'
    ]
  },
  GETSHEETNAME: {
    name: 'GETSHEETNAME',
    params: [],
    description: 'Get spreadsheet name',
    examples: [
      'name = GETSHEETNAME()'
    ]
  },
  SETSHEETNAME: {
    name: 'SETSHEETNAME',
    params: ['name'],
    description: 'Set spreadsheet name',
    examples: [
      'CALL SETSHEETNAME("Budget2024")'
    ]
  },
  EVALUATE: {
    name: 'EVALUATE',
    params: ['expression'],
    description: 'Evaluate RexxJS expression in spreadsheet context',
    examples: [
      'result = EVALUATE("A1 + A2")'
    ]
  },
  RECALCULATE: {
    name: 'RECALCULATE',
    params: [],
    description: 'Force recalculation of all formulas',
    examples: [
      'CALL RECALCULATE()'
    ]
  },
  GETSETUPSCRIPT: {
    name: 'GETSETUPSCRIPT',
    params: [],
    description: 'Get setup script',
    examples: [
      'script = GETSETUPSCRIPT()'
    ]
  },
  SETSETUPSCRIPT: {
    name: 'SETSETUPSCRIPT',
    params: ['script'],
    description: 'Set setup script',
    examples: [
      'CALL SETSETUPSCRIPT("LET TAX_RATE = 0.07")'
    ]
  },
  EXECUTESETUPSCRIPT: {
    name: 'EXECUTESETUPSCRIPT',
    params: [],
    description: 'Execute setup script',
    examples: [
      'CALL EXECUTESETUPSCRIPT()'
    ]
  },
  LISTCOMMANDS: {
    name: 'LISTCOMMANDS',
    params: [],
    description: 'Get list of available commands',
    examples: [
      'commands = LISTCOMMANDS()',
      'SAY commands.0 "commands available"'
    ]
  }
};
