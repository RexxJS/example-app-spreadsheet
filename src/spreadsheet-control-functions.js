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
     * FILLDOWN - Fill down from source to target range
     * Usage: CALL FILLDOWN("A1", "A2:A10")
     *        CALL FILLDOWN("A1:B1", "A2:B10")
     */
    FILLDOWN: async function(sourceRef, targetRangeRef) {
      if (!sourceRef || !targetRangeRef) {
        throw new Error('FILLDOWN requires source reference and target range (e.g., "A1", "A2:A10")');
      }

      model.fillDown(sourceRef, targetRangeRef, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * FILLRIGHT - Fill right from source to target range
     * Usage: CALL FILLRIGHT("A1", "B1:E1")
     *        CALL FILLRIGHT("A1:A2", "B1:E2")
     */
    FILLRIGHT: async function(sourceRef, targetRangeRef) {
      if (!sourceRef || !targetRangeRef) {
        throw new Error('FILLRIGHT requires source reference and target range (e.g., "A1", "B1:E1")');
      }

      model.fillRight(sourceRef, targetRangeRef, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * SORTRANGE - Sort a range of cells by a column
     * Usage: CALL SORTRANGE("A1:C10", "A", 1)  -- Sort by column A ascending
     *        CALL SORTRANGE("A1:C10", 1, 1)    -- Sort by column 1 ascending
     *        CALL SORTRANGE("A1:C10", "B", 0)  -- Sort by column B descending
     */
    SORTRANGE: async function(rangeRef, sortCol, ascending) {
      if (!rangeRef || !sortCol) {
        throw new Error('SORTRANGE requires range reference and sort column (e.g., "A1:C10", "A", 1)');
      }

      const isAscending = ascending === undefined || ascending === 1 || ascending === '1' || ascending === true;
      model.sortRange(rangeRef, sortCol, isAscending, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * SETWRAPTEXT - Set text wrapping for a cell
     * Usage: CALL SETWRAPTEXT("A1", 1)  -- Enable wrapping
     *        CALL SETWRAPTEXT("A1", 0)  -- Disable wrapping
     */
    SETWRAPTEXT: function(cellRef, wrap) {
      if (!cellRef) {
        throw new Error('SETWRAPTEXT requires cell reference');
      }

      const wrapValue = wrap === 1 || wrap === '1' || wrap === true;
      model.setCellMetadata(cellRef, { wrapText: wrapValue });

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * GETWRAPTEXT - Get text wrapping setting for a cell
     * Usage: wrapped = GETWRAPTEXT("A1")
     * Returns: 1 if wrapping enabled, 0 if disabled
     */
    GETWRAPTEXT: function(cellRef) {
      if (!cellRef) {
        throw new Error('GETWRAPTEXT requires cell reference');
      }

      const cell = model.getCell(cellRef);
      return cell.wrapText ? 1 : 0;
    },

    /**
     * FIND - Find cells matching criteria
     * Usage: results = FIND("searchValue")
     *        results = FIND("searchValue", matchCase, matchEntireCell, searchFormulas)
     * Returns: REXX stem array with matching cell references
     */
    FIND: function(searchValue, matchCase, matchEntireCell, searchFormulas) {
      if (!searchValue) {
        throw new Error('FIND requires search value as first argument');
      }

      const options = {
        matchCase: matchCase === 1 || matchCase === '1' || matchCase === true,
        matchEntireCell: matchEntireCell === 1 || matchEntireCell === '1' || matchEntireCell === true,
        searchFormulas: searchFormulas === 1 || searchFormulas === '1' || searchFormulas === true
      };

      const results = model.find(searchValue, options);

      // Return as REXX stem array
      const stemArray = { 0: results.length };
      results.forEach((ref, index) => {
        stemArray[index + 1] = ref;
      });

      return stemArray;
    },

    /**
     * REPLACE - Replace all occurrences of a value
     * Usage: count = REPLACE("oldValue", "newValue")
     *        count = REPLACE("oldValue", "newValue", matchCase, matchEntireCell, searchFormulas)
     * Returns: Number of replacements made
     */
    REPLACE: async function(searchValue, replaceValue, matchCase, matchEntireCell, searchFormulas) {
      if (!searchValue || replaceValue === undefined) {
        throw new Error('REPLACE requires search value and replace value');
      }

      const options = {
        matchCase: matchCase === 1 || matchCase === '1' || matchCase === true,
        matchEntireCell: matchEntireCell === 1 || matchEntireCell === '1' || matchEntireCell === true,
        searchFormulas: searchFormulas === 1 || searchFormulas === '1' || searchFormulas === true
      };

      const count = model.replace(searchValue, replaceValue, options, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return count;
    },

    /**
     * HIDEROW - Hide a row
     * Usage: CALL HIDEROW(5)
     */
    HIDEROW: function(rowNum) {
      if (!rowNum) {
        throw new Error('HIDEROW requires row number as argument');
      }

      const rowNumber = parseInt(rowNum, 10);
      if (isNaN(rowNumber)) {
        throw new Error('HIDEROW requires a valid row number');
      }

      model.hideRow(rowNumber);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * UNHIDEROW - Unhide a row
     * Usage: CALL UNHIDEROW(5)
     */
    UNHIDEROW: function(rowNum) {
      if (!rowNum) {
        throw new Error('UNHIDEROW requires row number as argument');
      }

      const rowNumber = parseInt(rowNum, 10);
      if (isNaN(rowNumber)) {
        throw new Error('UNHIDEROW requires a valid row number');
      }

      model.unhideRow(rowNumber);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * HIDECOLUMN - Hide a column
     * Usage: CALL HIDECOLUMN(5)
     *        CALL HIDECOLUMN("E")
     */
    HIDECOLUMN: function(colNum) {
      if (!colNum) {
        throw new Error('HIDECOLUMN requires column number or letter as argument');
      }

      model.hideColumn(colNum);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * UNHIDECOLUMN - Unhide a column
     * Usage: CALL UNHIDECOLUMN(5)
     *        CALL UNHIDECOLUMN("E")
     */
    UNHIDECOLUMN: function(colNum) {
      if (!colNum) {
        throw new Error('UNHIDECOLUMN requires column number or letter as argument');
      }

      model.unhideColumn(colNum);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * ISROWHIDDEN - Check if a row is hidden
     * Usage: hidden = ISROWHIDDEN(5)
     * Returns: 1 if hidden, 0 if not
     */
    ISROWHIDDEN: function(rowNum) {
      if (!rowNum) {
        throw new Error('ISROWHIDDEN requires row number as argument');
      }

      const rowNumber = parseInt(rowNum, 10);
      if (isNaN(rowNumber)) {
        throw new Error('ISROWHIDDEN requires a valid row number');
      }

      return model.isRowHidden(rowNumber) ? 1 : 0;
    },

    /**
     * ISCOLUMNHIDDEN - Check if a column is hidden
     * Usage: hidden = ISCOLUMNHIDDEN(5)
     *        hidden = ISCOLUMNHIDDEN("E")
     * Returns: 1 if hidden, 0 if not
     */
    ISCOLUMNHIDDEN: function(colNum) {
      if (!colNum) {
        throw new Error('ISCOLUMNHIDDEN requires column number or letter as argument');
      }

      return model.isColumnHidden(colNum) ? 1 : 0;
    },

    /**
     * SETCOLWIDTH - Set column width
     * Usage: CALL SETCOLWIDTH(5, 150)
     *        CALL SETCOLWIDTH("E", 150)
     */
    SETCOLWIDTH: function(colNum, width) {
      if (!colNum || !width) {
        throw new Error('SETCOLWIDTH requires column number/letter and width as arguments');
      }

      const widthNum = parseFloat(width);
      if (isNaN(widthNum) || widthNum <= 0) {
        throw new Error('SETCOLWIDTH width must be a positive number');
      }

      model.setColumnWidth(colNum, widthNum);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * GETCOLWIDTH - Get column width
     * Usage: width = GETCOLWIDTH(5)
     *        width = GETCOLWIDTH("E")
     * Returns: Width in pixels (default 100 if not set)
     */
    GETCOLWIDTH: function(colNum) {
      if (!colNum) {
        throw new Error('GETCOLWIDTH requires column number or letter as argument');
      }

      return model.getColumnWidth(colNum);
    },

    /**
     * SETROWHEIGHT - Set row height
     * Usage: CALL SETROWHEIGHT(5, 50)
     */
    SETROWHEIGHT: function(rowNum, height) {
      if (!rowNum || !height) {
        throw new Error('SETROWHEIGHT requires row number and height as arguments');
      }

      const rowNumber = parseInt(rowNum, 10);
      if (isNaN(rowNumber)) {
        throw new Error('SETROWHEIGHT requires a valid row number');
      }

      const heightNum = parseFloat(height);
      if (isNaN(heightNum) || heightNum <= 0) {
        throw new Error('SETROWHEIGHT height must be a positive number');
      }

      model.setRowHeight(rowNumber, heightNum);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * GETROWHEIGHT - Get row height
     * Usage: height = GETROWHEIGHT(5)
     * Returns: Height in pixels (default 32 if not set)
     */
    GETROWHEIGHT: function(rowNum) {
      if (!rowNum) {
        throw new Error('GETROWHEIGHT requires row number as argument');
      }

      const rowNumber = parseInt(rowNum, 10);
      if (isNaN(rowNumber)) {
        throw new Error('GETROWHEIGHT requires a valid row number');
      }

      return model.getRowHeight(rowNumber);
    },

    /**
     * DEFINENAMEDRANGE - Define a named range
     * Usage: CALL DEFINENAMEDRANGE("SalesData", "A1:B10")
     *        CALL DEFINENAMEDRANGE("TaxRate", "A1")
     */
    DEFINENAMEDRANGE: function(name, rangeRef) {
      if (!name || !rangeRef) {
        throw new Error('DEFINENAMEDRANGE requires name and range reference');
      }

      model.defineNamedRange(name, rangeRef);

      return 'OK';
    },

    /**
     * DELETENAMEDRANGE - Delete a named range
     * Usage: CALL DELETENAMEDRANGE("SalesData")
     */
    DELETENAMEDRANGE: function(name) {
      if (!name) {
        throw new Error('DELETENAMEDRANGE requires name as argument');
      }

      model.deleteNamedRange(name);

      return 'OK';
    },

    /**
     * GETNAMEDRANGE - Get a named range reference
     * Usage: rangeRef = GETNAMEDRANGE("SalesData")
     * Returns: Range reference or empty string if not found
     */
    GETNAMEDRANGE: function(name) {
      if (!name) {
        throw new Error('GETNAMEDRANGE requires name as argument');
      }

      return model.getNamedRange(name) || '';
    },

    /**
     * GETALLNAMEDRANGES - Get all named ranges
     * Usage: ranges = GETALLNAMEDRANGES()
     * Returns: REXX stem array with name.value pairs
     */
    GETALLNAMEDRANGES: function() {
      const ranges = model.getAllNamedRanges();
      const names = Object.keys(ranges);

      // Return as REXX stem array with compound variables
      const result = { 0: names.length };
      names.forEach((name, index) => {
        const idx = index + 1;
        result[`${idx}.NAME`] = name;
        result[`${idx}.RANGE`] = ranges[name];
      });

      return result;
    },

    /**
     * FREEZEPANES - Freeze rows and columns
     * Usage: CALL FREEZEPANES(2, 3)  -- Freeze 2 rows, 3 columns
     */
    FREEZEPANES: function(rows, cols) {
      const rowNum = parseInt(rows, 10) || 0;
      const colNum = parseInt(cols, 10) || 0;

      model.freezePanes(rowNum, colNum);

      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * UNFREEZEPANES - Unfreeze all panes
     * Usage: CALL UNFREEZEPANES()
     */
    UNFREEZEPANES: function() {
      model.unfreezePanes();

      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return 'OK';
    },

    /**
     * GETFROZENPANES - Get frozen pane settings
     * Usage: frozen = GETFROZENPANES()
     * Returns: REXX stem array with frozen.ROWS and frozen.COLUMNS
     */
    GETFROZENPANES: function() {
      const frozen = model.getFrozenPanes();
      return {
        ROWS: frozen.rows,
        COLUMNS: frozen.columns
      };
    },

    /**
     * SETCELLVALIDATION - Set validation rules for a cell
     * Usage: CALL SETCELLVALIDATION("A1", "list", "Red,Green,Blue")
     *        CALL SETCELLVALIDATION("B1", "number", 0, 100)
     */
    SETCELLVALIDATION: function(cellRef, type, ...args) {
      if (!cellRef || !type) {
        throw new Error('SETCELLVALIDATION requires cell reference and validation type');
      }

      let validation = { type };

      switch (type) {
        case 'list':
          // args[0] is comma-separated list
          validation.values = String(args[0] || '').split(',');
          break;
        case 'number':
          validation.min = args[0] !== undefined ? parseFloat(args[0]) : undefined;
          validation.max = args[1] !== undefined ? parseFloat(args[1]) : undefined;
          break;
        case 'text':
          validation.minLength = args[0] !== undefined ? parseInt(args[0], 10) : undefined;
          validation.maxLength = args[1] !== undefined ? parseInt(args[1], 10) : undefined;
          validation.pattern = args[2];
          break;
      }

      model.setCellValidation(cellRef, validation);
      return 'OK';
    },

    /**
     * CLEARCELLVALIDATION - Clear validation rules for a cell
     * Usage: CALL CLEARCELLVALIDATION("A1")
     */
    CLEARCELLVALIDATION: function(cellRef) {
      if (!cellRef) {
        throw new Error('CLEARCELLVALIDATION requires cell reference');
      }

      model.setCellValidation(cellRef, null);
      return 'OK';
    },

    /**
     * VALIDATECELL - Validate a cell value
     * Usage: result = VALIDATECELL("A1", "Red")
     * Returns: 1 if valid, 0 if invalid
     */
    VALIDATECELL: function(cellRef, value) {
      if (!cellRef) {
        throw new Error('VALIDATECELL requires cell reference');
      }

      const result = model.validateCellValue(cellRef, value);
      return result.valid ? 1 : 0;
    },

    /**
     * UNDO - Undo the last action
     * Usage: success = UNDO()
     * Returns: 1 if undo performed, 0 if nothing to undo
     */
    UNDO: async function() {
      const success = model.undo(adapter);

      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return success ? 1 : 0;
    },

    /**
     * REDO - Redo the last undone action
     * Usage: success = REDO()
     * Returns: 1 if redo performed, 0 if nothing to redo
     */
    REDO: async function() {
      const success = model.redo(adapter);

      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return success ? 1 : 0;
    },

    /**
     * CANUNDO - Check if undo is available
     * Usage: can = CANUNDO()
     * Returns: 1 if undo available, 0 if not
     */
    CANUNDO: function() {
      return model.canUndo() ? 1 : 0;
    },

    /**
     * CANREDO - Check if redo is available
     * Usage: can = CANREDO()
     * Returns: 1 if redo available, 0 if not
     */
    CANREDO: function() {
      return model.canRedo() ? 1 : 0;
    },

    /**
     * ADDSHEET - Add a new sheet to the spreadsheet
     * Usage: CALL ADDSHEET("Sheet2")
     *        sheetName = ADDSHEET("DataSheet")
     * Returns: Sheet name
     */
    ADDSHEET: function(sheetName) {
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('ADDSHEET requires sheet name as argument');
      }

      model.addSheet(sheetName);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return sheetName;
    },

    /**
     * DELETESHEET - Delete a sheet from the spreadsheet
     * Usage: CALL DELETESHEET("Sheet2")
     */
    DELETESHEET: function(sheetName) {
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('DELETESHEET requires sheet name as argument');
      }

      model.deleteSheet(sheetName);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return '';
    },

    /**
     * RENAMESHEET - Rename a sheet
     * Usage: CALL RENAMESHEET("Sheet1", "MainData")
     */
    RENAMESHEET: function(oldName, newName) {
      if (!oldName || typeof oldName !== 'string') {
        throw new Error('RENAMESHEET requires old sheet name as first argument');
      }
      if (!newName || typeof newName !== 'string') {
        throw new Error('RENAMESHEET requires new sheet name as second argument');
      }

      model.renameSheet(oldName, newName);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return newName;
    },

    /**
     * SETACTIVESHEET - Set the active sheet
     * Usage: CALL SETACTIVESHEET("Sheet2")
     */
    SETACTIVESHEET: function(sheetName) {
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('SETACTIVESHEET requires sheet name as argument');
      }

      model.setActiveSheet(sheetName);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return sheetName;
    },

    /**
     * GETACTIVESHEET - Get the name of the active sheet
     * Usage: sheetName = GETACTIVESHEET()
     */
    GETACTIVESHEET: function() {
      return model.getActiveSheetName();
    },

    /**
     * GETSHEETNAMES - Get list of all sheet names
     * Usage: sheets = GETSHEETNAMES()
     * Returns: REXX stem array with sheet names
     */
    GETSHEETNAMES: function() {
      const sheetNames = model.getSheetNames();

      // Return as REXX stem array
      const result = { 0: sheetNames.length };
      sheetNames.forEach((name, index) => {
        result[index + 1] = name;
      });

      return result;
    },

    /**
     * APPLYROWFILTER - Filter rows by column criteria
     * Usage: CALL APPLYROWFILTER("A", "apple")
     *        CALL APPLYROWFILTER(1, "test")
     */
    APPLYROWFILTER: function(columnNum, criteria) {
      if (columnNum === undefined || columnNum === null) {
        throw new Error('APPLYROWFILTER requires column as first argument (letter or number)');
      }
      if (criteria === undefined || criteria === null) {
        throw new Error('APPLYROWFILTER requires criteria as second argument');
      }

      model.applyRowFilter(columnNum, String(criteria));

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return '';
    },

    /**
     * CLEARROWFILTER - Clear all row filters
     * Usage: CALL CLEARROWFILTER()
     */
    CLEARROWFILTER: function() {
      model.clearRowFilter();

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return '';
    },

    /**
     * GETFILTERCRITERIA - Get current filter criteria
     * Usage: criteria = GETFILTERCRITERIA()
     * Returns: Filter criteria object or empty string if no filter
     */
    GETFILTERCRITERIA: function() {
      const criteria = model.getFilterCriteria();
      if (!criteria) {
        return '';
      }
      return JSON.stringify(criteria);
    },

    /**
     * ISROWVISIBLE - Check if a row is visible (not filtered out)
     * Usage: visible = ISROWVISIBLE(5)
     * Returns: 1 if visible, 0 if hidden by filter
     */
    ISROWVISIBLE: function(rowNum) {
      if (typeof rowNum !== 'number') {
        rowNum = parseInt(rowNum, 10);
      }
      if (isNaN(rowNum) || rowNum < 1) {
        throw new Error('ISROWVISIBLE requires row number as argument');
      }

      return model.isRowVisible(rowNum) ? '1' : '0';
    },

    /**
     * MOVECOLUMNLEFT - Move a column one position to the left
     * Usage: CALL MOVECOLUMNLEFT("B")
     *        CALL MOVECOLUMNLEFT(2)
     */
    MOVECOLUMNLEFT: async function(colNum) {
      if (colNum === undefined || colNum === null) {
        throw new Error('MOVECOLUMNLEFT requires column as argument (letter or number)');
      }

      await model.moveColumnLeft(colNum, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return '';
    },

    /**
     * MOVECOLUMNRIGHT - Move a column one position to the right
     * Usage: CALL MOVECOLUMNRIGHT("B")
     *        CALL MOVECOLUMNRIGHT(2)
     */
    MOVECOLUMNRIGHT: async function(colNum) {
      if (colNum === undefined || colNum === null) {
        throw new Error('MOVECOLUMNRIGHT requires column as argument (letter or number)');
      }

      await model.moveColumnRight(colNum, adapter);

      // Trigger UI update
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new CustomEvent('spreadsheet-update'));
      }

      return '';
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
        'FILLDOWN', 'FILLRIGHT', 'SORTRANGE', 'FIND', 'REPLACE',
        'SETWRAPTEXT', 'GETWRAPTEXT',
        'HIDEROW', 'UNHIDEROW', 'HIDECOLUMN', 'UNHIDECOLUMN',
        'ISROWHIDDEN', 'ISCOLUMNHIDDEN',
        'DEFINENAMEDRANGE', 'DELETENAMEDRANGE', 'GETNAMEDRANGE', 'GETALLNAMEDRANGES',
        'FREEZEPANES', 'UNFREEZEPANES', 'GETFROZENPANES',
        'SETCELLVALIDATION', 'CLEARCELLVALIDATION', 'VALIDATECELL',
        'UNDO', 'REDO', 'CANUNDO', 'CANREDO',
        'ADDSHEET', 'DELETESHEET', 'RENAMESHEET', 'SETACTIVESHEET', 'GETACTIVESHEET', 'GETSHEETNAMES',
        'APPLYROWFILTER', 'CLEARROWFILTER', 'GETFILTERCRITERIA', 'ISROWVISIBLE',
        'MOVECOLUMNLEFT', 'MOVECOLUMNRIGHT',
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
  ADDSHEET: {
    name: 'ADDSHEET',
    params: ['sheetName'],
    description: 'Add a new sheet to the spreadsheet',
    examples: [
      'CALL ADDSHEET("Sheet2")',
      'sheetName = ADDSHEET("DataSheet")'
    ]
  },
  DELETESHEET: {
    name: 'DELETESHEET',
    params: ['sheetName'],
    description: 'Delete a sheet from the spreadsheet',
    examples: [
      'CALL DELETESHEET("Sheet2")'
    ]
  },
  RENAMESHEET: {
    name: 'RENAMESHEET',
    params: ['oldName', 'newName'],
    description: 'Rename a sheet',
    examples: [
      'CALL RENAMESHEET("Sheet1", "MainData")'
    ]
  },
  SETACTIVESHEET: {
    name: 'SETACTIVESHEET',
    params: ['sheetName'],
    description: 'Set the active sheet',
    examples: [
      'CALL SETACTIVESHEET("Sheet2")'
    ]
  },
  GETACTIVESHEET: {
    name: 'GETACTIVESHEET',
    params: [],
    description: 'Get the name of the active sheet',
    examples: [
      'sheetName = GETACTIVESHEET()',
      'SAY "Current sheet:" GETACTIVESHEET()'
    ]
  },
  GETSHEETNAMES: {
    name: 'GETSHEETNAMES',
    params: [],
    description: 'Get list of all sheet names',
    examples: [
      'sheets = GETSHEETNAMES()',
      'SAY sheets.0 "sheets found"'
    ]
  },
  APPLYROWFILTER: {
    name: 'APPLYROWFILTER',
    params: ['columnNum', 'criteria'],
    description: 'Filter rows by column criteria (shows rows containing the text)',
    examples: [
      'CALL APPLYROWFILTER("A", "apple")',
      'CALL APPLYROWFILTER(1, "test")'
    ]
  },
  CLEARROWFILTER: {
    name: 'CLEARROWFILTER',
    params: [],
    description: 'Clear all row filters',
    examples: [
      'CALL CLEARROWFILTER()'
    ]
  },
  GETFILTERCRITERIA: {
    name: 'GETFILTERCRITERIA',
    params: [],
    description: 'Get current filter criteria (JSON string)',
    examples: [
      'criteria = GETFILTERCRITERIA()',
      'SAY "Filter:" GETFILTERCRITERIA()'
    ]
  },
  ISROWVISIBLE: {
    name: 'ISROWVISIBLE',
    params: ['rowNum'],
    description: 'Check if a row is visible (not filtered out)',
    examples: [
      'visible = ISROWVISIBLE(5)',
      'IF ISROWVISIBLE(10) THEN SAY "Row 10 is visible"'
    ]
  },
  MOVECOLUMNLEFT: {
    name: 'MOVECOLUMNLEFT',
    params: ['colNum'],
    description: 'Move a column one position to the left (swaps with previous column)',
    examples: [
      'CALL MOVECOLUMNLEFT("B")',
      'CALL MOVECOLUMNLEFT(2)'
    ]
  },
  MOVECOLUMNRIGHT: {
    name: 'MOVECOLUMNRIGHT',
    params: ['colNum'],
    description: 'Move a column one position to the right (swaps with next column)',
    examples: [
      'CALL MOVECOLUMNRIGHT("B")',
      'CALL MOVECOLUMNRIGHT(2)'
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
