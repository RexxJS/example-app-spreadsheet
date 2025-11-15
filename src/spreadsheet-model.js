/**
 * SpreadsheetModel - Core spreadsheet logic (no DOM, pure JS)
 *
 * Responsibilities:
 * - Cell storage and retrieval
 * - A1-style addressing (A1, B2, AA10, etc.)
 * - Track cell values vs expressions
 * - Dependency tracking for recalculation
 * - Evaluation order resolution
 */

class SpreadsheetModel {
    constructor(rows = 100, cols = 26) {
        this.rows = rows;
        this.cols = cols;

        // Multi-sheet support
        this.sheets = new Map(); // key: sheet name, value: sheet data
        this.activeSheetName = 'Sheet1'; // Current active sheet
        this.sheetOrder = ['Sheet1']; // Ordered list of sheet names

        // Initialize default sheet
        this._initializeSheet('Sheet1');

        this.setupScript = ''; // Page-level RexxJS code (REQUIRE statements, etc.)
        this.undoStack = []; // History for undo
        this.redoStack = []; // History for redo
        this.maxHistorySize = 100; // Limit undo/redo stack size
        this.recordingHistory = true; // Flag to enable/disable history recording
    }

    /**
     * Initialize a new sheet with default values
     */
    _initializeSheet(name) {
        this.sheets.set(name, {
            cells: new Map(), // key: "A1", value: {value, expression, dependencies}
            dependents: new Map(), // key: "A1", value: Set of cells that depend on A1
            evaluationInProgress: new Set(), // For circular reference detection
            hiddenRows: new Set(), // Set of hidden row numbers
            hiddenColumns: new Set(), // Set of hidden column numbers
            namedRanges: new Map(), // key: "MyRange", value: "A1:B5"
            frozenRows: 0, // Number of rows frozen at top
            frozenColumns: 0, // Number of columns frozen at left
            validations: new Map(), // key: "A1", value: validation rules
            filteredRows: null, // null = no filter, Set = visible rows
            filterCriteria: null, // Filter criteria object
            columnOrder: null // null = default order, Array = custom column order
        });
    }

    /**
     * Get the active sheet data
     */
    _getActiveSheet() {
        return this.sheets.get(this.activeSheetName);
    }

    /**
     * Get sheet data by name
     */
    _getSheet(sheetName) {
        return this.sheets.get(sheetName || this.activeSheetName);
    }

    /**
     * Legacy compatibility: expose cells, dependents, etc. from active sheet
     */
    get cells() {
        return this._getActiveSheet().cells;
    }
    set cells(value) {
        this._getActiveSheet().cells = value;
    }

    get dependents() {
        return this._getActiveSheet().dependents;
    }
    set dependents(value) {
        this._getActiveSheet().dependents = value;
    }

    get evaluationInProgress() {
        return this._getActiveSheet().evaluationInProgress;
    }
    set evaluationInProgress(value) {
        this._getActiveSheet().evaluationInProgress = value;
    }

    get hiddenRows() {
        return this._getActiveSheet().hiddenRows;
    }
    set hiddenRows(value) {
        this._getActiveSheet().hiddenRows = value;
    }

    get hiddenColumns() {
        return this._getActiveSheet().hiddenColumns;
    }
    set hiddenColumns(value) {
        this._getActiveSheet().hiddenColumns = value;
    }

    get namedRanges() {
        return this._getActiveSheet().namedRanges;
    }
    set namedRanges(value) {
        this._getActiveSheet().namedRanges = value;
    }

    get frozenRows() {
        return this._getActiveSheet().frozenRows;
    }
    set frozenRows(value) {
        this._getActiveSheet().frozenRows = value;
    }

    get frozenColumns() {
        return this._getActiveSheet().frozenColumns;
    }
    set frozenColumns(value) {
        this._getActiveSheet().frozenColumns = value;
    }

    get validations() {
        return this._getActiveSheet().validations;
    }
    set validations(value) {
        this._getActiveSheet().validations = value;
    }

    get filteredRows() {
        return this._getActiveSheet().filteredRows;
    }
    set filteredRows(value) {
        this._getActiveSheet().filteredRows = value;
    }

    get filterCriteria() {
        return this._getActiveSheet().filterCriteria;
    }
    set filterCriteria(value) {
        this._getActiveSheet().filterCriteria = value;
    }

    get columnOrder() {
        return this._getActiveSheet().columnOrder;
    }
    set columnOrder(value) {
        this._getActiveSheet().columnOrder = value;
    }

    /**
     * Convert column letter(s) to number (A=1, B=2, ..., Z=26, AA=27, etc.)
     */
    static colLetterToNumber(col) {
        col = col.toUpperCase();
        let result = 0;
        for (let i = 0; i < col.length; i++) {
            result = result * 26 + (col.charCodeAt(i) - 64);
        }
        return result;
    }

    /**
     * Convert column number to letter(s) (1=A, 2=B, ..., 26=Z, 27=AA, etc.)
     */
    static colNumberToLetter(num) {
        let result = '';
        while (num > 0) {
            let remainder = (num - 1) % 26;
            result = String.fromCharCode(65 + remainder) + result;
            num = Math.floor((num - 1) / 26);
        }
        return result;
    }

    /**
     * Parse cell reference like "A1" into {col: "A", row: 1}
     * Also supports cross-sheet references like "Sheet2.A1"
     */
    static parseCellRef(ref) {
        // Check for cross-sheet reference
        const sheetMatch = ref.match(/^([A-Za-z][A-Za-z0-9_]*)\.([A-Z]+\d+)$/i);
        if (sheetMatch) {
            return {
                sheet: sheetMatch[1],
                col: sheetMatch[2].match(/^([A-Z]+)(\d+)$/i)[1].toUpperCase(),
                row: parseInt(sheetMatch[2].match(/^([A-Z]+)(\d+)$/i)[2], 10)
            };
        }

        // Regular cell reference
        const match = ref.match(/^([A-Z]+)(\d+)$/i);
        if (!match) {
            throw new Error(`Invalid cell reference: ${ref}`);
        }
        return {
            col: match[1].toUpperCase(),
            row: parseInt(match[2], 10)
        };
    }

    /**
     * Format cell reference from {col, row, sheet?}
     */
    static formatCellRef(col, row, sheet) {
        if (typeof col === 'number') {
            col = SpreadsheetModel.colNumberToLetter(col);
        }
        const cellRef = `${col}${row}`;
        return sheet ? `${sheet}.${cellRef}` : cellRef;
    }

    /**
     * Validate sheet name (must be valid Rexx variable name)
     * - Must start with a letter
     * - Can contain only letters, numbers, and underscores
     * - No spaces or special characters
     */
    static isValidSheetName(name) {
        return /^[A-Za-z][A-Za-z0-9_]*$/.test(name);
    }

    /**
     * Sheet Management Methods
     */

    /**
     * Add a new sheet
     */
    addSheet(name) {
        if (!SpreadsheetModel.isValidSheetName(name)) {
            throw new Error('Sheet name must be a valid Rexx variable name (start with letter, no spaces)');
        }
        if (this.sheets.has(name)) {
            throw new Error(`Sheet "${name}" already exists`);
        }
        this._initializeSheet(name);
        this.sheetOrder.push(name);
        return name;
    }

    /**
     * Delete a sheet
     */
    deleteSheet(name) {
        if (this.sheets.size <= 1) {
            throw new Error('Cannot delete the last sheet');
        }
        if (!this.sheets.has(name)) {
            throw new Error(`Sheet "${name}" does not exist`);
        }

        this.sheets.delete(name);
        const index = this.sheetOrder.indexOf(name);
        if (index > -1) {
            this.sheetOrder.splice(index, 1);
        }

        // If we deleted the active sheet, switch to another one
        if (this.activeSheetName === name) {
            this.activeSheetName = this.sheetOrder[0];
        }
    }

    /**
     * Rename a sheet
     */
    renameSheet(oldName, newName) {
        if (!SpreadsheetModel.isValidSheetName(newName)) {
            throw new Error('Sheet name must be a valid Rexx variable name (start with letter, no spaces)');
        }
        if (!this.sheets.has(oldName)) {
            throw new Error(`Sheet "${oldName}" does not exist`);
        }
        if (this.sheets.has(newName)) {
            throw new Error(`Sheet "${newName}" already exists`);
        }

        const sheetData = this.sheets.get(oldName);
        this.sheets.delete(oldName);
        this.sheets.set(newName, sheetData);

        const index = this.sheetOrder.indexOf(oldName);
        if (index > -1) {
            this.sheetOrder[index] = newName;
        }

        if (this.activeSheetName === oldName) {
            this.activeSheetName = newName;
        }

        // TODO: Update formulas that reference this sheet
    }

    /**
     * Set the active sheet
     */
    setActiveSheet(name) {
        if (!this.sheets.has(name)) {
            throw new Error(`Sheet "${name}" does not exist`);
        }
        this.activeSheetName = name;
    }

    /**
     * Get all sheet names in order
     */
    getSheetNames() {
        return [...this.sheetOrder];
    }

    /**
     * Get the active sheet name
     */
    getActiveSheetName() {
        return this.activeSheetName;
    }

    /**
     * Get cell data
     */
    getCell(ref) {
        if (typeof ref === 'object') {
            ref = SpreadsheetModel.formatCellRef(ref.col, ref.row);
        }
        return this.cells.get(ref) || { value: '', expression: null, dependencies: [], chartScript: null, wrapText: false };
    }

    /**
     * Get cell value (computed result)
     */
    getCellValue(ref) {
        return this.getCell(ref).value;
    }

    /**
     * Get cell expression (formula)
     */
    getCellExpression(ref) {
        return this.getCell(ref).expression;
    }

    /**
     * Set cell content
     * If content starts with '=', treat as expression
     * Otherwise, treat as literal value
     */
    setCell(ref, content, rexxInterpreter = null, metadata = {}) {
        if (typeof ref === 'object') {
            ref = SpreadsheetModel.formatCellRef(ref.col, ref.row);
        }

        // Clear old dependencies
        const oldCell = this.cells.get(ref);
        if (oldCell && oldCell.dependencies) {
            oldCell.dependencies.forEach(dep => {
                const depSet = this.dependents.get(dep);
                if (depSet) {
                    depSet.delete(ref);
                }
            });
        }

        if (content === '' || content === null || content === undefined) {
            // Clear cell
            this.cells.delete(ref);
            this.propagateChanges(ref, rexxInterpreter);
            return;
        }

        const isExpression = typeof content === 'string' && content.trim().startsWith('=');

        if (isExpression) {
            // Store expression, evaluate later
            let expression = content.trim().substring(1).trim();

            // Resolve named ranges in the expression
            expression = this.resolveNamedRanges(expression);

            this.cells.set(ref, {
                value: '',
                expression: expression,
                dependencies: [],
                error: null,
                comment: metadata.comment || oldCell?.comment || '',
                format: metadata.format || oldCell?.format || '',
                chartScript: metadata.chartScript || oldCell?.chartScript || null,
                wrapText: metadata.wrapText !== undefined ? metadata.wrapText : (oldCell?.wrapText || false)
            });

            // Evaluate the expression
            if (rexxInterpreter) {
                this.evaluateCell(ref, rexxInterpreter);
            }
        } else {
            // Literal value
            this.cells.set(ref, {
                value: content,
                expression: null,
                dependencies: [],
                error: null,
                comment: metadata.comment || oldCell?.comment || '',
                format: metadata.format || oldCell?.format || '',
                chartScript: metadata.chartScript || oldCell?.chartScript || null,
                wrapText: metadata.wrapText !== undefined ? metadata.wrapText : (oldCell?.wrapText || false)
            });
        }

        // Propagate to dependent cells
        this.propagateChanges(ref, rexxInterpreter);
    }

    /**
     * Evaluate a cell's expression using RexxJS
     */
    async evaluateCell(ref, rexxInterpreter) {
        const cell = this.cells.get(ref);
        if (!cell || !cell.expression) {
            return;
        }

        // Check for circular references
        if (this.evaluationInProgress.has(ref)) {
            cell.error = 'Circular reference';
            cell.value = '#CIRCULAR!';
            return;
        }

        this.evaluationInProgress.add(ref);

        try {
            // Extract cell references from expression
            const dependencies = this.extractCellReferences(cell.expression);
            cell.dependencies = dependencies;

            // Update dependents map
            dependencies.forEach(dep => {
                if (!this.dependents.has(dep)) {
                    this.dependents.set(dep, new Set());
                }
                this.dependents.get(dep).add(ref);
            });

            // Evaluate expression via RexxJS
            const result = await rexxInterpreter.evaluate(cell.expression, this);
            cell.value = result;
            cell.error = null;
        } catch (error) {
            cell.error = error.message;
            cell.value = '#ERROR!';
        } finally {
            this.evaluationInProgress.delete(ref);
        }
    }

    /**
     * Extract cell references from an expression
     * Matches patterns like A1, B2, AA10, and cross-sheet references like Sheet2.A1
     */
    extractCellReferences(expression) {
        const crossSheetPattern = /\b([A-Za-z][A-Za-z0-9_]*)\.([A-Z]+\d+)\b/gi;
        const cellRefPattern = /\b([A-Z]+\d+)\b/g;

        const refs = new Set();

        // Extract cross-sheet references first
        let match;
        while ((match = crossSheetPattern.exec(expression)) !== null) {
            refs.add(match[0]); // Full match: Sheet2.A1
        }

        // Extract local cell references
        const cellMatches = expression.match(cellRefPattern);
        if (cellMatches) {
            cellMatches.forEach(ref => {
                // Only add if not already part of a cross-sheet reference
                if (!Array.from(refs).some(r => r.endsWith('.' + ref))) {
                    refs.add(ref);
                }
            });
        }

        return Array.from(refs);
    }

    /**
     * Propagate changes to dependent cells
     */
    async propagateChanges(ref, rexxInterpreter) {
        const deps = this.dependents.get(ref);
        if (!deps || !rexxInterpreter) {
            return;
        }

        for (const depRef of deps) {
            await this.evaluateCell(depRef, rexxInterpreter);
            await this.propagateChanges(depRef, rexxInterpreter);
        }
    }

    /**
     * Get all non-empty cells
     */
    getAllCells() {
        const result = [];
        for (const [ref, cell] of this.cells.entries()) {
            result.push({ ref, ...cell });
        }
        return result;
    }

    /**
     * Set cell metadata (comment, format, chartScript)
     */
    setCellMetadata(ref, metadata) {
        if (typeof ref === 'object') {
            ref = SpreadsheetModel.formatCellRef(ref.col, ref.row);
        }

        const cell = this.cells.get(ref);
        if (!cell) {
            // Create empty cell with metadata
            this.cells.set(ref, {
                value: '',
                expression: null,
                dependencies: [],
                error: null,
                comment: metadata.comment || '',
                format: metadata.format || '',
                chartScript: metadata.chartScript || null,
                wrapText: metadata.wrapText || false
            });
        } else {
            // Update existing cell
            if (metadata.comment !== undefined) {
                cell.comment = metadata.comment;
            }
            if (metadata.format !== undefined) {
                cell.format = metadata.format;
            }
            if (metadata.chartScript !== undefined) {
                cell.chartScript = metadata.chartScript;
            }
            if (metadata.wrapText !== undefined) {
                cell.wrapText = metadata.wrapText;
            }
        }
    }

    /**
     * Get setup script
     */
    getSetupScript() {
        return this.setupScript;
    }

    /**
     * Set setup script
     */
    setSetupScript(script) {
        this.setupScript = script || '';
    }

    /**
     * Export to JSON
     */
    toJSON() {
        const data = {
            version: 2, // Version 2 with multi-sheet support
            setupScript: this.setupScript,
            activeSheetName: this.activeSheetName,
            sheetOrder: [...this.sheetOrder],
            sheets: {}
        };

        // Export each sheet
        for (const [sheetName, sheet] of this.sheets.entries()) {
            const sheetData = {
                cells: {},
                hiddenRows: Array.from(sheet.hiddenRows),
                hiddenColumns: Array.from(sheet.hiddenColumns),
                namedRanges: Object.fromEntries(sheet.namedRanges),
                frozenRows: sheet.frozenRows,
                frozenColumns: sheet.frozenColumns,
                validations: Object.fromEntries(sheet.validations)
            };

            // Add filter criteria if present
            if (sheet.filterCriteria) {
                sheetData.filterCriteria = sheet.filterCriteria;
            }

            // Export cells
            for (const [ref, cell] of sheet.cells.entries()) {
                const cellData = {};

                if (cell.expression) {
                    cellData.content = '=' + cell.expression;
                } else if (cell.value !== '') {
                    cellData.content = cell.value;
                }

                // Add metadata if present
                if (cell.comment) {
                    cellData.comment = cell.comment;
                }
                if (cell.format) {
                    cellData.format = cell.format;
                }
                if (cell.chartScript) {
                    cellData.chartScript = cell.chartScript;
                }
                if (cell.wrapText) {
                    cellData.wrapText = cell.wrapText;
                }

                // Only store if there's content or metadata
                if (cellData.content || cellData.comment || cellData.format || cellData.chartScript || cellData.wrapText) {
                    // If only content, store as string for backward compatibility
                    if (Object.keys(cellData).length === 1 && cellData.content) {
                        sheetData.cells[ref] = cellData.content;
                    } else {
                        sheetData.cells[ref] = cellData;
                    }
                }
            }

            data.sheets[sheetName] = sheetData;
        }

        return data;
    }

    /**
     * Import from JSON
     */
    fromJSON(data, rexxInterpreter = null) {
        this.undoStack = [];
        this.redoStack = [];

        // Handle version 2 (multi-sheet) format
        if (data.version === 2 && data.sheets) {
            this.setupScript = data.setupScript || '';
            this.sheets.clear();
            this.sheetOrder = data.sheetOrder || Object.keys(data.sheets);
            this.activeSheetName = data.activeSheetName || this.sheetOrder[0] || 'Sheet1';

            // Import each sheet
            for (const [sheetName, sheetData] of Object.entries(data.sheets)) {
                this._initializeSheet(sheetName);
                const sheet = this.sheets.get(sheetName);

                // Restore hidden rows/columns
                if (sheetData.hiddenRows) {
                    sheetData.hiddenRows.forEach(row => sheet.hiddenRows.add(row));
                }
                if (sheetData.hiddenColumns) {
                    sheetData.hiddenColumns.forEach(col => sheet.hiddenColumns.add(col));
                }

                // Restore named ranges
                if (sheetData.namedRanges) {
                    Object.entries(sheetData.namedRanges).forEach(([name, range]) => {
                        sheet.namedRanges.set(name, range);
                    });
                }

                // Restore frozen panes
                sheet.frozenRows = sheetData.frozenRows || 0;
                sheet.frozenColumns = sheetData.frozenColumns || 0;

                // Restore validations
                if (sheetData.validations) {
                    Object.entries(sheetData.validations).forEach(([ref, validation]) => {
                        sheet.validations.set(ref, validation);
                    });
                }

                // Import cells first (filter needs cells to be present)
                const cells = sheetData.cells || {};
                const savedActive = this.activeSheetName;
                this.activeSheetName = sheetName;

                for (const [ref, cellData] of Object.entries(cells)) {
                    // Handle both string format and object format
                    if (typeof cellData === 'string') {
                        this.setCell(ref, cellData, rexxInterpreter);
                    } else {
                        const metadata = {
                            comment: cellData.comment || '',
                            format: cellData.format || '',
                            chartScript: cellData.chartScript || null,
                            wrapText: cellData.wrapText || false
                        };
                        this.setCell(ref, cellData.content || '', rexxInterpreter, metadata);
                    }
                }

                // Restore filter criteria after cells are imported
                if (sheetData.filterCriteria) {
                    this.applyRowFilter(sheetData.filterCriteria.column, sheetData.filterCriteria.criteria);
                }

                this.activeSheetName = savedActive;
            }
        }
        // Handle version 1 (single sheet) format - backward compatibility
        else if (data.setupScript !== undefined) {
            this.setupScript = data.setupScript || '';
            this.sheets.clear();
            this.sheetOrder = ['Sheet1'];
            this.activeSheetName = 'Sheet1';
            this._initializeSheet('Sheet1');

            // Restore hidden rows/columns
            if (data.hiddenRows) {
                data.hiddenRows.forEach(row => this.hiddenRows.add(row));
            }
            if (data.hiddenColumns) {
                data.hiddenColumns.forEach(col => this.hiddenColumns.add(col));
            }

            // Restore named ranges
            if (data.namedRanges) {
                Object.entries(data.namedRanges).forEach(([name, range]) => {
                    this.namedRanges.set(name, range);
                });
            }

            // Restore frozen panes
            this.frozenRows = data.frozenRows || 0;
            this.frozenColumns = data.frozenColumns || 0;

            // Restore validations
            if (data.validations) {
                Object.entries(data.validations).forEach(([ref, validation]) => {
                    this.validations.set(ref, validation);
                });
            }

            const cells = data.cells || {};
            for (const [ref, cellData] of Object.entries(cells)) {
                // Handle both string format and object format
                if (typeof cellData === 'string') {
                    this.setCell(ref, cellData, rexxInterpreter);
                } else {
                    const metadata = {
                        comment: cellData.comment || '',
                        format: cellData.format || '',
                        chartScript: cellData.chartScript || null,
                        wrapText: cellData.wrapText || false
                    };
                    this.setCell(ref, cellData.content || '', rexxInterpreter, metadata);
                }
            }
        } else {
            // Very old format - all entries are cells
            this.setupScript = '';
            this.sheets.clear();
            this.sheetOrder = ['Sheet1'];
            this.activeSheetName = 'Sheet1';
            this._initializeSheet('Sheet1');

            for (const [ref, content] of Object.entries(data)) {
                this.setCell(ref, content, rexxInterpreter);
            }
        }
    }

    /**
     * Insert a row at the specified position
     * Shifts all rows at or below the position down by 1
     */
    insertRow(rowNum, rexxInterpreter = null) {
        if (rowNum < 1 || rowNum > this.rows) {
            throw new Error(`Invalid row number: ${rowNum}. Must be between 1 and ${this.rows}`);
        }

        // Create a new map to hold updated cells
        const newCells = new Map();

        // Move cells down, starting from the bottom to avoid overwrites
        for (const [ref, cell] of this.cells.entries()) {
            const { col, row } = SpreadsheetModel.parseCellRef(ref);
            const cellCopy = { ...cell };

            // Adjust formula references if this cell has an expression
            if (cellCopy.expression) {
                cellCopy.expression = this._adjustCellReferencesInExpression(
                    cellCopy.expression,
                    'insertRow',
                    rowNum
                );
            }

            if (row >= rowNum) {
                // Shift this cell down
                const newRef = SpreadsheetModel.formatCellRef(col, row + 1);
                newCells.set(newRef, cellCopy);
            } else {
                // Keep this cell where it is
                newCells.set(ref, cellCopy);
            }
        }

        // Replace cells map
        this.cells = newCells;

        // Rebuild dependents map
        this._rebuildDependents();

        // Recalculate all formulas if interpreter provided
        if (rexxInterpreter) {
            this._recalculateAll(rexxInterpreter);
        }
    }

    /**
     * Delete a row at the specified position
     * Shifts all rows below the position up by 1
     */
    deleteRow(rowNum, rexxInterpreter = null) {
        if (rowNum < 1 || rowNum > this.rows) {
            throw new Error(`Invalid row number: ${rowNum}. Must be between 1 and ${this.rows}`);
        }

        // Create a new map to hold updated cells
        const newCells = new Map();

        // Delete row and shift cells up
        for (const [ref, cell] of this.cells.entries()) {
            const { col, row } = SpreadsheetModel.parseCellRef(ref);
            if (row === rowNum) {
                // Skip this cell (delete it)
                continue;
            }

            const cellCopy = { ...cell };

            // Adjust formula references if this cell has an expression
            if (cellCopy.expression) {
                cellCopy.expression = this._adjustCellReferencesInExpression(
                    cellCopy.expression,
                    'deleteRow',
                    rowNum
                );
            }

            if (row > rowNum) {
                // Shift this cell up
                const newRef = SpreadsheetModel.formatCellRef(col, row - 1);
                newCells.set(newRef, cellCopy);
            } else {
                // Keep this cell where it is
                newCells.set(ref, cellCopy);
            }
        }

        // Replace cells map
        this.cells = newCells;

        // Rebuild dependents map
        this._rebuildDependents();

        // Recalculate all formulas if interpreter provided
        if (rexxInterpreter) {
            this._recalculateAll(rexxInterpreter);
        }
    }

    /**
     * Insert a column at the specified position
     * Shifts all columns at or to the right of the position right by 1
     */
    insertColumn(colNum, rexxInterpreter = null) {
        // Convert column number to letter if needed
        if (typeof colNum === 'string') {
            colNum = SpreadsheetModel.colLetterToNumber(colNum);
        }

        if (colNum < 1 || colNum > this.cols) {
            throw new Error(`Invalid column number: ${colNum}. Must be between 1 and ${this.cols}`);
        }

        // Create a new map to hold updated cells
        const newCells = new Map();

        // Move cells right, processing from right to left to avoid overwrites
        for (const [ref, cell] of this.cells.entries()) {
            const { col, row } = SpreadsheetModel.parseCellRef(ref);
            const currentColNum = SpreadsheetModel.colLetterToNumber(col);
            const cellCopy = { ...cell };

            // Adjust formula references if this cell has an expression
            if (cellCopy.expression) {
                cellCopy.expression = this._adjustCellReferencesInExpression(
                    cellCopy.expression,
                    'insertColumn',
                    colNum
                );
            }

            if (currentColNum >= colNum) {
                // Shift this cell right
                const newRef = SpreadsheetModel.formatCellRef(currentColNum + 1, row);
                newCells.set(newRef, cellCopy);
            } else {
                // Keep this cell where it is
                newCells.set(ref, cellCopy);
            }
        }

        // Replace cells map
        this.cells = newCells;

        // Rebuild dependents map
        this._rebuildDependents();

        // Recalculate all formulas if interpreter provided
        if (rexxInterpreter) {
            this._recalculateAll(rexxInterpreter);
        }
    }

    /**
     * Delete a column at the specified position
     * Shifts all columns to the right of the position left by 1
     */
    deleteColumn(colNum, rexxInterpreter = null) {
        // Convert column number to letter if needed
        if (typeof colNum === 'string') {
            colNum = SpreadsheetModel.colLetterToNumber(colNum);
        }

        if (colNum < 1 || colNum > this.cols) {
            throw new Error(`Invalid column number: ${colNum}. Must be between 1 and ${this.cols}`);
        }

        // Create a new map to hold updated cells
        const newCells = new Map();

        // Delete column and shift cells left
        for (const [ref, cell] of this.cells.entries()) {
            const { col, row } = SpreadsheetModel.parseCellRef(ref);
            const currentColNum = SpreadsheetModel.colLetterToNumber(col);

            if (currentColNum === colNum) {
                // Skip this cell (delete it)
                continue;
            }

            const cellCopy = { ...cell };

            // Adjust formula references if this cell has an expression
            if (cellCopy.expression) {
                cellCopy.expression = this._adjustCellReferencesInExpression(
                    cellCopy.expression,
                    'deleteColumn',
                    colNum
                );
            }

            if (currentColNum > colNum) {
                // Shift this cell left
                const newRef = SpreadsheetModel.formatCellRef(currentColNum - 1, row);
                newCells.set(newRef, cellCopy);
            } else {
                // Keep this cell where it is
                newCells.set(ref, cellCopy);
            }
        }

        // Replace cells map
        this.cells = newCells;

        // Rebuild dependents map
        this._rebuildDependents();

        // Recalculate all formulas if interpreter provided
        if (rexxInterpreter) {
            this._recalculateAll(rexxInterpreter);
        }
    }

    /**
     * Rebuild the dependents map from scratch based on current cells
     */
    _rebuildDependents() {
        this.dependents.clear();

        for (const [ref, cell] of this.cells.entries()) {
            if (cell.expression) {
                const dependencies = this.extractCellReferences(cell.expression);
                cell.dependencies = dependencies;

                dependencies.forEach(dep => {
                    if (!this.dependents.has(dep)) {
                        this.dependents.set(dep, new Set());
                    }
                    this.dependents.get(dep).add(ref);
                });
            }
        }
    }

    /**
     * Recalculate all formula cells
     */
    async _recalculateAll(rexxInterpreter) {
        for (const [ref, cell] of this.cells.entries()) {
            if (cell.expression) {
                await this.evaluateCell(ref, rexxInterpreter);
            }
        }
    }

    /**
     * Adjust cell references in an expression based on row/column operations
     * @param {string} expression - The formula expression
     * @param {string} operation - 'insertRow', 'deleteRow', 'insertColumn', 'deleteColumn'
     * @param {number} position - The row/column number where the operation occurs
     * @returns {string} - The adjusted expression
     */
    _adjustCellReferencesInExpression(expression, operation, position) {
        if (!expression) return expression;

        // Match both single cell references (A1) and range references (A1:B5)
        // This regex captures cell references with optional $ for absolute references
        const cellRefPattern = /(\$?)([A-Z]+)(\$?)(\d+)/g;

        return expression.replace(cellRefPattern, (match, colAbs, col, rowAbs, row) => {
            const rowNum = parseInt(row, 10);
            const colNum = SpreadsheetModel.colLetterToNumber(col);

            let newRow = rowNum;
            let newCol = col;

            switch (operation) {
                case 'insertRow':
                    // If the reference is at or below the insertion point, shift down
                    if (!rowAbs && rowNum >= position) {
                        newRow = rowNum + 1;
                    }
                    break;

                case 'deleteRow':
                    // If the reference is the deleted row, mark as invalid
                    if (rowNum === position) {
                        return '#REF!';
                    }
                    // If the reference is below the deleted row, shift up
                    if (!rowAbs && rowNum > position) {
                        newRow = rowNum - 1;
                    }
                    break;

                case 'insertColumn':
                    // If the reference is at or right of the insertion point, shift right
                    if (!colAbs && colNum >= position) {
                        newCol = SpreadsheetModel.colNumberToLetter(colNum + 1);
                    }
                    break;

                case 'deleteColumn':
                    // If the reference is the deleted column, mark as invalid
                    if (colNum === position) {
                        return '#REF!';
                    }
                    // If the reference is right of the deleted column, shift left
                    if (!colAbs && colNum > position) {
                        newCol = SpreadsheetModel.colNumberToLetter(colNum - 1);
                    }
                    break;
            }

            // Reconstruct the cell reference with absolute markers if present
            return `${colAbs}${newCol}${rowAbs}${newRow}`;
        });
    }

    /**
     * Sort a range of cells by a specific column
     * @param {string} rangeRef - Range like "A1:C5"
     * @param {number|string} sortCol - Column to sort by (1-based index or letter)
     * @param {boolean} ascending - Sort direction (default: true)
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     */
    sortRange(rangeRef, sortCol, ascending = true, rexxInterpreter = null) {
        // Parse range reference
        const match = rangeRef.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
        if (!match) {
            throw new Error(`Invalid range reference: ${rangeRef}. Expected format: "A1:C5"`);
        }

        const startColLetter = match[1].toUpperCase();
        const startRow = parseInt(match[2], 10);
        const endColLetter = match[3].toUpperCase();
        const endRow = parseInt(match[4], 10);

        const startCol = SpreadsheetModel.colLetterToNumber(startColLetter);
        const endCol = SpreadsheetModel.colLetterToNumber(endColLetter);

        // Convert sortCol to number if needed
        const sortColNum = typeof sortCol === 'string' ?
            SpreadsheetModel.colLetterToNumber(sortCol) :
            sortCol;

        // Validate sort column is within range
        if (sortColNum < startCol || sortColNum > endCol) {
            throw new Error(`Sort column ${sortCol} is outside the range ${rangeRef}`);
        }

        // Extract all rows in the range
        const rows = [];
        for (let row = startRow; row <= endRow; row++) {
            const rowData = {
                rowNum: row,
                cells: {}
            };
            for (let col = startCol; col <= endCol; col++) {
                const ref = SpreadsheetModel.formatCellRef(col, row);
                const cell = this.getCell(ref);
                rowData.cells[col] = { ...cell };
            }
            rows.push(rowData);
        }

        // Sort rows by the sort column
        rows.sort((a, b) => {
            const aCell = a.cells[sortColNum];
            const bCell = b.cells[sortColNum];
            const aVal = aCell.value || '';
            const bVal = bCell.value || '';

            // Try numeric comparison first
            const aNum = parseFloat(aVal);
            const bNum = parseFloat(bVal);
            if (!isNaN(aNum) && !isNaN(bNum)) {
                return ascending ? aNum - bNum : bNum - aNum;
            }

            // Fall back to string comparison
            const aStr = String(aVal).toLowerCase();
            const bStr = String(bVal).toLowerCase();
            if (aStr < bStr) return ascending ? -1 : 1;
            if (aStr > bStr) return ascending ? 1 : -1;
            return 0;
        });

        // Write sorted rows back to the model
        for (let i = 0; i < rows.length; i++) {
            const targetRow = startRow + i;
            const sourceRow = rows[i];

            for (let col = startCol; col <= endCol; col++) {
                const targetRef = SpreadsheetModel.formatCellRef(col, targetRow);
                const sourceCell = sourceRow.cells[col];

                if (sourceCell.expression) {
                    // Preserve expression
                    const metadata = {
                        comment: sourceCell.comment || '',
                        format: sourceCell.format || '',
                        wrapText: sourceCell.wrapText || false
                    };
                    this.setCell(targetRef, '=' + sourceCell.expression, rexxInterpreter, metadata);
                } else if (sourceCell.value !== '') {
                    // Preserve value
                    const metadata = {
                        comment: sourceCell.comment || '',
                        format: sourceCell.format || '',
                        wrapText: sourceCell.wrapText || false
                    };
                    this.setCell(targetRef, sourceCell.value, rexxInterpreter, metadata);
                } else {
                    // Clear cell
                    this.setCell(targetRef, '', rexxInterpreter);
                }
            }
        }

        // Recalculate if interpreter provided
        if (rexxInterpreter) {
            this._recalculateAll(rexxInterpreter);
        }
    }

    /**
     * Fill down - Copy cell(s) down to a range
     * @param {string} sourceRef - Source cell or range (e.g., "A1" or "A1:B1")
     * @param {string} targetRangeRef - Target range (e.g., "A2:A10")
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     */
    fillDown(sourceRef, targetRangeRef, rexxInterpreter = null) {
        // Parse source range
        const sourceMatch = sourceRef.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
        if (!sourceMatch) {
            throw new Error(`Invalid source reference: ${sourceRef}`);
        }

        const sourceStartCol = sourceMatch[1].toUpperCase();
        const sourceStartRow = parseInt(sourceMatch[2], 10);
        const sourceEndCol = sourceMatch[3] ? sourceMatch[3].toUpperCase() : sourceStartCol;
        const sourceEndRow = sourceMatch[4] ? parseInt(sourceMatch[4], 10) : sourceStartRow;

        // Parse target range
        const targetMatch = targetRangeRef.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
        if (!targetMatch) {
            throw new Error(`Invalid target range: ${targetRangeRef}`);
        }

        const targetStartCol = targetMatch[1].toUpperCase();
        const targetStartRow = parseInt(targetMatch[2], 10);
        const targetEndCol = targetMatch[3].toUpperCase();
        const targetEndRow = parseInt(targetMatch[4], 10);

        const sourceColStart = SpreadsheetModel.colLetterToNumber(sourceStartCol);
        const sourceColEnd = SpreadsheetModel.colLetterToNumber(sourceEndCol);
        const targetColStart = SpreadsheetModel.colLetterToNumber(targetStartCol);
        const targetColEnd = SpreadsheetModel.colLetterToNumber(targetEndCol);

        const sourceWidth = sourceColEnd - sourceColStart + 1;
        const targetWidth = targetColEnd - targetColStart + 1;

        if (sourceWidth !== targetWidth) {
            throw new Error('Source and target must have the same number of columns');
        }

        // Fill down
        for (let row = targetStartRow; row <= targetEndRow; row++) {
            for (let colOffset = 0; colOffset < sourceWidth; colOffset++) {
                const sourceCol = sourceColStart + colOffset;
                const targetCol = targetColStart + colOffset;
                const sourceRowToUse = sourceStartRow + ((row - targetStartRow) % (sourceEndRow - sourceStartRow + 1));

                const sourceCellRef = SpreadsheetModel.formatCellRef(sourceCol, sourceRowToUse);
                const targetCellRef = SpreadsheetModel.formatCellRef(targetCol, row);

                const sourceCell = this.getCell(sourceCellRef);
                if (sourceCell.expression) {
                    // Adjust formula for new position
                    const rowOffset = row - sourceRowToUse;
                    const colOffset = targetCol - sourceCol;
                    const adjustedExpression = this._adjustFormulaForCopy(sourceCell.expression, rowOffset, colOffset);
                    this.setCell(targetCellRef, '=' + adjustedExpression, rexxInterpreter, {
                        format: sourceCell.format,
                        comment: sourceCell.comment,
                        wrapText: sourceCell.wrapText
                    });
                } else {
                    this.setCell(targetCellRef, sourceCell.value, rexxInterpreter, {
                        format: sourceCell.format,
                        comment: sourceCell.comment,
                        wrapText: sourceCell.wrapText
                    });
                }
            }
        }
    }

    /**
     * Fill right - Copy cell(s) right to a range
     * @param {string} sourceRef - Source cell or range (e.g., "A1" or "A1:A2")
     * @param {string} targetRangeRef - Target range (e.g., "B1:E1")
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     */
    fillRight(sourceRef, targetRangeRef, rexxInterpreter = null) {
        // Parse source range
        const sourceMatch = sourceRef.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
        if (!sourceMatch) {
            throw new Error(`Invalid source reference: ${sourceRef}`);
        }

        const sourceStartCol = sourceMatch[1].toUpperCase();
        const sourceStartRow = parseInt(sourceMatch[2], 10);
        const sourceEndCol = sourceMatch[3] ? sourceMatch[3].toUpperCase() : sourceStartCol;
        const sourceEndRow = sourceMatch[4] ? parseInt(sourceMatch[4], 10) : sourceStartRow;

        // Parse target range
        const targetMatch = targetRangeRef.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
        if (!targetMatch) {
            throw new Error(`Invalid target range: ${targetRangeRef}`);
        }

        const targetStartCol = targetMatch[1].toUpperCase();
        const targetStartRow = parseInt(targetMatch[2], 10);
        const targetEndCol = targetMatch[3].toUpperCase();
        const targetEndRow = parseInt(targetMatch[4], 10);

        const sourceColStart = SpreadsheetModel.colLetterToNumber(sourceStartCol);
        const sourceColEnd = SpreadsheetModel.colLetterToNumber(sourceEndCol);
        const targetColStart = SpreadsheetModel.colLetterToNumber(targetStartCol);
        const targetColEnd = SpreadsheetModel.colLetterToNumber(targetEndCol);

        const sourceHeight = sourceEndRow - sourceStartRow + 1;
        const targetHeight = targetEndRow - targetStartRow + 1;

        if (sourceHeight !== targetHeight) {
            throw new Error('Source and target must have the same number of rows');
        }

        // Fill right
        for (let col = targetColStart; col <= targetColEnd; col++) {
            for (let rowOffset = 0; rowOffset < sourceHeight; rowOffset++) {
                const sourceRow = sourceStartRow + rowOffset;
                const targetRow = targetStartRow + rowOffset;
                const sourceColToUse = sourceColStart + ((col - targetColStart) % (sourceColEnd - sourceColStart + 1));

                const sourceCellRef = SpreadsheetModel.formatCellRef(sourceColToUse, sourceRow);
                const targetCellRef = SpreadsheetModel.formatCellRef(col, targetRow);

                const sourceCell = this.getCell(sourceCellRef);
                if (sourceCell.expression) {
                    // Adjust formula for new position
                    const rowDiff = targetRow - sourceRow;
                    const colDiff = col - sourceColToUse;
                    const adjustedExpression = this._adjustFormulaForCopy(sourceCell.expression, rowDiff, colDiff);
                    this.setCell(targetCellRef, '=' + adjustedExpression, rexxInterpreter, {
                        format: sourceCell.format,
                        comment: sourceCell.comment,
                        wrapText: sourceCell.wrapText
                    });
                } else {
                    this.setCell(targetCellRef, sourceCell.value, rexxInterpreter, {
                        format: sourceCell.format,
                        comment: sourceCell.comment,
                        wrapText: sourceCell.wrapText
                    });
                }
            }
        }
    }

    /**
     * Adjust formula when copying (not inserting/deleting)
     * Only adjusts relative references
     */
    _adjustFormulaForCopy(expression, rowOffset, colOffset) {
        if (!expression) return expression;

        const cellRefPattern = /(\$?)([A-Z]+)(\$?)(\d+)/g;

        return expression.replace(cellRefPattern, (match, colAbs, col, rowAbs, row) => {
            const rowNum = parseInt(row, 10);
            const colNum = SpreadsheetModel.colLetterToNumber(col);

            const newRow = rowAbs ? rowNum : rowNum + rowOffset;
            const newCol = colAbs ? col : SpreadsheetModel.colNumberToLetter(colNum + colOffset);

            return `${colAbs}${newCol}${rowAbs}${newRow}`;
        });
    }

    /**
     * Find all cells matching criteria
     * @param {string} searchValue - Value to search for
     * @param {Object} options - { matchCase: boolean, matchEntireCell: boolean, searchFormulas: boolean }
     * @returns {Array} Array of cell references that match
     */
    find(searchValue, options = {}) {
        const {
            matchCase = false,
            matchEntireCell = false,
            searchFormulas = false
        } = options;

        const results = [];
        const searchStr = matchCase ? searchValue : searchValue.toLowerCase();

        for (const [ref, cell] of this.cells.entries()) {
            let textToSearch = searchFormulas && cell.expression ? cell.expression : cell.value;
            if (!matchCase) {
                textToSearch = String(textToSearch).toLowerCase();
            } else {
                textToSearch = String(textToSearch);
            }

            const matches = matchEntireCell
                ? textToSearch === searchStr
                : textToSearch.includes(searchStr);

            if (matches) {
                results.push(ref);
            }
        }

        return results;
    }

    /**
     * Replace all occurrences of a value
     * @param {string} searchValue - Value to search for
     * @param {string} replaceValue - Value to replace with
     * @param {Object} options - { matchCase, matchEntireCell, searchFormulas }
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     * @returns {number} Number of replacements made
     */
    replace(searchValue, replaceValue, options = {}, rexxInterpreter = null) {
        const {
            matchCase = false,
            matchEntireCell = false,
            searchFormulas = false
        } = options;

        let count = 0;

        for (const [ref, cell] of this.cells.entries()) {
            let textToReplace = searchFormulas && cell.expression ? cell.expression : cell.value;
            let originalText = textToReplace;

            if (matchEntireCell) {
                if (matchCase) {
                    if (textToReplace === searchValue) {
                        textToReplace = replaceValue;
                    }
                } else {
                    if (String(textToReplace).toLowerCase() === searchValue.toLowerCase()) {
                        textToReplace = replaceValue;
                    }
                }
            } else {
                if (matchCase) {
                    textToReplace = String(textToReplace).split(searchValue).join(replaceValue);
                } else {
                    const regex = new RegExp(searchValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
                    textToReplace = String(textToReplace).replace(regex, replaceValue);
                }
            }

            if (textToReplace !== originalText) {
                if (searchFormulas && cell.expression) {
                    this.setCell(ref, '=' + textToReplace, rexxInterpreter, {
                        format: cell.format,
                        comment: cell.comment,
                        wrapText: cell.wrapText
                    });
                } else {
                    this.setCell(ref, textToReplace, rexxInterpreter, {
                        format: cell.format,
                        comment: cell.comment,
                        wrapText: cell.wrapText
                    });
                }
                count++;
            }
        }

        return count;
    }

    /**
     * Hide a row
     * @param {number} rowNum - Row number to hide
     */
    hideRow(rowNum) {
        if (rowNum < 1 || rowNum > this.rows) {
            throw new Error(`Invalid row number: ${rowNum}`);
        }
        this.hiddenRows.add(rowNum);
    }

    /**
     * Unhide a row
     * @param {number} rowNum - Row number to unhide
     */
    unhideRow(rowNum) {
        this.hiddenRows.delete(rowNum);
    }

    /**
     * Hide a column
     * @param {number|string} colNum - Column number or letter to hide
     */
    hideColumn(colNum) {
        if (typeof colNum === 'string') {
            colNum = SpreadsheetModel.colLetterToNumber(colNum);
        }
        if (colNum < 1 || colNum > this.cols) {
            throw new Error(`Invalid column number: ${colNum}`);
        }
        this.hiddenColumns.add(colNum);
    }

    /**
     * Unhide a column
     * @param {number|string} colNum - Column number or letter to unhide
     */
    unhideColumn(colNum) {
        if (typeof colNum === 'string') {
            colNum = SpreadsheetModel.colLetterToNumber(colNum);
        }
        this.hiddenColumns.delete(colNum);
    }

    /**
     * Check if a row is hidden
     */
    isRowHidden(rowNum) {
        return this.hiddenRows.has(rowNum);
    }

    /**
     * Check if a column is hidden
     */
    isColumnHidden(colNum) {
        if (typeof colNum === 'string') {
            colNum = SpreadsheetModel.colLetterToNumber(colNum);
        }
        return this.hiddenColumns.has(colNum);
    }

    /**
     * Define a named range
     * @param {string} name - Name for the range (e.g., "SalesData")
     * @param {string} rangeRef - Range reference (e.g., "A1:B10")
     */
    defineNamedRange(name, rangeRef) {
        // Validate name (alphanumeric, underscore, must start with letter)
        if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(name)) {
            throw new Error('Named range must start with a letter and contain only letters, numbers, and underscores');
        }

        // Validate range reference
        const match = rangeRef.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
        if (!match) {
            throw new Error(`Invalid range reference: ${rangeRef}`);
        }

        this.namedRanges.set(name, rangeRef.toUpperCase());
    }

    /**
     * Delete a named range
     * @param {string} name - Name of the range to delete
     */
    deleteNamedRange(name) {
        this.namedRanges.delete(name);
    }

    /**
     * Get a named range reference
     * @param {string} name - Name of the range
     * @returns {string|null} Range reference or null if not found
     */
    getNamedRange(name) {
        return this.namedRanges.get(name) || null;
    }

    /**
     * Get all named ranges
     * @returns {Object} Object with name -> range mappings
     */
    getAllNamedRanges() {
        return Object.fromEntries(this.namedRanges);
    }

    /**
     * Resolve named ranges in an expression
     * @param {string} expression - Expression that may contain named ranges
     * @returns {string} Expression with named ranges replaced by cell references
     */
    resolveNamedRanges(expression) {
        if (!expression) return expression;

        // Replace named ranges with their cell references
        // Named ranges should match pattern: word boundary + name + word boundary
        for (const [name, rangeRef] of this.namedRanges.entries()) {
            const regex = new RegExp(`\\b${name}\\b`, 'g');
            expression = expression.replace(regex, rangeRef);
        }

        return expression;
    }

    /**
     * Freeze panes - lock rows and columns in place
     * @param {number} rows - Number of rows to freeze from top
     * @param {number} cols - Number of columns to freeze from left
     */
    freezePanes(rows, cols) {
        if (rows < 0 || rows > this.rows) {
            throw new Error(`Invalid rows to freeze: ${rows}`);
        }
        if (cols < 0 || cols > this.cols) {
            throw new Error(`Invalid columns to freeze: ${cols}`);
        }
        this.frozenRows = rows;
        this.frozenColumns = cols;
    }

    /**
     * Unfreeze all panes
     */
    unfreezePanes() {
        this.frozenRows = 0;
        this.frozenColumns = 0;
    }

    /**
     * Get frozen pane settings
     * @returns {Object} {rows, columns}
     */
    getFrozenPanes() {
        return {
            rows: this.frozenRows,
            columns: this.frozenColumns
        };
    }

    /**
     * Set cell validation rules
     * @param {string} cellRef - Cell reference
     * @param {Object} validation - Validation rules
     */
    setCellValidation(cellRef, validation) {
        if (!validation) {
            this.validations.delete(cellRef);
            return;
        }

        // Validate the validation object
        if (!validation.type) {
            throw new Error('Validation must have a type');
        }

        const validTypes = ['list', 'number', 'date', 'text', 'custom'];
        if (!validTypes.includes(validation.type)) {
            throw new Error(`Invalid validation type: ${validation.type}`);
        }

        this.validations.set(cellRef, validation);
    }

    /**
     * Get cell validation rules
     * @param {string} cellRef - Cell reference
     * @returns {Object|null} Validation rules or null
     */
    getCellValidation(cellRef) {
        return this.validations.get(cellRef) || null;
    }

    /**
     * Validate a cell value against its validation rules
     * @param {string} cellRef - Cell reference
     * @param {string} value - Value to validate
     * @returns {Object} {valid: boolean, message: string}
     */
    validateCellValue(cellRef, value) {
        const validation = this.validations.get(cellRef);
        if (!validation) {
            return { valid: true, message: '' };
        }

        switch (validation.type) {
            case 'list':
                if (!validation.values || !Array.isArray(validation.values)) {
                    return { valid: false, message: 'Invalid list validation configuration' };
                }
                const valid = validation.values.includes(value);
                return {
                    valid,
                    message: valid ? '' : `Value must be one of: ${validation.values.join(', ')}`
                };

            case 'number':
                const num = parseFloat(value);
                if (isNaN(num)) {
                    return { valid: false, message: 'Value must be a number' };
                }
                if (validation.min !== undefined && num < validation.min) {
                    return { valid: false, message: `Value must be >= ${validation.min}` };
                }
                if (validation.max !== undefined && num > validation.max) {
                    return { valid: false, message: `Value must be <= ${validation.max}` };
                }
                return { valid: true, message: '' };

            case 'text':
                const text = String(value);
                if (validation.minLength !== undefined && text.length < validation.minLength) {
                    return { valid: false, message: `Text must be at least ${validation.minLength} characters` };
                }
                if (validation.maxLength !== undefined && text.length > validation.maxLength) {
                    return { valid: false, message: `Text must be at most ${validation.maxLength} characters` };
                }
                if (validation.pattern) {
                    const regex = new RegExp(validation.pattern);
                    if (!regex.test(text)) {
                        return { valid: false, message: `Text must match pattern: ${validation.pattern}` };
                    }
                }
                return { valid: true, message: '' };

            case 'custom':
                // Custom validation with a function (stored as string)
                if (!validation.formula) {
                    return { valid: false, message: 'Custom validation requires a formula' };
                }
                // For now, return valid - would need interpreter to evaluate
                return { valid: true, message: '' };

            default:
                return { valid: true, message: '' };
        }
    }

    /**
     * Row Filtering Methods
     */

    /**
     * Apply a filter to show only rows that match criteria
     * @param {number} columnNum - Column to filter by (1-based)
     * @param {string|Function} criteria - Filter criteria (string, number, or function)
     */
    applyRowFilter(columnNum, criteria) {
        const sheet = this._getActiveSheet();
        const filteredRows = new Set();

        // Convert column number to letter if needed
        const colLetter = typeof columnNum === 'number' ?
            SpreadsheetModel.colNumberToLetter(columnNum) :
            columnNum;

        // Iterate through all rows
        for (let row = 1; row <= this.rows; row++) {
            const cellRef = SpreadsheetModel.formatCellRef(colLetter, row);
            const cell = this.getCell(cellRef);
            const value = cell.value;

            let matches = false;

            if (typeof criteria === 'function') {
                matches = criteria(value, row);
            } else if (typeof criteria === 'string') {
                // String matching (case-insensitive)
                matches = String(value).toLowerCase().includes(criteria.toLowerCase());
            } else {
                // Exact match
                matches = value == criteria;
            }

            if (matches) {
                filteredRows.add(row);
            }
        }

        sheet.filteredRows = filteredRows;
        sheet.filterCriteria = { column: columnNum, criteria };
    }

    /**
     * Clear all row filters
     */
    clearRowFilter() {
        const sheet = this._getActiveSheet();
        sheet.filteredRows = null;
        sheet.filterCriteria = null;
    }

    /**
     * Check if a row is visible (not filtered out)
     */
    isRowVisible(rowNum) {
        const sheet = this._getActiveSheet();
        if (sheet.filteredRows === null) {
            return true; // No filter applied
        }
        return sheet.filteredRows.has(rowNum);
    }

    /**
     * Get filter criteria
     */
    getFilterCriteria() {
        return this._getActiveSheet().filterCriteria;
    }

    /**
     * Column Reordering Methods
     */

    /**
     * Move a column left (swap with previous column)
     * @param {number|string} colNum - Column to move (1-based number or letter)
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     */
    moveColumnLeft(colNum, rexxInterpreter = null) {
        if (typeof colNum === 'string') {
            colNum = SpreadsheetModel.colLetterToNumber(colNum);
        }

        if (colNum <= 1) {
            throw new Error('Cannot move first column left');
        }

        this._swapColumns(colNum, colNum - 1, rexxInterpreter);
    }

    /**
     * Move a column right (swap with next column)
     * @param {number|string} colNum - Column to move (1-based number or letter)
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     */
    moveColumnRight(colNum, rexxInterpreter = null) {
        if (typeof colNum === 'string') {
            colNum = SpreadsheetModel.colLetterToNumber(colNum);
        }

        if (colNum >= this.cols) {
            throw new Error('Cannot move last column right');
        }

        this._swapColumns(colNum, colNum + 1, rexxInterpreter);
    }

    /**
     * Swap two columns
     * @param {number} col1 - First column number (1-based)
     * @param {number} col2 - Second column number (1-based)
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     */
    _swapColumns(col1, col2, rexxInterpreter = null) {
        const newCells = new Map();
        const col1Letter = SpreadsheetModel.colNumberToLetter(col1);
        const col2Letter = SpreadsheetModel.colNumberToLetter(col2);

        // Swap all cells in the two columns
        for (const [ref, cell] of this.cells.entries()) {
            const { col, row } = SpreadsheetModel.parseCellRef(ref);
            const colNum = SpreadsheetModel.colLetterToNumber(col);
            const cellCopy = { ...cell };

            // Adjust formula references if this cell has an expression
            if (cellCopy.expression) {
                cellCopy.expression = this._adjustCellReferencesForColumnSwap(
                    cellCopy.expression,
                    col1,
                    col2
                );
            }

            let newRef;
            if (colNum === col1) {
                // Move from col1 to col2
                newRef = SpreadsheetModel.formatCellRef(col2Letter, row);
            } else if (colNum === col2) {
                // Move from col2 to col1
                newRef = SpreadsheetModel.formatCellRef(col1Letter, row);
            } else {
                // Keep in same position
                newRef = ref;
            }

            newCells.set(newRef, cellCopy);
        }

        // Replace cells map
        this.cells = newCells;

        // Rebuild dependents map
        this._rebuildDependents();

        // Recalculate all formulas if interpreter provided
        if (rexxInterpreter) {
            this._recalculateAll(rexxInterpreter);
        }
    }

    /**
     * Adjust cell references in formulas when columns are swapped
     */
    _adjustCellReferencesForColumnSwap(expression, col1, col2) {
        if (!expression) return expression;

        const cellRefPattern = /(\$?)([A-Z]+)(\$?)(\d+)/g;

        return expression.replace(cellRefPattern, (match, colAbs, col, rowAbs, row) => {
            const colNum = SpreadsheetModel.colLetterToNumber(col);
            let newCol = col;

            if (!colAbs) { // Only adjust relative references
                if (colNum === col1) {
                    newCol = SpreadsheetModel.colNumberToLetter(col2);
                } else if (colNum === col2) {
                    newCol = SpreadsheetModel.colNumberToLetter(col1);
                }
            }

            return `${colAbs}${newCol}${rowAbs}${row}`;
        });
    }

    /**
     * Record a snapshot for undo
     */
    _recordSnapshot(action) {
        if (!this.recordingHistory) return;

        const snapshot = {
            action,
            cells: new Map(this.cells),
            hiddenRows: new Set(this.hiddenRows),
            hiddenColumns: new Set(this.hiddenColumns),
            namedRanges: new Map(this.namedRanges),
            frozenRows: this.frozenRows,
            frozenColumns: this.frozenColumns,
            validations: new Map(this.validations),
            timestamp: Date.now()
        };

        this.undoStack.push(snapshot);

        // Limit stack size
        if (this.undoStack.length > this.maxHistorySize) {
            this.undoStack.shift();
        }

        // Clear redo stack when new action is recorded
        this.redoStack = [];
    }

    /**
     * Undo the last action
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     * @returns {boolean} True if undo was performed
     */
    undo(rexxInterpreter = null) {
        if (this.undoStack.length === 0) {
            return false;
        }

        // Save current state to redo stack
        const currentSnapshot = {
            action: 'redo',
            cells: new Map(this.cells),
            hiddenRows: new Set(this.hiddenRows),
            hiddenColumns: new Set(this.hiddenColumns),
            namedRanges: new Map(this.namedRanges),
            frozenRows: this.frozenRows,
            frozenColumns: this.frozenColumns,
            validations: new Map(this.validations),
            timestamp: Date.now()
        };
        this.redoStack.push(currentSnapshot);

        // Restore previous state
        const snapshot = this.undoStack.pop();
        this.recordingHistory = false;

        this.cells = new Map(snapshot.cells);
        this.hiddenRows = new Set(snapshot.hiddenRows);
        this.hiddenColumns = new Set(snapshot.hiddenColumns);
        this.namedRanges = new Map(snapshot.namedRanges);
        this.frozenRows = snapshot.frozenRows;
        this.frozenColumns = snapshot.frozenColumns;
        this.validations = new Map(snapshot.validations);

        this._rebuildDependents();
        if (rexxInterpreter) {
            this._recalculateAll(rexxInterpreter);
        }

        this.recordingHistory = true;
        return true;
    }

    /**
     * Redo the last undone action
     * @param {Object} rexxInterpreter - Optional interpreter for recalculation
     * @returns {boolean} True if redo was performed
     */
    redo(rexxInterpreter = null) {
        if (this.redoStack.length === 0) {
            return false;
        }

        // Save current state to undo stack
        const currentSnapshot = {
            action: 'undo',
            cells: new Map(this.cells),
            hiddenRows: new Set(this.hiddenRows),
            hiddenColumns: new Set(this.hiddenColumns),
            namedRanges: new Map(this.namedRanges),
            frozenRows: this.frozenRows,
            frozenColumns: this.frozenColumns,
            validations: new Map(this.validations),
            timestamp: Date.now()
        };
        this.undoStack.push(currentSnapshot);

        // Restore redo state
        const snapshot = this.redoStack.pop();
        this.recordingHistory = false;

        this.cells = new Map(snapshot.cells);
        this.hiddenRows = new Set(snapshot.hiddenRows);
        this.hiddenColumns = new Set(snapshot.hiddenColumns);
        this.namedRanges = new Map(snapshot.namedRanges);
        this.frozenRows = snapshot.frozenRows;
        this.frozenColumns = snapshot.frozenColumns;
        this.validations = new Map(snapshot.validations);

        this._rebuildDependents();
        if (rexxInterpreter) {
            this._recalculateAll(rexxInterpreter);
        }

        this.recordingHistory = true;
        return true;
    }

    /**
     * Clear undo/redo history
     */
    clearHistory() {
        this.undoStack = [];
        this.redoStack = [];
    }

    /**
     * Check if undo is available
     */
    canUndo() {
        return this.undoStack.length > 0;
    }

    /**
     * Check if redo is available
     */
    canRedo() {
        return this.redoStack.length > 0;
    }
}

// Export for Node.js (Jest), ES6 modules, and browser
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SpreadsheetModel;
}

export default SpreadsheetModel;
