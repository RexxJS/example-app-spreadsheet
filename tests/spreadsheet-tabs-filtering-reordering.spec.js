/**
 * Tests for spreadsheet tabs, filtering, and column reordering features
 */

const SpreadsheetModel = require('../src/spreadsheet-model.js');

describe('SpreadsheetModel - Tabs, Filtering, and Reordering', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
    });

    describe('Multi-Sheet Support', () => {
        it('should start with a default Sheet1', () => {
            expect(model.getActiveSheetName()).toBe('Sheet1');
            expect(model.getSheetNames()).toEqual(['Sheet1']);
        });

        it('should add a new sheet', () => {
            model.addSheet('Sheet2');
            expect(model.getSheetNames()).toEqual(['Sheet1', 'Sheet2']);
        });

        it('should reject invalid sheet names', () => {
            expect(() => model.addSheet('Sheet 2')).toThrow('valid Rexx variable name');
            expect(() => model.addSheet('2Sheet')).toThrow('valid Rexx variable name');
            expect(() => model.addSheet('Sheet-2')).toThrow('valid Rexx variable name');
        });

        it('should accept valid Rexx sheet names', () => {
            expect(() => model.addSheet('Sheet2')).not.toThrow();
            expect(() => model.addSheet('Data_Sheet')).not.toThrow();
            expect(() => model.addSheet('Report_2024')).not.toThrow();
        });

        it('should prevent duplicate sheet names', () => {
            model.addSheet('Sheet2');
            expect(() => model.addSheet('Sheet2')).toThrow('already exists');
        });

        it('should delete a sheet', () => {
            model.addSheet('Sheet2');
            model.deleteSheet('Sheet2');
            expect(model.getSheetNames()).toEqual(['Sheet1']);
        });

        it('should prevent deleting the last sheet', () => {
            expect(() => model.deleteSheet('Sheet1')).toThrow('Cannot delete the last sheet');
        });

        it('should switch active sheet when deleting current sheet', () => {
            model.addSheet('Sheet2');
            model.setActiveSheet('Sheet2');
            model.deleteSheet('Sheet2');
            expect(model.getActiveSheetName()).toBe('Sheet1');
        });

        it('should rename a sheet', () => {
            model.renameSheet('Sheet1', 'MainData');
            expect(model.getSheetNames()).toEqual(['MainData']);
            expect(model.getActiveSheetName()).toBe('MainData');
        });

        it('should switch active sheet', () => {
            model.addSheet('Sheet2');
            model.setActiveSheet('Sheet2');
            expect(model.getActiveSheetName()).toBe('Sheet2');
        });

        it('should keep data isolated between sheets', () => {
            model.setCell('A1', 'Sheet1 Data');
            model.addSheet('Sheet2');
            model.setActiveSheet('Sheet2');
            model.setCell('A1', 'Sheet2 Data');

            expect(model.getCellValue('A1')).toBe('Sheet2 Data');
            model.setActiveSheet('Sheet1');
            expect(model.getCellValue('A1')).toBe('Sheet1 Data');
        });

        it('should export and import multi-sheet data', () => {
            model.setCell('A1', '10');
            model.addSheet('Sheet2');
            model.setActiveSheet('Sheet2');
            model.setCell('A1', '20');

            const json = model.toJSON();
            expect(json.version).toBe(2);
            expect(json.sheets).toBeDefined();
            expect(json.sheets.Sheet1).toBeDefined();
            expect(json.sheets.Sheet2).toBeDefined();

            const model2 = new SpreadsheetModel(100, 26);
            model2.fromJSON(json);

            expect(model2.getSheetNames()).toEqual(['Sheet1', 'Sheet2']);
            expect(model2.getActiveSheetName()).toBe('Sheet2');

            model2.setActiveSheet('Sheet1');
            expect(model2.getCellValue('A1')).toBe('10');
            model2.setActiveSheet('Sheet2');
            expect(model2.getCellValue('A1')).toBe('20');
        });

        it('should import legacy single-sheet format', () => {
            const legacyData = {
                setupScript: '',
                cells: {
                    'A1': '10',
                    'B1': '=A1 * 2'
                },
                hiddenRows: [],
                hiddenColumns: []
            };

            const model2 = new SpreadsheetModel(100, 26);
            model2.fromJSON(legacyData);

            expect(model2.getSheetNames()).toEqual(['Sheet1']);
            expect(model2.getCellValue('A1')).toBe('10');
        });
    });

    describe('Row Filtering', () => {
        beforeEach(() => {
            // Setup test data
            model.setCell('A1', 'Apple');
            model.setCell('A2', 'Banana');
            model.setCell('A3', 'Apple Pie');
            model.setCell('A4', 'Cherry');
            model.setCell('A5', 'Apple Juice');
        });

        it('should filter rows by text criteria', () => {
            model.applyRowFilter('A', 'Apple');

            expect(model.isRowVisible(1)).toBe(true);  // Apple
            expect(model.isRowVisible(2)).toBe(false); // Banana
            expect(model.isRowVisible(3)).toBe(true);  // Apple Pie
            expect(model.isRowVisible(4)).toBe(false); // Cherry
            expect(model.isRowVisible(5)).toBe(true);  // Apple Juice
        });

        it('should filter case-insensitively', () => {
            model.applyRowFilter('A', 'apple');

            expect(model.isRowVisible(1)).toBe(true);  // Apple
            expect(model.isRowVisible(3)).toBe(true);  // Apple Pie
        });

        it('should clear filters', () => {
            model.applyRowFilter('A', 'Apple');
            model.clearRowFilter();

            expect(model.isRowVisible(1)).toBe(true);
            expect(model.isRowVisible(2)).toBe(true);
            expect(model.isRowVisible(3)).toBe(true);
            expect(model.isRowVisible(4)).toBe(true);
            expect(model.isRowVisible(5)).toBe(true);
        });

        it('should return filter criteria', () => {
            model.applyRowFilter('A', 'Apple');
            const criteria = model.getFilterCriteria();

            expect(criteria).toBeDefined();
            expect(criteria.column).toBe('A');
            expect(criteria.criteria).toBe('Apple');
        });

        it('should persist filter criteria in JSON', () => {
            model.applyRowFilter('A', 'Apple');
            const json = model.toJSON();

            expect(json.sheets.Sheet1.filterCriteria).toBeDefined();
            expect(json.sheets.Sheet1.filterCriteria.column).toBe('A');
            expect(json.sheets.Sheet1.filterCriteria.criteria).toBe('Apple');
        });

        it('should restore filter on import', () => {
            model.applyRowFilter('A', 'Apple');
            const json = model.toJSON();

            const model2 = new SpreadsheetModel(100, 26);
            model2.setCell('A1', 'Apple');
            model2.setCell('A2', 'Banana');
            model2.fromJSON(json);

            expect(model2.isRowVisible(1)).toBe(true);
            expect(model2.isRowVisible(2)).toBe(false);
        });
    });

    describe('Column Reordering', () => {
        beforeEach(() => {
            // Setup test data in columns A, B, C
            model.setCell('A1', '10');
            model.setCell('B1', '20');
            model.setCell('C1', '30');
            model.setCell('A2', '11');
            model.setCell('B2', '21');
            model.setCell('C2', '31');
        });

        it('should move column right', () => {
            model.moveColumnRight('A');

            // After move: B, A, C
            expect(model.getCellValue('A1')).toBe('20');
            expect(model.getCellValue('B1')).toBe('10');
            expect(model.getCellValue('C1')).toBe('30');
        });

        it('should move column left', () => {
            model.moveColumnLeft('B');

            // After move: B, A, C
            expect(model.getCellValue('A1')).toBe('20');
            expect(model.getCellValue('B1')).toBe('10');
            expect(model.getCellValue('C1')).toBe('30');
        });

        it('should prevent moving first column left', () => {
            expect(() => model.moveColumnLeft('A')).toThrow('Cannot move first column left');
        });

        it('should prevent moving last column right', () => {
            expect(() => model.moveColumnRight('Z')).toThrow('Cannot move last column right');
        });

        it('should update formulas when columns are swapped', () => {
            model.setCell('C1', '=A1 + B1');

            model.moveColumnRight('A');

            // After move: B, A, C
            // Formula in C1 should now be =B1 + A1 (columns swapped)
            expect(model.getCell('C1').expression).toBe('B1 + A1');
        });

        it('should preserve absolute references when swapping columns', () => {
            model.setCell('C1', '=$A$1 + B1');

            model.moveColumnRight('A');

            // Absolute reference should not change
            expect(model.getCell('C1').expression).toBe('$A$1 + B1');
        });

        it('should handle complex formulas when swapping', () => {
            model.setCell('D1', '=A1 + B1 * C1 / A2');

            model.moveColumnRight('A');

            // After swap A and B: =B1 + A1 * C1 / B2
            const expr = model.getCell('D1').expression;
            expect(expr).toBe('B1 + A1 * C1 / B2');
        });

        it('should work with column numbers', () => {
            model.moveColumnRight(1); // Move column 1 (A) right

            expect(model.getCellValue('A1')).toBe('20');
            expect(model.getCellValue('B1')).toBe('10');
        });
    });

    describe('Cross-Sheet Formula References', () => {
        it('should support extracting cross-sheet references', () => {
            const expr = 'Sheet2.A1 + Sheet3.B2';
            const refs = model.extractCellReferences(expr);

            expect(refs).toContain('Sheet2.A1');
            expect(refs).toContain('Sheet3.B2');
        });

        it('should extract both local and cross-sheet references', () => {
            const expr = 'A1 + Sheet2.B1 + C1';
            const refs = model.extractCellReferences(expr);

            expect(refs).toContain('A1');
            expect(refs).toContain('Sheet2.B1');
            expect(refs).toContain('C1');
        });

        it('should not duplicate references that are part of sheet names', () => {
            const expr = 'Sheet2.A1 + B1';
            const refs = model.extractCellReferences(expr);

            // Should not extract 'A1' separately when it's part of 'Sheet2.A1'
            expect(refs.filter(r => r === 'A1').length).toBe(0);
            expect(refs).toContain('Sheet2.A1');
            expect(refs).toContain('B1');
        });
    });
});
