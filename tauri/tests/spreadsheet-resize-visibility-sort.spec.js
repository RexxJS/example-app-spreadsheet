/**
 * Tests for Column/Row Resize, Visibility, and Sort Features
 */

const SpreadsheetModel = require('../src/spreadsheet-model');

describe('Column Width and Row Height', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
    });

    describe('Column Width', () => {
        it('should set and get column width by number', () => {
            model.setColumnWidth(1, 150);
            expect(model.getColumnWidth(1)).toBe(150);
        });

        it('should set and get column width by letter', () => {
            model.setColumnWidth('B', 200);
            expect(model.getColumnWidth('B')).toBe(200);
        });

        it('should return default width (100) for unset columns', () => {
            expect(model.getColumnWidth(5)).toBe(100);
        });

        it('should throw error for invalid column number', () => {
            expect(() => model.setColumnWidth(0, 100)).toThrow('Invalid column number');
            expect(() => model.setColumnWidth(100, 100)).toThrow('Invalid column number');
        });

        it('should throw error for invalid width', () => {
            expect(() => model.setColumnWidth(1, 0)).toThrow('Width must be greater than 0');
            expect(() => model.setColumnWidth(1, -10)).toThrow('Width must be greater than 0');
        });

        it('should handle multiple column widths', () => {
            model.setColumnWidth(1, 80);
            model.setColumnWidth(2, 120);
            model.setColumnWidth(3, 150);

            expect(model.getColumnWidth(1)).toBe(80);
            expect(model.getColumnWidth(2)).toBe(120);
            expect(model.getColumnWidth(3)).toBe(150);
        });

        it('should update column width', () => {
            model.setColumnWidth(1, 100);
            expect(model.getColumnWidth(1)).toBe(100);

            model.setColumnWidth(1, 200);
            expect(model.getColumnWidth(1)).toBe(200);
        });
    });

    describe('Row Height', () => {
        it('should set and get row height', () => {
            model.setRowHeight(1, 50);
            expect(model.getRowHeight(1)).toBe(50);
        });

        it('should return default height (32) for unset rows', () => {
            expect(model.getRowHeight(5)).toBe(32);
        });

        it('should throw error for invalid row number', () => {
            expect(() => model.setRowHeight(0, 50)).toThrow('Invalid row number');
            expect(() => model.setRowHeight(200, 50)).toThrow('Invalid row number');
        });

        it('should throw error for invalid height', () => {
            expect(() => model.setRowHeight(1, 0)).toThrow('Height must be greater than 0');
            expect(() => model.setRowHeight(1, -10)).toThrow('Height must be greater than 0');
        });

        it('should handle multiple row heights', () => {
            model.setRowHeight(1, 40);
            model.setRowHeight(2, 60);
            model.setRowHeight(3, 80);

            expect(model.getRowHeight(1)).toBe(40);
            expect(model.getRowHeight(2)).toBe(60);
            expect(model.getRowHeight(3)).toBe(80);
        });

        it('should update row height', () => {
            model.setRowHeight(1, 32);
            expect(model.getRowHeight(1)).toBe(32);

            model.setRowHeight(1, 64);
            expect(model.getRowHeight(1)).toBe(64);
        });
    });

    describe('Persistence', () => {
        it('should serialize and deserialize column widths', () => {
            model.setColumnWidth(1, 150);
            model.setColumnWidth('B', 200);
            model.setColumnWidth(5, 120);

            const json = model.toJSON();
            const newModel = new SpreadsheetModel(100, 26);
            newModel.fromJSON(json);

            expect(newModel.getColumnWidth(1)).toBe(150);
            expect(newModel.getColumnWidth(2)).toBe(200);
            expect(newModel.getColumnWidth(5)).toBe(120);
        });

        it('should serialize and deserialize row heights', () => {
            model.setRowHeight(1, 50);
            model.setRowHeight(3, 75);
            model.setRowHeight(10, 100);

            const json = model.toJSON();
            const newModel = new SpreadsheetModel(100, 26);
            newModel.fromJSON(json);

            expect(newModel.getRowHeight(1)).toBe(50);
            expect(newModel.getRowHeight(3)).toBe(75);
            expect(newModel.getRowHeight(10)).toBe(100);
        });

        it('should handle empty widths/heights in serialization', () => {
            const json = model.toJSON();
            const newModel = new SpreadsheetModel(100, 26);
            newModel.fromJSON(json);

            expect(newModel.getColumnWidth(1)).toBe(100); // default
            expect(newModel.getRowHeight(1)).toBe(32); // default
        });
    });
});

describe('Column and Row Visibility', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
    });

    describe('Column Visibility', () => {
        it('should hide column', () => {
            model.hideColumn(1);
            expect(model.isColumnHidden(1)).toBe(true);
        });

        it('should hide column by letter', () => {
            model.hideColumn('B');
            expect(model.isColumnHidden('B')).toBe(true);
        });

        it('should unhide column', () => {
            model.hideColumn(1);
            expect(model.isColumnHidden(1)).toBe(true);

            model.unhideColumn(1);
            expect(model.isColumnHidden(1)).toBe(false);
        });

        it('should handle multiple hidden columns', () => {
            model.hideColumn(1);
            model.hideColumn('C');
            model.hideColumn(5);

            expect(model.isColumnHidden(1)).toBe(true);
            expect(model.isColumnHidden(3)).toBe(true);
            expect(model.isColumnHidden(5)).toBe(true);
            expect(model.isColumnHidden(2)).toBe(false);
        });

        it('should persist hidden columns', () => {
            model.hideColumn(1);
            model.hideColumn(2);

            const json = model.toJSON();
            const newModel = new SpreadsheetModel(100, 26);
            newModel.fromJSON(json);

            expect(newModel.isColumnHidden(1)).toBe(true);
            expect(newModel.isColumnHidden(2)).toBe(true);
        });
    });

    describe('Row Visibility', () => {
        it('should hide row', () => {
            model.hideRow(1);
            expect(model.isRowHidden(1)).toBe(true);
        });

        it('should unhide row', () => {
            model.hideRow(1);
            expect(model.isRowHidden(1)).toBe(true);

            model.unhideRow(1);
            expect(model.isRowHidden(1)).toBe(false);
        });

        it('should handle multiple hidden rows', () => {
            model.hideRow(1);
            model.hideRow(3);
            model.hideRow(5);

            expect(model.isRowHidden(1)).toBe(true);
            expect(model.isRowHidden(3)).toBe(true);
            expect(model.isRowHidden(5)).toBe(true);
            expect(model.isRowHidden(2)).toBe(false);
        });

        it('should persist hidden rows', () => {
            model.hideRow(1);
            model.hideRow(2);

            const json = model.toJSON();
            const newModel = new SpreadsheetModel(100, 26);
            newModel.fromJSON(json);

            expect(newModel.isRowHidden(1)).toBe(true);
            expect(newModel.isRowHidden(2)).toBe(true);
        });
    });

    describe('Combined Visibility', () => {
        it('should handle both hidden rows and columns', () => {
            model.hideRow(1);
            model.hideRow(3);
            model.hideColumn(2);
            model.hideColumn('D');

            const json = model.toJSON();
            const newModel = new SpreadsheetModel(100, 26);
            newModel.fromJSON(json);

            expect(newModel.isRowHidden(1)).toBe(true);
            expect(newModel.isRowHidden(3)).toBe(true);
            expect(newModel.isColumnHidden(2)).toBe(true);
            expect(newModel.isColumnHidden(4)).toBe(true);
        });
    });
});

describe('Row Sorting', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
    });

    it('should sort range in ascending order', () => {
        model.setCell('A1', '3');
        model.setCell('A2', '1');
        model.setCell('A3', '2');
        model.setCell('B1', 'Charlie');
        model.setCell('B2', 'Alice');
        model.setCell('B3', 'Bob');

        model.sortRange('A1:B3', 'A', true);

        expect(model.getCellValue('A1')).toBe('1');
        expect(model.getCellValue('A2')).toBe('2');
        expect(model.getCellValue('A3')).toBe('3');
        expect(model.getCellValue('B1')).toBe('Alice');
        expect(model.getCellValue('B2')).toBe('Bob');
        expect(model.getCellValue('B3')).toBe('Charlie');
    });

    it('should sort range in descending order', () => {
        model.setCell('A1', '1');
        model.setCell('A2', '2');
        model.setCell('A3', '3');
        model.setCell('B1', 'Alice');
        model.setCell('B2', 'Bob');
        model.setCell('B3', 'Charlie');

        model.sortRange('A1:B3', 'A', false);

        expect(model.getCellValue('A1')).toBe('3');
        expect(model.getCellValue('A2')).toBe('2');
        expect(model.getCellValue('A3')).toBe('1');
        expect(model.getCellValue('B1')).toBe('Charlie');
        expect(model.getCellValue('B2')).toBe('Bob');
        expect(model.getCellValue('B3')).toBe('Alice');
    });

    it('should sort by text column', () => {
        model.setCell('A1', 'Charlie');
        model.setCell('A2', 'Alice');
        model.setCell('A3', 'Bob');
        model.setCell('B1', '3');
        model.setCell('B2', '1');
        model.setCell('B3', '2');

        model.sortRange('A1:B3', 'A', true);

        expect(model.getCellValue('A1')).toBe('Alice');
        expect(model.getCellValue('A2')).toBe('Bob');
        expect(model.getCellValue('A3')).toBe('Charlie');
        expect(model.getCellValue('B1')).toBe('1');
        expect(model.getCellValue('B2')).toBe('2');
        expect(model.getCellValue('B3')).toBe('3');
    });

    it('should handle numeric sorting correctly', () => {
        model.setCell('A1', '10');
        model.setCell('A2', '2');
        model.setCell('A3', '100');
        model.setCell('A4', '20');

        model.sortRange('A1:A4', 'A', true);

        expect(model.getCellValue('A1')).toBe('2');
        expect(model.getCellValue('A2')).toBe('10');
        expect(model.getCellValue('A3')).toBe('20');
        expect(model.getCellValue('A4')).toBe('100');
    });

    it('should handle empty cells in sort', () => {
        model.setCell('A1', '3');
        model.setCell('A2', '');
        model.setCell('A3', '1');
        model.setCell('A4', '2');

        model.sortRange('A1:A4', 'A', true);

        // Empty cells sort to the beginning in the current implementation
        expect(model.getCellValue('A1')).toBe('');
        expect(model.getCellValue('A2')).toBe('1');
        expect(model.getCellValue('A3')).toBe('2');
        expect(model.getCellValue('A4')).toBe('3');
    });

    it('should move formulas with their rows during sort', () => {
        model.setCell('A1', '3');
        model.setCell('A2', '1');
        model.setCell('A3', '2');
        model.setCell('B1', '=A1*10');
        model.setCell('B2', '=A2*10');
        model.setCell('B3', '=A3*10');

        model.sortRange('A1:B3', 'A', true);

        expect(model.getCellValue('A1')).toBe('1');
        expect(model.getCellValue('A2')).toBe('2');
        expect(model.getCellValue('A3')).toBe('3');

        // Formulas move with their rows, so the formula that was with '1' is now in B1
        const cell1 = model.getCell('B1');
        const cell2 = model.getCell('B2');
        const cell3 = model.getCell('B3');

        // The formula from row with value '1' (originally A2) should now reference A1
        expect(cell1.expression).toBe('A2*10');
        expect(cell2.expression).toBe('A3*10');
        expect(cell3.expression).toBe('A1*10');
    });
});

describe('Multi-Sheet Support for Resize and Visibility', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
        model.addSheet('Sheet2');
    });

    it('should maintain separate column widths per sheet', () => {
        model.activeSheetName = 'Sheet1';
        model.setColumnWidth(1, 150);

        model.activeSheetName = 'Sheet2';
        model.setColumnWidth(1, 200);

        model.activeSheetName = 'Sheet1';
        expect(model.getColumnWidth(1)).toBe(150);

        model.activeSheetName = 'Sheet2';
        expect(model.getColumnWidth(1)).toBe(200);
    });

    it('should maintain separate row heights per sheet', () => {
        model.activeSheetName = 'Sheet1';
        model.setRowHeight(1, 50);

        model.activeSheetName = 'Sheet2';
        model.setRowHeight(1, 75);

        model.activeSheetName = 'Sheet1';
        expect(model.getRowHeight(1)).toBe(50);

        model.activeSheetName = 'Sheet2';
        expect(model.getRowHeight(1)).toBe(75);
    });

    it('should maintain separate hidden columns per sheet', () => {
        model.activeSheetName = 'Sheet1';
        model.hideColumn(1);

        model.activeSheetName = 'Sheet2';
        expect(model.isColumnHidden(1)).toBe(false);

        model.hideColumn(2);

        model.activeSheetName = 'Sheet1';
        expect(model.isColumnHidden(1)).toBe(true);
        expect(model.isColumnHidden(2)).toBe(false);
    });

    it('should persist all sheets with widths and visibility', () => {
        model.activeSheetName = 'Sheet1';
        model.setColumnWidth(1, 150);
        model.hideRow(2);

        model.activeSheetName = 'Sheet2';
        model.setRowHeight(3, 60);
        model.hideColumn(4);

        const json = model.toJSON();
        const newModel = new SpreadsheetModel(100, 26);
        newModel.fromJSON(json);

        newModel.activeSheetName = 'Sheet1';
        expect(newModel.getColumnWidth(1)).toBe(150);
        expect(newModel.isRowHidden(2)).toBe(true);

        newModel.activeSheetName = 'Sheet2';
        expect(newModel.getRowHeight(3)).toBe(60);
        expect(newModel.isColumnHidden(4)).toBe(true);
    });
});
