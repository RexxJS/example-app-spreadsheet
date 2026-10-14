/**
 * Tests for new spreadsheet features: Fill, Find/Replace, Hide/Unhide, Named Ranges
 */

const SpreadsheetModel = require('../src/spreadsheet-model.js');

describe('SpreadsheetModel - New Features', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
    });

    describe('Fill Down', () => {
        it('should fill down values from a single cell', () => {
            model.setCell('A1', '10');
            model.fillDown('A1', 'A2:A5');

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('A2')).toBe('10');
            expect(model.getCellValue('A3')).toBe('10');
            expect(model.getCellValue('A4')).toBe('10');
            expect(model.getCellValue('A5')).toBe('10');
        });

        it('should fill down formulas and adjust references', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('B1', '=A1 + 5');

            model.fillDown('B1', 'B2:B5');

            expect(model.getCell('B1').expression).toBe('A1 + 5');
            expect(model.getCell('B2').expression).toBe('A2 + 5');
            expect(model.getCell('B3').expression).toBe('A3 + 5');
            expect(model.getCell('B4').expression).toBe('A4 + 5');
            expect(model.getCell('B5').expression).toBe('A5 + 5');
        });

        it('should preserve absolute references when filling down', () => {
            model.setCell('A1', '10');
            model.setCell('B1', '=$A$1 + A1');

            model.fillDown('B1', 'B2:B3');

            expect(model.getCell('B1').expression).toBe('$A$1 + A1');
            expect(model.getCell('B2').expression).toBe('$A$1 + A2');
            expect(model.getCell('B3').expression).toBe('$A$1 + A3');
        });

        it('should fill down multiple columns', () => {
            model.setCell('A1', '10');
            model.setCell('B1', '20');
            model.fillDown('A1:B1', 'A2:B5');

            expect(model.getCellValue('A2')).toBe('10');
            expect(model.getCellValue('B2')).toBe('20');
            expect(model.getCellValue('A5')).toBe('10');
            expect(model.getCellValue('B5')).toBe('20');
        });

        it('should throw error for mismatched column widths', () => {
            model.setCell('A1', '10');
            expect(() => model.fillDown('A1', 'A2:B5')).toThrow('same number of columns');
        });

        it('should copy cell formatting when filling down', () => {
            model.setCell('A1', '10', null, { format: 'bold', comment: 'Test' });
            model.fillDown('A1', 'A2:A3');

            const cell2 = model.getCell('A2');
            expect(cell2.value).toBe('10');
            expect(cell2.format).toBe('bold');
            expect(cell2.comment).toBe('Test');
        });
    });

    describe('Fill Right', () => {
        it('should fill right values from a single cell', () => {
            model.setCell('A1', '10');
            model.fillRight('A1', 'B1:E1');

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('B1')).toBe('10');
            expect(model.getCellValue('C1')).toBe('10');
            expect(model.getCellValue('D1')).toBe('10');
            expect(model.getCellValue('E1')).toBe('10');
        });

        it('should fill right formulas and adjust references', () => {
            model.setCell('A1', '10');
            model.setCell('B1', '20');
            model.setCell('A2', '=A1 + 5');

            model.fillRight('A2', 'B2:E2');

            expect(model.getCell('A2').expression).toBe('A1 + 5');
            expect(model.getCell('B2').expression).toBe('B1 + 5');
            expect(model.getCell('C2').expression).toBe('C1 + 5');
            expect(model.getCell('D2').expression).toBe('D1 + 5');
            expect(model.getCell('E2').expression).toBe('E1 + 5');
        });

        it('should preserve absolute references when filling right', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '=$A$1 + A1');

            model.fillRight('A2', 'B2:C2');

            expect(model.getCell('A2').expression).toBe('$A$1 + A1');
            expect(model.getCell('B2').expression).toBe('$A$1 + B1');
            expect(model.getCell('C2').expression).toBe('$A$1 + C1');
        });

        it('should fill right multiple rows', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.fillRight('A1:A2', 'B1:E2');

            expect(model.getCellValue('B1')).toBe('10');
            expect(model.getCellValue('B2')).toBe('20');
            expect(model.getCellValue('E1')).toBe('10');
            expect(model.getCellValue('E2')).toBe('20');
        });

        it('should throw error for mismatched row heights', () => {
            model.setCell('A1', '10');
            expect(() => model.fillRight('A1', 'B1:B5')).toThrow('same number of rows');
        });

        it('should copy cell formatting when filling right', () => {
            model.setCell('A1', '10', null, { format: 'bold;color:red', comment: 'Important' });
            model.fillRight('A1', 'B1:C1');

            const cellB = model.getCell('B1');
            expect(cellB.value).toBe('10');
            expect(cellB.format).toBe('bold;color:red');
            expect(cellB.comment).toBe('Important');
        });
    });

    describe('Find', () => {
        beforeEach(() => {
            model.setCell('A1', 'Hello');
            model.setCell('A2', 'World');
            model.setCell('B1', 'hello');
            model.setCell('B2', 'WORLD');
            model.setCell('C1', '=A1 & " " & A2');
        });

        it('should find cells with matching values (case insensitive)', () => {
            const results = model.find('hello', { matchCase: false });
            expect(results).toContain('A1');
            expect(results).toContain('B1');
            expect(results.length).toBeGreaterThanOrEqual(2);
        });

        it('should find cells with matching values (case sensitive)', () => {
            const results = model.find('Hello', { matchCase: true });
            expect(results).toContain('A1');
            expect(results).not.toContain('B1');
        });

        it('should find partial matches', () => {
            const results = model.find('ell', { matchCase: false });
            expect(results).toContain('A1');
            expect(results).toContain('B1');
        });

        it('should find exact matches only', () => {
            const results = model.find('Hello', { matchCase: true, matchEntireCell: true });
            expect(results).toContain('A1');
            expect(results).not.toContain('B1');
            expect(results.length).toBe(1);
        });

        it('should search formulas when specified', () => {
            const results = model.find('A1', { searchFormulas: true });
            expect(results).toContain('C1');
        });

        it('should not search formulas by default', () => {
            const results = model.find('A1', { searchFormulas: false });
            expect(results).not.toContain('C1');
        });

        it('should return empty array when no matches', () => {
            const results = model.find('xyz123', {});
            expect(results).toEqual([]);
        });
    });

    describe('Replace', () => {
        beforeEach(() => {
            model.setCell('A1', 'Hello');
            model.setCell('A2', 'World');
            model.setCell('B1', 'Hello World');
            model.setCell('B2', 'hello');
        });

        it('should replace all occurrences (case insensitive)', () => {
            const count = model.replace('hello', 'Hi', { matchCase: false });
            expect(count).toBe(3);
            expect(model.getCellValue('A1')).toBe('Hi');
            expect(model.getCellValue('B1')).toBe('Hi World');
            expect(model.getCellValue('B2')).toBe('Hi');
        });

        it('should replace case sensitive matches only', () => {
            const count = model.replace('Hello', 'Hi', { matchCase: true });
            expect(count).toBe(2);
            expect(model.getCellValue('A1')).toBe('Hi');
            expect(model.getCellValue('B1')).toBe('Hi World');
            expect(model.getCellValue('B2')).toBe('hello'); // Not replaced
        });

        it('should replace entire cell only', () => {
            const count = model.replace('Hello', 'Hi', { matchCase: false, matchEntireCell: true });
            expect(count).toBe(2);
            expect(model.getCellValue('A1')).toBe('Hi');
            expect(model.getCellValue('B1')).toBe('Hello World'); // Not replaced (not entire cell)
            expect(model.getCellValue('B2')).toBe('Hi');
        });

        it('should replace in formulas when specified', () => {
            model.setCell('C1', '=A1 + A2');
            const count = model.replace('A1', 'B1', { searchFormulas: true });
            expect(count).toBeGreaterThan(0);
            expect(model.getCell('C1').expression).toBe('B1 + A2');
        });

        it('should preserve cell metadata when replacing', () => {
            model.setCell('A1', 'Test', null, { format: 'bold', comment: 'Note' });
            model.replace('Test', 'Changed', {});
            const cell = model.getCell('A1');
            expect(cell.value).toBe('Changed');
            expect(cell.format).toBe('bold');
            expect(cell.comment).toBe('Note');
        });

        it('should return 0 when no replacements made', () => {
            const count = model.replace('xyz123', 'abc', {});
            expect(count).toBe(0);
        });
    });

    describe('Hide/Unhide Rows', () => {
        it('should hide a row', () => {
            model.hideRow(5);
            expect(model.isRowHidden(5)).toBe(true);
        });

        it('should unhide a row', () => {
            model.hideRow(5);
            model.unhideRow(5);
            expect(model.isRowHidden(5)).toBe(false);
        });

        it('should track multiple hidden rows', () => {
            model.hideRow(2);
            model.hideRow(5);
            model.hideRow(10);
            expect(model.isRowHidden(2)).toBe(true);
            expect(model.isRowHidden(5)).toBe(true);
            expect(model.isRowHidden(10)).toBe(true);
            expect(model.isRowHidden(3)).toBe(false);
        });

        it('should throw error for invalid row number', () => {
            expect(() => model.hideRow(0)).toThrow('Invalid row number');
            expect(() => model.hideRow(101)).toThrow('Invalid row number');
        });
    });

    describe('Hide/Unhide Columns', () => {
        it('should hide a column by number', () => {
            model.hideColumn(5);
            expect(model.isColumnHidden(5)).toBe(true);
        });

        it('should hide a column by letter', () => {
            model.hideColumn('E');
            expect(model.isColumnHidden('E')).toBe(true);
            expect(model.isColumnHidden(5)).toBe(true);
        });

        it('should unhide a column', () => {
            model.hideColumn(5);
            model.unhideColumn(5);
            expect(model.isColumnHidden(5)).toBe(false);
        });

        it('should track multiple hidden columns', () => {
            model.hideColumn('A');
            model.hideColumn('C');
            model.hideColumn('Z');
            expect(model.isColumnHidden(1)).toBe(true);
            expect(model.isColumnHidden(3)).toBe(true);
            expect(model.isColumnHidden(26)).toBe(true);
            expect(model.isColumnHidden(2)).toBe(false);
        });

        it('should throw error for invalid column number', () => {
            expect(() => model.hideColumn(0)).toThrow('Invalid column number');
            expect(() => model.hideColumn(27)).toThrow('Invalid column number');
        });
    });

    describe('Named Ranges', () => {
        it('should define a named range', () => {
            model.defineNamedRange('SalesData', 'A1:B10');
            expect(model.getNamedRange('SalesData')).toBe('A1:B10');
        });

        it('should define named range for single cell', () => {
            model.defineNamedRange('TaxRate', 'A1');
            expect(model.getNamedRange('TaxRate')).toBe('A1');
        });

        it('should delete a named range', () => {
            model.defineNamedRange('SalesData', 'A1:B10');
            model.deleteNamedRange('SalesData');
            expect(model.getNamedRange('SalesData')).toBeNull();
        });

        it('should get all named ranges', () => {
            model.defineNamedRange('Sales', 'A1:B10');
            model.defineNamedRange('Expenses', 'C1:D10');
            const ranges = model.getAllNamedRanges();
            expect(ranges.Sales).toBe('A1:B10');
            expect(ranges.Expenses).toBe('C1:D10');
        });

        it('should validate named range name format', () => {
            expect(() => model.defineNamedRange('123Invalid', 'A1:B10'))
                .toThrow('must start with a letter');
            expect(() => model.defineNamedRange('Invalid Name', 'A1:B10'))
                .toThrow('must start with a letter');
            expect(() => model.defineNamedRange('Invalid-Name', 'A1:B10'))
                .toThrow('must start with a letter');
        });

        it('should validate range reference format', () => {
            expect(() => model.defineNamedRange('ValidName', 'Invalid'))
                .toThrow('Invalid range reference');
        });

        it('should allow valid named range names', () => {
            expect(() => model.defineNamedRange('Sales2023', 'A1:B10')).not.toThrow();
            expect(() => model.defineNamedRange('Total_Sales', 'A1:B10')).not.toThrow();
            expect(() => model.defineNamedRange('Q1', 'A1:B10')).not.toThrow();
        });

        it('should resolve named ranges in expressions', () => {
            model.defineNamedRange('Data', 'A1:A5');
            const expression = 'SUM_RANGE("Data") + 10';
            const resolved = model.resolveNamedRanges(expression);
            expect(resolved).toBe('SUM_RANGE("A1:A5") + 10');
        });

        it('should use named ranges in formulas', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('A3', '30');
            model.defineNamedRange('MyData', 'A1:A3');
            model.setCell('B1', '=SUM_RANGE("MyData")');

            // The expression should have resolved MyData to A1:A3
            const cell = model.getCell('B1');
            expect(cell.expression).toContain('A1:A3');
        });

        it('should handle multiple named ranges in one expression', () => {
            model.defineNamedRange('Sales', 'A1:A10');
            model.defineNamedRange('Expenses', 'B1:B10');
            const expression = 'SUM_RANGE("Sales") - SUM_RANGE("Expenses")';
            const resolved = model.resolveNamedRanges(expression);
            expect(resolved).toBe('SUM_RANGE("A1:A10") - SUM_RANGE("B1:B10")');
        });
    });

    describe('toJSON/fromJSON with new features', () => {
        it('should export and import hidden rows', () => {
            model.hideRow(2);
            model.hideRow(5);
            const json = model.toJSON();

            const newModel = new SpreadsheetModel();
            newModel.fromJSON(json);

            expect(newModel.isRowHidden(2)).toBe(true);
            expect(newModel.isRowHidden(5)).toBe(true);
            expect(newModel.isRowHidden(3)).toBe(false);
        });

        it('should export and import hidden columns', () => {
            model.hideColumn('A');
            model.hideColumn('C');
            const json = model.toJSON();

            const newModel = new SpreadsheetModel();
            newModel.fromJSON(json);

            expect(newModel.isColumnHidden(1)).toBe(true);
            expect(newModel.isColumnHidden(3)).toBe(true);
            expect(newModel.isColumnHidden(2)).toBe(false);
        });

        it('should export and import named ranges', () => {
            model.defineNamedRange('Sales', 'A1:B10');
            model.defineNamedRange('Expenses', 'C1:D10');
            const json = model.toJSON();

            const newModel = new SpreadsheetModel();
            newModel.fromJSON(json);

            expect(newModel.getNamedRange('Sales')).toBe('A1:B10');
            expect(newModel.getNamedRange('Expenses')).toBe('C1:D10');
        });

        it('should handle backward compatibility with old format', () => {
            const oldFormat = {
                A1: '10',
                A2: '20'
            };

            model.fromJSON(oldFormat);
            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('A2')).toBe('20');
        });
    });
});
