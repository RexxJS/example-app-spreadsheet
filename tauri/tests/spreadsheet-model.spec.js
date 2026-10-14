/**
 * SpreadsheetModel Tests
 *
 * Tests for the core spreadsheet model without DOM dependencies.
 */

import SpreadsheetModel from '../src/spreadsheet-model.js';

describe('SpreadsheetModel', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
    });

    describe('Column letter conversion', () => {
        it('should convert column letters to numbers', () => {
            expect(SpreadsheetModel.colLetterToNumber('A')).toBe(1);
            expect(SpreadsheetModel.colLetterToNumber('B')).toBe(2);
            expect(SpreadsheetModel.colLetterToNumber('Z')).toBe(26);
            expect(SpreadsheetModel.colLetterToNumber('AA')).toBe(27);
            expect(SpreadsheetModel.colLetterToNumber('AB')).toBe(28);
            expect(SpreadsheetModel.colLetterToNumber('AZ')).toBe(52);
            expect(SpreadsheetModel.colLetterToNumber('BA')).toBe(53);
        });

        it('should convert column numbers to letters', () => {
            expect(SpreadsheetModel.colNumberToLetter(1)).toBe('A');
            expect(SpreadsheetModel.colNumberToLetter(2)).toBe('B');
            expect(SpreadsheetModel.colNumberToLetter(26)).toBe('Z');
            expect(SpreadsheetModel.colNumberToLetter(27)).toBe('AA');
            expect(SpreadsheetModel.colNumberToLetter(28)).toBe('AB');
            expect(SpreadsheetModel.colNumberToLetter(52)).toBe('AZ');
            expect(SpreadsheetModel.colNumberToLetter(53)).toBe('BA');
        });

        it('should round-trip column conversions', () => {
            for (let i = 1; i <= 100; i++) {
                const letter = SpreadsheetModel.colNumberToLetter(i);
                const number = SpreadsheetModel.colLetterToNumber(letter);
                expect(number).toBe(i);
            }
        });
    });

    describe('Cell reference parsing', () => {
        it('should parse valid cell references', () => {
            const ref1 = SpreadsheetModel.parseCellRef('A1');
            expect(ref1.col).toBe('A');
            expect(ref1.row).toBe(1);

            const ref2 = SpreadsheetModel.parseCellRef('B10');
            expect(ref2.col).toBe('B');
            expect(ref2.row).toBe(10);

            const ref3 = SpreadsheetModel.parseCellRef('AA99');
            expect(ref3.col).toBe('AA');
            expect(ref3.row).toBe(99);
        });

        it('should throw error for invalid cell references', () => {
            expect(() => SpreadsheetModel.parseCellRef('1A')).toThrow();
            expect(() => SpreadsheetModel.parseCellRef('A')).toThrow();
            expect(() => SpreadsheetModel.parseCellRef('123')).toThrow();
            expect(() => SpreadsheetModel.parseCellRef('')).toThrow();
        });

        it('should format cell references', () => {
            expect(SpreadsheetModel.formatCellRef('A', 1)).toBe('A1');
            expect(SpreadsheetModel.formatCellRef('Z', 99)).toBe('Z99');
            expect(SpreadsheetModel.formatCellRef('AA', 10)).toBe('AA10');
            expect(SpreadsheetModel.formatCellRef(1, 1)).toBe('A1');
            expect(SpreadsheetModel.formatCellRef(26, 99)).toBe('Z99');
        });
    });

    describe('Cell value operations', () => {
        it('should store and retrieve literal values', () => {
            model.setCell('A1', '10');
            expect(model.getCellValue('A1')).toBe('10');

            model.setCell('B2', 'Hello');
            expect(model.getCellValue('B2')).toBe('Hello');
        });

        it('should return empty string for unset cells', () => {
            expect(model.getCellValue('A1')).toBe('');
            expect(model.getCellValue('Z99')).toBe('');
        });

        it('should clear cells when set to empty', () => {
            model.setCell('A1', '10');
            expect(model.getCellValue('A1')).toBe('10');

            model.setCell('A1', '');
            expect(model.getCellValue('A1')).toBe('');

            model.setCell('A1', null);
            expect(model.getCellValue('A1')).toBe('');
        });

        it('should distinguish between value and expression', () => {
            model.setCell('A1', '10');
            const cell1 = model.getCell('A1');
            expect(cell1.value).toBe('10');
            expect(cell1.expression).toBeNull();

            model.setCell('A2', '=A1 + 5');
            const cell2 = model.getCell('A2');
            expect(cell2.expression).toBe('A1 + 5');
        });
    });

    describe('Expression detection', () => {
        it('should detect expressions starting with =', () => {
            model.setCell('A1', '=10 + 20');
            const cell = model.getCell('A1');
            expect(cell.expression).toBe('10 + 20');
        });

        it('should trim whitespace around =', () => {
            model.setCell('A1', '  = 10 + 20  ');
            const cell = model.getCell('A1');
            expect(cell.expression).toBe('10 + 20');
        });

        it('should not treat = in middle of string as expression', () => {
            model.setCell('A1', 'x=10');
            const cell = model.getCell('A1');
            expect(cell.value).toBe('x=10');
            expect(cell.expression).toBeNull();
        });
    });

    describe('Cell reference extraction', () => {
        it('should extract cell references from expressions', () => {
            const refs1 = model.extractCellReferences('A1 + B2');
            expect(refs1).toEqual(['A1', 'B2']);

            const refs2 = model.extractCellReferences('SUM(A1, A2, A3)');
            expect(refs2).toEqual(['A1', 'A2', 'A3']);

            const refs3 = model.extractCellReferences('A1 + A1 + B1');
            expect(refs3).toEqual(['A1', 'B1']); // Duplicates removed
        });

        it('should handle multi-letter column references', () => {
            const refs = model.extractCellReferences('AA10 + AB20 + Z99');
            expect(refs).toEqual(['AA10', 'AB20', 'Z99']);
        });

        it('should return empty array for expressions with no references', () => {
            const refs = model.extractCellReferences('10 + 20');
            expect(refs).toEqual([]);
        });

        it('should not extract partial matches', () => {
            const refs = model.extractCellReferences('BA1D + A1');
            // BA1D should not match, but A1 should
            expect(refs).toContain('A1');
            expect(refs).not.toContain('BA1D');
        });
    });

    describe('JSON export/import', () => {
        it('should export cells to JSON', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('A3', '=A1 + A2');

            const json = model.toJSON();
            // Version 2 format uses sheets
            expect(json.version).toBe(2);
            expect(json.sheets).toBeDefined();
            expect(json.sheets.Sheet1).toBeDefined();
            expect(json.sheets.Sheet1.cells).toBeDefined();
            expect(json.sheets.Sheet1.cells.A1).toBe('10');
            expect(json.sheets.Sheet1.cells.A2).toBe('20');
            expect(json.sheets.Sheet1.cells.A3).toBe('=A1 + A2');
        });

        it('should import cells from JSON', () => {
            const data = {
                version: 2,
                sheets: {
                    Sheet1: {
                        cells: {
                            'A1': '10',
                            'B1': 'Hello',
                            'C1': '=A1 + 5'
                        }
                    }
                }
            };

            model.fromJSON(data);

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('B1')).toBe('Hello');
            expect(model.getCell('C1').expression).toBe('A1 + 5');
        });

        it('should clear model before import', () => {
            model.setCell('A1', '999');
            model.setCell('B1', '888');

            model.fromJSON({ 'C1': '777' });

            expect(model.getCellValue('A1')).toBe('');
            expect(model.getCellValue('B1')).toBe('');
            expect(model.getCellValue('C1')).toBe('777');
        });
    });

    describe('Get all cells', () => {
        it('should return all non-empty cells', () => {
            model.setCell('A1', '10');
            model.setCell('B2', 'Hello');
            model.setCell('C3', '=A1 + 5');

            const cells = model.getAllCells();
            expect(cells).toHaveLength(3);

            const refs = cells.map(c => c.ref).sort();
            expect(refs).toEqual(['A1', 'B2', 'C3']);
        });

        it('should not include cleared cells', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('A1', ''); // Clear A1

            const cells = model.getAllCells();
            expect(cells).toHaveLength(1);
            expect(cells[0].ref).toBe('A2');
        });
    });

    describe('Cell addressing with objects', () => {
        it('should accept {col, row} objects for getCell', () => {
            model.setCell('A1', '10');
            const cell = model.getCell({ col: 'A', row: 1 });
            expect(cell.value).toBe('10');
        });

        it('should accept {col, row} objects for setCell', () => {
            model.setCell({ col: 'B', row: 2 }, 'Hello');
            expect(model.getCellValue('B2')).toBe('Hello');
        });

        it('should accept numeric column in objects', () => {
            model.setCell({ col: 1, row: 1 }, 'Test');
            expect(model.getCellValue('A1')).toBe('Test');
        });
    });

    describe('Cell metadata and formatting', () => {
        it('should store and retrieve cell format', () => {
            model.setCell('A1', '100', null, { format: 'bold' });
            const cell = model.getCell('A1');
            expect(cell.format).toBe('bold');
            expect(cell.value).toBe('100');
        });

        it('should store and retrieve cell comment', () => {
            model.setCell('A1', '100', null, { comment: 'Important value' });
            const cell = model.getCell('A1');
            expect(cell.comment).toBe('Important value');
            expect(cell.value).toBe('100');
        });

        it('should store both format and comment', () => {
            model.setCell('A1', '100', null, {
                format: 'bold;color:red',
                comment: 'Critical value'
            });
            const cell = model.getCell('A1');
            expect(cell.format).toBe('bold;color:red');
            expect(cell.comment).toBe('Critical value');
        });

        it('should update format via setCellMetadata', () => {
            model.setCell('A1', '100');
            model.setCellMetadata('A1', { format: 'italic' });
            const cell = model.getCell('A1');
            expect(cell.format).toBe('italic');
            expect(cell.value).toBe('100');
        });

        it('should preserve existing metadata when updating cell value', () => {
            model.setCell('A1', '100', null, { format: 'bold', comment: 'Note' });
            model.setCell('A1', '200');
            const cell = model.getCell('A1');
            expect(cell.value).toBe('200');
            expect(cell.format).toBe('bold');
            expect(cell.comment).toBe('Note');
        });

        it('should support alignment formats', () => {
            model.setCell('A1', 'Left', null, { format: 'align:left' });
            model.setCell('A2', 'Center', null, { format: 'align:center' });
            model.setCell('A3', 'Right', null, { format: 'align:right' });

            expect(model.getCell('A1').format).toBe('align:left');
            expect(model.getCell('A2').format).toBe('align:center');
            expect(model.getCell('A3').format).toBe('align:right');
        });

        it('should support number formats', () => {
            model.setCell('A1', '123.456', null, { format: 'number:0.00' });
            model.setCell('A2', '1234.56', null, { format: 'currency:USD' });
            model.setCell('A3', '0.125', null, { format: 'percent:0.0%' });

            expect(model.getCell('A1').format).toBe('number:0.00');
            expect(model.getCell('A2').format).toBe('currency:USD');
            expect(model.getCell('A3').format).toBe('percent:0.0%');
        });

        it('should support combined formats', () => {
            model.setCell('A1', '999.99', null, {
                format: 'bold;color:red;currency:USD'
            });
            const cell = model.getCell('A1');
            expect(cell.format).toBe('bold;color:red;currency:USD');
        });

        it('should export and import cell metadata', () => {
            model.setCell('A1', '100', null, {
                format: 'bold;number:0.00',
                comment: 'Test note'
            });

            const json = model.toJSON();
            const newModel = new SpreadsheetModel(100, 26);
            newModel.fromJSON(json);

            const cell = newModel.getCell('A1');
            expect(cell.value).toBe('100');
            expect(cell.format).toBe('bold;number:0.00');
            expect(cell.comment).toBe('Test note');
        });
    });

    describe('Row operations', () => {
        it('should insert a row and shift cells down', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('A3', '30');

            model.insertRow(2);

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('A2')).toBe(''); // New empty row
            expect(model.getCellValue('A3')).toBe('20'); // Shifted down
            expect(model.getCellValue('A4')).toBe('30'); // Shifted down
        });

        it('should delete a row and shift cells up', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('A3', '30');
            model.setCell('A4', '40');

            model.deleteRow(2);

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('A2')).toBe('30'); // Shifted up from A3
            expect(model.getCellValue('A3')).toBe('40'); // Shifted up from A4
            expect(model.getCellValue('A4')).toBe(''); // Now empty
        });

        it('should handle row insert at beginning', () => {
            model.setCell('A1', '10');
            model.setCell('A2', '20');

            model.insertRow(1);

            expect(model.getCellValue('A1')).toBe(''); // New row
            expect(model.getCellValue('A2')).toBe('10'); // Shifted
            expect(model.getCellValue('A3')).toBe('20'); // Shifted
        });

        it('should preserve cell metadata during row operations', () => {
            model.setCell('A1', '100', null, { format: 'bold', comment: 'Note' });
            model.setCell('A2', '200');

            model.insertRow(2);

            const cell = model.getCell('A1');
            expect(cell.value).toBe('100');
            expect(cell.format).toBe('bold');
            expect(cell.comment).toBe('Note');
        });

        it('should throw error for invalid row number', () => {
            expect(() => model.insertRow(0)).toThrow('Invalid row number');
            expect(() => model.insertRow(101)).toThrow('Invalid row number');
            expect(() => model.deleteRow(0)).toThrow('Invalid row number');
            expect(() => model.deleteRow(101)).toThrow('Invalid row number');
        });

        it('should handle multiple columns during row insert', () => {
            model.setCell('A1', 'A1-val');
            model.setCell('B1', 'B1-val');
            model.setCell('C1', 'C1-val');
            model.setCell('A2', 'A2-val');
            model.setCell('B2', 'B2-val');
            model.setCell('C2', 'C2-val');

            model.insertRow(2);

            expect(model.getCellValue('A1')).toBe('A1-val');
            expect(model.getCellValue('B1')).toBe('B1-val');
            expect(model.getCellValue('C1')).toBe('C1-val');
            expect(model.getCellValue('A2')).toBe('');
            expect(model.getCellValue('B2')).toBe('');
            expect(model.getCellValue('C2')).toBe('');
            expect(model.getCellValue('A3')).toBe('A2-val');
            expect(model.getCellValue('B3')).toBe('B2-val');
            expect(model.getCellValue('C3')).toBe('C2-val');
        });
    });

    describe('Column operations', () => {
        it('should insert a column and shift cells right', () => {
            model.setCell('A1', '10');
            model.setCell('B1', '20');
            model.setCell('C1', '30');

            model.insertColumn(2); // Insert before column B

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('B1')).toBe(''); // New empty column
            expect(model.getCellValue('C1')).toBe('20'); // Shifted right (was B1)
            expect(model.getCellValue('D1')).toBe('30'); // Shifted right (was C1)
        });

        it('should delete a column and shift cells left', () => {
            model.setCell('A1', '10');
            model.setCell('B1', '20');
            model.setCell('C1', '30');
            model.setCell('D1', '40');

            model.deleteColumn(2); // Delete column B

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('B1')).toBe('30'); // Shifted left from C1
            expect(model.getCellValue('C1')).toBe('40'); // Shifted left from D1
            expect(model.getCellValue('D1')).toBe(''); // Now empty
        });

        it('should accept column letter for insert', () => {
            model.setCell('A1', '10');
            model.setCell('B1', '20');

            model.insertColumn('B');

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('B1')).toBe(''); // New column
            expect(model.getCellValue('C1')).toBe('20'); // Shifted right
        });

        it('should accept column letter for delete', () => {
            model.setCell('A1', '10');
            model.setCell('B1', '20');
            model.setCell('C1', '30');

            model.deleteColumn('B');

            expect(model.getCellValue('A1')).toBe('10');
            expect(model.getCellValue('B1')).toBe('30'); // Shifted left from C1
            expect(model.getCellValue('C1')).toBe(''); // Now empty
        });

        it('should preserve cell metadata during column operations', () => {
            model.setCell('A1', '100', null, { format: 'bold', comment: 'Note' });
            model.setCell('B1', '200');

            model.insertColumn(2);

            const cell = model.getCell('A1');
            expect(cell.value).toBe('100');
            expect(cell.format).toBe('bold');
            expect(cell.comment).toBe('Note');
        });

        it('should throw error for invalid column number', () => {
            expect(() => model.insertColumn(0)).toThrow('Invalid column number');
            expect(() => model.insertColumn(27)).toThrow('Invalid column number');
            expect(() => model.deleteColumn(0)).toThrow('Invalid column number');
            expect(() => model.deleteColumn(27)).toThrow('Invalid column number');
        });

        it('should handle multiple rows during column insert', () => {
            model.setCell('A1', 'A1-val');
            model.setCell('B1', 'B1-val');
            model.setCell('A2', 'A2-val');
            model.setCell('B2', 'B2-val');

            model.insertColumn(2);

            expect(model.getCellValue('A1')).toBe('A1-val');
            expect(model.getCellValue('B1')).toBe('');
            expect(model.getCellValue('C1')).toBe('B1-val');
            expect(model.getCellValue('A2')).toBe('A2-val');
            expect(model.getCellValue('B2')).toBe('');
            expect(model.getCellValue('C2')).toBe('B2-val');
        });
    });

    describe('Formula adjustment during row/column operations', () => {
        describe('Insert row - formula adjustment', () => {
            it('should adjust cell references in formulas when inserting a row', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '=A1 + A2');

                model.insertRow(2);

                // A1 stays the same (10)
                expect(model.getCellValue('A1')).toBe('10');
                // A2 is now empty (new row)
                expect(model.getCellValue('A2')).toBe('');
                // A3 now has the value that was in A2 (20)
                expect(model.getCellValue('A3')).toBe('20');
                // A4 has the formula, and it should be adjusted to =A1 + A3
                const cell = model.getCell('A4');
                expect(cell.expression).toBe('A1 + A3');
            });

            it('should adjust range references when inserting a row', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '30');
                model.setCell('B1', '=SUM_RANGE("A1:A3")');

                model.insertRow(2);

                // Formula should adjust range to A1:A4
                const cell = model.getCell('B1');
                expect(cell.expression).toBe('SUM_RANGE("A1:A4")');
            });

            it('should handle multiple cell references in formula when inserting row', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '30');
                model.setCell('A4', '=A1 + A2 + A3');

                model.insertRow(2);

                const cell = model.getCell('A5');
                expect(cell.expression).toBe('A1 + A3 + A4');
            });

            it('should not adjust absolute row references when inserting row', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '=A1 + A$2');

                model.insertRow(2);

                const cell = model.getCell('A4');
                // A$2 should stay as A$2 because it's absolute
                expect(cell.expression).toBe('A1 + A$2');
            });
        });

        describe('Delete row - formula adjustment', () => {
            it('should adjust cell references in formulas when deleting a row', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '30');
                model.setCell('A4', '=A1 + A2 + A3');

                model.deleteRow(2);

                // A2 now has what was in A3 (30)
                expect(model.getCellValue('A2')).toBe('30');
                // A3 has the formula, adjusted to =A1 + #REF! + A2
                const cell = model.getCell('A3');
                expect(cell.expression).toBe('A1 + #REF! + A2');
            });

            it('should mark deleted cell references as #REF! when deleting row', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '=A2');

                model.deleteRow(2);

                const cell = model.getCell('A2');
                expect(cell.expression).toBe('#REF!');
            });

            it('should adjust range references when deleting a row', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '30');
                model.setCell('A4', '40');
                model.setCell('B1', '=SUM_RANGE("A1:A4")');

                model.deleteRow(2);

                // Formula should adjust range to A1:A3
                const cell = model.getCell('B1');
                expect(cell.expression).toBe('SUM_RANGE("A1:A3")');
            });

            it('should not adjust absolute row references when deleting row', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '30');
                model.setCell('A4', '=A1 + A$3');

                model.deleteRow(2);

                const cell = model.getCell('A3');
                // A$3 should stay as A$3 because it's absolute
                expect(cell.expression).toBe('A1 + A$3');
            });
        });

        describe('Insert column - formula adjustment', () => {
            it('should adjust cell references in formulas when inserting a column', () => {
                model.setCell('A1', '10');
                model.setCell('B1', '20');
                model.setCell('C1', '=A1 + B1');

                model.insertColumn(2); // Insert before column B

                // A1 stays the same (10)
                expect(model.getCellValue('A1')).toBe('10');
                // B1 is now empty (new column)
                expect(model.getCellValue('B1')).toBe('');
                // C1 now has the value that was in B1 (20)
                expect(model.getCellValue('C1')).toBe('20');
                // D1 has the formula, and it should be adjusted to =A1 + C1
                const cell = model.getCell('D1');
                expect(cell.expression).toBe('A1 + C1');
            });

            it('should adjust range references when inserting a column', () => {
                model.setCell('A1', '10');
                model.setCell('B1', '20');
                model.setCell('C1', '30');
                model.setCell('A2', '=SUM_RANGE("A1:C1")');

                model.insertColumn(2);

                // Formula should adjust range to A1:D1
                const cell = model.getCell('A2');
                expect(cell.expression).toBe('SUM_RANGE("A1:D1")');
            });

            it('should not adjust absolute column references when inserting column', () => {
                model.setCell('A1', '10');
                model.setCell('B1', '20');
                model.setCell('C1', '=$A1 + B1');

                model.insertColumn(2);

                const cell = model.getCell('D1');
                // $A1 should stay as $A1 because it's absolute
                expect(cell.expression).toBe('$A1 + C1');
            });
        });

        describe('Delete column - formula adjustment', () => {
            it('should adjust cell references in formulas when deleting a column', () => {
                model.setCell('A1', '10');
                model.setCell('B1', '20');
                model.setCell('C1', '30');
                model.setCell('D1', '=A1 + B1 + C1');

                model.deleteColumn(2); // Delete column B

                // B1 now has what was in C1 (30)
                expect(model.getCellValue('B1')).toBe('30');
                // C1 has the formula, adjusted to =A1 + #REF! + B1
                const cell = model.getCell('C1');
                expect(cell.expression).toBe('A1 + #REF! + B1');
            });

            it('should mark deleted cell references as #REF! when deleting column', () => {
                model.setCell('A1', '10');
                model.setCell('B1', '20');
                model.setCell('C1', '=B1');

                model.deleteColumn('B');

                const cell = model.getCell('B1');
                expect(cell.expression).toBe('#REF!');
            });

            it('should adjust range references when deleting a column', () => {
                model.setCell('A1', '10');
                model.setCell('B1', '20');
                model.setCell('C1', '30');
                model.setCell('D1', '40');
                model.setCell('A2', '=SUM_RANGE("A1:D1")');

                model.deleteColumn('B');

                // Formula should adjust range to A1:C1
                const cell = model.getCell('A2');
                expect(cell.expression).toBe('SUM_RANGE("A1:C1")');
            });

            it('should not adjust absolute column references when deleting column', () => {
                model.setCell('A1', '10');
                model.setCell('B1', '20');
                model.setCell('C1', '30');
                model.setCell('D1', '=$A1 + C1');

                model.deleteColumn('B');

                const cell = model.getCell('C1');
                // $A1 should stay as $A1 because it's absolute
                expect(cell.expression).toBe('$A1 + B1');
            });
        });

        describe('Complex formula adjustment scenarios', () => {
            it('should handle mixed absolute and relative references', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '30');
                model.setCell('A4', '=A1 + $A$2 + A3');

                model.insertRow(2);

                const cell = model.getCell('A5');
                // A1 becomes A1, $A$2 stays $A$2, A3 becomes A4
                expect(cell.expression).toBe('A1 + $A$2 + A4');
            });

            it('should handle formulas with functions and complex expressions', () => {
                model.setCell('A1', '10');
                model.setCell('A2', '20');
                model.setCell('A3', '=SUM_RANGE("A1:A2") * 2 + A1');

                model.insertRow(2);

                const cell = model.getCell('A4');
                expect(cell.expression).toBe('SUM_RANGE("A1:A3") * 2 + A1');
            });

            it('should preserve formulas that do not reference affected cells', () => {
                model.setCell('A1', '10');
                model.setCell('A5', '50');
                model.setCell('A6', '=A5 * 2');

                model.insertRow(2);

                const cell = model.getCell('A7');
                // A5 becomes A6 because it's shifted down
                expect(cell.expression).toBe('A6 * 2');
            });
        });
    });

    describe('Batch Control Bus Operations (Priority 2)', () => {
        let adapter;
        let controlFunctions;

        beforeEach(async () => {
            model = new SpreadsheetModel(100, 26);

            // Dynamically import ES modules
            const SpreadsheetRexxAdapterModule = await import('../src/spreadsheet-rexx-adapter.js');
            const controlFunctionsModule = await import('../src/spreadsheet-control-functions.js');

            const SpreadsheetRexxAdapter = SpreadsheetRexxAdapterModule.default;
            const { createSpreadsheetControlFunctions } = controlFunctionsModule;

            adapter = new SpreadsheetRexxAdapter(model);
            controlFunctions = createSpreadsheetControlFunctions(model, adapter);
        });

        describe('BATCH_SET_CELLS', () => {
            it('should set multiple cells with individual addresses', async () => {
                const updates = [
                    { address: 'A1', value: '100' },
                    { address: 'B2', value: '200' },
                    { address: 'C3', value: '300' }
                ];

                const result = await controlFunctions.BATCH_SET_CELLS(updates);
                const resultObj = JSON.parse(result);

                expect(resultObj.total).toBe(3);
                expect(resultObj.success).toBe(3);
                expect(resultObj.errors).toBe(0);

                expect(model.getCellValue('A1')).toBe('100');
                expect(model.getCellValue('B2')).toBe('200');
                expect(model.getCellValue('C3')).toBe('300');
            });

            it('should accept JSON string', async () => {
                const updatesJson = '[{"address":"A1","value":"100"},{"address":"B2","value":"200"}]';

                const result = await controlFunctions.BATCH_SET_CELLS(updatesJson);
                const resultObj = JSON.parse(result);

                expect(resultObj.total).toBe(2);
                expect(resultObj.success).toBe(2);

                expect(model.getCellValue('A1')).toBe('100');
                expect(model.getCellValue('B2')).toBe('200');
            });

            it('should handle mixed success and errors', async () => {
                const updates = [
                    { address: 'A1', value: '100' },
                    { address: 'INVALID', value: '200' },  // Invalid address
                    { address: 'C3', value: '300' }
                ];

                const result = await controlFunctions.BATCH_SET_CELLS(updates);
                const resultObj = JSON.parse(result);

                expect(resultObj.total).toBe(3);
                expect(resultObj.success).toBe(2); // A1 and C3 succeeded
                expect(resultObj.errors).toBe(1);
                expect(resultObj.errorDetails).toBeDefined();
                expect(resultObj.errorDetails.length).toBe(1);

                expect(model.getCellValue('A1')).toBe('100');
                expect(model.getCellValue('C3')).toBe('300');
            });

            it('should validate update structure', async () => {
                await expect(async () => {
                    await controlFunctions.BATCH_SET_CELLS([{ value: '100' }]); // Missing address
                }).rejects.toThrow('address');
            });

            it('should handle formulas in batch', async () => {
                const updates = [
                    { address: 'A1', value: '10' },
                    { address: 'A2', value: '20' },
                    { address: 'A3', value: '=A1+A2' }
                ];

                await controlFunctions.BATCH_SET_CELLS(updates);

                expect(model.getCellValue('A1')).toBe('10');
                expect(model.getCellValue('A2')).toBe('20');
                // Expression is stored without the '=' prefix
                expect(model.getCellExpression('A3')).toBe('A1+A2');
            });
        });

        describe('BATCH_EXECUTE', () => {
            it('should execute multiple commands in sequence', async () => {
                const commands = [
                    { command: 'SETCELL', args: ['A1', '100'] },
                    { command: 'SETCELL', args: ['A2', '200'] },
                    { command: 'SETCELL', args: ['A3', '=A1+A2'] }
                ];

                const result = await controlFunctions.BATCH_EXECUTE(commands);
                const results = JSON.parse(result);

                expect(results.length).toBe(3);
                expect(results[0].success).toBe(true);
                expect(results[1].success).toBe(true);
                expect(results[2].success).toBe(true);

                expect(model.getCellValue('A1')).toBe('100');
                expect(model.getCellValue('A2')).toBe('200');
            });

            it('should accept JSON string', async () => {
                const commandsJson = '[{"command":"SETCELL","args":["A1","100"]},{"command":"GETCELL","args":["A1"]}]';

                const result = await controlFunctions.BATCH_EXECUTE(commandsJson);
                const results = JSON.parse(result);

                expect(results.length).toBe(2);
                expect(results[0].success).toBe(true);
                expect(results[1].success).toBe(true);
            });

            it('should handle mixed command successes and failures', async () => {
                const commands = [
                    { command: 'SETCELL', args: ['A1', '100'] },
                    { command: 'INVALID_COMMAND', args: [] },
                    { command: 'SETCELL', args: ['A2', '200'] }
                ];

                const result = await controlFunctions.BATCH_EXECUTE(commands);
                const results = JSON.parse(result);

                expect(results.length).toBe(3);
                expect(results[0].success).toBe(true);
                expect(results[1].success).toBe(false);
                expect(results[1].error).toContain('Unknown command');
                expect(results[2].success).toBe(true);

                // First and third commands should have succeeded
                expect(model.getCellValue('A1')).toBe('100');
                expect(model.getCellValue('A2')).toBe('200');
            });

            it('should validate command structure', async () => {
                const commands = [
                    { args: ['A1', '100'] }  // Missing command field
                ];

                const result = await controlFunctions.BATCH_EXECUTE(commands);
                const results = JSON.parse(result);

                expect(results[0].success).toBe(false);
                expect(results[0].error).toContain('command');
            });

            it('should include command results in response', async () => {
                const commands = [
                    { command: 'SETCELL', args: ['A1', '100'] },
                    { command: 'GETCELL', args: ['A1'] }
                ];

                const result = await controlFunctions.BATCH_EXECUTE(commands);
                const results = JSON.parse(result);

                expect(results[1].result).toBe('100');
            });

            it('should handle complex batch operations', async () => {
                const commands = [
                    { command: 'SETCELL', args: ['A1', '10'] },
                    { command: 'SETCELL', args: ['A2', '20'] },
                    { command: 'SETCELL', args: ['A3', '30'] },
                    { command: 'SUM_RANGE', args: ['A1:A3'] }
                ];

                const result = await controlFunctions.BATCH_EXECUTE(commands);
                const results = JSON.parse(result);

                expect(results.length).toBe(4);
                expect(results[3].success).toBe(true);
                expect(results[3].result).toBe(60);
            });
        });

        describe('Performance benefits', () => {
            it('should update UI once for batch operations', async () => {
                const updates = [];
                for (let i = 1; i <= 100; i++) {
                    updates.push({ address: `A${i}`, value: String(i) });
                }

                const startTime = Date.now();
                await controlFunctions.BATCH_SET_CELLS(updates);
                const endTime = Date.now();

                // Batch operation should complete reasonably fast
                expect(endTime - startTime).toBeLessThan(1000);

                // Verify all cells were set
                expect(model.getCellValue('A1')).toBe('1');
                expect(model.getCellValue('A50')).toBe('50');
                expect(model.getCellValue('A100')).toBe('100');
            });
        });
    });

    describe('Auto-Increment Row IDs (Priority 4)', () => {
        describe('Configuration', () => {
            it('should configure auto-ID column', () => {
                model.configureAutoId('A', 1, 'ID-');

                expect(model.autoIdColumn).toBe('A');
                expect(model.nextId).toBe(1);
                expect(model.idPrefix).toBe('ID-');
            });

            it('should disable auto-ID when passed null', () => {
                model.configureAutoId('A', 1);
                model.configureAutoId(null);

                expect(model.autoIdColumn).toBeNull();
            });

            it('should throw error for invalid column letter', () => {
                expect(() => {
                    model.configureAutoId('123');
                }).toThrow('Invalid column letter');
            });

            it('should normalize column letter to uppercase', () => {
                model.configureAutoId('b', 1);
                expect(model.autoIdColumn).toBe('B');
            });
        });

        describe('Auto-generation on row insert', () => {
            beforeEach(() => {
                model.configureAutoId('A', 1);
            });

            it('should auto-generate ID when inserting a row', () => {
                model.insertRow(1);

                expect(model.getCellValue('A1')).toBe('1');
                expect(model.nextId).toBe(2);
            });

            it('should auto-generate sequential IDs for multiple rows', () => {
                model.insertRow(1);
                model.insertRow(2);
                model.insertRow(3);

                expect(model.getCellValue('A1')).toBe('1');
                expect(model.getCellValue('A2')).toBe('2');
                expect(model.getCellValue('A3')).toBe('3');
                expect(model.nextId).toBe(4);
            });

            it('should include prefix in generated IDs', () => {
                model.configureAutoId('A', 1000, 'ORDER-');

                model.insertRow(1);
                model.insertRow(2);

                expect(model.getCellValue('A1')).toBe('ORDER-1000');
                expect(model.getCellValue('A2')).toBe('ORDER-1001');
                expect(model.nextId).toBe(1002);
            });

            it('should not generate ID when auto-ID is disabled', () => {
                model.configureAutoId(null);

                model.insertRow(1);

                expect(model.getCellValue('A1')).toBe('');
            });

            it('should support different ID columns', () => {
                model.configureAutoId('C', 1);

                model.insertRow(1);

                expect(model.getCellValue('C1')).toBe('1');
                expect(model.getCellValue('A1')).toBe('');
            });

            it('should start from custom start ID', () => {
                model.configureAutoId('A', 5000);

                model.insertRow(1);
                model.insertRow(2);

                expect(model.getCellValue('A1')).toBe('5000');
                expect(model.getCellValue('A2')).toBe('5001');
            });
        });

        describe('Finding rows by ID', () => {
            beforeEach(() => {
                model.configureAutoId('A', 100);
                model.insertRow(1);
                model.insertRow(2);
                model.insertRow(3);
            });

            it('should find row by ID', () => {
                const rowNum = model.findRowById('100');
                expect(rowNum).toBe(1);

                const rowNum2 = model.findRowById('101');
                expect(rowNum2).toBe(2);
            });

            it('should handle numeric ID values', () => {
                const rowNum = model.findRowById(100);
                expect(rowNum).toBe(1);
            });

            it('should return null when ID not found', () => {
                const rowNum = model.findRowById('999');
                expect(rowNum).toBeNull();
            });

            it('should throw error when auto-ID not configured', () => {
                model.configureAutoId(null);

                expect(() => {
                    model.findRowById('100');
                }).toThrow('Auto-ID column is not configured');
            });

            it('should find rows with prefixed IDs', () => {
                model.configureAutoId('A', 1, 'USER-');
                model.insertRow(1);

                const rowNum = model.findRowById('USER-1');
                expect(rowNum).toBe(1);
            });
        });

        describe('Get next ID', () => {
            it('should return next ID without prefix', () => {
                model.configureAutoId('A', 42);
                expect(model.getNextId()).toBe('42');
            });

            it('should return next ID with prefix', () => {
                model.configureAutoId('A', 100, 'ORD-');
                expect(model.getNextId()).toBe('ORD-100');
            });

            it('should update after row insert', () => {
                model.configureAutoId('A', 1);

                expect(model.getNextId()).toBe('1');
                model.insertRow(1);
                expect(model.getNextId()).toBe('2');
            });
        });

        describe('Control Bus Commands', () => {
            let adapter;
            let controlFunctions;

            beforeEach(async () => {
                model = new SpreadsheetModel(100, 26);

                const SpreadsheetRexxAdapterModule = await import('../src/spreadsheet-rexx-adapter.js');
                const controlFunctionsModule = await import('../src/spreadsheet-control-functions.js');

                const SpreadsheetRexxAdapter = SpreadsheetRexxAdapterModule.default;
                const { createSpreadsheetControlFunctions } = controlFunctionsModule;

                adapter = new SpreadsheetRexxAdapter(model);
                controlFunctions = createSpreadsheetControlFunctions(model, adapter);
            });

            it('should configure auto-ID via Control Bus', async () => {
                const result = await controlFunctions.CONFIGURE_AUTO_ID('A', 100, 'ID-');

                expect(result).toContain('Auto-ID configured');
                expect(model.autoIdColumn).toBe('A');
                expect(model.nextId).toBe(100);
                expect(model.idPrefix).toBe('ID-');
            });

            it('should disable auto-ID via Control Bus', async () => {
                await controlFunctions.CONFIGURE_AUTO_ID('A', 1);
                const result = await controlFunctions.CONFIGURE_AUTO_ID(null);

                expect(result).toBe('Auto-ID disabled');
                expect(model.autoIdColumn).toBeNull();
            });

            it('should find row by ID via Control Bus', async () => {
                model.configureAutoId('A', 100);
                model.insertRow(1);
                model.insertRow(2);

                const rowNum = controlFunctions.FIND_ROW_BY_ID('100');
                expect(rowNum).toBe('1');

                const rowNum2 = controlFunctions.FIND_ROW_BY_ID(101);
                expect(rowNum2).toBe('2');
            });

            it('should return empty string when ID not found', () => {
                model.configureAutoId('A', 100);

                const rowNum = controlFunctions.FIND_ROW_BY_ID('999');
                expect(rowNum).toBe('');
            });

            it('should update row by ID via Control Bus', async () => {
                model.configureAutoId('A', 100);
                model.insertRow(1);
                model.setCell('B1', 'OldName');
                model.setCell('C1', '25');

                const updates = [
                    { column: 'B', value: 'NewName' },
                    { column: 'C', value: '30' }
                ];

                const rowNum = await controlFunctions.UPDATE_ROW_BY_ID('100', updates);

                expect(rowNum).toBe('1');
                expect(model.getCellValue('B1')).toBe('NewName');
                expect(model.getCellValue('C1')).toBe('30');
            });

            it('should accept JSON string for UPDATE_ROW_BY_ID', async () => {
                model.configureAutoId('A', 100);
                model.insertRow(1);

                const updatesJson = '[{"column":"B","value":"Test"}]';
                await controlFunctions.UPDATE_ROW_BY_ID('100', updatesJson);

                expect(model.getCellValue('B1')).toBe('Test');
            });

            it('should throw error when updating non-existent ID', async () => {
                model.configureAutoId('A', 100);

                await expect(async () => {
                    await controlFunctions.UPDATE_ROW_BY_ID('999', [{ column: 'B', value: 'Test' }]);
                }).rejects.toThrow('Row with ID 999 not found');
            });

            it('should get next ID via Control Bus', () => {
                model.configureAutoId('A', 5000, 'CUST-');

                const nextId = controlFunctions.GET_NEXT_ID();
                expect(nextId).toBe('CUST-5000');
            });

            it('should throw error when getting next ID without configuration', () => {
                expect(() => {
                    controlFunctions.GET_NEXT_ID();
                }).toThrow('Auto-ID column is not configured');
            });
        });

        describe('Integration with other features', () => {
            it('should preserve IDs when sorting', () => {
                model.configureAutoId('A', 1);
                model.insertRow(1);
                model.insertRow(2);
                model.insertRow(3);

                model.setCell('B1', '30');
                model.setCell('B2', '10');
                model.setCell('B3', '20');

                model.sortRange('A1:B3', 'B', true);

                // IDs should stay with their rows
                expect(model.getCellValue('A1')).toBe('2'); // Row with B=10
                expect(model.getCellValue('A2')).toBe('3'); // Row with B=20
                expect(model.getCellValue('A3')).toBe('1'); // Row with B=30
            });

            it('should work across multiple sheets', () => {
                model.addSheet('Sheet2');

                // Configure different auto-IDs for each sheet
                model.setActiveSheet('Sheet1');
                model.configureAutoId('A', 1);
                model.insertRow(1);

                model.setActiveSheet('Sheet2');
                model.configureAutoId('A', 100, 'S2-');
                model.insertRow(1);

                // Verify each sheet has independent IDs
                model.setActiveSheet('Sheet1');
                expect(model.getCellValue('A1')).toBe('1');

                model.setActiveSheet('Sheet2');
                expect(model.getCellValue('A1')).toBe('S2-100');
            });
        });
    });

    describe('Table Metadata Control Bus Commands (Priority 5)', () => {
        let model, controlFunctions;

        beforeEach(async () => {
            model = new SpreadsheetModel(100, 26);

            // Set up sample data
            model.setCell('A1', 'id');
            model.setCell('B1', 'name');
            model.setCell('C1', 'amount');

            model.setCell('A2', '1');
            model.setCell('B2', 'Alice');
            model.setCell('C2', '100');

            model.setCell('A3', '2');
            model.setCell('B3', 'Bob');
            model.setCell('C3', '200');

            // Import control functions and adapter
            const SpreadsheetRexxAdapterModule = await import('../src/spreadsheet-rexx-adapter.js');
            const SpreadsheetRexxAdapter = SpreadsheetRexxAdapterModule.default;
            const adapter = new SpreadsheetRexxAdapter(model);

            const controlFunctionsModule = await import('../src/spreadsheet-control-functions.js');
            const createSpreadsheetControlFunctions = controlFunctionsModule.createSpreadsheetControlFunctions;
            controlFunctions = createSpreadsheetControlFunctions(model, adapter);
        });

        it('should set table metadata via Control Bus', () => {
            const metadata = {
                range: 'A1:C3',
                columns: {
                    id: 'A',
                    name: 'B',
                    amount: 'C'
                },
                hasHeader: true
            };

            const result = controlFunctions.SET_TABLE_METADATA('TestTable', JSON.stringify(metadata));

            expect(result).toContain("Table metadata set for 'TestTable'");
            expect(model.getTableMetadata('TestTable')).toBeDefined();
        });

        it('should get table metadata via Control Bus', () => {
            model.setTableMetadata('TestTable', {
                range: 'A1:C3',
                columns: {
                    id: 'A',
                    name: 'B',
                    amount: 'C'
                },
                hasHeader: true
            });

            const result = controlFunctions.GET_TABLE_METADATA('TestTable');
            const parsed = JSON.parse(result);

            expect(parsed.range).toBe('A1:C3');
            expect(parsed.columns.id).toBe('A');
            expect(parsed.hasHeader).toBe(true);
        });

        it('should return empty string for non-existent table', () => {
            const result = controlFunctions.GET_TABLE_METADATA('NonExistent');
            expect(result).toBe('');
        });

        it('should delete table metadata via Control Bus', () => {
            model.setTableMetadata('TestTable', {
                range: 'A1:C3',
                columns: { id: 'A' }
            });

            const result = controlFunctions.DELETE_TABLE_METADATA('TestTable');

            expect(result).toContain("Table metadata deleted for 'TestTable'");
            expect(model.getTableMetadata('TestTable')).toBeNull();
        });

        it('should list all tables via Control Bus', () => {
            model.setTableMetadata('Table1', {
                range: 'A1:C3',
                columns: { id: 'A' }
            });

            model.setTableMetadata('Table2', {
                range: 'D1:F5',
                columns: { x: 'D' }
            });

            const result = controlFunctions.LIST_TABLES();
            const parsed = JSON.parse(result);

            expect(parsed).toContain('Table1');
            expect(parsed).toContain('Table2');
            expect(parsed.length).toBe(2);
        });

        it('should accept metadata as object (not just JSON string)', () => {
            const metadata = {
                range: 'A1:C3',
                columns: { id: 'A', name: 'B' },
                hasHeader: true
            };

            const result = controlFunctions.SET_TABLE_METADATA('TestTable', metadata);

            expect(result).toContain("Table metadata set for 'TestTable'");
            expect(model.getTableMetadata('TestTable')).toBeDefined();
        });

        it('should reject invalid table name', () => {
            expect(() => {
                controlFunctions.SET_TABLE_METADATA('123Invalid', '{"range":"A1:C3","columns":{"id":"A"}}');
            }).toThrow('Table name must start with a letter');
        });

        it('should reject invalid JSON', () => {
            expect(() => {
                controlFunctions.SET_TABLE_METADATA('TestTable', '{invalid json}');
            }).toThrow('Invalid JSON for table metadata');
        });

        it('should reject missing table name parameter', () => {
            expect(() => {
                controlFunctions.SET_TABLE_METADATA('', '{"range":"A1:C3","columns":{"id":"A"}}');
            }).toThrow('Table name is required');
        });
    });
});
