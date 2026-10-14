/**
 * Tests for Merge Cells and Custom Cell Editors
 */

import SpreadsheetModel from '../src/spreadsheet-model.js';

describe('Merge Cells Functionality', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
    });

    describe('mergeCells', () => {
        it('should merge cells in a range', () => {
            const topLeft = model.mergeCells('A1:C3');
            expect(topLeft).toBe('A1');

            const mergeInfo = model.getMergedRange('A1');
            expect(mergeInfo).toBeTruthy();
            expect(mergeInfo.topLeft).toBe('A1');
            expect(mergeInfo.bottomRight).toBe('C3');
            expect(mergeInfo.range).toBe('A1:C3');
        });

        it('should detect merged cells anywhere in the range', () => {
            model.mergeCells('B2:D4');

            expect(model.isCellMerged('B2')).toBe(true);
            expect(model.isCellMerged('C3')).toBe(true);
            expect(model.isCellMerged('D4')).toBe(true);
            expect(model.isCellMerged('A1')).toBe(false);
            expect(model.isCellMerged('E5')).toBe(false);
        });

        it('should return same merge info for any cell in the range', () => {
            model.mergeCells('A1:C3');

            const info1 = model.getMergedRange('A1');
            const info2 = model.getMergedRange('B2');
            const info3 = model.getMergedRange('C3');

            expect(info1).toEqual(info2);
            expect(info2).toEqual(info3);
        });

        it('should throw error if range is invalid', () => {
            expect(() => model.mergeCells('InvalidRange')).toThrow();
        });

        it('should throw error if cells are already merged', () => {
            model.mergeCells('A1:C3');
            expect(() => model.mergeCells('B2:D4')).toThrow();
        });

        it('should allow merging non-overlapping ranges', () => {
            model.mergeCells('A1:B2');
            model.mergeCells('D1:E2');

            expect(model.isCellMerged('A1')).toBe(true);
            expect(model.isCellMerged('D1')).toBe(true);
            expect(model.isCellMerged('C1')).toBe(false);
        });
    });

    describe('unmergeCells', () => {
        it('should unmerge cells', () => {
            model.mergeCells('A1:C3');
            expect(model.isCellMerged('A1')).toBe(true);

            model.unmergeCells('A1');
            expect(model.isCellMerged('A1')).toBe(false);
            expect(model.isCellMerged('B2')).toBe(false);
            expect(model.isCellMerged('C3')).toBe(false);
        });

        it('should unmerge by any cell in the range', () => {
            model.mergeCells('A1:C3');

            model.unmergeCells('B2'); // Middle cell
            expect(model.isCellMerged('A1')).toBe(false);
        });

        it('should throw error if cell is not merged', () => {
            expect(() => model.unmergeCells('A1')).toThrow();
        });
    });

    describe('getMergedRange', () => {
        it('should return null for non-merged cells', () => {
            const info = model.getMergedRange('A1');
            expect(info).toBeNull();
        });

        it('should return merge info for merged cells', () => {
            model.mergeCells('B2:D5');

            const info = model.getMergedRange('C3');
            expect(info).toBeTruthy();
            expect(info.topLeft).toBe('B2');
            expect(info.bottomRight).toBe('D5');
            expect(info.range).toBe('B2:D5');
        });
    });

    describe('JSON serialization', () => {
        it('should export merged cells to JSON', () => {
            model.mergeCells('A1:C3');
            model.mergeCells('E5:F6');

            const json = model.toJSON();
            expect(json.sheets.Sheet1.mergedCells).toBeDefined();
            expect(json.sheets.Sheet1.mergedCells.A1).toBe('C3');
            expect(json.sheets.Sheet1.mergedCells.E5).toBe('F6');
        });

        it('should import merged cells from JSON', () => {
            const data = {
                version: 2,
                activeSheetName: 'Sheet1',
                sheetOrder: ['Sheet1'],
                sheets: {
                    Sheet1: {
                        cells: {},
                        mergedCells: {
                            'A1': 'C3',
                            'E5': 'F6'
                        },
                        cellEditors: {}
                    }
                }
            };

            model.fromJSON(data);

            expect(model.isCellMerged('A1')).toBe(true);
            expect(model.isCellMerged('B2')).toBe(true);
            expect(model.isCellMerged('E5')).toBe(true);
            expect(model.isCellMerged('F6')).toBe(true);
        });
    });

    describe('Multi-sheet support', () => {
        it('should handle merged cells per sheet', () => {
            model.addSheet('Sheet2');

            model.setActiveSheet('Sheet1');
            model.mergeCells('A1:B2');

            model.setActiveSheet('Sheet2');
            model.mergeCells('A1:C3');

            model.setActiveSheet('Sheet1');
            const info1 = model.getMergedRange('A1');
            expect(info1.range).toBe('A1:B2');

            model.setActiveSheet('Sheet2');
            const info2 = model.getMergedRange('A1');
            expect(info2.range).toBe('A1:C3');
        });
    });
});

describe('Custom Cell Editors', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
    });

    describe('setCellEditor', () => {
        it('should set checkbox editor', () => {
            model.setCellEditor('A1', 'checkbox');

            const editor = model.getCellEditor('A1');
            expect(editor).toBeTruthy();
            expect(editor.type).toBe('checkbox');
        });

        it('should set dropdown editor with options', () => {
            model.setCellEditor('B1', 'dropdown', { options: ['Red', 'Green', 'Blue'] });

            const editor = model.getCellEditor('B1');
            expect(editor).toBeTruthy();
            expect(editor.type).toBe('dropdown');
            expect(editor.config.options).toEqual(['Red', 'Green', 'Blue']);
        });

        it('should set date editor with format', () => {
            model.setCellEditor('C1', 'date', { format: 'YYYY-MM-DD' });

            const editor = model.getCellEditor('C1');
            expect(editor).toBeTruthy();
            expect(editor.type).toBe('date');
            expect(editor.config.format).toBe('YYYY-MM-DD');
        });

        it('should throw error for invalid editor type', () => {
            expect(() => model.setCellEditor('A1', 'invalid')).toThrow();
        });

        it('should throw error for dropdown without options', () => {
            expect(() => model.setCellEditor('A1', 'dropdown')).toThrow();
        });

        it('should allow dropdown with empty options array', () => {
            model.setCellEditor('A1', 'dropdown', { options: [] });
            const editor = model.getCellEditor('A1');
            expect(editor.config.options).toEqual([]);
        });
    });

    describe('getCellEditor', () => {
        it('should return null for cells without custom editor', () => {
            const editor = model.getCellEditor('A1');
            expect(editor).toBeNull();
        });

        it('should return editor info for cells with custom editor', () => {
            model.setCellEditor('A1', 'checkbox', { checked: true });

            const editor = model.getCellEditor('A1');
            expect(editor.type).toBe('checkbox');
            expect(editor.config.checked).toBe(true);
        });
    });

    describe('removeCellEditor', () => {
        it('should remove custom editor from cell', () => {
            model.setCellEditor('A1', 'checkbox');
            expect(model.getCellEditor('A1')).toBeTruthy();

            model.removeCellEditor('A1');
            expect(model.getCellEditor('A1')).toBeNull();
        });
    });

    describe('JSON serialization', () => {
        it('should export cell editors to JSON', () => {
            model.setCellEditor('A1', 'checkbox');
            model.setCellEditor('B1', 'dropdown', { options: ['Yes', 'No'] });
            model.setCellEditor('C1', 'date', { format: 'MM/DD/YYYY' });

            const json = model.toJSON();
            expect(json.sheets.Sheet1.cellEditors).toBeDefined();
            expect(json.sheets.Sheet1.cellEditors.A1.type).toBe('checkbox');
            expect(json.sheets.Sheet1.cellEditors.B1.type).toBe('dropdown');
            expect(json.sheets.Sheet1.cellEditors.C1.type).toBe('date');
        });

        it('should import cell editors from JSON', () => {
            const data = {
                version: 2,
                activeSheetName: 'Sheet1',
                sheetOrder: ['Sheet1'],
                sheets: {
                    Sheet1: {
                        cells: {},
                        mergedCells: {},
                        cellEditors: {
                            'A1': { type: 'checkbox', config: {} },
                            'B1': { type: 'dropdown', config: { options: ['A', 'B', 'C'] } },
                            'C1': { type: 'date', config: { format: 'YYYY-MM-DD' } }
                        }
                    }
                }
            };

            model.fromJSON(data);

            expect(model.getCellEditor('A1').type).toBe('checkbox');
            expect(model.getCellEditor('B1').type).toBe('dropdown');
            expect(model.getCellEditor('B1').config.options).toEqual(['A', 'B', 'C']);
            expect(model.getCellEditor('C1').type).toBe('date');
        });
    });

    describe('Multi-sheet support', () => {
        it('should handle cell editors per sheet', () => {
            model.addSheet('Sheet2');

            model.setActiveSheet('Sheet1');
            model.setCellEditor('A1', 'checkbox');

            model.setActiveSheet('Sheet2');
            model.setCellEditor('A1', 'dropdown', { options: ['X', 'Y'] });

            model.setActiveSheet('Sheet1');
            expect(model.getCellEditor('A1').type).toBe('checkbox');

            model.setActiveSheet('Sheet2');
            expect(model.getCellEditor('A1').type).toBe('dropdown');
        });
    });

    describe('Complex configurations', () => {
        it('should support dropdown with multiple configuration options', () => {
            model.setCellEditor('A1', 'dropdown', {
                options: ['Option 1', 'Option 2', 'Option 3'],
                allowCustom: true,
                placeholder: 'Select an option'
            });

            const editor = model.getCellEditor('A1');
            expect(editor.config.allowCustom).toBe(true);
            expect(editor.config.placeholder).toBe('Select an option');
        });

        it('should support date picker with various formats', () => {
            model.setCellEditor('A1', 'date', {
                format: 'DD/MM/YYYY',
                minDate: '2020-01-01',
                maxDate: '2025-12-31'
            });

            const editor = model.getCellEditor('A1');
            expect(editor.config.format).toBe('DD/MM/YYYY');
            expect(editor.config.minDate).toBe('2020-01-01');
            expect(editor.config.maxDate).toBe('2025-12-31');
        });

        it('should support checkbox with default state', () => {
            model.setCellEditor('A1', 'checkbox', {
                defaultChecked: true,
                label: 'Enable feature'
            });

            const editor = model.getCellEditor('A1');
            expect(editor.config.defaultChecked).toBe(true);
            expect(editor.config.label).toBe('Enable feature');
        });
    });
});
