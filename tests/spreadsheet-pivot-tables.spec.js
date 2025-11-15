/**
 * Tests for Pivot Table Functionality
 */

import SpreadsheetModel from '../src/spreadsheet-model.js';

describe('Pivot Table Functionality', () => {
    let model;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);

        // Set up sample data for pivot tables
        // Headers
        model.setCell('A1', 'Category');
        model.setCell('B1', 'Month');
        model.setCell('C1', 'Region');
        model.setCell('D1', 'Sales');

        // Data rows
        model.setCell('A2', 'Electronics');
        model.setCell('B2', 'Jan');
        model.setCell('C2', 'East');
        model.setCell('D2', '100');

        model.setCell('A3', 'Electronics');
        model.setCell('B3', 'Jan');
        model.setCell('C3', 'West');
        model.setCell('D3', '150');

        model.setCell('A4', 'Electronics');
        model.setCell('B4', 'Feb');
        model.setCell('C4', 'East');
        model.setCell('D4', '200');

        model.setCell('A5', 'Furniture');
        model.setCell('B5', 'Jan');
        model.setCell('C5', 'East');
        model.setCell('D5', '300');

        model.setCell('A6', 'Furniture');
        model.setCell('B6', 'Feb');
        model.setCell('C6', 'West');
        model.setCell('D6', '250');
    });

    describe('createPivotTable', () => {
        it('should create a basic pivot table with row grouping', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            const pivotId = model.createPivotTable('test_pivot', config);
            expect(pivotId).toBe('test_pivot');

            // Verify pivot was stored
            const stored = model.getPivotTable('test_pivot');
            expect(stored).toBeTruthy();
            expect(stored.sourceRange).toBe('A1:D6');
            expect(stored.aggFunction).toBe('SUM');
        });

        it('should create pivot table with row and column grouping', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: ['Month'],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('pivot2', config);

            // Check that output cells are populated
            const categoryHeader = model.getCellValue('F1');
            expect(categoryHeader).toBe('Category');

            // Verify data is written to cells
            const pivotResult = model.getCellValue('G2'); // Should contain aggregated value
            expect(pivotResult).not.toBe('');
        });

        it('should support SUM aggregation', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('sum_pivot', config);

            // Electronics: 100 + 150 + 200 = 450
            // Furniture: 300 + 250 = 550
            // Verify the sums are correct (exact cell positions depend on pivot layout)
            const cells = model.getAllCells();
            const values = Array.from(cells.values()).map(c => c.value);

            // Should include the summed values
            expect(values.includes('450')).toBe(true);
            expect(values.includes('550')).toBe(true);
        });

        it('should support COUNT aggregation', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'COUNT',
                outputCell: 'F1'
            };

            model.createPivotTable('count_pivot', config);

            const cells = model.getAllCells();
            const values = Array.from(cells.values()).map(c => c.value);

            // Electronics appears 3 times, Furniture 2 times
            expect(values.includes('3')).toBe(true);
            expect(values.includes('2')).toBe(true);
        });

        it('should support AVERAGE aggregation', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'AVERAGE',
                outputCell: 'F1'
            };

            model.createPivotTable('avg_pivot', config);

            const cells = model.getAllCells();
            const values = Array.from(cells.values()).map(c => c.value);

            // Electronics average: (100+150+200)/3 = 150
            // Furniture average: (300+250)/2 = 275
            expect(values.includes('150')).toBe(true);
            expect(values.includes('275')).toBe(true);
        });

        it('should support MIN aggregation', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'MIN',
                outputCell: 'F1'
            };

            model.createPivotTable('min_pivot', config);

            const cells = model.getAllCells();
            const values = Array.from(cells.values()).map(c => c.value);

            // Electronics min: 100
            // Furniture min: 250
            expect(values.includes('100')).toBe(true);
            expect(values.includes('250')).toBe(true);
        });

        it('should support MAX aggregation', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'MAX',
                outputCell: 'F1'
            };

            model.createPivotTable('max_pivot', config);

            const cells = model.getAllCells();
            const values = Array.from(cells.values()).map(c => c.value);

            // Electronics max: 200
            // Furniture max: 300
            expect(values.includes('200')).toBe(true);
            expect(values.includes('300')).toBe(true);
        });

        it('should throw error for invalid aggregation function', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'INVALID',
                outputCell: 'F1'
            };

            expect(() => model.createPivotTable('bad_pivot', config)).toThrow();
        });

        it('should throw error for missing required fields', () => {
            const config = {
                sourceRange: 'A1:D6',
                outputCell: 'F1'
                // Missing valueField and aggFunction
            };

            expect(() => model.createPivotTable('bad_pivot', config)).toThrow();
        });
    });

    describe('updatePivotTable', () => {
        it('should refresh pivot table when source data changes', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('update_pivot', config);

            // Get initial Electronics sum (450)
            const cells1 = model.getAllCells();
            const values1 = Array.from(cells1.values()).map(c => c.value);
            expect(values1.includes('450')).toBe(true);

            // Update source data
            model.setCell('D2', '500'); // Change 100 to 500

            // Refresh pivot
            model.updatePivotTable('update_pivot');

            // New Electronics sum should be 850 (500 + 150 + 200)
            const cells2 = model.getAllCells();
            const values2 = Array.from(cells2.values()).map(c => c.value);
            expect(values2.includes('850')).toBe(true);
        });

        it('should throw error for non-existent pivot', () => {
            expect(() => model.updatePivotTable('nonexistent')).toThrow();
        });
    });

    describe('deletePivotTable', () => {
        it('should delete a pivot table', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('delete_pivot', config);
            expect(model.getPivotTable('delete_pivot')).toBeTruthy();

            model.deletePivotTable('delete_pivot');
            expect(model.getPivotTable('delete_pivot')).toBeNull();
        });

        it('should throw error for non-existent pivot', () => {
            expect(() => model.deletePivotTable('nonexistent')).toThrow();
        });
    });

    describe('getPivotTable', () => {
        it('should return null for non-existent pivot', () => {
            expect(model.getPivotTable('nonexistent')).toBeNull();
        });

        it('should return pivot configuration', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: ['Month'],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('get_pivot', config);

            const retrieved = model.getPivotTable('get_pivot');
            expect(retrieved.sourceRange).toBe('A1:D6');
            expect(retrieved.rowFields).toEqual(['Category']);
            expect(retrieved.colFields).toEqual(['Month']);
            expect(retrieved.valueField).toBe('Sales');
            expect(retrieved.aggFunction).toBe('SUM');
        });
    });

    describe('getAllPivotTables', () => {
        it('should return empty array when no pivots exist', () => {
            const pivots = model.getAllPivotTables();
            expect(pivots).toEqual([]);
        });

        it('should return all pivot tables', () => {
            const config1 = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            const config2 = {
                sourceRange: 'A1:D6',
                rowFields: ['Region'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'COUNT',
                outputCell: 'J1'
            };

            model.createPivotTable('pivot1', config1);
            model.createPivotTable('pivot2', config2);

            const pivots = model.getAllPivotTables();
            expect(pivots.length).toBe(2);
            expect(pivots.map(p => p.id)).toContain('pivot1');
            expect(pivots.map(p => p.id)).toContain('pivot2');
        });
    });

    describe('JSON serialization', () => {
        it('should export pivot tables to JSON', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: ['Month'],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('export_pivot', config);

            const json = model.toJSON();
            expect(json.sheets.Sheet1.pivotTables).toBeDefined();
            expect(json.sheets.Sheet1.pivotTables.export_pivot).toBeDefined();
            expect(json.sheets.Sheet1.pivotTables.export_pivot.sourceRange).toBe('A1:D6');
        });

        it('should import pivot tables from JSON', () => {
            const data = {
                version: 2,
                activeSheetName: 'Sheet1',
                sheetOrder: ['Sheet1'],
                sheets: {
                    Sheet1: {
                        cells: {},
                        pivotTables: {
                            'import_pivot': {
                                sourceRange: 'A1:D10',
                                rowFields: ['Product'],
                                colFields: ['Quarter'],
                                valueField: 'Revenue',
                                aggFunction: 'AVERAGE',
                                outputCell: 'F1',
                                filters: {}
                            }
                        }
                    }
                }
            };

            model.fromJSON(data);

            const pivot = model.getPivotTable('import_pivot');
            expect(pivot).toBeTruthy();
            expect(pivot.sourceRange).toBe('A1:D10');
            expect(pivot.aggFunction).toBe('AVERAGE');
        });
    });

    describe('Multi-sheet support', () => {
        it('should handle pivot tables per sheet', () => {
            const config1 = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('sheet1_pivot', config1);

            model.addSheet('Sheet2');
            model.setActiveSheet('Sheet2');

            const config2 = {
                sourceRange: 'A1:C5',
                rowFields: ['Product'],
                colFields: [],
                valueField: 'Quantity',
                aggFunction: 'COUNT',
                outputCell: 'E1'
            };

            model.createPivotTable('sheet2_pivot', config2);

            // Verify Sheet1 pivot
            model.setActiveSheet('Sheet1');
            expect(model.getPivotTable('sheet1_pivot')).toBeTruthy();
            expect(model.getPivotTable('sheet2_pivot')).toBeNull();

            // Verify Sheet2 pivot
            model.setActiveSheet('Sheet2');
            expect(model.getPivotTable('sheet2_pivot')).toBeTruthy();
            expect(model.getPivotTable('sheet1_pivot')).toBeNull();
        });
    });

    describe('Complex pivot scenarios', () => {
        it('should handle multiple row fields', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category', 'Month'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('multi_row_pivot', config);

            // Verify headers are written
            expect(model.getCellValue('F1')).toBe('Category');
            expect(model.getCellValue('G1')).toBe('Month');
        });

        it('should handle multiple column fields', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: ['Month', 'Region'],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1'
            };

            model.createPivotTable('multi_col_pivot', config);

            // Should create pivot with multi-level column headers
            const pivot = model.getPivotTable('multi_col_pivot');
            expect(pivot.colFields).toEqual(['Month', 'Region']);
        });

        it('should handle filters in pivot configuration', () => {
            const config = {
                sourceRange: 'A1:D6',
                rowFields: ['Category'],
                colFields: [],
                valueField: 'Sales',
                aggFunction: 'SUM',
                outputCell: 'F1',
                filters: { Region: 'East' }
            };

            model.createPivotTable('filtered_pivot', config);

            // With Region=East filter:
            // Electronics (East): 100 + 200 = 300
            // Furniture (East): 300
            const cells = model.getAllCells();
            const values = Array.from(cells.values()).map(c => c.value);

            expect(values.includes('300')).toBe(true);
        });
    });
});
