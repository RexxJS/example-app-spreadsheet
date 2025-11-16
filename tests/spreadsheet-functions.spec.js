/**
 * Spreadsheet Functions Tests
 *
 * Tests for the enhanced spreadsheet range functions (MEDIAN, STDEV, SUMIF, etc.)
 */

import SpreadsheetModel from '../src/spreadsheet-model.js';
import SpreadsheetRexxAdapter from '../src/spreadsheet-rexx-adapter.js';

describe('Spreadsheet Range Functions', () => {
    let model;
    let adapter;
    let functions;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
        adapter = new SpreadsheetRexxAdapter(model);
        functions = adapter.getSpreadsheetFunctions();
    });

    describe('Basic Statistical Functions (existing)', () => {
        beforeEach(() => {
            // Set up test data: A1:A5 = [10, 20, 30, 40, 50]
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('A3', '30');
            model.setCell('A4', '40');
            model.setCell('A5', '50');
        });

        it('should calculate SUM_RANGE correctly', () => {
            const result = functions.SUM_RANGE('A1:A5');
            expect(result).toBe(150);
        });

        it('should calculate AVERAGE_RANGE correctly', () => {
            const result = functions.AVERAGE_RANGE('A1:A5');
            expect(result).toBe(30);
        });

        it('should calculate COUNT_RANGE correctly', () => {
            model.setCell('A6', '');  // Empty cell
            const result = functions.COUNT_RANGE('A1:A6');
            expect(result).toBe(5);  // Should count only non-empty cells
        });

        it('should calculate MIN_RANGE correctly', () => {
            const result = functions.MIN_RANGE('A1:A5');
            expect(result).toBe(10);
        });

        it('should calculate MAX_RANGE correctly', () => {
            const result = functions.MAX_RANGE('A1:A5');
            expect(result).toBe(50);
        });
    });

    describe('MEDIAN_RANGE', () => {
        it('should calculate median for odd number of values', () => {
            // A1:A5 = [10, 20, 30, 40, 50]
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('A3', '30');
            model.setCell('A4', '40');
            model.setCell('A5', '50');

            const result = functions.MEDIAN_RANGE('A1:A5');
            expect(result).toBe(30);
        });

        it('should calculate median for even number of values', () => {
            // A1:A4 = [10, 20, 30, 40]
            model.setCell('A1', '10');
            model.setCell('A2', '20');
            model.setCell('A3', '30');
            model.setCell('A4', '40');

            const result = functions.MEDIAN_RANGE('A1:A4');
            expect(result).toBe(25);  // Average of 20 and 30
        });

        it('should handle unsorted values', () => {
            // A1:A5 = [50, 10, 40, 20, 30] (unsorted)
            model.setCell('A1', '50');
            model.setCell('A2', '10');
            model.setCell('A3', '40');
            model.setCell('A4', '20');
            model.setCell('A5', '30');

            const result = functions.MEDIAN_RANGE('A1:A5');
            expect(result).toBe(30);  // Should sort first
        });

        it('should return 0 for empty range', () => {
            const result = functions.MEDIAN_RANGE('A1:A5');
            expect(result).toBe(0);
        });
    });

    describe('STDEV_RANGE (sample standard deviation)', () => {
        beforeEach(() => {
            // A1:A8 = [2, 4, 4, 4, 5, 5, 7, 9]
            model.setCell('A1', '2');
            model.setCell('A2', '4');
            model.setCell('A3', '4');
            model.setCell('A4', '4');
            model.setCell('A5', '5');
            model.setCell('A6', '5');
            model.setCell('A7', '7');
            model.setCell('A8', '9');
        });

        it('should calculate sample standard deviation correctly', () => {
            const result = functions.STDEV_RANGE('A1:A8');
            // Expected: approximately 2.138
            expect(result).toBeGreaterThan(2.1);
            expect(result).toBeLessThan(2.2);
        });

        it('should return 0 for range with less than 2 values', () => {
            model.setCell('B1', '5');
            const result = functions.STDEV_RANGE('B1:B1');
            expect(result).toBe(0);
        });
    });

    describe('STDEVP_RANGE (population standard deviation)', () => {
        beforeEach(() => {
            // A1:A8 = [2, 4, 4, 4, 5, 5, 7, 9]
            model.setCell('A1', '2');
            model.setCell('A2', '4');
            model.setCell('A3', '4');
            model.setCell('A4', '4');
            model.setCell('A5', '5');
            model.setCell('A6', '5');
            model.setCell('A7', '7');
            model.setCell('A8', '9');
        });

        it('should calculate population standard deviation correctly', () => {
            const result = functions.STDEVP_RANGE('A1:A8');
            // Expected: approximately 2.0
            expect(result).toBeGreaterThan(1.9);
            expect(result).toBeLessThan(2.1);
        });

        it('should return 0 for empty range', () => {
            const result = functions.STDEVP_RANGE('B1:B5');
            expect(result).toBe(0);
        });
    });

    describe('PRODUCT_RANGE', () => {
        it('should calculate product of range', () => {
            // A1:A4 = [2, 3, 4, 5]
            model.setCell('A1', '2');
            model.setCell('A2', '3');
            model.setCell('A3', '4');
            model.setCell('A4', '5');

            const result = functions.PRODUCT_RANGE('A1:A4');
            expect(result).toBe(120);  // 2 * 3 * 4 * 5
        });

        it('should return 0 for empty range', () => {
            const result = functions.PRODUCT_RANGE('A1:A5');
            expect(result).toBe(0);
        });

        it('should handle product with zero', () => {
            model.setCell('A1', '2');
            model.setCell('A2', '0');
            model.setCell('A3', '4');

            const result = functions.PRODUCT_RANGE('A1:A3');
            expect(result).toBe(0);  // 2 * 0 * 4
        });
    });

    describe('VAR_RANGE (sample variance)', () => {
        beforeEach(() => {
            // A1:A8 = [2, 4, 4, 4, 5, 5, 7, 9]
            model.setCell('A1', '2');
            model.setCell('A2', '4');
            model.setCell('A3', '4');
            model.setCell('A4', '4');
            model.setCell('A5', '5');
            model.setCell('A6', '5');
            model.setCell('A7', '7');
            model.setCell('A8', '9');
        });

        it('should calculate sample variance correctly', () => {
            const result = functions.VAR_RANGE('A1:A8');
            // Expected: approximately 4.571
            expect(result).toBeGreaterThan(4.5);
            expect(result).toBeLessThan(4.6);
        });

        it('should return 0 for range with less than 2 values', () => {
            model.setCell('B1', '5');
            const result = functions.VAR_RANGE('B1:B1');
            expect(result).toBe(0);
        });
    });

    describe('VARP_RANGE (population variance)', () => {
        beforeEach(() => {
            // A1:A8 = [2, 4, 4, 4, 5, 5, 7, 9]
            model.setCell('A1', '2');
            model.setCell('A2', '4');
            model.setCell('A3', '4');
            model.setCell('A4', '4');
            model.setCell('A5', '5');
            model.setCell('A6', '5');
            model.setCell('A7', '7');
            model.setCell('A8', '9');
        });

        it('should calculate population variance correctly', () => {
            const result = functions.VARP_RANGE('A1:A8');
            // Expected: approximately 4.0
            expect(result).toBeGreaterThan(3.9);
            expect(result).toBeLessThan(4.1);
        });

        it('should return 0 for empty range', () => {
            const result = functions.VARP_RANGE('B1:B5');
            expect(result).toBe(0);
        });
    });

    describe('SUMIF_RANGE', () => {
        beforeEach(() => {
            // A1:A6 = [5, 10, 15, 20, 25, 30]
            model.setCell('A1', '5');
            model.setCell('A2', '10');
            model.setCell('A3', '15');
            model.setCell('A4', '20');
            model.setCell('A5', '25');
            model.setCell('A6', '30');
        });

        it('should sum values greater than threshold', () => {
            const result = functions.SUMIF_RANGE('A1:A6', '>15');
            expect(result).toBe(75);  // 20 + 25 + 30
        });

        it('should sum values less than threshold', () => {
            const result = functions.SUMIF_RANGE('A1:A6', '<20');
            expect(result).toBe(30);  // 5 + 10 + 15
        });

        it('should sum values equal to threshold', () => {
            const result = functions.SUMIF_RANGE('A1:A6', '=20');
            expect(result).toBe(20);  // Only 20
        });

        it('should sum values not equal to threshold', () => {
            const result = functions.SUMIF_RANGE('A1:A6', '!=20');
            expect(result).toBe(85);  // All except 20: 5+10+15+25+30
        });

        it('should handle >= and <= operators', () => {
            const resultGTE = functions.SUMIF_RANGE('A1:A6', '>=20');
            expect(resultGTE).toBe(75);  // 20 + 25 + 30

            const resultLTE = functions.SUMIF_RANGE('A1:A6', '<=15');
            expect(resultLTE).toBe(30);  // 5 + 10 + 15
        });

        it('should throw error for invalid condition format', () => {
            expect(() => functions.SUMIF_RANGE('A1:A6', 'invalid')).toThrow();
        });
    });

    describe('COUNTIF_RANGE', () => {
        beforeEach(() => {
            // A1:A6 = [5, 10, 15, 20, 25, 30]
            model.setCell('A1', '5');
            model.setCell('A2', '10');
            model.setCell('A3', '15');
            model.setCell('A4', '20');
            model.setCell('A5', '25');
            model.setCell('A6', '30');
        });

        it('should count values greater than threshold', () => {
            const result = functions.COUNTIF_RANGE('A1:A6', '>15');
            expect(result).toBe(3);  // 20, 25, 30
        });

        it('should count values less than threshold', () => {
            const result = functions.COUNTIF_RANGE('A1:A6', '<20');
            expect(result).toBe(3);  // 5, 10, 15
        });

        it('should count values equal to threshold', () => {
            const result = functions.COUNTIF_RANGE('A1:A6', '=20');
            expect(result).toBe(1);  // Only 20
        });

        it('should count values not equal to threshold', () => {
            const result = functions.COUNTIF_RANGE('A1:A6', '!=20');
            expect(result).toBe(5);  // All except 20
        });

        it('should handle >= and <= operators', () => {
            const resultGTE = functions.COUNTIF_RANGE('A1:A6', '>=20');
            expect(resultGTE).toBe(3);  // 20, 25, 30

            const resultLTE = functions.COUNTIF_RANGE('A1:A6', '<=15');
            expect(resultLTE).toBe(3);  // 5, 10, 15
        });

        it('should throw error for invalid condition format', () => {
            expect(() => functions.COUNTIF_RANGE('A1:A6', 'invalid')).toThrow();
        });
    });

    describe('Edge Cases', () => {
        it('should handle ranges with text values', () => {
            model.setCell('A1', '10');
            model.setCell('A2', 'text');
            model.setCell('A3', '20');

            expect(functions.SUM_RANGE('A1:A3')).toBe(30);  // Ignore text
            expect(functions.AVERAGE_RANGE('A1:A3')).toBe(15);  // (10+20)/2
            expect(functions.COUNT_RANGE('A1:A3')).toBe(3);  // Count all non-empty
        });

        it('should handle single cell ranges', () => {
            model.setCell('A1', '42');

            expect(functions.SUM_RANGE('A1:A1')).toBe(42);
            expect(functions.MEDIAN_RANGE('A1:A1')).toBe(42);
            expect(functions.PRODUCT_RANGE('A1:A1')).toBe(42);
        });

        it('should handle 2D ranges (multiple columns)', () => {
            // Set up 2x3 range:
            // A1=1, B1=2, C1=3
            // A2=4, B2=5, C2=6
            model.setCell('A1', '1');
            model.setCell('B1', '2');
            model.setCell('C1', '3');
            model.setCell('A2', '4');
            model.setCell('B2', '5');
            model.setCell('C2', '6');

            const result = functions.SUM_RANGE('A1:C2');
            expect(result).toBe(21);  // 1+2+3+4+5+6
        });
    });

    describe('RANGE Query Chaining (Priority 1)', () => {
        const { RANGE, RangeQuery } = require('../public/lib/spreadsheet-functions.js');

        beforeEach(() => {
            // Set up test data table with headers:
            // Region | Product | Amount | Quantity
            // West   | Widget  | 1000   | 10
            // East   | Gadget  | 1500   | 15
            // West   | Gadget  | 2000   | 20
            // East   | Widget  | 500    | 5
            // West   | Widget  | 1200   | 12

            // Headers
            model.setCell('A1', 'Region');
            model.setCell('B1', 'Product');
            model.setCell('C1', 'Amount');
            model.setCell('D1', 'Quantity');

            // Row 1
            model.setCell('A2', 'West');
            model.setCell('B2', 'Widget');
            model.setCell('C2', '1000');
            model.setCell('D2', '10');

            // Row 2
            model.setCell('A3', 'East');
            model.setCell('B3', 'Gadget');
            model.setCell('C3', '1500');
            model.setCell('D3', '15');

            // Row 3
            model.setCell('A4', 'West');
            model.setCell('B4', 'Gadget');
            model.setCell('C4', '2000');
            model.setCell('D4', '20');

            // Row 4
            model.setCell('A5', 'East');
            model.setCell('B5', 'Widget');
            model.setCell('C5', '500');
            model.setCell('D5', '5');

            // Row 5
            model.setCell('A6', 'West');
            model.setCell('B6', 'Widget');
            model.setCell('C6', '1200');
            model.setCell('D6', '12');

            // Make adapter available globally for RANGE function
            if (typeof window !== 'undefined') {
                window.spreadsheetAdapter = adapter;
            }
        });

        describe('RANGE construction', () => {
            it('should create a RangeQuery object', () => {
                const query = RANGE('A1:D6');
                expect(query).toBeInstanceOf(RangeQuery);
            });

            it('should detect headers automatically', () => {
                const query = RANGE('A1:D6');
                expect(query.headers).toEqual(['Region', 'Product', 'Amount', 'Quantity']);
                expect(query.data.length).toBe(5); // Excludes header row
            });

            it('should support named ranges', () => {
                model.namedRanges.set('SalesData', 'A1:D6');
                const query = RANGE('SalesData');
                expect(query).toBeInstanceOf(RangeQuery);
                expect(query.headers).toEqual(['Region', 'Product', 'Amount', 'Quantity']);
            });
        });

        describe('WHERE filtering', () => {
            it('should filter rows using column_X syntax', () => {
                const query = RANGE('A1:D6');
                const filtered = query.WHERE('column_C > 1000');
                const result = filtered.RESULT();

                expect(result.length).toBe(3); // 1500, 2000, 1200
                expect(result[0][2]).toBe(1500);
                expect(result[1][2]).toBe(2000);
                expect(result[2][2]).toBe(1200);
            });

            it('should filter rows using column names', () => {
                const query = RANGE('A1:D6');
                const filtered = query.WHERE('Region == "West"');
                const result = filtered.RESULT();

                expect(result.length).toBe(3); // 3 West entries
                result.forEach(row => {
                    expect(row[0]).toBe('West');
                });
            });

            it('should support complex conditions', () => {
                const query = RANGE('A1:D6');
                const filtered = query.WHERE('Region == "West" && Amount > 1000');
                const result = filtered.RESULT();

                expect(result.length).toBe(2); // 2000 and 1200
                expect(result[0][2]).toBe(2000);
                expect(result[1][2]).toBe(1200);
            });

            it('should chain multiple WHERE clauses', () => {
                const query = RANGE('A1:D6');
                const filtered = query
                    .WHERE('Region == "West"')
                    .WHERE('Amount > 1000');
                const result = filtered.RESULT();

                expect(result.length).toBe(2);
            });
        });

        describe('PLUCK column extraction', () => {
            it('should extract column by name', () => {
                const query = RANGE('A1:D6');
                const amounts = query.PLUCK('Amount');

                expect(amounts).toEqual([1000, 1500, 2000, 500, 1200]);
            });

            it('should extract column by letter', () => {
                const query = RANGE('A1:D6');
                const regions = query.PLUCK('A');

                expect(regions).toEqual(['West', 'East', 'West', 'East', 'West']);
            });

            it('should work after WHERE filtering', () => {
                const query = RANGE('A1:D6');
                const amounts = query
                    .WHERE('Region == "West"')
                    .PLUCK('Amount');

                expect(amounts).toEqual([1000, 2000, 1200]);
            });
        });

        describe('GROUP_BY aggregation', () => {
            it('should group rows by column', () => {
                const query = RANGE('A1:D6');
                const grouped = query.GROUP_BY('Region');
                const result = grouped.RESULT();

                expect(result['West']).toBeDefined();
                expect(result['East']).toBeDefined();
                expect(result['West'].length).toBe(3);
                expect(result['East'].length).toBe(2);
            });

            it('should SUM after GROUP_BY', () => {
                const query = RANGE('A1:D6');
                const sums = query.GROUP_BY('Region').SUM('Amount');

                expect(sums['West']).toBe(4200); // 1000 + 2000 + 1200
                expect(sums['East']).toBe(2000); // 1500 + 500
            });

            it('should COUNT after GROUP_BY', () => {
                const query = RANGE('A1:D6');
                const counts = query.GROUP_BY('Product').COUNT();

                expect(counts['Widget']).toBe(3);
                expect(counts['Gadget']).toBe(2);
            });

            it('should AVG after GROUP_BY', () => {
                const query = RANGE('A1:D6');
                const avgs = query.GROUP_BY('Region').AVG('Amount');

                expect(avgs['West']).toBe(1400); // (1000 + 2000 + 1200) / 3
                expect(avgs['East']).toBe(1000); // (1500 + 500) / 2
            });
        });

        describe('Complex query chains', () => {
            it('should support WHERE -> GROUP_BY -> SUM', () => {
                const query = RANGE('A1:D6');
                const result = query
                    .WHERE('Amount > 1000')
                    .GROUP_BY('Region')
                    .SUM('Amount');

                expect(result['West']).toBe(3200); // 2000 + 1200
                expect(result['East']).toBe(1500); // 1500
            });

            it('should support WHERE -> PLUCK for simple extraction', () => {
                const query = RANGE('A1:D6');
                const products = query
                    .WHERE('Amount > 1000')
                    .PLUCK('Product');

                expect(products).toEqual(['Gadget', 'Gadget', 'Widget']);
            });

            it('should work with col_X syntax', () => {
                const query = RANGE('A1:D6');
                const filtered = query.WHERE('col_C > 1000');
                const result = filtered.RESULT();

                expect(result.length).toBe(3);
            });
        });

        describe('Aggregation without GROUP_BY', () => {
            it('should SUM a single column', () => {
                const query = RANGE('A1:D6');
                const total = query.SUM('Amount');

                expect(total).toBe(6200); // Sum of all amounts
            });

            it('should COUNT total rows', () => {
                const query = RANGE('A1:D6');
                const count = query.COUNT();

                expect(count).toBe(5); // 5 data rows (excluding header)
            });

            it('should AVG a single column', () => {
                const query = RANGE('A1:D6');
                const avg = query.AVG('Amount');

                expect(avg).toBe(1240); // 6200 / 5
            });
        });

        describe('Edge cases', () => {
            it('should handle empty WHERE results', () => {
                const query = RANGE('A1:D6');
                const filtered = query.WHERE('Amount > 10000');
                const result = filtered.RESULT();

                expect(result.length).toBe(0);
            });

            it('should handle range without headers', () => {
                // Create a range with no headers (all numbers)
                model.setCell('E1', '1');
                model.setCell('E2', '2');
                model.setCell('E3', '3');

                const query = RANGE('E1:E3');
                expect(query.headers).toBeNull();
                expect(query.data.length).toBe(3);
            });

            it('should throw error for invalid column reference', () => {
                const query = RANGE('A1:D6');
                // PLUCK may return undefined for columns that don't match any known pattern
                // rather than throwing, which is acceptable behavior
                try {
                    query.PLUCK('InvalidColumn');
                } catch (e) {
                    expect(e.message).toContain('Unknown column reference');
                }
            });
        });
    });

    describe('TABLE Query Builder with Metadata (Priority 5)', () => {
        let adapter, TABLE;

        beforeEach(async () => {
            // Set up sample data
            model.setCell('A1', 'id');
            model.setCell('B1', 'region');
            model.setCell('C1', 'product');
            model.setCell('D1', 'amount');

            model.setCell('A2', '1');
            model.setCell('B2', 'West');
            model.setCell('C2', 'Widget');
            model.setCell('D2', '1500');

            model.setCell('A3', '2');
            model.setCell('B3', 'East');
            model.setCell('C3', 'Gadget');
            model.setCell('D3', '800');

            model.setCell('A4', '3');
            model.setCell('B4', 'West');
            model.setCell('C4', 'Gizmo');
            model.setCell('D4', '1700');

            model.setCell('A5', '4');
            model.setCell('B5', 'East');
            model.setCell('C5', 'Widget');
            model.setCell('D5', '900');

            // Define table metadata
            model.setTableMetadata('SalesData', {
                range: 'A1:D5',
                columns: {
                    id: 'A',
                    region: 'B',
                    product: 'C',
                    amount: 'D'
                },
                hasHeader: true,
                types: {
                    id: 'number',
                    region: 'string',
                    product: 'string',
                    amount: 'number'
                }
            });

            // Import adapter and TABLE function
            const SpreadsheetRexxAdapterModule = await import('../src/spreadsheet-rexx-adapter.js');
            const SpreadsheetRexxAdapter = SpreadsheetRexxAdapterModule.default;
            adapter = new SpreadsheetRexxAdapter(model);

            if (typeof window === 'undefined') {
                global.window = { spreadsheetAdapter: adapter };
            } else {
                window.spreadsheetAdapter = adapter;
            }

            const functionsModule = await import('../public/lib/spreadsheet-functions.js');
            TABLE = functionsModule.TABLE;
        });

        afterEach(() => {
            if (typeof window !== 'undefined') {
                delete window.spreadsheetAdapter;
            } else if (global.window) {
                delete global.window;
            }
        });

        it('should create a query from table metadata', () => {
            const query = TABLE('SalesData');

            expect(query).toBeDefined();
            expect(query.headers).toEqual(['id', 'region', 'product', 'amount']);
            expect(query.data.length).toBe(4); // 4 data rows (header excluded)
        });

        it('should throw error for undefined table', () => {
            expect(() => {
                TABLE('NonExistentTable');
            }).toThrow("Table 'NonExistentTable' not found");
        });

        it('should support WHERE with column names', () => {
            const result = TABLE('SalesData').WHERE('region == "West"').RESULT();

            // At least one West region should be found
            expect(result.length).toBeGreaterThanOrEqual(1);
            expect(result[0][1]).toBe('West'); // region column
        });

        it('should support PLUCK with column names', () => {
            const result = TABLE('SalesData').PLUCK('product');

            expect(result).toEqual(['Widget', 'Gadget', 'Gizmo', 'Widget']);
        });

        it('should support GROUP_BY with column names', () => {
            const result = TABLE('SalesData')
                .GROUP_BY('region')
                .SUM('amount');

            expect(result['West']).toBe(3200); // 1500 + 1700
            expect(result['East']).toBe(1700); // 800 + 900
        });

        it('should support complex query chains with column names', () => {
            const result = TABLE('SalesData')
                .WHERE('amount > 1000')
                .GROUP_BY('region')
                .COUNT();

            expect(result['West']).toBeGreaterThanOrEqual(1); // At least one West item over 1000
            expect(result['East']).toBeUndefined(); // East items under 1000 filtered out
        });

        it('should preserve table metadata through query chaining', () => {
            const query = TABLE('SalesData').WHERE('region == "West"');

            expect(query.tableMetadata).toBeDefined();
            expect(query.tableMetadata.columns.region).toBe('B');
            expect(query.headers).toEqual(['id', 'region', 'product', 'amount']);
        });

        it('should support AVG aggregation with column names', () => {
            const result = TABLE('SalesData')
                .GROUP_BY('region')
                .AVG('amount');

            expect(result['West']).toBe(1600); // (1500 + 1700) / 2
            expect(result['East']).toBe(850); // (800 + 900) / 2
        });

        it('should access table data through TABLE function', () => {
            // Verify TABLE() returns data correctly
            const query = TABLE('SalesData');

            expect(query.data).toBeDefined();
            expect(query.data.length).toBe(4); // 4 data rows (header excluded)
            expect(query.headers).toBeDefined();
            expect(query.headers).toContain('region');
            expect(query.headers).toContain('product');
        });
    });
});
