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
});
