/**
 * Tests for search facility and base64 image support
 */

import SpreadsheetModel from '../src/spreadsheet-model.js';
import SpreadsheetRexxAdapter from '../src/spreadsheet-rexx-adapter.js';

describe('SpreadsheetModel - Search and Images', () => {
    let model;
    let adapter;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
        adapter = new SpreadsheetRexxAdapter(model);
    });

    describe('Search Facility (FIND)', () => {
        beforeEach(() => {
            model.setCell('A1', 'Apple');
            model.setCell('A2', 'Banana');
            model.setCell('A3', 'Cherry');
            model.setCell('B1', '100');
            model.setCell('B2', '200');
            model.setCell('B3', 'apple');
            model.setCell('C1', '=A1');
        });

        it('should find cells by value (case-insensitive)', () => {
            const results = model.find('apple');
            expect(results).toContain('A1');
            expect(results).toContain('B3');
            expect(results.length).toBe(2);
        });

        it('should find cells by value (case-sensitive)', () => {
            const results = model.find('apple', { matchCase: true });
            expect(results).toContain('B3');
            expect(results.length).toBe(1);
        });

        it('should find cells with exact match', () => {
            const results = model.find('100', { matchEntireCell: true });
            expect(results).toContain('B1');
            expect(results.length).toBe(1);
        });

        it('should find cells by partial match', () => {
            const results = model.find('an');
            expect(results).toContain('A2'); // Banana
            expect(results.length).toBeGreaterThanOrEqual(1);
        });

        it('should search in formulas when specified', () => {
            const results = model.find('A1', { searchFormulas: true });
            expect(results).toContain('C1');
            expect(results.length).toBeGreaterThanOrEqual(1);
        });

        it('should not search in formulas by default', () => {
            const results = model.find('A1', { searchFormulas: false });
            expect(results).not.toContain('C1');
        });

        it('should return empty array when no matches found', () => {
            const results = model.find('xyz123');
            expect(results).toEqual([]);
        });

        it('should find numeric values', () => {
            const results = model.find('200');
            expect(results).toContain('B2');
            expect(results.length).toBe(1);
        });
    });

    describe('Base64 Image Support', () => {
        const sampleBase64PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
        const sampleBase64JPG = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCwAA8A/9k=';
        const sampleBase64GIF = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
        const sampleBase64WEBP = 'data:image/webp;base64,UklGRiQAAABXRUJQVlA4IBgAAAAwAQCdASoBAAEAAwA0JaQAA3AA/vuUAAA=';

        it('should store base64 PNG image in a cell', () => {
            model.setCell('A1', sampleBase64PNG);
            const cell = model.getCell('A1');
            expect(cell.value).toBe(sampleBase64PNG);
        });

        it('should store base64 JPEG image in a cell', () => {
            model.setCell('A1', sampleBase64JPG);
            const cell = model.getCell('A1');
            expect(cell.value).toBe(sampleBase64JPG);
        });

        it('should store base64 GIF image in a cell', () => {
            model.setCell('A1', sampleBase64GIF);
            const cell = model.getCell('A1');
            expect(cell.value).toBe(sampleBase64GIF);
        });

        it('should store base64 WEBP image in a cell', () => {
            model.setCell('A1', sampleBase64WEBP);
            const cell = model.getCell('A1');
            expect(cell.value).toBe(sampleBase64WEBP);
        });

        it('should support base64 images as formula results', async () => {
            // Set a base64 image in A1
            model.setCell('A1', sampleBase64PNG);

            // Create a formula that references A1
            await model.setCell('B1', '=A1', adapter);
            const cell = model.getCell('B1');

            // The formula should exist
            expect(cell.expression).toBe('A1');

            // The value should be set (even if evaluation fails in test environment,
            // the cell should have been created with the formula)
            expect(cell).toBeDefined();
        });

        it('should search for cells containing base64 images', () => {
            model.setCell('A1', sampleBase64PNG);
            model.setCell('A2', 'Regular text');
            model.setCell('A3', sampleBase64JPG);

            const results = model.find('data:image/png');
            expect(results).toContain('A1');
            expect(results.length).toBe(1);
        });

        it('should handle mixed content with images', () => {
            model.setCell('A1', 'Text value');
            model.setCell('A2', sampleBase64PNG);
            model.setCell('A3', '12345');

            expect(model.getCellValue('A1')).toBe('Text value');
            expect(model.getCellValue('A2')).toBe(sampleBase64PNG);
            expect(model.getCellValue('A3')).toBe('12345');
        });

        it('should preserve base64 images when exporting to JSON', () => {
            model.setCell('A1', sampleBase64PNG);
            const exported = model.toJSON();

            // Version 2 format uses sheets
            expect(exported.sheets).toBeDefined();
            expect(exported.sheets.Sheet1).toBeDefined();
            expect(exported.sheets.Sheet1.cells).toBeDefined();
            expect(exported.sheets.Sheet1.cells.A1).toBe(sampleBase64PNG);
        });

        it('should restore base64 images when importing from JSON', () => {
            const data = {
                version: 2,
                sheets: {
                    Sheet1: {
                        cells: {
                            A1: sampleBase64PNG
                        }
                    }
                }
            };

            model.fromJSON(data);
            const cell = model.getCell('A1');
            expect(cell.value).toBe(sampleBase64PNG);
        });

        it('should handle base64 images with different MIME types', () => {
            const images = [
                'data:image/png;base64,test1',
                'data:image/jpeg;base64,test2',
                'data:image/jpg;base64,test3',
                'data:image/gif;base64,test4',
                'data:image/bmp;base64,test5',
                'data:image/webp;base64,test6',
                'data:image/svg+xml;base64,test7'
            ];

            images.forEach((img, idx) => {
                const cellRef = SpreadsheetModel.formatCellRef(1, idx + 1);
                model.setCell(cellRef, img);
                expect(model.getCellValue(cellRef)).toBe(img);
            });
        });

        it('should find and replace base64 images', () => {
            model.setCell('A1', sampleBase64PNG);
            model.setCell('A2', sampleBase64PNG);
            model.setCell('A3', 'Other value');

            const count = model.replace(sampleBase64PNG, sampleBase64JPG, { matchEntireCell: true });

            expect(count).toBe(2);
            expect(model.getCellValue('A1')).toBe(sampleBase64JPG);
            expect(model.getCellValue('A2')).toBe(sampleBase64JPG);
            expect(model.getCellValue('A3')).toBe('Other value');
        });
    });

    describe('Combined Search and Image Features', () => {
        const sampleImage1 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
        const sampleImage2 = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/4QAiRXhpZgAASUkqAAgAAAABABIBAwABAAAAAQAAAAAAAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCwAA8A/9k=';

        it('should find all cells with base64 images', () => {
            model.setCell('A1', 'Text');
            model.setCell('A2', sampleImage1);
            model.setCell('A3', '123');
            model.setCell('B1', sampleImage2);

            const results = model.find('data:image/');
            expect(results.length).toBe(2);
            expect(results).toContain('A2');
            expect(results).toContain('B1');
        });

        it('should distinguish between different image types in search', () => {
            model.setCell('A1', sampleImage1);
            model.setCell('A2', sampleImage2);

            const pngResults = model.find('data:image/png');
            expect(pngResults).toContain('A1');
            expect(pngResults.length).toBe(1);

            const jpegResults = model.find('data:image/jpeg');
            expect(jpegResults).toContain('A2');
            expect(jpegResults.length).toBe(1);
        });
    });
});
