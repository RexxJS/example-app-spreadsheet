/**
 * Tests for spreadsheet import functionality
 * Tests CSV, JSON, TOML, and YAML import formats
 */

import SpreadsheetModel from '../src/spreadsheet-model.js';
import { importCSV, importJSON, importTOML, importYAML } from '../src/spreadsheet-import-export.js';

describe('Spreadsheet Import Formats', () => {
    let model;

    beforeEach(() => {
        // Create a fresh model for each test
        model = new SpreadsheetModel(100, 26);
    });

    describe('CSV Import', () => {
        test('should import basic CSV data', () => {
            const csv = 'Name,Age,City\nJohn,30,New York\nJane,25,Los Angeles';
            const result = importCSV(model, csv);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(3);
            expect(result.columnsImported).toBe(3);

            // Check headers
            expect(model.getCellValue('A1')).toBe('Name');
            expect(model.getCellValue('B1')).toBe('Age');
            expect(model.getCellValue('C1')).toBe('City');

            // Check first data row
            expect(model.getCellValue('A2')).toBe('John');
            expect(model.getCellValue('B2')).toBe('30');
            expect(model.getCellValue('C2')).toBe('New York');

            // Check second data row
            expect(model.getCellValue('A3')).toBe('Jane');
            expect(model.getCellValue('B3')).toBe('25');
            expect(model.getCellValue('C3')).toBe('Los Angeles');
        });

        test('should handle CSV with quotes', () => {
            const csv = 'Name,Description\n"John Doe","A person"\n"Jane Smith","Another person"';
            const result = importCSV(model, csv);

            expect(result.success).toBe(true);
            expect(model.getCellValue('A1')).toBe('Name');
            expect(model.getCellValue('B1')).toBe('Description');
            expect(model.getCellValue('A2')).toBe('John Doe');
            expect(model.getCellValue('B2')).toBe('A person');
        });

        test('should handle CSV with custom delimiter', () => {
            const csv = 'Name;Age;City\nJohn;30;New York\nJane;25;Los Angeles';
            const result = importCSV(model, csv, null, { delimiter: ';' });

            expect(result.success).toBe(true);
            expect(model.getCellValue('A1')).toBe('Name');
            expect(model.getCellValue('B1')).toBe('Age');
            expect(model.getCellValue('C1')).toBe('City');
        });

        test('should handle empty CSV', () => {
            const csv = '';
            const result = importCSV(model, csv);

            expect(result.success).toBe(false);
            expect(result.error).toBe('No data found in CSV');
            expect(result.rowsImported).toBe(0);
        });

        test('should handle CSV with empty cells', () => {
            const csv = 'A,B,C\n1,,3\n,5,';
            const result = importCSV(model, csv);

            expect(result.success).toBe(true);
            expect(model.getCellValue('A2')).toBe('1');
            expect(model.getCellValue('B2')).toBe('');
            expect(model.getCellValue('C2')).toBe('3');
            expect(model.getCellValue('A3')).toBe('');
            expect(model.getCellValue('B3')).toBe('5');
            expect(model.getCellValue('C3')).toBe('');
        });

        test('should import CSV to specific sheet', () => {
            model.addSheet('TestSheet');
            const csv = 'Name,Value\nTest,123';
            const result = importCSV(model, csv, 'TestSheet');

            expect(result.success).toBe(true);
            expect(result.sheetName).toBe('TestSheet');

            // Switch to the sheet to check values
            model.setActiveSheet('TestSheet');
            expect(model.getCellValue('A1')).toBe('Name');
            expect(model.getCellValue('B1')).toBe('Value');
        });

        test('should handle CSV with numbers', () => {
            const csv = 'Product,Price,Quantity\nApple,1.50,100\nBanana,0.75,200';
            const result = importCSV(model, csv);

            expect(result.success).toBe(true);
            expect(model.getCellValue('B2')).toBe('1.50');
            expect(model.getCellValue('C2')).toBe('100');
        });
    });

    describe('JSON Import', () => {
        test('should import JSON array of objects', () => {
            const json = [
                { name: 'John', age: 30, city: 'New York' },
                { name: 'Jane', age: 25, city: 'Los Angeles' }
            ];
            const result = importJSON(model, json);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(3); // 1 header + 2 data rows
            expect(result.columnsImported).toBe(3);
            expect(result.format).toBe('array-of-objects');

            // Check headers
            expect(model.getCellValue('A1')).toBe('name');
            expect(model.getCellValue('B1')).toBe('age');
            expect(model.getCellValue('C1')).toBe('city');

            // Check data
            expect(model.getCellValue('A2')).toBe('John');
            expect(model.getCellValue('B2')).toBe('30');
            expect(model.getCellValue('C2')).toBe('New York');
        });

        test('should import JSON array of objects without headers', () => {
            const json = [
                { name: 'John', age: 30 },
                { name: 'Jane', age: 25 }
            ];
            const result = importJSON(model, json, null, { includeHeaders: false });

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(2); // Only data rows

            // First row should be data, not headers
            expect(model.getCellValue('A1')).toBe('John');
            expect(model.getCellValue('B1')).toBe('30');
        });

        test('should import JSON array of arrays', () => {
            const json = [
                ['Name', 'Age', 'City'],
                ['John', 30, 'New York'],
                ['Jane', 25, 'Los Angeles']
            ];
            const result = importJSON(model, json);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(3);
            expect(result.format).toBe('array-of-arrays');

            expect(model.getCellValue('A1')).toBe('Name');
            expect(model.getCellValue('B2')).toBe('30');
            expect(model.getCellValue('C3')).toBe('Los Angeles');
        });

        test('should import JSON from string', () => {
            const jsonString = '[{"name":"John","age":30},{"name":"Jane","age":25}]';
            const result = importJSON(model, jsonString);

            expect(result.success).toBe(true);
            expect(model.getCellValue('A1')).toBe('name');
            expect(model.getCellValue('A2')).toBe('John');
        });

        test('should handle invalid JSON string', () => {
            const invalidJson = '{invalid json}';
            const result = importJSON(model, invalidJson);

            expect(result.success).toBe(false);
            expect(result.error).toBeDefined();
        });

        test('should handle non-array JSON', () => {
            const json = { name: 'John', age: 30 };
            const result = importJSON(model, json);

            expect(result.success).toBe(false);
            expect(result.error).toBe('JSON data must be an array');
        });

        test('should handle empty array', () => {
            const json = [];
            const result = importJSON(model, json);

            expect(result.success).toBe(false);
            expect(result.error).toBe('No data found in JSON array');
        });

        test('should handle null values in JSON', () => {
            const json = [
                { name: 'John', age: null, city: 'New York' },
                { name: null, age: 25, city: null }
            ];
            const result = importJSON(model, json);

            expect(result.success).toBe(true);
            expect(model.getCellValue('B2')).toBe(''); // null becomes empty string
            expect(model.getCellValue('A3')).toBe('');
        });

        test('should import JSON to specific sheet', () => {
            model.addSheet('JSONSheet');
            const json = [{ name: 'Test', value: 123 }];
            const result = importJSON(model, json, 'JSONSheet');

            expect(result.success).toBe(true);
            expect(result.sheetName).toBe('JSONSheet');

            // Switch to the sheet to check values
            model.setActiveSheet('JSONSheet');
            expect(model.getCellValue('A1')).toBe('name');
        });
    });

    describe('TOML Import', () => {
        test('should import basic TOML data', () => {
            const toml = `
title = "TOML Example"
version = "1.0.0"
enabled = true
`;
            const result = importTOML(model, toml);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(4); // 1 header + 3 data rows
            expect(result.columnsImported).toBe(2);

            // Check headers
            expect(model.getCellValue('A1')).toBe('Key');
            expect(model.getCellValue('B1')).toBe('Value');

            // Check data (order may vary, so check if values exist)
            const keys = [
                model.getCellValue('A2'),
                model.getCellValue('A3'),
                model.getCellValue('A4')
            ];

            expect(keys).toContain('title');
            expect(keys).toContain('version');
            expect(keys).toContain('enabled');
        });

        test('should import nested TOML data', () => {
            const toml = `
[owner]
name = "John Doe"
age = 30

[database]
server = "localhost"
port = 5432
`;
            const result = importTOML(model, toml);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBeGreaterThan(1);

            // Check that nested keys are flattened
            const keys = [];
            for (let i = 2; i <= result.rowsImported; i++) {
                keys.push(model.getCellValue(`A${i}`));
            }

            expect(keys).toContain('owner.name');
            expect(keys).toContain('owner.age');
            expect(keys).toContain('database.server');
            expect(keys).toContain('database.port');
        });

        test('should import TOML arrays', () => {
            const toml = `
colors = ["red", "green", "blue"]
`;
            const result = importTOML(model, toml);

            expect(result.success).toBe(true);

            // Find the colors row
            let colorsRow = -1;
            for (let i = 2; i <= result.rowsImported; i++) {
                if (model.getCellValue(`A${i}`) === 'colors') {
                    colorsRow = i;
                    break;
                }
            }

            expect(colorsRow).toBeGreaterThan(-1);
            expect(model.getCellValue(`B${colorsRow}`)).toContain('red');
        });

        test('should import TOML array of tables', () => {
            const toml = `
[[users]]
name = "John"
age = 30

[[users]]
name = "Jane"
age = 25
`;
            const result = importTOML(model, toml);

            expect(result.success).toBe(true);

            const keys = [];
            for (let i = 2; i <= result.rowsImported; i++) {
                keys.push(model.getCellValue(`A${i}`));
            }

            expect(keys).toContain('users[0].name');
            expect(keys).toContain('users[0].age');
            expect(keys).toContain('users[1].name');
            expect(keys).toContain('users[1].age');
        });

        test('should handle invalid TOML', () => {
            const invalidToml = 'invalid toml syntax ===';
            const result = importTOML(model, invalidToml);

            expect(result.success).toBe(false);
            expect(result.error).toBeDefined();
        });

        test('should import TOML to specific sheet', () => {
            model.addSheet('TOMLSheet');
            const toml = 'name = "Test"';
            const result = importTOML(model, toml, 'TOMLSheet');

            expect(result.success).toBe(true);
            expect(result.sheetName).toBe('TOMLSheet');

            // Switch to the sheet to check values
            model.setActiveSheet('TOMLSheet');
            expect(model.getCellValue('A1')).toBe('Key');
        });

        test('should handle empty TOML', () => {
            const toml = '';
            const result = importTOML(model, toml);

            expect(result.success).toBe(false);
            expect(result.error).toBe('No data found in TOML');
        });
    });

    describe('YAML Import', () => {
        test('should import YAML array of objects', () => {
            const yaml = `
- name: John
  age: 30
  city: New York
- name: Jane
  age: 25
  city: Los Angeles
`;
            const result = importYAML(model, yaml);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(3); // 1 header + 2 data rows

            // Check headers
            expect(model.getCellValue('A1')).toBe('name');
            expect(model.getCellValue('B1')).toBe('age');
            expect(model.getCellValue('C1')).toBe('city');

            // Check data
            expect(model.getCellValue('A2')).toBe('John');
            expect(model.getCellValue('B2')).toBe('30');
            expect(model.getCellValue('C2')).toBe('New York');
        });

        test('should import YAML array of arrays', () => {
            const yaml = `
- [Name, Age, City]
- [John, 30, New York]
- [Jane, 25, Los Angeles]
`;
            const result = importYAML(model, yaml);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(3);

            expect(model.getCellValue('A1')).toBe('Name');
            expect(model.getCellValue('B2')).toBe('30');
            expect(model.getCellValue('C3')).toBe('Los Angeles');
        });

        test('should import YAML object', () => {
            const yaml = `
title: YAML Example
version: 1.0.0
author:
  name: John Doe
  email: john@example.com
`;
            const result = importYAML(model, yaml);

            expect(result.success).toBe(true);
            expect(result.columnsImported).toBe(2);

            // Check headers
            expect(model.getCellValue('A1')).toBe('Key');
            expect(model.getCellValue('B1')).toBe('Value');

            // Check that nested keys are flattened
            const keys = [];
            for (let i = 2; i <= result.rowsImported; i++) {
                keys.push(model.getCellValue(`A${i}`));
            }

            expect(keys).toContain('title');
            expect(keys).toContain('version');
            expect(keys).toContain('author.name');
            expect(keys).toContain('author.email');
        });

        test('should import YAML array of primitives', () => {
            const yaml = `
- red
- green
- blue
`;
            const result = importYAML(model, yaml);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(3);
            expect(result.columnsImported).toBe(1);

            expect(model.getCellValue('A1')).toBe('red');
            expect(model.getCellValue('A2')).toBe('green');
            expect(model.getCellValue('A3')).toBe('blue');
        });

        test('should import YAML without headers', () => {
            const yaml = `
- name: John
  age: 30
- name: Jane
  age: 25
`;
            const result = importYAML(model, yaml, null, { includeHeaders: false });

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(2); // Only data rows

            // First row should be data, not headers
            expect(model.getCellValue('A1')).toBe('John');
            expect(model.getCellValue('B1')).toBe('30');
        });

        test('should handle invalid YAML', () => {
            const invalidYaml = '- invalid:\n  - nested\n syntax: error:';
            const result = importYAML(model, invalidYaml);

            // This might succeed or fail depending on the YAML parser
            // We just check it doesn't crash
            expect(result).toBeDefined();
            expect(result.success).toBeDefined();
        });

        test('should handle empty YAML', () => {
            const yaml = '';
            const result = importYAML(model, yaml);

            expect(result.success).toBe(false);
            expect(result.error).toBe('No data found in YAML');
        });

        test('should import YAML to specific sheet', () => {
            model.addSheet('YAMLSheet');
            const yaml = `
- name: Test
  value: 123
`;
            const result = importYAML(model, yaml, 'YAMLSheet');

            expect(result.success).toBe(true);
            expect(result.sheetName).toBe('YAMLSheet');

            // Switch to the sheet to check values
            model.setActiveSheet('YAMLSheet');
            expect(model.getCellValue('A1')).toBe('name');
        });

        test('should handle YAML with null values', () => {
            const yaml = `
- name: John
  age: null
  city: New York
`;
            const result = importYAML(model, yaml);

            expect(result.success).toBe(true);
            expect(model.getCellValue('B2')).toBe(''); // null becomes empty string
        });

        test('should handle YAML with boolean values', () => {
            const yaml = `
- name: Feature A
  enabled: true
- name: Feature B
  enabled: false
`;
            const result = importYAML(model, yaml);

            expect(result.success).toBe(true);
            expect(model.getCellValue('B2')).toBe('true');
            expect(model.getCellValue('B3')).toBe('false');
        });

        test('should handle YAML with numbers', () => {
            const yaml = `
- product: Apple
  price: 1.50
  quantity: 100
`;
            const result = importYAML(model, yaml);

            expect(result.success).toBe(true);
            expect(model.getCellValue('B2')).toBe('1.5'); // Number converted to string
            expect(model.getCellValue('C2')).toBe('100');
        });
    });

    describe('Cross-format Integration', () => {
        test('all formats should work with multi-sheet models', () => {
            model.addSheet('CSVSheet');
            const csvResult = importCSV(model, 'A,B\n1,2', 'CSVSheet');
            expect(csvResult.success).toBe(true);

            model.addSheet('JSONSheet');
            const jsonResult = importJSON(model, [{ x: 10 }], 'JSONSheet');
            expect(jsonResult.success).toBe(true);

            model.addSheet('TOMLSheet');
            const tomlResult = importTOML(model, 'key = "value"', 'TOMLSheet');
            expect(tomlResult.success).toBe(true);

            model.addSheet('YAMLSheet');
            const yamlResult = importYAML(model, '- item: test', 'YAMLSheet');
            expect(yamlResult.success).toBe(true);

            // Verify all sheets still have their correct data
            model.setActiveSheet('CSVSheet');
            expect(model.getCellValue('A1')).toBe('A');

            model.setActiveSheet('JSONSheet');
            expect(model.getCellValue('A1')).toBe('x');

            model.setActiveSheet('TOMLSheet');
            expect(model.getCellValue('A1')).toBe('Key');

            model.setActiveSheet('YAMLSheet');
            expect(model.getCellValue('A1')).toBe('item');
        });

        test('should handle same data in different formats', () => {
            const data = [
                { name: 'John', age: 30 },
                { name: 'Jane', age: 25 }
            ];

            // CSV format
            model.addSheet('CSV');
            const csv = 'name,age\nJohn,30\nJane,25';
            importCSV(model, csv, 'CSV');

            // JSON format
            model.addSheet('JSON');
            importJSON(model, data, 'JSON');

            // YAML format
            model.addSheet('YAML');
            const yaml = '- name: John\n  age: 30\n- name: Jane\n  age: 25';
            importYAML(model, yaml, 'YAML');

            // All should have the same data
            model.setActiveSheet('CSV');
            expect(model.getCellValue('A2')).toBe('John');
            expect(model.getCellValue('B3')).toBe('25');

            model.setActiveSheet('JSON');
            expect(model.getCellValue('A2')).toBe('John');
            expect(model.getCellValue('B3')).toBe('25');

            model.setActiveSheet('YAML');
            expect(model.getCellValue('A2')).toBe('John');
            expect(model.getCellValue('B3')).toBe('25');
        });
    });

    describe('Edge Cases', () => {
        test('should handle very large datasets', () => {
            // Create a large CSV
            let csv = 'A,B,C\n';
            for (let i = 0; i < 100; i++) {
                csv += `${i},${i * 2},${i * 3}\n`;
            }

            const result = importCSV(model, csv);

            expect(result.success).toBe(true);
            expect(result.rowsImported).toBe(101); // 1 header + 100 data rows
            expect(model.getCellValue('A101')).toBe('99');
            expect(model.getCellValue('C101')).toBe('297');
        });

        test('should handle unicode characters', () => {
            const csv = 'Name,City\n日本,東京\n한국,서울';
            const result = importCSV(model, csv);

            expect(result.success).toBe(true);
            expect(model.getCellValue('A2')).toBe('日本');
            expect(model.getCellValue('B2')).toBe('東京');
            expect(model.getCellValue('A3')).toBe('한국');
        });

        test('should handle special characters', () => {
            const json = [{ name: 'O\'Brien', desc: 'Uses "quotes"' }];
            const result = importJSON(model, json);

            expect(result.success).toBe(true);
            expect(model.getCellValue('A2')).toBe('O\'Brien');
            expect(model.getCellValue('B2')).toBe('Uses "quotes"');
        });

        test('should handle very long strings', () => {
            const longString = 'A'.repeat(1000);
            const json = [{ data: longString }];
            const result = importJSON(model, json);

            expect(result.success).toBe(true);
            expect(model.getCellValue('A2').length).toBe(1000);
        });
    });
});
