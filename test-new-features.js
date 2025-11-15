#!/usr/bin/env node
/**
 * Quick manual test of new features
 */

import SpreadsheetModel from './src/spreadsheet-model.js';

console.log('Testing Multi-Sheet Support, Row Filtering, and Column Reordering...\n');

let passedTests = 0;
let failedTests = 0;

function test(name, fn) {
    try {
        fn();
        console.log(`✓ ${name}`);
        passedTests++;
    } catch (error) {
        console.log(`✗ ${name}`);
        console.log(`  Error: ${error.message}`);
        failedTests++;
    }
}

function assertEqual(actual, expected, message) {
    if (actual !== expected) {
        throw new Error(`Expected ${expected}, got ${actual}. ${message || ''}`);
    }
}

function assertArrayEqual(actual, expected, message) {
    if (JSON.stringify(actual) !== JSON.stringify(expected)) {
        throw new Error(`Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}. ${message || ''}`);
    }
}

// Test 1: Multi-Sheet Support
test('Should start with default Sheet1', () => {
    const model = new SpreadsheetModel(100, 26);
    assertEqual(model.getActiveSheetName(), 'Sheet1');
    assertArrayEqual(model.getSheetNames(), ['Sheet1']);
});

test('Should add new sheet', () => {
    const model = new SpreadsheetModel(100, 26);
    model.addSheet('Sheet2');
    assertArrayEqual(model.getSheetNames(), ['Sheet1', 'Sheet2']);
});

test('Should reject invalid sheet names (with spaces)', () => {
    const model = new SpreadsheetModel(100, 26);
    try {
        model.addSheet('Sheet 2');
        throw new Error('Should have thrown');
    } catch (e) {
        if (!e.message.includes('valid Rexx variable name')) throw e;
    }
});

test('Should accept valid Rexx sheet names', () => {
    const model = new SpreadsheetModel(100, 26);
    model.addSheet('Sheet2');
    model.addSheet('Data_Sheet');
    model.addSheet('Report_2024');
    assertArrayEqual(model.getSheetNames(), ['Sheet1', 'Sheet2', 'Data_Sheet', 'Report_2024']);
});

test('Should keep data isolated between sheets', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', 'Sheet1 Data');
    model.addSheet('Sheet2');
    model.setActiveSheet('Sheet2');
    model.setCell('A1', 'Sheet2 Data');
    assertEqual(model.getCellValue('A1'), 'Sheet2 Data');
    model.setActiveSheet('Sheet1');
    assertEqual(model.getCellValue('A1'), 'Sheet1 Data');
});

// Test 2: Row Filtering
test('Should filter rows by text criteria', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', 'Apple');
    model.setCell('A2', 'Banana');
    model.setCell('A3', 'Apple Pie');
    model.setCell('A4', 'Cherry');
    model.setCell('A5', 'Apple Juice');

    model.applyRowFilter('A', 'Apple');

    assertEqual(model.isRowVisible(1), true, 'Row 1 should be visible');
    assertEqual(model.isRowVisible(2), false, 'Row 2 should be hidden');
    assertEqual(model.isRowVisible(3), true, 'Row 3 should be visible');
    assertEqual(model.isRowVisible(4), false, 'Row 4 should be hidden');
    assertEqual(model.isRowVisible(5), true, 'Row 5 should be visible');
});

test('Should clear filters', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', 'Apple');
    model.setCell('A2', 'Banana');
    model.applyRowFilter('A', 'Apple');
    model.clearRowFilter();

    assertEqual(model.isRowVisible(1), true);
    assertEqual(model.isRowVisible(2), true);
});

// Test 3: Column Reordering
test('Should move column right', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', '10');
    model.setCell('B1', '20');
    model.setCell('C1', '30');

    model.moveColumnRight('A');

    // After move: B, A, C
    assertEqual(model.getCellValue('A1'), '20');
    assertEqual(model.getCellValue('B1'), '10');
    assertEqual(model.getCellValue('C1'), '30');
});

test('Should move column left', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', '10');
    model.setCell('B1', '20');
    model.setCell('C1', '30');

    model.moveColumnLeft('B');

    // After move: B, A, C
    assertEqual(model.getCellValue('A1'), '20');
    assertEqual(model.getCellValue('B1'), '10');
    assertEqual(model.getCellValue('C1'), '30');
});

test('Should update formulas when columns are swapped', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', '10');
    model.setCell('B1', '20');
    model.setCell('C1', '=A1 + B1');

    model.moveColumnRight('A');

    // After move: B, A, C
    // Formula in C1 should now be =B1 + A1 (columns swapped)
    const expr = model.getCell('C1').expression;
    assertEqual(expr, 'B1 + A1', 'Formula should be updated');
});

test('Should handle complex formulas when swapping columns', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', '10');
    model.setCell('A2', '11');
    model.setCell('B1', '20');
    model.setCell('C1', '30');
    model.setCell('D1', '=A1 + B1 * C1 / A2');

    model.moveColumnRight('A');

    // After swap A and B: =B1 + A1 * C1 / B2
    const expr = model.getCell('D1').expression;
    assertEqual(expr, 'B1 + A1 * C1 / B2', 'Complex formula should be updated');
});

test('Should preserve absolute references when swapping columns', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', '10');
    model.setCell('B1', '20');
    model.setCell('C1', '=$A$1 + B1');

    model.moveColumnRight('A');

    // After swap: B, A, C
    // $A$1 stays pointing to column A (absolute), but B1 becomes A1 (relative)
    const expr = model.getCell('C1').expression;
    assertEqual(expr, '$A$1 + A1', 'Absolute ref stays at column A, relative B1 becomes A1');
});

// Test 4: Cross-Sheet References
test('Should extract cross-sheet references', () => {
    const model = new SpreadsheetModel(100, 26);
    const expr = 'Sheet2.A1 + Sheet3.B2';
    const refs = model.extractCellReferences(expr);

    if (!refs.includes('Sheet2.A1')) {
        throw new Error('Should include Sheet2.A1');
    }
    if (!refs.includes('Sheet3.B2')) {
        throw new Error('Should include Sheet3.B2');
    }
});

test('Should extract both local and cross-sheet references', () => {
    const model = new SpreadsheetModel(100, 26);
    const expr = 'A1 + Sheet2.B1 + C1';
    const refs = model.extractCellReferences(expr);

    if (!refs.includes('A1')) throw new Error('Should include A1');
    if (!refs.includes('Sheet2.B1')) throw new Error('Should include Sheet2.B1');
    if (!refs.includes('C1')) throw new Error('Should include C1');
});

test('Should export and import multi-sheet JSON', () => {
    const model = new SpreadsheetModel(100, 26);
    model.setCell('A1', '10');
    model.addSheet('Sheet2');
    model.setActiveSheet('Sheet2');
    model.setCell('A1', '20');

    const json = model.toJSON();
    assertEqual(json.version, 2, 'Should have version 2');

    const model2 = new SpreadsheetModel(100, 26);
    model2.fromJSON(json);

    assertArrayEqual(model2.getSheetNames(), ['Sheet1', 'Sheet2']);
    assertEqual(model2.getActiveSheetName(), 'Sheet2');

    model2.setActiveSheet('Sheet1');
    assertEqual(model2.getCellValue('A1'), '10');
    model2.setActiveSheet('Sheet2');
    assertEqual(model2.getCellValue('A1'), '20');
});

console.log(`\n${passedTests} passed, ${failedTests} failed`);
process.exit(failedTests > 0 ? 1 : 0);
