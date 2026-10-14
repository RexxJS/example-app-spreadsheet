/**
 * Tests for Table Management, Column Sorting, and Autofill Features
 */

const SpreadsheetModel = require('../src/spreadsheet-model');
const SpreadsheetRexxAdapter = require('../src/spreadsheet-rexx-adapter');
const { createSpreadsheetControlFunctions } = require('../src/spreadsheet-control-functions');

describe('Table Management', () => {
    let model;
    let adapter;
    let functions;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
        adapter = new SpreadsheetRexxAdapter(model);
        functions = createSpreadsheetControlFunctions(model, adapter);
    });

    describe('Table Definition', () => {
        it('should define a table with headers', () => {
            // Set up table data
            model.setCell('A1', 'Name');
            model.setCell('B1', 'Age');
            model.setCell('C1', 'City');
            model.setCell('A2', 'Alice');
            model.setCell('B2', '30');
            model.setCell('C2', 'NYC');
            model.setCell('A3', 'Bob');
            model.setCell('B3', '25');
            model.setCell('C3', 'LA');

            const table = model.defineTable('People', 'A1:C3', true);

            expect(table.name).toBe('People');
            expect(table.range).toBe('A1:C3');
            expect(table.hasHeader).toBe(true);
            expect(table.headers).toEqual(['Name', 'Age', 'City']);
            expect(table.originalOrder).toEqual([2, 3]); // Data rows
            expect(table.currentSortColumn).toBeNull();
            expect(table.currentSortDirection).toBeNull();
        });

        it('should define a table without headers', () => {
            model.setCell('A1', 'Alice');
            model.setCell('B1', '30');
            model.setCell('A2', 'Bob');
            model.setCell('B2', '25');

            const table = model.defineTable('Data', 'A1:B2', false);

            expect(table.hasHeader).toBe(false);
            expect(table.originalOrder).toEqual([1, 2]); // All rows are data
        });

        it('should throw error for invalid range format', () => {
            expect(() => model.defineTable('Test', 'A1-C3')).toThrow('Invalid range reference');
        });

        it('should get table metadata', () => {
            model.setCell('A1', 'Header');
            model.defineTable('Test', 'A1:A5');

            const table = model.getTable('Test');
            expect(table.name).toBe('Test');
            expect(table.range).toBe('A1:A5');
        });

        it('should return undefined for non-existent table', () => {
            const table = model.getTable('NonExistent');
            expect(table).toBeUndefined();
        });

        it('should list all tables', () => {
            model.defineTable('Table1', 'A1:B5');
            model.defineTable('Table2', 'D1:E10');
            model.defineTable('Table3', 'G1:H3');

            const tables = model.getTables();
            expect(tables).toEqual(['Table1', 'Table2', 'Table3']);
        });

        it('should delete a table', () => {
            model.defineTable('Test', 'A1:B5');
            expect(model.getTables()).toContain('Test');

            model.deleteTable('Test');
            expect(model.getTables()).not.toContain('Test');
        });

        it('should throw error when deleting non-existent table', () => {
            expect(() => model.deleteTable('NonExistent')).toThrow('Table NonExistent does not exist');
        });
    });

    describe('REXX Table Commands', () => {
        it('should define table via DEFINETABLE command', () => {
            model.setCell('A1', 'Name');
            model.setCell('B1', 'Score');
            model.setCell('A2', 'Alice');
            model.setCell('B2', '95');

            const result = functions.DEFINETABLE('Scores', 'A1:B3', 1);
            const table = JSON.parse(result);

            expect(table.name).toBe('Scores');
            expect(table.headers).toEqual(['Name', 'Score']);
        });

        it('should list tables via LISTTABLES command', () => {
            model.defineTable('Table1', 'A1:B5');
            model.defineTable('Table2', 'C1:D5');

            const result = functions.LISTTABLES();

            expect(result[0]).toBe(2); // Count
            expect(result[1]).toBe('Table1');
            expect(result[2]).toBe('Table2');
        });

        it('should get table via GETTABLE command', () => {
            model.setCell('A1', 'ID');
            model.defineTable('Test', 'A1:A5');

            const result = functions.GETTABLE('Test');
            const table = JSON.parse(result);

            expect(table.name).toBe('Test');
        });

        it('should delete table via DELETETABLE command', () => {
            model.defineTable('Test', 'A1:B5');
            functions.DELETETABLE('Test');

            expect(model.getTables()).not.toContain('Test');
        });
    });
});

describe('Table Column Sorting', () => {
    let model;
    let adapter;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
        adapter = new SpreadsheetRexxAdapter(model);

        // Set up sample table
        model.setCell('A1', 'Name');
        model.setCell('B1', 'Age');
        model.setCell('C1', 'Score');
        model.setCell('A2', 'Charlie');
        model.setCell('B2', '30');
        model.setCell('C2', '85');
        model.setCell('A3', 'Alice');
        model.setCell('B3', '25');
        model.setCell('C3', '95');
        model.setCell('A4', 'Bob');
        model.setCell('B4', '35');
        model.setCell('C4', '75');

        model.defineTable('People', 'A1:C4', true);
    });

    describe('Sort State Cycling', () => {
        it('should cycle through sort states: original → asc → desc → original', () => {
            const table = model.getTable('People');

            // First click: original → ascending
            let state = model.sortTableByColumn('People', 'A', adapter);
            expect(state).toBe('asc');
            expect(table.currentSortColumn).toBe('A');
            expect(table.currentSortDirection).toBe('asc');
            expect(model.getCellValue('A2')).toBe('Alice');
            expect(model.getCellValue('A3')).toBe('Bob');
            expect(model.getCellValue('A4')).toBe('Charlie');

            // Second click: ascending → descending
            state = model.sortTableByColumn('People', 'A', adapter);
            expect(state).toBe('desc');
            expect(table.currentSortDirection).toBe('desc');
            expect(model.getCellValue('A2')).toBe('Charlie');
            expect(model.getCellValue('A3')).toBe('Bob');
            expect(model.getCellValue('A4')).toBe('Alice');

            // Third click: descending → original
            state = model.sortTableByColumn('People', 'A', adapter);
            expect(state).toBe('original');
            expect(table.currentSortColumn).toBeNull();
            expect(table.currentSortDirection).toBeNull();
            expect(model.getCellValue('A2')).toBe('Charlie');
            expect(model.getCellValue('A3')).toBe('Alice');
            expect(model.getCellValue('A4')).toBe('Bob');
        });

        it('should sort numerically when values are numbers', () => {
            model.sortTableByColumn('People', 'B', adapter);

            expect(model.getCellValue('B2')).toBe('25'); // Alice
            expect(model.getCellValue('B3')).toBe('30'); // Charlie
            expect(model.getCellValue('B4')).toBe('35'); // Bob
            expect(model.getCellValue('A2')).toBe('Alice');
            expect(model.getCellValue('A3')).toBe('Charlie');
            expect(model.getCellValue('A4')).toBe('Bob');
        });

        it('should reset to ascending when sorting by different column', () => {
            // Sort by column A (ascending)
            model.sortTableByColumn('People', 'A', adapter);
            expect(model.getCellValue('A2')).toBe('Alice');

            // Sort by column B (should start with ascending)
            const state = model.sortTableByColumn('People', 'B', adapter);
            expect(state).toBe('asc');
            expect(model.getCellValue('B2')).toBe('25'); // Ascending by age
        });

        it('should throw error for column outside table range', () => {
            expect(() => model.sortTableByColumn('People', 'D', adapter))
                .toThrow('Column D is outside table range');
        });

        it('should throw error for non-existent table', () => {
            expect(() => model.sortTableByColumn('NonExistent', 'A', adapter))
                .toThrow('Table NonExistent does not exist');
        });
    });

    describe('Sorting with Formulas', () => {
        it('should preserve formulas when sorting', () => {
            // Add formulas in column D (without evaluating)
            model.setCell('D1', 'Double');
            model.setCell('D2', '=C2*2');
            model.setCell('D3', '=C3*2');
            model.setCell('D4', '=C4*2');

            // Redefine table to include column D
            model.defineTable('People', 'A1:D4', true);

            // Sort by Name (ascending)
            model.sortTableByColumn('People', 'A');

            // Check that formulas were preserved and moved with rows
            expect(model.getCellValue('A2')).toBe('Alice');
            expect(model.getCellExpression('D2')).toBe('C2*2');

            expect(model.getCellValue('A3')).toBe('Bob');
            expect(model.getCellExpression('D3')).toBe('C3*2');

            expect(model.getCellValue('A4')).toBe('Charlie');
            expect(model.getCellExpression('D4')).toBe('C4*2');
        });

        it('should update formula references when rows move', () => {
            // Add a sum formula below the table (without evaluating)
            model.setCell('C5', '=C2+C3+C4');

            // Redefine table to only include data rows
            model.defineTable('People', 'A1:C4', true);

            // Sort by Name
            model.sortTableByColumn('People', 'A');

            // Formula references should be updated
            // After sort: A2=Alice(95), A3=Bob(75), A4=Charlie(85)
            // Original: C2(85), C3(95), C4(75)
            // The formula should still reference the same row numbers
            expect(model.getCellExpression('C5')).toBe('C2+C3+C4');
        });

        it('should handle absolute references correctly', () => {
            // Add formula with absolute reference
            model.setCell('D2', '=$C$2');

            model.defineTable('People', 'A1:D4', true);
            model.sortTableByColumn('People', 'A');

            // Absolute reference should remain unchanged
            expect(model.getCellExpression('D2')).toContain('$C$2');
        });
    });

    describe('REXX Sort Command', () => {
        it('should sort table via SORTTABLEBYCOLUMN command', () => {
            const functions = createSpreadsheetControlFunctions(model, adapter);

            const state = functions.SORTTABLEBYCOLUMN('People', 'A');
            expect(state).toBe('asc');
            expect(model.getCellValue('A2')).toBe('Alice');
            expect(model.getCellValue('A3')).toBe('Bob');
            expect(model.getCellValue('A4')).toBe('Charlie');
        });
    });
});

describe('Autofill Features', () => {
    let model;
    let adapter;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
        adapter = new SpreadsheetRexxAdapter(model);
    });

    describe('Fill Up', () => {
        it('should fill values upward', () => {
            model.setCell('A5', '100');
            model.fillUp('A5', 'A1:A4', adapter);

            expect(model.getCellValue('A1')).toBe('100');
            expect(model.getCellValue('A2')).toBe('100');
            expect(model.getCellValue('A3')).toBe('100');
            expect(model.getCellValue('A4')).toBe('100');
            expect(model.getCellValue('A5')).toBe('100');
        });

        it('should fill formulas upward with adjusted references', () => {
            model.setCell('A5', '10');
            model.setCell('B5', '=A5*2');

            model.fillUp('B5', 'B1:B4');

            expect(model.getCellExpression('B1')).toBe('A1*2');
            expect(model.getCellExpression('B2')).toBe('A2*2');
            expect(model.getCellExpression('B3')).toBe('A3*2');
            expect(model.getCellExpression('B4')).toBe('A4*2');
        });

        it('should fill multiple columns upward', () => {
            model.setCell('A5', 'X');
            model.setCell('B5', 'Y');
            model.fillUp('A5:B5', 'A1:B4', adapter);

            expect(model.getCellValue('A1')).toBe('X');
            expect(model.getCellValue('B1')).toBe('Y');
            expect(model.getCellValue('A4')).toBe('X');
            expect(model.getCellValue('B4')).toBe('Y');
        });

        it('should throw error for width mismatch', () => {
            model.setCell('A5', '100');
            expect(() => model.fillUp('A5', 'A1:B4', adapter))
                .toThrow('Source and target must have the same number of columns');
        });
    });

    describe('Fill Left', () => {
        it('should fill values leftward', () => {
            model.setCell('E1', 'Right');
            model.fillLeft('E1', 'A1:D1', adapter);

            expect(model.getCellValue('A1')).toBe('Right');
            expect(model.getCellValue('B1')).toBe('Right');
            expect(model.getCellValue('C1')).toBe('Right');
            expect(model.getCellValue('D1')).toBe('Right');
            expect(model.getCellValue('E1')).toBe('Right');
        });

        it('should fill formulas leftward with adjusted references', () => {
            model.setCell('E1', '5');
            model.setCell('E2', '=E1+10');

            model.fillLeft('E2', 'A2:D2');

            expect(model.getCellExpression('A2')).toBe('A1+10');
            expect(model.getCellExpression('B2')).toBe('B1+10');
            expect(model.getCellExpression('C2')).toBe('C1+10');
            expect(model.getCellExpression('D2')).toBe('D1+10');
        });

        it('should fill multiple rows leftward', () => {
            model.setCell('E1', 'Top');
            model.setCell('E2', 'Bottom');
            model.fillLeft('E1:E2', 'A1:D2', adapter);

            expect(model.getCellValue('A1')).toBe('Top');
            expect(model.getCellValue('A2')).toBe('Bottom');
            expect(model.getCellValue('D1')).toBe('Top');
            expect(model.getCellValue('D2')).toBe('Bottom');
        });

        it('should throw error for height mismatch', () => {
            model.setCell('E1', 'Test');
            expect(() => model.fillLeft('E1', 'A1:D2', adapter))
                .toThrow('Source and target must have the same number of rows');
        });
    });

    describe('Autofill (Direction Detection)', () => {
        it('should detect downward fill', () => {
            model.setCell('A1', '10');
            model.autofill('A1:A1', 'A2:A5', adapter);

            expect(model.getCellValue('A2')).toBe('10');
            expect(model.getCellValue('A3')).toBe('10');
            expect(model.getCellValue('A4')).toBe('10');
            expect(model.getCellValue('A5')).toBe('10');
        });

        it('should detect upward fill', () => {
            model.setCell('A5', '20');
            model.autofill('A5:A5', 'A1:A4', adapter);

            expect(model.getCellValue('A1')).toBe('20');
            expect(model.getCellValue('A2')).toBe('20');
            expect(model.getCellValue('A3')).toBe('20');
            expect(model.getCellValue('A4')).toBe('20');
        });

        it('should detect rightward fill', () => {
            model.setCell('A1', 'Left');
            model.autofill('A1:A1', 'B1:E1', adapter);

            expect(model.getCellValue('B1')).toBe('Left');
            expect(model.getCellValue('C1')).toBe('Left');
            expect(model.getCellValue('D1')).toBe('Left');
            expect(model.getCellValue('E1')).toBe('Left');
        });

        it('should detect leftward fill', () => {
            model.setCell('E1', 'Right');
            model.autofill('E1:E1', 'A1:D1', adapter);

            expect(model.getCellValue('A1')).toBe('Right');
            expect(model.getCellValue('B1')).toBe('Right');
            expect(model.getCellValue('C1')).toBe('Right');
            expect(model.getCellValue('D1')).toBe('Right');
        });

        it('should throw error if ranges are not adjacent', () => {
            model.setCell('A1', 'Test');
            expect(() => model.autofill('A1:A1', 'C3:C5', adapter))
                .toThrow('Target range must be adjacent to source range');
        });
    });

    describe('REXX Autofill Commands', () => {
        it('should fill up via FILLUP command', () => {
            const functions = createSpreadsheetControlFunctions(model, adapter);
            model.setCell('A5', '99');

            functions.FILLUP('A5', 'A1:A4');

            expect(model.getCellValue('A1')).toBe('99');
            expect(model.getCellValue('A2')).toBe('99');
            expect(model.getCellValue('A3')).toBe('99');
            expect(model.getCellValue('A4')).toBe('99');
        });

        it('should fill left via FILLLEFT command', () => {
            const functions = createSpreadsheetControlFunctions(model, adapter);
            model.setCell('E1', 'Test');

            functions.FILLLEFT('E1', 'A1:D1');

            expect(model.getCellValue('A1')).toBe('Test');
            expect(model.getCellValue('B1')).toBe('Test');
            expect(model.getCellValue('C1')).toBe('Test');
            expect(model.getCellValue('D1')).toBe('Test');
        });

        it('should autofill via AUTOFILL command', () => {
            const functions = createSpreadsheetControlFunctions(model, adapter);
            model.setCell('A1', '50');

            functions.AUTOFILL('A1:A1', 'A2:A5');

            expect(model.getCellValue('A2')).toBe('50');
            expect(model.getCellValue('A3')).toBe('50');
            expect(model.getCellValue('A4')).toBe('50');
            expect(model.getCellValue('A5')).toBe('50');
        });
    });

    describe('Formula Adjustments in Autofill', () => {
        it('should adjust relative references when filling down', () => {
            model.setCell('A1', '10');
            model.setCell('B1', '=A1*2');

            model.fillDown('B1', 'B2:B5');

            expect(model.getCellExpression('B2')).toBe('A2*2');
            expect(model.getCellExpression('B3')).toBe('A3*2');
            expect(model.getCellExpression('B4')).toBe('A4*2');
            expect(model.getCellExpression('B5')).toBe('A5*2');
        });

        it('should preserve absolute references when filling', () => {
            model.setCell('A1', '100');
            model.setCell('B1', '=$A$1*2');

            model.fillDown('B1', 'B2:B5');

            expect(model.getCellExpression('B2')).toBe('$A$1*2');
            expect(model.getCellExpression('B3')).toBe('$A$1*2');
            expect(model.getCellExpression('B4')).toBe('$A$1*2');
            expect(model.getCellExpression('B5')).toBe('$A$1*2');
        });

        it('should handle mixed absolute/relative references', () => {
            model.setCell('A1', '5');
            model.setCell('B1', '=$A1+B$1');

            model.fillDown('B1', 'B2:B3');

            expect(model.getCellExpression('B2')).toBe('$A2+B$1');
            expect(model.getCellExpression('B3')).toBe('$A3+B$1');
        });
    });

    describe('Existing Fill Methods Compatibility', () => {
        it('should still support fillDown', () => {
            model.setCell('A1', '=10+5');
            model.fillDown('A1', 'A2:A5');

            expect(model.getCellExpression('A2')).toBe('10+5');
        });

        it('should still support fillRight', () => {
            model.setCell('A1', '=20*3');
            model.fillRight('A1', 'B1:E1');

            expect(model.getCellExpression('B1')).toBe('20*3');
        });
    });
});

describe('Integration: Tables with Autofill', () => {
    let model;
    let adapter;
    let functions;

    beforeEach(() => {
        model = new SpreadsheetModel(100, 26);
        adapter = new SpreadsheetRexxAdapter(model);
        functions = createSpreadsheetControlFunctions(model, adapter);
    });

    it('should create table, autofill formulas, and sort', () => {
        // Create initial data
        functions.SETCELL('A1', 'Product');
        functions.SETCELL('B1', 'Price');
        functions.SETCELL('C1', 'Tax');
        functions.SETCELL('A2', 'Widget');
        functions.SETCELL('B2', '100');
        functions.SETCELL('C2', '=B2*0.1');

        // Autofill formula down
        functions.FILLDOWN('C2', 'C3:C5');

        // Fill in more product data
        functions.SETCELL('A3', 'Gadget');
        functions.SETCELL('B3', '50');
        functions.SETCELL('A4', 'Doohickey');
        functions.SETCELL('B4', '75');
        functions.SETCELL('A5', 'Thingamajig');
        functions.SETCELL('B5', '125');

        // Define as table
        functions.DEFINETABLE('Products', 'A1:C5', 1);

        // Verify formulas were created (not evaluated yet, just check expressions)
        expect(model.getCellExpression('C2')).toBe('B2*0.1');
        expect(model.getCellExpression('C3')).toBe('B3*0.1');
        expect(model.getCellExpression('C4')).toBe('B4*0.1');
        expect(model.getCellExpression('C5')).toBe('B5*0.1');

        // Sort by price (ascending)
        functions.SORTTABLEBYCOLUMN('Products', 'B');

        // Verify sorted order
        expect(model.getCellValue('A2')).toBe('Gadget');      // Price: 50
        expect(model.getCellValue('A3')).toBe('Doohickey');   // Price: 75
        expect(model.getCellValue('A4')).toBe('Widget');      // Price: 100
        expect(model.getCellValue('A5')).toBe('Thingamajig'); // Price: 125

        // Verify formulas were preserved
        expect(model.getCellExpression('C2')).toBe('B2*0.1');
        expect(model.getCellExpression('C3')).toBe('B3*0.1');
        expect(model.getCellExpression('C4')).toBe('B4*0.1');
        expect(model.getCellExpression('C5')).toBe('B5*0.1');
    });
});
