/**
 * Tests for range function Rexx commands
 * SUM_RANGE, AVERAGE_RANGE, COUNT_RANGE, MIN_RANGE, MAX_RANGE, SUMIF_RANGE, COUNTIF_RANGE
 */

import SpreadsheetModel from '../src/spreadsheet-model.js';
import SpreadsheetRexxAdapter from '../src/spreadsheet-rexx-adapter.js';
import { createSpreadsheetControlFunctions } from '../src/spreadsheet-control-functions.js';

describe('Range Function Rexx Commands', () => {
  let model;
  let adapter;
  let functions;

  beforeEach(() => {
    model = new SpreadsheetModel();
    adapter = new SpreadsheetRexxAdapter(model);
    functions = createSpreadsheetControlFunctions(model, adapter);
  });

  describe('SUM_RANGE', () => {
    test('should sum numeric values in range', async () => {
      await functions.SETCELL('A1', '10');
      await functions.SETCELL('A2', '20');
      await functions.SETCELL('A3', '30');

      const result = functions.SUM_RANGE('A1:A3');
      expect(result).toBe(60);
    });

    test('should ignore non-numeric values', async () => {
      await functions.SETCELL('A1', '10');
      await functions.SETCELL('A2', 'text');
      await functions.SETCELL('A3', '30');

      const result = functions.SUM_RANGE('A1:A3');
      expect(result).toBe(40);
    });

    test('should return 0 for empty range', () => {
      const result = functions.SUM_RANGE('A1:A3');
      expect(result).toBe(0);
    });
  });

  describe('AVERAGE_RANGE', () => {
    test('should calculate average of numeric values', async () => {
      await functions.SETCELL('B1', '10');
      await functions.SETCELL('B2', '20');
      await functions.SETCELL('B3', '30');

      const result = functions.AVERAGE_RANGE('B1:B3');
      expect(result).toBe(20);
    });

    test('should ignore non-numeric values when calculating', async () => {
      await functions.SETCELL('B1', '10');
      await functions.SETCELL('B2', 'text');
      await functions.SETCELL('B3', '30');

      const result = functions.AVERAGE_RANGE('B1:B3');
      expect(result).toBe(20);
    });

    test('should return 0 for empty range', () => {
      const result = functions.AVERAGE_RANGE('B1:B3');
      expect(result).toBe(0);
    });
  });

  describe('COUNT_RANGE', () => {
    test('should count non-empty cells', async () => {
      await functions.SETCELL('C1', '10');
      await functions.SETCELL('C2', 'text');
      await functions.SETCELL('C3', '');
      await functions.SETCELL('C4', '0');

      const result = functions.COUNT_RANGE('C1:C4');
      expect(result).toBe(3);
    });

    test('should return 0 for empty range', () => {
      const result = functions.COUNT_RANGE('C1:C3');
      expect(result).toBe(0);
    });
  });

  describe('MIN_RANGE', () => {
    test('should find minimum value in range', async () => {
      await functions.SETCELL('D1', '50');
      await functions.SETCELL('D2', '10');
      await functions.SETCELL('D3', '30');

      const result = functions.MIN_RANGE('D1:D3');
      expect(result).toBe(10);
    });

    test('should ignore non-numeric values', async () => {
      await functions.SETCELL('D1', '50');
      await functions.SETCELL('D2', 'text');
      await functions.SETCELL('D3', '10');

      const result = functions.MIN_RANGE('D1:D3');
      expect(result).toBe(10);
    });

    test('should return 0 for empty range', () => {
      const result = functions.MIN_RANGE('D1:D3');
      expect(result).toBe(0);
    });
  });

  describe('MAX_RANGE', () => {
    test('should find maximum value in range', async () => {
      await functions.SETCELL('E1', '50');
      await functions.SETCELL('E2', '100');
      await functions.SETCELL('E3', '30');

      const result = functions.MAX_RANGE('E1:E3');
      expect(result).toBe(100);
    });

    test('should ignore non-numeric values', async () => {
      await functions.SETCELL('E1', '50');
      await functions.SETCELL('E2', 'text');
      await functions.SETCELL('E3', '100');

      const result = functions.MAX_RANGE('E1:E3');
      expect(result).toBe(100);
    });

    test('should return 0 for empty range', () => {
      const result = functions.MAX_RANGE('E1:E3');
      expect(result).toBe(0);
    });
  });

  describe('SUMIF_RANGE', () => {
    beforeEach(async () => {
      await functions.SETCELL('F1', '5');
      await functions.SETCELL('F2', '10');
      await functions.SETCELL('F3', '15');
      await functions.SETCELL('F4', '20');
    });

    test('should sum values greater than threshold', () => {
      const result = functions.SUMIF_RANGE('F1:F4', '>10');
      expect(result).toBe(35); // 15 + 20
    });

    test('should sum values less than threshold', () => {
      const result = functions.SUMIF_RANGE('F1:F4', '<15');
      expect(result).toBe(15); // 5 + 10
    });

    test('should sum values equal to threshold', () => {
      const result = functions.SUMIF_RANGE('F1:F4', '=10');
      expect(result).toBe(10);
    });

    test('should sum values not equal to threshold', () => {
      const result = functions.SUMIF_RANGE('F1:F4', '!=10');
      expect(result).toBe(40); // 5 + 15 + 20
    });

    test('should sum values greater than or equal to threshold', () => {
      const result = functions.SUMIF_RANGE('F1:F4', '>=10');
      expect(result).toBe(45); // 10 + 15 + 20
    });

    test('should sum values less than or equal to threshold', () => {
      const result = functions.SUMIF_RANGE('F1:F4', '<=10');
      expect(result).toBe(15); // 5 + 10
    });

    test('should throw error for invalid condition format', () => {
      expect(() => functions.SUMIF_RANGE('F1:F4', 'invalid')).toThrow('Invalid condition format');
    });
  });

  describe('COUNTIF_RANGE', () => {
    beforeEach(async () => {
      await functions.SETCELL('G1', '5');
      await functions.SETCELL('G2', '10');
      await functions.SETCELL('G3', '15');
      await functions.SETCELL('G4', '20');
    });

    test('should count values greater than threshold', () => {
      const result = functions.COUNTIF_RANGE('G1:G4', '>10');
      expect(result).toBe(2); // 15, 20
    });

    test('should count values less than threshold', () => {
      const result = functions.COUNTIF_RANGE('G1:G4', '<15');
      expect(result).toBe(2); // 5, 10
    });

    test('should count values equal to threshold', () => {
      const result = functions.COUNTIF_RANGE('G1:G4', '=10');
      expect(result).toBe(1);
    });

    test('should count values not equal to threshold', () => {
      const result = functions.COUNTIF_RANGE('G1:G4', '!=10');
      expect(result).toBe(3); // 5, 15, 20
    });

    test('should count values greater than or equal to threshold', () => {
      const result = functions.COUNTIF_RANGE('G1:G4', '>=10');
      expect(result).toBe(3); // 10, 15, 20
    });

    test('should count values less than or equal to threshold', () => {
      const result = functions.COUNTIF_RANGE('G1:G4', '<=10');
      expect(result).toBe(2); // 5, 10
    });

    test('should throw error for invalid condition format', () => {
      expect(() => functions.COUNTIF_RANGE('G1:G4', 'invalid')).toThrow('Invalid condition format');
    });
  });

  describe('Error handling', () => {
    test('should throw error when range reference is missing', () => {
      expect(() => functions.SUM_RANGE()).toThrow('requires range reference');
      expect(() => functions.AVERAGE_RANGE()).toThrow('requires range reference');
      expect(() => functions.COUNT_RANGE()).toThrow('requires range reference');
      expect(() => functions.MIN_RANGE()).toThrow('requires range reference');
      expect(() => functions.MAX_RANGE()).toThrow('requires range reference');
    });

    test('should throw error when condition is missing for conditional functions', () => {
      expect(() => functions.SUMIF_RANGE('A1:A10')).toThrow('requires condition');
      expect(() => functions.COUNTIF_RANGE('A1:A10')).toThrow('requires condition');
    });
  });

  describe('LISTCOMMANDS integration', () => {
    test('should include new range commands in LISTCOMMANDS', () => {
      const commands = functions.LISTCOMMANDS();

      // Convert stem array to regular array for easier testing
      const commandList = [];
      for (let i = 1; i <= commands[0]; i++) {
        commandList.push(commands[i]);
      }

      expect(commandList).toContain('SUM_RANGE');
      expect(commandList).toContain('AVERAGE_RANGE');
      expect(commandList).toContain('COUNT_RANGE');
      expect(commandList).toContain('MIN_RANGE');
      expect(commandList).toContain('MAX_RANGE');
      expect(commandList).toContain('SUMIF_RANGE');
      expect(commandList).toContain('COUNTIF_RANGE');
    });
  });
});
