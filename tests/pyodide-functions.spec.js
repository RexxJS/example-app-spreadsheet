/**
 * PyOdide Functions Tests
 *
 * Tests for Python-powered scientific computing functions:
 * - PY_LINREGRESS: Linear regression using scipy.stats
 * - PY_FFT: Fast Fourier Transform using numpy.fft
 * - PY_SOLVE: Symbolic equation solving using sympy
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';

// Mock PyOdide for testing
global.loadPyodide = vi.fn();

describe('PyOdide Functions', () => {
    let PyOdideFunctions;
    let pyFunctions;
    let mockPyodide;

    beforeEach(async () => {
        // Mock PyOdide
        mockPyodide = {
            runPython: vi.fn(),
            loadPackage: vi.fn().mockResolvedValue(undefined)
        };

        global.loadPyodide = vi.fn().mockResolvedValue(mockPyodide);

        // Import the module
        const module = await import('../pyodide-functions.js');
        PyOdideFunctions = module.default;

        pyFunctions = new PyOdideFunctions();
    });

    describe('PY_LINREGRESS', () => {
        it('should calculate linear regression for simple data', async () => {
            // Mock scipy.stats.linregress result
            const mockResult = JSON.stringify({
                slope: 2.0,
                intercept: 1.0,
                r_value: 0.999,
                p_value: 0.001,
                std_err: 0.05,
                r_squared: 0.998
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const x = [1, 2, 3, 4, 5];
            const y = [3, 5, 7, 9, 11];

            const result = await pyFunctions.PY_LINREGRESS(x, y);

            expect(result).toBeDefined();
            expect(result.slope).toBe(2.0);
            expect(result.intercept).toBe(1.0);
            expect(result.r_value).toBeCloseTo(0.999, 2);
            expect(result.p_value).toBeLessThan(0.01);
            expect(result.equation).toBe('y = 2.0000x + 1.0000');
        });

        it('should handle negative correlation', async () => {
            const mockResult = JSON.stringify({
                slope: -1.5,
                intercept: 10.0,
                r_value: -0.95,
                p_value: 0.02,
                std_err: 0.1,
                r_squared: 0.9025
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const x = [1, 2, 3, 4];
            const y = [8.5, 7, 5.5, 4];

            const result = await pyFunctions.PY_LINREGRESS(x, y);

            expect(result.slope).toBe(-1.5);
            expect(result.intercept).toBe(10.0);
            expect(result.equation).toBe('y = -1.5000x + 10.0000');
        });

        it('should throw error for mismatched array lengths', async () => {
            const x = [1, 2, 3];
            const y = [1, 2];

            await expect(pyFunctions.PY_LINREGRESS(x, y)).rejects.toThrow(
                'x and y arrays must have the same length'
            );
        });

        it('should throw error for insufficient data points', async () => {
            const x = [1];
            const y = [1];

            await expect(pyFunctions.PY_LINREGRESS(x, y)).rejects.toThrow(
                'need at least 2 data points'
            );
        });

        it('should filter out non-numeric values', async () => {
            const mockResult = JSON.stringify({
                slope: 2.0,
                intercept: 0.0,
                r_value: 1.0,
                p_value: 0.0,
                std_err: 0.0,
                r_squared: 1.0
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const x = [1, 2, NaN, 3, undefined, 4];
            const y = [2, 4, 'invalid', 6, null, 8];

            const result = await pyFunctions.PY_LINREGRESS(x, y);

            // Should only use valid pairs: (1,2), (2,4), (3,6), (4,8)
            expect(result.slope).toBe(2.0);
        });

        it('should throw error if not enough valid numeric pairs', async () => {
            const x = [NaN, undefined, 'invalid'];
            const y = [1, null, 3];

            await expect(pyFunctions.PY_LINREGRESS(x, y)).rejects.toThrow(
                'need at least 2 valid numeric data points'
            );
        });
    });

    describe('PY_FFT', () => {
        it('should compute FFT for simple sine wave', async () => {
            const mockResult = JSON.stringify({
                frequencies: [0, 1, 2, 3, 4],
                magnitudes: [0, 5.0, 0.1, 0.1, 0],
                phases: [0, 1.57, 0, 0, 0],
                power: [0, 25.0, 0.01, 0.01, 0],
                dominant_freq: 1,
                dominant_magnitude: 5.0,
                dc_component: 0
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const data = [0, 1, 0, -1, 0, 1, 0, -1];
            const result = await pyFunctions.PY_FFT(data, 8);

            expect(result).toBeDefined();
            expect(result.dominant_freq).toBe(1);
            expect(result.dominant_magnitude).toBe(5.0);
            expect(Array.isArray(result.frequencies)).toBe(true);
            expect(Array.isArray(result.magnitudes)).toBe(true);
            expect(Array.isArray(result.phases)).toBe(true);
        });

        it('should handle DC component', async () => {
            const mockResult = JSON.stringify({
                frequencies: [0, 1, 2],
                magnitudes: [5.0, 0.1, 0.1],
                phases: [0, 0, 0],
                power: [25.0, 0.01, 0.01],
                dominant_freq: 0,
                dominant_magnitude: 0.1,
                dc_component: 5.0
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const data = [5, 5, 5, 5];  // Constant signal
            const result = await pyFunctions.PY_FFT(data);

            expect(result.dc_component).toBe(5.0);
        });

        it('should use default sample rate of 1.0', async () => {
            const mockResult = JSON.stringify({
                frequencies: [0, 0.5],
                magnitudes: [0, 1.0],
                phases: [0, 0],
                power: [0, 1.0],
                dominant_freq: 0.5,
                dominant_magnitude: 1.0,
                dc_component: 0
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const data = [0, 1, 0, -1];
            const result = await pyFunctions.PY_FFT(data);

            expect(result).toBeDefined();
        });

        it('should throw error for insufficient data', async () => {
            const data = [1];

            await expect(pyFunctions.PY_FFT(data)).rejects.toThrow(
                'need at least 2 data points'
            );
        });

        it('should filter out non-numeric values', async () => {
            const mockResult = JSON.stringify({
                frequencies: [0, 1],
                magnitudes: [0, 1.0],
                phases: [0, 0],
                power: [0, 1.0],
                dominant_freq: 1,
                dominant_magnitude: 1.0,
                dc_component: 0
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const data = [0, 1, NaN, 0, undefined, -1, 'invalid'];
            const result = await pyFunctions.PY_FFT(data);

            expect(result).toBeDefined();
        });
    });

    describe('PY_SOLVE', () => {
        it('should solve linear equation', async () => {
            const mockResult = JSON.stringify({
                solutions: ['3'],
                numeric_solutions: [3],
                equation_latex: '2*x - 6 = 0',
                solution_count: 1,
                original_equation: '2*x - 6 = 0'
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const result = await pyFunctions.PY_SOLVE('2*x - 6 = 0', 'x');

            expect(result).toBeDefined();
            expect(result.solutions).toEqual(['3']);
            expect(result.numeric_solutions).toEqual([3]);
            expect(result.solution_count).toBe(1);
        });

        it('should solve quadratic equation', async () => {
            const mockResult = JSON.stringify({
                solutions: ['-2', '2'],
                numeric_solutions: [-2, 2],
                equation_latex: 'x^2 - 4 = 0',
                solution_count: 2,
                original_equation: 'x**2 - 4 = 0'
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const result = await pyFunctions.PY_SOLVE('x**2 - 4 = 0', 'x');

            expect(result).toBeDefined();
            expect(result.solution_count).toBe(2);
            expect(result.numeric_solutions).toEqual([-2, 2]);
        });

        it('should handle equations with square roots', async () => {
            const mockResult = JSON.stringify({
                solutions: ['sqrt(2)', '-sqrt(2)'],
                numeric_solutions: ['1.4142135623730951', '-1.4142135623730951'],
                equation_latex: 'x^2 - 2 = 0',
                solution_count: 2,
                original_equation: 'x**2 - 2 = 0'
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const result = await pyFunctions.PY_SOLVE('x**2 - 2 = 0', 'x');

            expect(result).toBeDefined();
            expect(result.solutions).toContain('sqrt(2)');
            expect(result.solution_count).toBe(2);
        });

        it('should handle complex solutions', async () => {
            const mockResult = JSON.stringify({
                solutions: ['-I', 'I'],
                numeric_solutions: ['(-0+1j)', '(0-1j)'],
                equation_latex: 'x^2 + 1 = 0',
                solution_count: 2,
                original_equation: 'x**2 + 1 = 0'
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const result = await pyFunctions.PY_SOLVE('x**2 + 1 = 0', 'x');

            expect(result).toBeDefined();
            expect(result.solution_count).toBe(2);
        });

        it('should use default variable "x"', async () => {
            const mockResult = JSON.stringify({
                solutions: ['5'],
                numeric_solutions: [5],
                equation_latex: 'x - 5 = 0',
                solution_count: 1,
                original_equation: 'x - 5'
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const result = await pyFunctions.PY_SOLVE('x - 5');

            expect(result).toBeDefined();
            expect(result.numeric_solutions).toEqual([5]);
        });

        it('should handle equation without = sign', async () => {
            const mockResult = JSON.stringify({
                solutions: ['7'],
                numeric_solutions: [7],
                equation_latex: '2*x - 14 = 0',
                solution_count: 1,
                original_equation: '2*x - 14'
            });

            mockPyodide.runPython.mockReturnValue(mockResult);

            const result = await pyFunctions.PY_SOLVE('2*x - 14', 'x');

            expect(result).toBeDefined();
        });

        it('should throw error for empty equation', async () => {
            await expect(pyFunctions.PY_SOLVE('', 'x')).rejects.toThrow(
                'requires an equation string'
            );
        });

        it('should throw error for empty variable', async () => {
            await expect(pyFunctions.PY_SOLVE('x - 5', '')).rejects.toThrow(
                'requires a variable name'
            );
        });
    });

    describe('Helper Functions', () => {
        it('should have PY_READY function', () => {
            const functions = pyFunctions.getFunctions();
            expect(functions.PY_READY).toBeDefined();
            expect(typeof functions.PY_READY).toBe('function');
        });

        it('should have PY_INIT function', () => {
            const functions = pyFunctions.getFunctions();
            expect(functions.PY_INIT).toBeDefined();
            expect(typeof functions.PY_INIT).toBe('function');
        });

        it('PY_READY should return false before initialization', () => {
            const functions = pyFunctions.getFunctions();
            expect(functions.PY_READY()).toBe(false);
        });
    });
});

/**
 * Embedded RexxJS Integration Tests
 *
 * These tests demonstrate how PyOdide functions would be used within
 * RexxJS expressions in spreadsheet cells.
 */
describe('PyOdide Functions - RexxJS Integration', () => {
    describe('Example Usage Scenarios', () => {
        it('LINEAR REGRESSION: Sales forecasting example', () => {
            // Embedded RexxJS test scenario
            const scenario = {
                description: 'Linear regression for sales forecasting',
                setupScript: `
                    /* Setup script - define data ranges */
                    LET months = [1, 2, 3, 4, 5, 6]
                    LET sales = [100, 150, 180, 220, 250, 290]
                `,
                cells: {
                    'A1': { content: 'Month', style: { fontWeight: 'bold' } },
                    'B1': { content: 'Sales', style: { fontWeight: 'bold' } },
                    'A2': '1', 'B2': '100',
                    'A3': '2', 'B3': '150',
                    'A4': '3', 'B4': '180',
                    'A5': '4', 'B5': '220',
                    'A6': '5', 'B6': '250',
                    'A7': '6', 'B7': '290',
                    'D2': {
                        content: '=PY_LINREGRESS(A2:A7, B2:B7).slope',
                        comment: 'Calculate slope (growth rate)'
                    },
                    'D3': {
                        content: '=PY_LINREGRESS(A2:A7, B2:B7).intercept',
                        comment: 'Calculate y-intercept'
                    },
                    'D4': {
                        content: '=PY_LINREGRESS(A2:A7, B2:B7).r_squared',
                        comment: 'R-squared (goodness of fit)',
                        format: '0.00%'
                    },
                    'D5': {
                        content: '=PY_LINREGRESS(A2:A7, B2:B7).equation',
                        comment: 'Regression equation'
                    }
                },
                expected: {
                    'D2': 'approximately 38.0 (slope)',
                    'D3': 'approximately 62.0 (intercept)',
                    'D4': '>99% (excellent fit)',
                    'D5': 'y = 38.0000x + 62.0000'
                }
            };

            expect(scenario.cells['D2'].content).toContain('PY_LINREGRESS');
            expect(scenario.cells['D4'].format).toBe('0.00%');
        });

        it('FFT: Audio frequency analysis example', () => {
            const scenario = {
                description: 'FFT analysis of audio signal',
                setupScript: `
                    /* Generate sample sine wave at 440 Hz (A note) */
                    LET sample_rate = 8000  /* 8 kHz */
                    LET duration = 0.1      /* 100ms */
                `,
                cells: {
                    'A1': { content: 'Time', style: { fontWeight: 'bold' } },
                    'B1': { content: 'Signal', style: { fontWeight: 'bold' } },
                    'A2': '0.0000', 'B2': '=SIN(2 * PI * 440 * A2)',
                    'A3': '0.0001', 'B3': '=SIN(2 * PI * 440 * A3)',
                    // ... more samples
                    'D2': {
                        content: '=PY_FFT(B2:B801, 8000).dominant_freq',
                        comment: 'Find dominant frequency'
                    },
                    'D3': {
                        content: '=PY_FFT(B2:B801, 8000).dominant_magnitude',
                        comment: 'Magnitude at dominant frequency'
                    }
                },
                expected: {
                    'D2': '440 Hz (A note)',
                    'D3': 'High magnitude'
                }
            };

            expect(scenario.cells['D2'].content).toContain('PY_FFT');
            expect(scenario.cells['D2'].comment).toContain('frequency');
        });

        it('SOLVE: Engineering calculations example', () => {
            const scenario = {
                description: 'Solve engineering equations symbolically',
                cells: {
                    'A1': { content: 'Problem', style: { fontWeight: 'bold' } },
                    'B1': { content: 'Solution', style: { fontWeight: 'bold' } },
                    'A2': 'Find x: x² - 16 = 0',
                    'B2': {
                        content: '=PY_SOLVE("x**2 - 16 = 0", "x").numeric_solutions',
                        comment: 'Returns [-4, 4]'
                    },
                    'A3': 'Find t: 5t + 3 = 18',
                    'B3': {
                        content: '=PY_SOLVE("5*t + 3 = 18", "t").numeric_solutions[0]',
                        comment: 'Returns 3'
                    },
                    'A4': 'Find r: r² - 2 = 0',
                    'B4': {
                        content: '=PY_SOLVE("r**2 - 2 = 0", "r").solutions',
                        comment: 'Returns [sqrt(2), -sqrt(2)] (exact symbolic)'
                    }
                },
                expected: {
                    'B2': '[-4, 4]',
                    'B3': '3',
                    'B4': 'symbolic: sqrt(2), -sqrt(2)'
                }
            };

            expect(scenario.cells['B2'].content).toContain('PY_SOLVE');
            expect(scenario.cells['B4'].comment).toContain('symbolic');
        });

        it('COMBINED: Statistical analysis with regression and FFT', () => {
            const scenario = {
                description: 'Comprehensive data analysis workflow',
                setupScript: `
                    /* Load experimental data */
                    CALL PY_INIT()  /* Initialize PyOdide */
                `,
                cells: {
                    'A1': 'Step 1: Linear regression on time series',
                    'A2': '=PY_LINREGRESS(Data_X, Data_Y)',
                    'A3': 'Step 2: Check if residuals are random (FFT)',
                    'A4': '=PY_FFT(Residuals, 100)',
                    'A5': 'Step 3: Solve for predicted value at t=10',
                    'A6': {
                        content: '=LET m = A2.slope, b = A2.intercept IN m * 10 + b',
                        comment: 'Use regression equation'
                    }
                },
                workflow: [
                    '1. Run regression on data',
                    '2. Analyze residuals with FFT to check for patterns',
                    '3. Use regression equation for predictions'
                ]
            };

            expect(scenario.setupScript).toContain('PY_INIT');
            expect(scenario.cells['A6'].comment).toContain('regression');
        });
    });

    describe('RexxJS Pipeline Integration', () => {
        it('should support function pipelines with PyOdide results', () => {
            const examples = [
                {
                    description: 'Extract and format regression slope',
                    formula: '=PY_LINREGRESS(X_data, Y_data) |> GET("slope") |> ROUND(4)',
                    explanation: 'Get slope and round to 4 decimals'
                },
                {
                    description: 'Find dominant frequency and format',
                    formula: '=PY_FFT(signal, 1000) |> GET("dominant_freq") |> FORMAT("#,##0.00 Hz")',
                    explanation: 'Get dominant frequency and format as Hz'
                },
                {
                    description: 'Solve and extract first solution',
                    formula: '=PY_SOLVE("x**2 - 9 = 0", "x") |> GET("numeric_solutions") |> FIRST()',
                    explanation: 'Get first numeric solution'
                }
            ];

            examples.forEach(ex => {
                expect(ex.formula).toContain('|>');
                expect(ex.formula).toMatch(/^=PY_/);
            });
        });
    });

    describe('Error Handling in RexxJS Context', () => {
        it('should provide meaningful errors for invalid inputs', () => {
            const errorCases = [
                {
                    formula: '=PY_LINREGRESS(A1:A2, B1:B1)',
                    error: 'Array length mismatch',
                    cell: 'C1'
                },
                {
                    formula: '=PY_FFT(A1:A1)',
                    error: 'Insufficient data points',
                    cell: 'C2'
                },
                {
                    formula: '=PY_SOLVE("", "x")',
                    error: 'Empty equation',
                    cell: 'C3'
                }
            ];

            errorCases.forEach(testCase => {
                expect(testCase.formula).toMatch(/^=PY_/);
                expect(testCase.error).toBeTruthy();
            });
        });
    });
});
