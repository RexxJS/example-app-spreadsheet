/**
 * PyOdide Functions Library for RexxSheet
 *
 * Provides Python-powered scientific computing functions via PyOdide:
 * - PY_LINREGRESS: Linear regression using scipy.stats
 * - PY_FFT: Fast Fourier Transform using numpy.fft
 * - PY_SOLVE: Symbolic equation solving using sympy
 */

class PyOdideFunctions {
    constructor() {
        this.pyodide = null;
        this.initialized = false;
        this.initializing = false;
        this.initPromise = null;
    }

    /**
     * Initialize PyOdide and load required packages
     */
    async initialize() {
        if (this.initialized) {
            return this.pyodide;
        }

        if (this.initializing) {
            return this.initPromise;
        }

        this.initializing = true;
        this.initPromise = (async () => {
            try {
                // Load PyOdide from CDN
                if (typeof loadPyodide === 'undefined') {
                    throw new Error('PyOdide not loaded. Include <script src="https://cdn.jsdelivr.net/pyodide/v0.24.1/full/pyodide.js"></script>');
                }

                this.pyodide = await loadPyodide({
                    indexURL: 'https://cdn.jsdelivr.net/pyodide/v0.24.1/full/'
                });

                // Load required Python packages
                await this.pyodide.loadPackage(['numpy', 'scipy', 'sympy']);

                this.initialized = true;
                this.initializing = false;
                return this.pyodide;
            } catch (error) {
                this.initializing = false;
                throw new Error(`PyOdide initialization failed: ${error.message}`);
            }
        })();

        return this.initPromise;
    }

    /**
     * PY_LINREGRESS - Linear regression using scipy.stats.linregress
     *
     * @param {Array<number>} x_values - X values (independent variable)
     * @param {Array<number>} y_values - Y values (dependent variable)
     * @returns {Object} - Regression statistics
     *   {
     *     slope: number,
     *     intercept: number,
     *     r_value: number,      // Correlation coefficient
     *     p_value: number,      // Two-sided p-value for hypothesis test
     *     std_err: number,      // Standard error of the estimate
     *     equation: string      // "y = mx + b" format
     *   }
     */
    async PY_LINREGRESS(x_values, y_values) {
        await this.initialize();

        if (!Array.isArray(x_values) || !Array.isArray(y_values)) {
            throw new Error('PY_LINREGRESS requires two arrays of numbers');
        }

        if (x_values.length !== y_values.length) {
            throw new Error('PY_LINREGRESS: x and y arrays must have the same length');
        }

        if (x_values.length < 2) {
            throw new Error('PY_LINREGRESS: need at least 2 data points');
        }

        // Filter out non-numeric values
        const pairs = x_values.map((x, i) => [x, y_values[i]])
            .filter(([x, y]) => typeof x === 'number' && typeof y === 'number' && !isNaN(x) && !isNaN(y));

        if (pairs.length < 2) {
            throw new Error('PY_LINREGRESS: need at least 2 valid numeric data points');
        }

        const filtered_x = pairs.map(p => p[0]);
        const filtered_y = pairs.map(p => p[1]);

        // Execute Python code
        const code = `
import numpy as np
from scipy import stats

x = np.array(${JSON.stringify(filtered_x)})
y = np.array(${JSON.stringify(filtered_y)})

result = stats.linregress(x, y)

{
    "slope": float(result.slope),
    "intercept": float(result.intercept),
    "r_value": float(result.rvalue),
    "p_value": float(result.pvalue),
    "std_err": float(result.stderr),
    "r_squared": float(result.rvalue ** 2)
}
`;

        const resultJSON = this.pyodide.runPython(code);
        const result = JSON.parse(resultJSON);

        // Add equation string
        const sign = result.intercept >= 0 ? '+' : '';
        result.equation = `y = ${result.slope.toFixed(4)}x ${sign} ${result.intercept.toFixed(4)}`;

        return result;
    }

    /**
     * PY_FFT - Fast Fourier Transform using numpy.fft.fft
     *
     * @param {Array<number>} data - Time-domain signal data
     * @param {number} sample_rate - Sample rate in Hz (optional, default: 1.0)
     * @returns {Object} - FFT results
     *   {
     *     frequencies: Array<number>,  // Frequency bins
     *     magnitudes: Array<number>,   // Magnitude spectrum
     *     phases: Array<number>,       // Phase spectrum (radians)
     *     power: Array<number>,        // Power spectrum
     *     dominant_freq: number,       // Frequency with highest magnitude
     *     dominant_magnitude: number   // Magnitude at dominant frequency
     *   }
     */
    async PY_FFT(data, sample_rate = 1.0) {
        await this.initialize();

        if (!Array.isArray(data)) {
            throw new Error('PY_FFT requires an array of numbers');
        }

        if (data.length < 2) {
            throw new Error('PY_FFT: need at least 2 data points');
        }

        // Filter out non-numeric values
        const filtered_data = data.filter(x => typeof x === 'number' && !isNaN(x));

        if (filtered_data.length < 2) {
            throw new Error('PY_FFT: need at least 2 valid numeric data points');
        }

        const code = `
import numpy as np

data = np.array(${JSON.stringify(filtered_data)})
sample_rate = ${sample_rate}
n = len(data)

# Compute FFT
fft_result = np.fft.fft(data)
frequencies = np.fft.fftfreq(n, d=1.0/sample_rate)

# Only take positive frequencies (first half)
positive_freq_indices = np.arange(n // 2)
frequencies_positive = frequencies[positive_freq_indices]
fft_positive = fft_result[positive_freq_indices]

# Calculate magnitude, phase, and power
magnitudes = np.abs(fft_positive)
phases = np.angle(fft_positive)
power = magnitudes ** 2

# Find dominant frequency
dominant_idx = np.argmax(magnitudes[1:]) + 1  # Skip DC component
dominant_freq = float(frequencies_positive[dominant_idx])
dominant_magnitude = float(magnitudes[dominant_idx])

{
    "frequencies": frequencies_positive.tolist(),
    "magnitudes": magnitudes.tolist(),
    "phases": phases.tolist(),
    "power": power.tolist(),
    "dominant_freq": dominant_freq,
    "dominant_magnitude": dominant_magnitude,
    "dc_component": float(magnitudes[0])
}
`;

        const resultJSON = this.pyodide.runPython(code);
        return JSON.parse(resultJSON);
    }

    /**
     * PY_SOLVE - Symbolic equation solving using sympy
     *
     * @param {string} equation - Equation to solve (e.g., "x**2 - 4 = 0" or "2*x + 3")
     * @param {string} variable - Variable to solve for (default: "x")
     * @returns {Object} - Solution results
     *   {
     *     solutions: Array<string>,    // Solutions as strings (may include complex numbers, sqrt, etc.)
     *     numeric_solutions: Array<number>, // Numeric approximations where possible
     *     equation_latex: string,      // LaTeX representation of equation
     *     solution_count: number
     *   }
     */
    async PY_SOLVE(equation, variable = 'x') {
        await this.initialize();

        if (typeof equation !== 'string' || equation.trim() === '') {
            throw new Error('PY_SOLVE requires an equation string');
        }

        if (typeof variable !== 'string' || variable.trim() === '') {
            throw new Error('PY_SOLVE requires a variable name');
        }

        // Sanitize inputs for Python
        const eq = equation.trim();
        const var_name = variable.trim();

        const code = `
import sympy as sp
from sympy import symbols, solve, latex, N

# Define the variable
${var_name} = symbols('${var_name}')

# Parse the equation
equation_str = """${eq}"""

# Handle different equation formats
if '=' in equation_str:
    # Split by = and create equation
    left, right = equation_str.split('=', 1)
    expr = sp.sympify(left.strip()) - sp.sympify(right.strip())
else:
    # Assume equation is already in form expr = 0
    expr = sp.sympify(equation_str)

# Solve the equation
solutions = solve(expr, ${var_name})

# Convert solutions to strings and numeric values
solution_strings = [str(sol) for sol in solutions]
numeric_solutions = []
for sol in solutions:
    try:
        # Try to convert to float
        numeric_val = complex(N(sol))
        if numeric_val.imag == 0:
            numeric_solutions.append(float(numeric_val.real))
        else:
            numeric_solutions.append(str(numeric_val))
    except:
        numeric_solutions.append(str(sol))

# Get LaTeX representation
equation_latex = latex(expr) + " = 0"

{
    "solutions": solution_strings,
    "numeric_solutions": numeric_solutions,
    "equation_latex": equation_latex,
    "solution_count": len(solutions),
    "original_equation": equation_str
}
`;

        try {
            const resultJSON = this.pyodide.runPython(code);
            return JSON.parse(resultJSON);
        } catch (error) {
            throw new Error(`PY_SOLVE failed: ${error.message}`);
        }
    }

    /**
     * Get all PyOdide functions for RexxJS integration
     */
    getFunctions() {
        const self = this;

        return {
            PY_LINREGRESS: async function(x_values, y_values) {
                return await self.PY_LINREGRESS(x_values, y_values);
            },

            PY_FFT: async function(data, sample_rate) {
                return await self.PY_FFT(data, sample_rate);
            },

            PY_SOLVE: async function(equation, variable) {
                return await self.PY_SOLVE(equation, variable);
            },

            // Helper function to check if PyOdide is initialized
            PY_READY: function() {
                return self.initialized;
            },

            // Helper function to initialize PyOdide manually
            PY_INIT: async function() {
                await self.initialize();
                return 'PyOdide initialized';
            }
        };
    }
}

// Export for Node.js (Jest) and browser
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PyOdideFunctions;
}
