/**
 * PyOdide SciPy Functions Library for RexxSheet
 *
 * Provides Python-powered SciPy functions via PyOdide:
 * - PY_LINREGRESS: Linear regression using scipy.stats
 *
 * Dependencies: numpy, scipy
 */

class PyOdideScipyFunctions {
    constructor() {
        this.pyodide = null;
        this.initialized = false;
        this.initializing = false;
        this.initPromise = null;
    }

    /**
     * Initialize PyOdide and load scipy (and numpy)
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

                // Load required Python packages (numpy is a dependency of scipy)
                await this.pyodide.loadPackage(['numpy', 'scipy']);

                this.initialized = true;
                this.initializing = false;
                return this.pyodide;
            } catch (error) {
                this.initializing = false;
                throw new Error(`PyOdide SciPy initialization failed: ${error.message}`);
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
     * Get all PyOdide SciPy functions for RexxJS integration
     */
    getFunctions() {
        const self = this;

        return {
            PY_LINREGRESS: async function(x_values, y_values) {
                return await self.PY_LINREGRESS(x_values, y_values);
            },

            // Helper function to check if PyOdide is initialized
            PY_SCIPY_READY: function() {
                return self.initialized;
            },

            // Helper function to initialize PyOdide manually
            PY_SCIPY_INIT: async function() {
                await self.initialize();
                return 'PyOdide SciPy initialized';
            }
        };
    }
}

// Export for Node.js (Jest) and browser
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PyOdideScipyFunctions;
}
