/**
 * PyOdide SymPy Functions Library for RexxSheet
 *
 * Provides Python-powered SymPy functions via PyOdide:
 * - PY_SOLVE: Symbolic equation solving using sympy
 *
 * Dependencies: sympy
 */

class PyOdideSympyFunctions {
    constructor() {
        this.pyodide = null;
        this.initialized = false;
        this.initializing = false;
        this.initPromise = null;
    }

    /**
     * Initialize PyOdide and load sympy
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
                await this.pyodide.loadPackage(['sympy']);

                this.initialized = true;
                this.initializing = false;
                return this.pyodide;
            } catch (error) {
                this.initializing = false;
                throw new Error(`PyOdide SymPy initialization failed: ${error.message}`);
            }
        })();

        return this.initPromise;
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
     * Get all PyOdide SymPy functions for RexxJS integration
     */
    getFunctions() {
        const self = this;

        return {
            PY_SOLVE: async function(equation, variable) {
                return await self.PY_SOLVE(equation, variable);
            },

            // Helper function to check if PyOdide is initialized
            PY_SYMPY_READY: function() {
                return self.initialized;
            },

            // Helper function to initialize PyOdide manually
            PY_SYMPY_INIT: async function() {
                await self.initialize();
                return 'PyOdide SymPy initialized';
            }
        };
    }
}

// Export for Node.js (Jest) and browser
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PyOdideSympyFunctions;
}
