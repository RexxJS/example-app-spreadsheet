/**
 * PyOdide NumPy Functions Library for RexxSheet
 *
 * Provides Python-powered NumPy functions via PyOdide:
 * - PY_FFT: Fast Fourier Transform using numpy.fft
 *
 * Dependencies: numpy
 */

class PyOdideNumpyFunctions {
    constructor() {
        this.pyodide = null;
        this.initialized = false;
        this.initializing = false;
        this.initPromise = null;
    }

    /**
     * Initialize PyOdide and load numpy
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
                await this.pyodide.loadPackage(['numpy']);

                this.initialized = true;
                this.initializing = false;
                return this.pyodide;
            } catch (error) {
                this.initializing = false;
                throw new Error(`PyOdide NumPy initialization failed: ${error.message}`);
            }
        })();

        return this.initPromise;
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
     * Get all PyOdide NumPy functions for RexxJS integration
     */
    getFunctions() {
        const self = this;

        return {
            PY_FFT: async function(data, sample_rate) {
                return await self.PY_FFT(data, sample_rate);
            },

            // Helper function to check if PyOdide is initialized
            PY_NUMPY_READY: function() {
                return self.initialized;
            },

            // Helper function to initialize PyOdide manually
            PY_NUMPY_INIT: async function() {
                await self.initialize();
                return 'PyOdide NumPy initialized';
            }
        };
    }
}

// Export for Node.js (Jest) and browser
if (typeof module !== 'undefined' && module.exports) {
    module.exports = PyOdideNumpyFunctions;
}
