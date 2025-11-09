/*!
 * Spreadsheet Functions Library for RexxJS
 * Provides Excel-like range functions for spreadsheet formulas
 * @rexxjs-meta=SPREADSHEET_FUNCTIONS_META
 */

// Get the spreadsheet adapter from global context (set by the app)
function getAdapter() {
    if (typeof window !== 'undefined' && window.spreadsheetAdapter) {
        return window.spreadsheetAdapter;
    }
    throw new Error('Spreadsheet adapter not available');
}

// SUM a range of cells
function SUM_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    return values.reduce((sum, val) => {
        const num = parseFloat(val);
        return sum + (isNaN(num) ? 0 : num);
    }, 0);
}

// AVERAGE a range of cells
function AVERAGE_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v)));
    if (numbers.length === 0) return 0;
    const sum = numbers.reduce((s, v) => s + parseFloat(v), 0);
    return sum / numbers.length;
}

// COUNT non-empty cells in range
function COUNT_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    return values.filter(v => v !== '' && v !== null && v !== undefined).length;
}

// MIN value in range
function MIN_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    return numbers.length > 0 ? Math.min(...numbers) : 0;
}

// MAX value in range
function MAX_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    return numbers.length > 0 ? Math.max(...numbers) : 0;
}

// Get cell value by reference
function CELL(ref) {
    const adapter = getAdapter();
    const value = adapter.model.getCellValue(ref);
    const numValue = parseFloat(value);
    return isNaN(numValue) ? value : numValue;
}

// MEDIAN value in range
function MEDIAN_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    if (numbers.length === 0) return 0;

    numbers.sort((a, b) => a - b);
    const mid = Math.floor(numbers.length / 2);

    if (numbers.length % 2 === 0) {
        return (numbers[mid - 1] + numbers[mid]) / 2;
    } else {
        return numbers[mid];
    }
}

// STDEV (standard deviation) of range - sample standard deviation
function STDEV_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    if (numbers.length < 2) return 0;

    const mean = numbers.reduce((s, v) => s + v, 0) / numbers.length;
    const squaredDiffs = numbers.map(v => Math.pow(v - mean, 2));
    const variance = squaredDiffs.reduce((s, v) => s + v, 0) / (numbers.length - 1);

    return Math.sqrt(variance);
}

// STDEVP (standard deviation) of range - population standard deviation
function STDEVP_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    if (numbers.length === 0) return 0;

    const mean = numbers.reduce((s, v) => s + v, 0) / numbers.length;
    const squaredDiffs = numbers.map(v => Math.pow(v - mean, 2));
    const variance = squaredDiffs.reduce((s, v) => s + v, 0) / numbers.length;

    return Math.sqrt(variance);
}

// PRODUCT of range
function PRODUCT_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    if (numbers.length === 0) return 0;

    return numbers.reduce((product, v) => product * v, 1);
}

// VAR (variance) of range - sample variance
function VAR_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    if (numbers.length < 2) return 0;

    const mean = numbers.reduce((s, v) => s + v, 0) / numbers.length;
    const squaredDiffs = numbers.map(v => Math.pow(v - mean, 2));

    return squaredDiffs.reduce((s, v) => s + v, 0) / (numbers.length - 1);
}

// VARP (variance) of range - population variance
function VARP_RANGE(rangeRef) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);
    const numbers = values.filter(v => !isNaN(parseFloat(v))).map(v => parseFloat(v));
    if (numbers.length === 0) return 0;

    const mean = numbers.reduce((s, v) => s + v, 0) / numbers.length;
    const squaredDiffs = numbers.map(v => Math.pow(v - mean, 2));

    return squaredDiffs.reduce((s, v) => s + v, 0) / numbers.length;
}

// SUMIF - Sum cells in range that meet a condition
function SUMIF_RANGE(rangeRef, condition) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);

    // Parse condition (e.g., ">5", "=10", "<100")
    const match = condition.match(/^([><=!]+)(.+)$/);
    if (!match) {
        throw new Error('Invalid condition format. Use: ">5", "=10", "<100", etc.');
    }

    const operator = match[1];
    const threshold = parseFloat(match[2]);

    if (isNaN(threshold)) {
        throw new Error('Condition value must be a number');
    }

    return values.reduce((sum, val) => {
        const num = parseFloat(val);
        if (isNaN(num)) return sum;

        let matches = false;
        switch (operator) {
            case '>': matches = num > threshold; break;
            case '>=': matches = num >= threshold; break;
            case '<': matches = num < threshold; break;
            case '<=': matches = num <= threshold; break;
            case '=': case '==': matches = num === threshold; break;
            case '!=': case '<>': matches = num !== threshold; break;
            default: throw new Error('Unknown operator: ' + operator);
        }

        return matches ? sum + num : sum;
    }, 0);
}

// COUNTIF - Count cells in range that meet a condition
function COUNTIF_RANGE(rangeRef, condition) {
    const adapter = getAdapter();
    const values = adapter.getCellRange(rangeRef);

    // Parse condition (e.g., ">5", "=10", "<100")
    const match = condition.match(/^([><=!]+)(.+)$/);
    if (!match) {
        throw new Error('Invalid condition format. Use: ">5", "=10", "<100", etc.');
    }

    const operator = match[1];
    const threshold = parseFloat(match[2]);

    if (isNaN(threshold)) {
        throw new Error('Condition value must be a number');
    }

    return values.reduce((count, val) => {
        const num = parseFloat(val);
        if (isNaN(num)) return count;

        let matches = false;
        switch (operator) {
            case '>': matches = num > threshold; break;
            case '>=': matches = num >= threshold; break;
            case '<': matches = num < threshold; break;
            case '<=': matches = num <= threshold; break;
            case '=': case '==': matches = num === threshold; break;
            case '!=': case '<>': matches = num !== threshold; break;
            default: throw new Error('Unknown operator: ' + operator);
        }

        return matches ? count + 1 : count;
    }, 0);
}

// RexxJS library metadata function
function SPREADSHEET_FUNCTIONS_META() {
    return {
        name: 'spreadsheet-functions',
        version: '2.0.0',
        type: 'functions',
        description: 'Comprehensive spreadsheet range functions for RexxJS - Excel-like statistical and conditional functions',
        functions: [
            'SUM_RANGE',
            'AVERAGE_RANGE',
            'COUNT_RANGE',
            'MIN_RANGE',
            'MAX_RANGE',
            'MEDIAN_RANGE',
            'STDEV_RANGE',
            'STDEVP_RANGE',
            'PRODUCT_RANGE',
            'VAR_RANGE',
            'VARP_RANGE',
            'SUMIF_RANGE',
            'COUNTIF_RANGE',
            'CELL'
        ]
    };
}

// Export for Node.js/CommonJS
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        SUM_RANGE,
        AVERAGE_RANGE,
        COUNT_RANGE,
        MIN_RANGE,
        MAX_RANGE,
        MEDIAN_RANGE,
        STDEV_RANGE,
        STDEVP_RANGE,
        PRODUCT_RANGE,
        VAR_RANGE,
        VARP_RANGE,
        SUMIF_RANGE,
        COUNTIF_RANGE,
        CELL,
        SPREADSHEET_FUNCTIONS_META
    };
}

// Export for browser/window (required for RexxJS REQUIRE in web mode)
if (typeof window !== 'undefined') {
    window.SUM_RANGE = SUM_RANGE;
    window.AVERAGE_RANGE = AVERAGE_RANGE;
    window.COUNT_RANGE = COUNT_RANGE;
    window.MIN_RANGE = MIN_RANGE;
    window.MAX_RANGE = MAX_RANGE;
    window.MEDIAN_RANGE = MEDIAN_RANGE;
    window.STDEV_RANGE = STDEV_RANGE;
    window.STDEVP_RANGE = STDEVP_RANGE;
    window.PRODUCT_RANGE = PRODUCT_RANGE;
    window.VAR_RANGE = VAR_RANGE;
    window.VARP_RANGE = VARP_RANGE;
    window.SUMIF_RANGE = SUMIF_RANGE;
    window.COUNTIF_RANGE = COUNTIF_RANGE;
    window.CELL = CELL;
    window.SPREADSHEET_FUNCTIONS_META = SPREADSHEET_FUNCTIONS_META;
}
