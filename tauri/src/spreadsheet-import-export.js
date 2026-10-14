/**
 * Spreadsheet Import/Export Module
 * Provides utilities for importing CSV, JSON, TOML, and YAML files into spreadsheet models
 */

import Papa from 'papaparse';
import yaml from 'js-yaml';
import toml from '@iarna/toml';

/**
 * Converts column index to letter notation (0 -> A, 1 -> B, 25 -> Z, 26 -> AA, etc.)
 * @param {number} index - Column index (0-based)
 * @returns {string} Column letter(s)
 */
function columnIndexToLetter(index) {
    let letter = '';
    while (index >= 0) {
        letter = String.fromCharCode((index % 26) + 65) + letter;
        index = Math.floor(index / 26) - 1;
    }
    return letter;
}

/**
 * Imports CSV data into a spreadsheet model
 * @param {SpreadsheetModel} model - The spreadsheet model to import into
 * @param {string} csvData - CSV data as a string
 * @param {string|null} sheetName - Name of the sheet to import into (null for active sheet)
 * @param {Object} options - Import options
 * @param {boolean} options.hasHeaders - Whether the first row contains headers (default: true)
 * @param {string} options.delimiter - CSV delimiter (default: auto-detect)
 * @returns {Object} Result with success status and metadata
 */
export function importCSV(model, csvData, sheetName = null, options = {}) {
    const { hasHeaders = true, delimiter = '' } = options;

    try {
        // Parse CSV using PapaParse
        const parseConfig = {
            skipEmptyLines: true,
        };

        if (delimiter) {
            parseConfig.delimiter = delimiter;
        }

        const result = Papa.parse(csvData, parseConfig);

        if (result.errors.length > 0) {
            console.warn('CSV parsing warnings:', result.errors);
        }

        const data = result.data;

        if (!data || data.length === 0) {
            return {
                success: false,
                error: 'No data found in CSV',
                rowsImported: 0,
                columnsImported: 0
            };
        }

        // Import data into model
        const targetSheet = sheetName || model.activeSheetName;

        // Save current active sheet and switch to target sheet
        const originalSheet = model.activeSheetName;
        if (targetSheet !== originalSheet) {
            model.setActiveSheet(targetSheet);
        }

        let rowsImported = 0;
        let columnsImported = 0;

        for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
            const row = data[rowIndex];
            if (!Array.isArray(row)) continue;

            for (let colIndex = 0; colIndex < row.length; colIndex++) {
                const cellValue = row[colIndex];
                const cellRef = `${columnIndexToLetter(colIndex)}${rowIndex + 1}`;

                // Set cell value
                model.setCell(cellRef, cellValue || '');

                columnsImported = Math.max(columnsImported, colIndex + 1);
            }
            rowsImported++;
        }

        // Restore original active sheet
        if (targetSheet !== originalSheet) {
            model.setActiveSheet(originalSheet);
        }

        return {
            success: true,
            rowsImported,
            columnsImported,
            sheetName: targetSheet,
            hasHeaders
        };

    } catch (error) {
        // Restore original active sheet in case of error
        const originalSheet = model.activeSheetName;
        if (sheetName && sheetName !== originalSheet) {
            try {
                model.setActiveSheet(originalSheet);
            } catch (e) {
                // Ignore restoration errors
            }
        }

        return {
            success: false,
            error: error.message,
            rowsImported: 0,
            columnsImported: 0
        };
    }
}

/**
 * Imports JSON data into a spreadsheet model
 * Supports both array of objects and array of arrays formats
 * @param {SpreadsheetModel} model - The spreadsheet model to import into
 * @param {string|Object} jsonData - JSON data as a string or parsed object
 * @param {string|null} sheetName - Name of the sheet to import into (null for active sheet)
 * @param {Object} options - Import options
 * @param {boolean} options.includeHeaders - Whether to include property names as headers (default: true)
 * @returns {Object} Result with success status and metadata
 */
export function importJSON(model, jsonData, sheetName = null, options = {}) {
    const { includeHeaders = true } = options;

    try {
        // Parse JSON if it's a string
        const data = typeof jsonData === 'string' ? JSON.parse(jsonData) : jsonData;

        if (!Array.isArray(data)) {
            return {
                success: false,
                error: 'JSON data must be an array',
                rowsImported: 0,
                columnsImported: 0
            };
        }

        if (data.length === 0) {
            return {
                success: false,
                error: 'No data found in JSON array',
                rowsImported: 0,
                columnsImported: 0
            };
        }

        const targetSheet = sheetName || model.activeSheetName;

        // Save current active sheet and switch to target sheet
        const originalSheet = model.activeSheetName;
        if (targetSheet !== originalSheet) {
            model.setActiveSheet(targetSheet);
        }

        let rowsImported = 0;
        let columnsImported = 0;

        // Check if data is array of arrays or array of objects
        const isArrayOfArrays = Array.isArray(data[0]);

        if (isArrayOfArrays) {
            // Array of arrays format
            for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
                const row = data[rowIndex];
                if (!Array.isArray(row)) continue;

                for (let colIndex = 0; colIndex < row.length; colIndex++) {
                    const cellValue = row[colIndex];
                    const cellRef = `${columnIndexToLetter(colIndex)}${rowIndex + 1}`;

                    // Convert value to string if it's not null/undefined
                    const stringValue = cellValue != null ? String(cellValue) : '';
                    model.setCell(cellRef, stringValue);

                    columnsImported = Math.max(columnsImported, colIndex + 1);
                }
                rowsImported++;
            }
        } else {
            // Array of objects format
            const keys = Object.keys(data[0]);
            let currentRow = 1;

            // Add headers if requested
            if (includeHeaders) {
                for (let colIndex = 0; colIndex < keys.length; colIndex++) {
                    const cellRef = `${columnIndexToLetter(colIndex)}1`;
                    model.setCell(cellRef, keys[colIndex]);
                }
                currentRow = 2;
                rowsImported++;
            }

            // Add data rows
            for (let objIndex = 0; objIndex < data.length; objIndex++) {
                const obj = data[objIndex];

                for (let colIndex = 0; colIndex < keys.length; colIndex++) {
                    const key = keys[colIndex];
                    const cellValue = obj[key];
                    const cellRef = `${columnIndexToLetter(colIndex)}${currentRow}`;

                    // Convert value to string if it's not null/undefined
                    const stringValue = cellValue != null ? String(cellValue) : '';
                    model.setCell(cellRef, stringValue);
                }

                currentRow++;
                rowsImported++;
            }

            columnsImported = keys.length;
        }

        // Restore original active sheet
        if (targetSheet !== originalSheet) {
            model.setActiveSheet(originalSheet);
        }

        return {
            success: true,
            rowsImported,
            columnsImported,
            sheetName: targetSheet,
            format: isArrayOfArrays ? 'array-of-arrays' : 'array-of-objects'
        };

    } catch (error) {
        // Restore original active sheet in case of error
        const originalSheet = model.activeSheetName;
        if (sheetName && sheetName !== originalSheet) {
            try {
                model.setActiveSheet(originalSheet);
            } catch (e) {
                // Ignore restoration errors
            }
        }

        return {
            success: false,
            error: error.message,
            rowsImported: 0,
            columnsImported: 0
        };
    }
}

/**
 * Imports TOML data into a spreadsheet model
 * Flattens nested TOML structure into rows with dot-notation keys
 * @param {SpreadsheetModel} model - The spreadsheet model to import into
 * @param {string} tomlData - TOML data as a string
 * @param {string|null} sheetName - Name of the sheet to import into (null for active sheet)
 * @param {Object} options - Import options
 * @param {boolean} options.flattenObjects - Whether to flatten nested objects (default: true)
 * @returns {Object} Result with success status and metadata
 */
export function importTOML(model, tomlData, sheetName = null, options = {}) {
    const { flattenObjects = true } = options;

    try {
        // Parse TOML
        const data = toml.parse(tomlData);

        if (!data || typeof data !== 'object') {
            return {
                success: false,
                error: 'Invalid TOML data',
                rowsImported: 0,
                columnsImported: 0
            };
        }

        const targetSheet = sheetName || model.activeSheetName;

        // Save current active sheet and switch to target sheet
        const originalSheet = model.activeSheetName;
        if (targetSheet !== originalSheet) {
            model.setActiveSheet(targetSheet);
        }

        // Flatten the TOML object into key-value pairs
        const flattenObject = (obj, prefix = '') => {
            const result = [];

            for (const [key, value] of Object.entries(obj)) {
                const fullKey = prefix ? `${prefix}.${key}` : key;

                if (Array.isArray(value)) {
                    // Handle arrays
                    if (value.length > 0 && typeof value[0] === 'object' && value[0] !== null) {
                        // Array of objects
                        for (let i = 0; i < value.length; i++) {
                            const itemKey = `${fullKey}[${i}]`;
                            if (flattenObjects) {
                                result.push(...flattenObject(value[i], itemKey));
                            } else {
                                result.push([itemKey, JSON.stringify(value[i])]);
                            }
                        }
                    } else {
                        // Array of primitives
                        result.push([fullKey, JSON.stringify(value)]);
                    }
                } else if (typeof value === 'object' && value !== null) {
                    // Nested object
                    if (flattenObjects) {
                        result.push(...flattenObject(value, fullKey));
                    } else {
                        result.push([fullKey, JSON.stringify(value)]);
                    }
                } else {
                    // Primitive value
                    result.push([fullKey, String(value)]);
                }
            }

            return result;
        };

        const flatData = flattenObject(data);

        if (flatData.length === 0) {
            return {
                success: false,
                error: 'No data found in TOML',
                rowsImported: 0,
                columnsImported: 0
            };
        }

        // Add headers
        model.setCell('A1', 'Key');
        model.setCell('B1', 'Value');

        // Add data rows
        for (let i = 0; i < flatData.length; i++) {
            const [key, value] = flatData[i];
            const rowNum = i + 2; // +2 because row 1 is headers

            model.setCell(`A${rowNum}`, key);
            model.setCell(`B${rowNum}`, value);
        }

        // Restore original active sheet
        if (targetSheet !== originalSheet) {
            model.setActiveSheet(originalSheet);
        }

        return {
            success: true,
            rowsImported: flatData.length + 1, // +1 for header row
            columnsImported: 2,
            sheetName: targetSheet
        };

    } catch (error) {
        // Restore original active sheet in case of error
        const originalSheet = model.activeSheetName;
        if (sheetName && sheetName !== originalSheet) {
            try {
                model.setActiveSheet(originalSheet);
            } catch (e) {
                // Ignore restoration errors
            }
        }

        return {
            success: false,
            error: error.message,
            rowsImported: 0,
            columnsImported: 0
        };
    }
}

/**
 * Imports YAML data into a spreadsheet model
 * Supports both array and object formats
 * @param {SpreadsheetModel} model - The spreadsheet model to import into
 * @param {string} yamlData - YAML data as a string
 * @param {string|null} sheetName - Name of the sheet to import into (null for active sheet)
 * @param {Object} options - Import options
 * @param {boolean} options.includeHeaders - Whether to include property names as headers (default: true)
 * @param {boolean} options.flattenObjects - Whether to flatten nested objects (default: false for arrays, true for single object)
 * @returns {Object} Result with success status and metadata
 */
export function importYAML(model, yamlData, sheetName = null, options = {}) {
    const { includeHeaders = true } = options;

    try {
        // Parse YAML
        const data = yaml.load(yamlData);

        if (!data) {
            return {
                success: false,
                error: 'No data found in YAML',
                rowsImported: 0,
                columnsImported: 0
            };
        }

        const targetSheet = sheetName || model.activeSheetName;

        // Save current active sheet and switch to target sheet
        const originalSheet = model.activeSheetName;
        if (targetSheet !== originalSheet) {
            model.setActiveSheet(targetSheet);
        }

        let rowsImported = 0;
        let columnsImported = 0;

        // Handle different YAML formats
        if (Array.isArray(data)) {
            // Array format - similar to JSON import
            if (data.length === 0) {
                return {
                    success: false,
                    error: 'Empty array in YAML',
                    rowsImported: 0,
                    columnsImported: 0
                };
            }

            const isArrayOfArrays = Array.isArray(data[0]);

            if (isArrayOfArrays) {
                // Array of arrays
                for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
                    const row = data[rowIndex];
                    if (!Array.isArray(row)) continue;

                    for (let colIndex = 0; colIndex < row.length; colIndex++) {
                        const cellValue = row[colIndex];
                        const cellRef = `${columnIndexToLetter(colIndex)}${rowIndex + 1}`;

                        const stringValue = cellValue != null ? String(cellValue) : '';
                        model.setCell(cellRef, stringValue);

                        columnsImported = Math.max(columnsImported, colIndex + 1);
                    }
                    rowsImported++;
                }
            } else if (typeof data[0] === 'object' && data[0] !== null) {
                // Array of objects
                const keys = Object.keys(data[0]);
                let currentRow = 1;

                if (includeHeaders) {
                    for (let colIndex = 0; colIndex < keys.length; colIndex++) {
                        const cellRef = `${columnIndexToLetter(colIndex)}1`;
                        model.setCell(cellRef, keys[colIndex]);
                    }
                    currentRow = 2;
                    rowsImported++;
                }

                for (let objIndex = 0; objIndex < data.length; objIndex++) {
                    const obj = data[objIndex];

                    for (let colIndex = 0; colIndex < keys.length; colIndex++) {
                        const key = keys[colIndex];
                        const cellValue = obj[key];
                        const cellRef = `${columnIndexToLetter(colIndex)}${currentRow}`;

                        const stringValue = cellValue != null ? String(cellValue) : '';
                        model.setCell(cellRef, stringValue);
                    }

                    currentRow++;
                    rowsImported++;
                }

                columnsImported = keys.length;
            } else {
                // Array of primitives
                for (let i = 0; i < data.length; i++) {
                    const cellRef = `A${i + 1}`;
                    const stringValue = data[i] != null ? String(data[i]) : '';
                    model.setCell(cellRef, stringValue);
                    rowsImported++;
                }
                columnsImported = 1;
            }
        } else if (typeof data === 'object') {
            // Single object - flatten to key-value pairs
            const flattenYamlObject = options.flattenObjects !== false;

            const flattenObject = (obj, prefix = '') => {
                const result = [];

                for (const [key, value] of Object.entries(obj)) {
                    const fullKey = prefix ? `${prefix}.${key}` : key;

                    if (Array.isArray(value)) {
                        if (value.length > 0 && typeof value[0] === 'object' && value[0] !== null) {
                            for (let i = 0; i < value.length; i++) {
                                const itemKey = `${fullKey}[${i}]`;
                                if (flattenYamlObject) {
                                    result.push(...flattenObject(value[i], itemKey));
                                } else {
                                    result.push([itemKey, JSON.stringify(value[i])]);
                                }
                            }
                        } else {
                            result.push([fullKey, JSON.stringify(value)]);
                        }
                    } else if (typeof value === 'object' && value !== null) {
                        if (flattenYamlObject) {
                            result.push(...flattenObject(value, fullKey));
                        } else {
                            result.push([fullKey, JSON.stringify(value)]);
                        }
                    } else {
                        result.push([fullKey, String(value)]);
                    }
                }

                return result;
            };

            const flatData = flattenObject(data);

            if (flatData.length === 0) {
                return {
                    success: false,
                    error: 'No data found in YAML object',
                    rowsImported: 0,
                    columnsImported: 0
                };
            }

            // Add headers
            model.setCell('A1', 'Key');
            model.setCell('B1', 'Value');

            // Add data rows
            for (let i = 0; i < flatData.length; i++) {
                const [key, value] = flatData[i];
                const rowNum = i + 2;

                model.setCell(`A${rowNum}`, key);
                model.setCell(`B${rowNum}`, value);
            }

            rowsImported = flatData.length + 1;
            columnsImported = 2;
        } else {
            return {
                success: false,
                error: 'Unsupported YAML format',
                rowsImported: 0,
                columnsImported: 0
            };
        }

        // Restore original active sheet
        if (targetSheet !== originalSheet) {
            model.setActiveSheet(originalSheet);
        }

        return {
            success: true,
            rowsImported,
            columnsImported,
            sheetName: targetSheet
        };

    } catch (error) {
        // Restore original active sheet in case of error
        const originalSheet = model.activeSheetName;
        if (sheetName && sheetName !== originalSheet) {
            try {
                model.setActiveSheet(originalSheet);
            } catch (e) {
                // Ignore restoration errors
            }
        }

        return {
            success: false,
            error: error.message,
            rowsImported: 0,
            columnsImported: 0
        };
    }
}

export default {
    importCSV,
    importJSON,
    importTOML,
    importYAML
};
