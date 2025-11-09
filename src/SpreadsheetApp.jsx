/**
 * SpreadsheetApp - React components for the spreadsheet UI
 *
 * Components:
 * - Cell: Individual spreadsheet cell
 * - Grid: Spreadsheet grid
 * - FormulaBar: Formula/value editor
 * - App: Main application with state management
 */

import React, { useState, useEffect, useRef, useCallback } from 'react';
import { listen } from '@tauri-apps/api/event';
import { createSpreadsheetControlFunctions } from './spreadsheet-control-functions';

/**
 * Cell Component
 */
function Cell({ cellRef, cell, isSelected, isInSelection, onSelect, onEdit, onStartEdit, viewMode, onMouseDown, onMouseEnter, onContextMenu, bufferedKeysRef, isTransitioningRef }) {
    const inputRef = useRef(null);
    const [isEditing, setIsEditing] = useState(false);
    const [editValue, setEditValue] = useState('');

    useEffect(() => {
        if (isEditing && inputRef.current) {
            inputRef.current.focus();

            // Apply buffered keys if any
            if (bufferedKeysRef && bufferedKeysRef.current.length > 0) {
                const buffered = bufferedKeysRef.current.join('');
                setEditValue(buffered);
                bufferedKeysRef.current = [];
                if (isTransitioningRef) {
                    isTransitioningRef.current = false;
                }
                // Set cursor to end
                setTimeout(() => {
                    if (inputRef.current) {
                        inputRef.current.setSelectionRange(buffered.length, buffered.length);
                    }
                }, 0);
            } else {
                inputRef.current.select();
            }
        }
    }, [isEditing, bufferedKeysRef, isTransitioningRef]);

    // Start editing when parent triggers it
    useEffect(() => {
        if (isSelected && onStartEdit) {
            const unsubscribe = onStartEdit((initialChar) => {
                setIsEditing(true);
                if (initialChar) {
                    setEditValue(initialChar);
                } else {
                    setEditValue(cell.expression ? '=' + cell.expression : cell.value);
                }
            });
            return unsubscribe;
        }
    }, [isSelected, cell, onStartEdit]);

    const handleDoubleClick = () => {
        setIsEditing(true);
        setEditValue(cell.expression ? '=' + cell.expression : cell.value);
    };

    const handleKeyDown = (e) => {
        if (e.key === 'Enter' || e.key === 'Tab' || e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
            e.preventDefault();
            onEdit(cellRef, editValue, e.key, e.shiftKey);
            setIsEditing(false);
        } else if (e.key === 'Escape') {
            setIsEditing(false);
        }
    };

    const handleBlur = () => {
        if (isEditing) {
            onEdit(cellRef, editValue);
            setIsEditing(false);
        }
    };

    // Determine what to display based on view mode
    let displayValue = '';
    let showCell = true;

    if (viewMode === 'values') {
        displayValue = cell.expression ? '' : (cell.value || '');
        showCell = !cell.expression || cell.value === '';
    } else if (viewMode === 'expressions') {
        displayValue = cell.expression ? '=' + cell.expression : '';
        showCell = !!cell.expression;
    } else if (viewMode === 'formats') {
        displayValue = cell.format || '';
        showCell = !!cell.format;
    } else {
        // Normal mode - show value with number formatting if applicable
        displayValue = cell.error ? cell.value : (cell.value || '');
        // Apply number formatting if present
        if (!cell.error && cell.format) {
            displayValue = applyNumberFormat(displayValue, cell.format);
        }
    }

    const hasError = !!cell.error;
    const hasFormula = !!cell.expression;
    const hasFormat = !!cell.format;
    const hasComment = !!cell.comment;

    // Build title attribute
    let title = '';
    if (cell.error) {
        title = cell.error;
    } else if (cell.expression) {
        title = '=' + cell.expression;
    }
    if (cell.comment) {
        title += (title ? '\n' : '') + '💬 ' + cell.comment;
    }
    if (cell.format) {
        title += (title ? '\n' : '') + '📊 ' + cell.format;
    }

    // Apply formatting styles
    const formatStyles = cell.format ? parseFormatToCSS(cell.format) : {};

    return (
        <div
            className={`cell ${isSelected ? 'selected' : ''} ${isInSelection ? 'in-selection' : ''} ${hasError ? 'error' : ''} ${hasFormula ? 'formula' : ''} ${hasFormat ? 'formatted' : ''} ${hasComment ? 'commented' : ''} ${viewMode !== 'normal' ? 'view-mode-' + viewMode : ''}`}
            onClick={(e) => onSelect(cellRef, e.shiftKey)}
            onDoubleClick={handleDoubleClick}
            onMouseDown={onMouseDown}
            onMouseEnter={onMouseEnter}
            onContextMenu={onContextMenu ? (e) => onContextMenu(e, cellRef) : undefined}
            title={title}
            style={formatStyles}
        >
            {isEditing ? (
                <input
                    ref={inputRef}
                    type="text"
                    value={editValue}
                    onChange={(e) => setEditValue(e.target.value)}
                    onKeyDown={handleKeyDown}
                    onBlur={handleBlur}
                    className="cell-input"
                />
            ) : (
                showCell && <span className="cell-value">{displayValue}</span>
            )}
        </div>
    );
}

/**
 * Column Header Component
 */
function ColumnHeader({ col }) {
    return (
        <div className="column-header">
            {SpreadsheetModel.colNumberToLetter(col)}
        </div>
    );
}

/**
 * Row Header Component
 */
function RowHeader({ row }) {
    return (
        <div className="row-header">
            {row}
        </div>
    );
}

/**
 * Grid Component
 */
function Grid({ model, selectedCell, selectionRange, onSelectCell, onEditCell, onStartCellEdit, visibleRows, visibleCols, viewMode, onSelectionStart, onSelectionMove, onSelectionEnd, onContextMenu, bufferedKeysRef, isTransitioningRef }) {
    const rows = [];

    // Header row with column letters
    const headerRow = (
        <div key="header" className="grid-row header-row">
            <div className="corner-cell"></div>
            {Array.from({ length: visibleCols }, (_, i) => (
                <ColumnHeader key={i} col={i + 1} />
            ))}
        </div>
    );
    rows.push(headerRow);

    // Data rows
    for (let row = 1; row <= visibleRows; row++) {
        const cells = [];

        // Row header
        cells.push(<RowHeader key={`row-${row}`} row={row} />);

        // Data cells
        for (let col = 1; col <= visibleCols; col++) {
            const cellRef = SpreadsheetModel.formatCellRef(col, row);
            const cell = model.getCell(cellRef);
            const isSelected = selectedCell === cellRef;

            // Check if cell is in selection range
            const isInSelection = selectionRange && isCellInRange(cellRef, selectionRange);

            cells.push(
                <Cell
                    key={`${cellRef}-${cell.value}-${cell.expression}`}
                    cellRef={cellRef}
                    cell={cell}
                    isSelected={isSelected}
                    isInSelection={isInSelection}
                    onSelect={onSelectCell}
                    onEdit={onEditCell}
                    onStartEdit={onStartCellEdit}
                    viewMode={viewMode}
                    onMouseDown={() => onSelectionStart(cellRef)}
                    onMouseEnter={() => onSelectionMove(cellRef)}
                    onContextMenu={onContextMenu}
                    bufferedKeysRef={bufferedKeysRef}
                    isTransitioningRef={isTransitioningRef}
                />
            );
        }

        rows.push(
            <div key={`row-${row}`} className="grid-row">
                {cells}
            </div>
        );
    }

    return <div className="grid" onMouseUp={onSelectionEnd}>{rows}</div>;
}

/**
 * Check if cell is in selection range
 */
function isCellInRange(cellRef, range) {
    if (!range) return false;

    const { startCol, startRow, endCol, endRow } = range;
    const { col, row } = parseCellRef(cellRef);

    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);

    return col >= minCol && col <= maxCol && row >= minRow && row <= maxRow;
}

/**
 * Parse cell reference like "A1" to {col: 1, row: 1}
 */
function parseCellRef(ref) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) return { col: 1, row: 1 };

    const colLetter = match[1];
    const row = parseInt(match[2], 10);

    // Convert column letter to number (A=1, B=2, etc.)
    let col = 0;
    for (let i = 0; i < colLetter.length; i++) {
        col = col * 26 + (colLetter.charCodeAt(i) - 64);
    }

    return { col, row };
}

/**
 * Parse format string to CSS styles
 * Format string: "bold;italic;color:red;background:yellow;align:center;number:0.00"
 */
function parseFormatToCSS(formatString) {
    if (!formatString) return {};

    const styles = {};
    const parts = formatString.split(';').filter(p => p.trim());

    for (const part of parts) {
        const trimmed = part.trim();
        if (trimmed === 'bold') {
            styles.fontWeight = 'bold';
        } else if (trimmed === 'italic') {
            styles.fontStyle = 'italic';
        } else if (trimmed === 'underline') {
            styles.textDecoration = 'underline';
        } else if (trimmed.startsWith('color:')) {
            styles.color = trimmed.substring(6);
        } else if (trimmed.startsWith('background:')) {
            styles.backgroundColor = trimmed.substring(11);
        } else if (trimmed.startsWith('font-size:')) {
            styles.fontSize = trimmed.substring(10);
        } else if (trimmed.startsWith('align:')) {
            const alignment = trimmed.substring(6);
            styles.textAlign = alignment; // left, center, right, justify
        }
        // Note: number formatting is handled in display logic, not CSS
    }

    return styles;
}

/**
 * Apply number formatting to a value
 * Format examples: "number:0", "number:0.00", "currency:USD", "percent:0.0%"
 */
function applyNumberFormat(value, formatString) {
    if (!formatString || value === '' || value === null || value === undefined) {
        return value;
    }

    // Extract number format from format string
    const parts = formatString.split(';').filter(p => p.trim());
    let numberFormat = null;

    for (const part of parts) {
        const trimmed = part.trim();
        if (trimmed.startsWith('number:') || trimmed.startsWith('currency:') || trimmed.startsWith('percent:')) {
            numberFormat = trimmed;
            break;
        }
    }

    if (!numberFormat) {
        return value;
    }

    // Try to parse as number
    const numValue = typeof value === 'number' ? value : parseFloat(value);
    if (isNaN(numValue)) {
        return value; // Not a number, return as-is
    }

    if (numberFormat.startsWith('number:')) {
        // Extract decimal places from format like "number:0.00"
        const formatPattern = numberFormat.substring(7);
        const decimalMatch = formatPattern.match(/0\.(\d+)/);
        const decimalPlaces = decimalMatch ? decimalMatch[1].length : 0;
        return numValue.toFixed(decimalPlaces);
    } else if (numberFormat.startsWith('currency:')) {
        // Extract currency code from format like "currency:USD"
        const currencyCode = numberFormat.substring(9);
        try {
            return new Intl.NumberFormat('en-US', {
                style: 'currency',
                currency: currencyCode || 'USD'
            }).format(numValue);
        } catch (e) {
            // Fallback if currency code invalid
            return '$' + numValue.toFixed(2);
        }
    } else if (numberFormat.startsWith('percent:')) {
        // Extract decimal places from format like "percent:0.0%"
        const formatPattern = numberFormat.substring(8);
        const decimalMatch = formatPattern.match(/0\.(\d+)/);
        const decimalPlaces = decimalMatch ? decimalMatch[1].length : 0;
        return (numValue * 100).toFixed(decimalPlaces) + '%';
    }

    return value;
}

/**
 * Formula Bar Component
 */
function FormulaBar({ selectedCell, model, onEdit }) {
    const [editValue, setEditValue] = useState('');
    const inputRef = useRef(null);

    useEffect(() => {
        if (selectedCell) {
            const cell = model.getCell(selectedCell);
            setEditValue(cell.expression ? '=' + cell.expression : cell.value);
        } else {
            setEditValue('');
        }
    }, [selectedCell, model]);

    const handleKeyDown = (e) => {
        if (e.key === 'Enter') {
            if (selectedCell) {
                onEdit(selectedCell, editValue);
            }
        }
    };

    const handleChange = (e) => {
        setEditValue(e.target.value);
    };

    return (
        <div className="formula-bar">
            <label className="cell-ref-label">{selectedCell || 'No cell selected'}</label>
            <input
                ref={inputRef}
                type="text"
                className="formula-input"
                value={editValue}
                onChange={handleChange}
                onKeyDown={handleKeyDown}
                placeholder="Enter value or =formula"
                disabled={!selectedCell}
            />
        </div>
    );
}

/**
 * Info Panel Component - Shows selected cell details
 */
function InfoPanel({ selectedCell, selectionRange, model, viewMode }) {
    const cell = selectedCell && model ? model.getCell(selectedCell) : null;

    if (!selectedCell || !cell) {
        return (
            <div className="info-panel">
                <h3>Cell Details</h3>
                <p className="no-selection">No cell selected</p>
                <div className="help-section">
                    <p><strong>Navigation:</strong></p>
                    <ul>
                        <li><kbd>Tab</kbd> - Move right</li>
                        <li><kbd>Enter</kbd> / <kbd>↓</kbd> - Move down</li>
                        <li><kbd>↑</kbd> <kbd>←</kbd> <kbd>→</kbd> - Navigate</li>
                        <li><kbd>Ctrl+C</kbd> - Copy selection</li>
                    </ul>
                    <p><strong>Hotkeys (hold to view):</strong></p>
                    <ul>
                        <li><kbd>V</kbd> - Peek at values only</li>
                        <li><kbd>E</kbd> - Peek at expressions only</li>
                        <li><kbd>F</kbd> - Peek at formats only</li>
                    </ul>
                    <p><strong>Named Variables:</strong></p>
                    <p>Use <strong>⚙️ Setup</strong> to define:</p>
                    <code style={{display: 'block', marginTop: '5px'}}>LET TAX_RATE = 0.07</code>
                </div>
            </div>
        );
    }

    const cellType = cell.expression ? 'Formula' : (cell.value ? 'Value' : 'Empty');
    const valueType = typeof cell.value === 'number' ? 'Number' :
                      cell.value ? 'String' : 'Empty';

    // Get dependents
    const dependents = [];
    if (model.dependents.has(selectedCell)) {
        dependents.push(...model.dependents.get(selectedCell));
    }

    // Show selection info if present
    let selectionInfo = null;
    if (selectionRange) {
        const { startCol, startRow, endCol, endRow } = selectionRange;
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const cellCount = (maxCol - minCol + 1) * (maxRow - minRow + 1);

        selectionInfo = (
            <div className="cell-detail-section">
                <p><strong>Selection:</strong> {cellCount} cells</p>
                <p className="selection-range">
                    {SpreadsheetModel.colNumberToLetter(minCol)}{minRow}:
                    {SpreadsheetModel.colNumberToLetter(maxCol)}{maxRow}
                </p>
            </div>
        );
    }

    return (
        <div className="info-panel">
            <h3>Cell: {selectedCell}</h3>

            {selectionInfo}

            <div className="cell-detail-section">
                <p><strong>Type:</strong> {cellType}</p>

                {cell.expression && (
                    <p><strong>Formula:</strong><br/>
                    <code className="formula-display">={cell.expression}</code></p>
                )}

                <p><strong>Value:</strong> {cell.value || <em>(empty)</em>}</p>

                <p><strong>Value Type:</strong> {valueType}</p>

                {cell.error && (
                    <p className="error-display"><strong>Error:</strong> {cell.error}</p>
                )}
            </div>

            {cell.dependencies && cell.dependencies.length > 0 && (
                <div className="cell-detail-section">
                    <p><strong>Dependencies:</strong></p>
                    <p className="dependency-list">{cell.dependencies.join(', ')}</p>
                </div>
            )}

            {dependents.length > 0 && (
                <div className="cell-detail-section">
                    <p><strong>Used By:</strong></p>
                    <p className="dependency-list">{dependents.join(', ')}</p>
                </div>
            )}

            {cell.comment && (
                <div className="cell-detail-section">
                    <p><strong>Comment:</strong></p>
                    <p className="comment-display">{cell.comment}</p>
                </div>
            )}

            {cell.format && (
                <div className="cell-detail-section">
                    <p><strong>Format:</strong></p>
                    <p className="format-display">{cell.format}</p>
                </div>
            )}

            <div className="cell-detail-section view-mode-indicator">
                <p><strong>View Mode:</strong> {viewMode.toUpperCase()}</p>
            </div>
        </div>
    );
}

/**
 * Context Menu Component
 */
function ContextMenu({ x, y, cellRef, onClose, onFormat, onCut, onCopy, onPaste, onPasteValues }) {
    const menuRef = useRef(null);

    useEffect(() => {
        const handleClickOutside = (e) => {
            if (menuRef.current && !menuRef.current.contains(e.target)) {
                onClose();
            }
        };

        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, [onClose]);

    const [formatMenuOpen, setFormatMenuOpen] = useState(false);
    const [alignMenuOpen, setAlignMenuOpen] = useState(false);
    const [numberFormatMenuOpen, setNumberFormatMenuOpen] = useState(false);

    const handleFormatClick = (format) => {
        onFormat(format);
        onClose();
    };

    return (
        <div
            ref={menuRef}
            className="context-menu"
            style={{ left: x, top: y }}
        >
            <div className="context-menu-item" onClick={() => onCut()}>
                <span>✂️ Cut</span>
            </div>
            <div className="context-menu-item" onClick={() => onCopy()}>
                <span>📋 Copy</span>
            </div>
            <div className="context-menu-item" onClick={() => onPaste()}>
                <span>📄 Paste</span>
            </div>
            <div className="context-menu-item" onClick={() => onPasteValues()}>
                <span>📝 Paste Values Only</span>
            </div>
            <div className="context-menu-separator"></div>
            <div
                className="context-menu-item context-menu-submenu"
                onMouseEnter={() => setFormatMenuOpen(true)}
                onMouseLeave={() => setFormatMenuOpen(false)}
            >
                <span>🎨 Format</span>
                <span className="submenu-arrow">▶</span>
                {formatMenuOpen && (
                    <div className="context-submenu">
                        <div className="context-menu-item" onClick={() => handleFormatClick('bold')}>
                            <span><strong>Bold</strong></span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('italic')}>
                            <span><em>Italic</em></span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('underline')}>
                            <span><u>Underline</u></span>
                        </div>
                        <div className="context-menu-separator"></div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('font-size:14px')}>
                            <span>Font Size: Small</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('font-size:16px')}>
                            <span>Font Size: Medium</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('font-size:20px')}>
                            <span>Font Size: Large</span>
                        </div>
                        <div className="context-menu-separator"></div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('color:red')}>
                            <span style={{color: 'red'}}>Text: Red</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('color:blue')}>
                            <span style={{color: 'blue'}}>Text: Blue</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('color:green')}>
                            <span style={{color: 'green'}}>Text: Green</span>
                        </div>
                        <div className="context-menu-separator"></div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('background:yellow')}>
                            <span style={{background: 'yellow'}}>Background: Yellow</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('background:lightblue')}>
                            <span style={{background: 'lightblue'}}>Background: Light Blue</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('background:lightgreen')}>
                            <span style={{background: 'lightgreen'}}>Background: Light Green</span>
                        </div>
                        <div className="context-menu-separator"></div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('')}>
                            <span>Clear Format</span>
                        </div>
                    </div>
                )}
            </div>
            <div
                className="context-menu-item context-menu-submenu"
                onMouseEnter={() => setAlignMenuOpen(true)}
                onMouseLeave={() => setAlignMenuOpen(false)}
            >
                <span>↔️ Align</span>
                <span className="submenu-arrow">▶</span>
                {alignMenuOpen && (
                    <div className="context-submenu">
                        <div className="context-menu-item" onClick={() => handleFormatClick('align:left')}>
                            <span>⬅️ Align Left</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('align:center')}>
                            <span>↔️ Align Center</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('align:right')}>
                            <span>➡️ Align Right</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('align:justify')}>
                            <span>⬌ Justify</span>
                        </div>
                    </div>
                )}
            </div>
            <div
                className="context-menu-item context-menu-submenu"
                onMouseEnter={() => setNumberFormatMenuOpen(true)}
                onMouseLeave={() => setNumberFormatMenuOpen(false)}
            >
                <span>🔢 Number Format</span>
                <span className="submenu-arrow">▶</span>
                {numberFormatMenuOpen && (
                    <div className="context-submenu">
                        <div className="context-menu-item" onClick={() => handleFormatClick('number:0')}>
                            <span>Integer (0)</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('number:0.0')}>
                            <span>1 Decimal (0.0)</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('number:0.00')}>
                            <span>2 Decimals (0.00)</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('number:0.000')}>
                            <span>3 Decimals (0.000)</span>
                        </div>
                        <div className="context-menu-separator"></div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('currency:USD')}>
                            <span>💵 Currency (USD)</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('currency:EUR')}>
                            <span>💶 Currency (EUR)</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('currency:GBP')}>
                            <span>💷 Currency (GBP)</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('currency:JPY')}>
                            <span>💴 Currency (JPY)</span>
                        </div>
                        <div className="context-menu-separator"></div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('percent:0%')}>
                            <span>Percent (0%)</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('percent:0.0%')}>
                            <span>Percent (0.0%)</span>
                        </div>
                        <div className="context-menu-item" onClick={() => handleFormatClick('percent:0.00%')}>
                            <span>Percent (0.00%)</span>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
}

/**
 * Settings Modal Component
 */
function SettingsModal({ isOpen, onClose, model, adapter, onScriptExecuted }) {
    const [setupScript, setSetupScript] = useState('');
    const [executeMessage, setExecuteMessage] = useState('');

    useEffect(() => {
        if (isOpen && model) {
            setSetupScript(model.getSetupScript());
            setExecuteMessage('');
        }
    }, [isOpen, model]);

    const handleSave = async () => {
        if (model && adapter) {
            model.setSetupScript(setupScript);

            // Execute the setup script
            const result = await adapter.executeSetupScript(setupScript);

            if (result.success) {
                setExecuteMessage('✅ ' + result.message);
                setTimeout(() => {
                    onScriptExecuted();
                    onClose();
                }, 1000);
            } else {
                setExecuteMessage('❌ ' + result.message);
            }
        }
    };

    if (!isOpen) return null;

    return (
        <div className="modal-overlay" onClick={onClose}>
            <div className="modal-content" onClick={(e) => e.stopPropagation()}>
                <div className="modal-header">
                    <h2>Spreadsheet Setup</h2>
                    <button className="close-button" onClick={onClose}>×</button>
                </div>
                <div className="modal-body">
                    <p>Page-level RexxJS code that runs once when the spreadsheet loads.</p>
                    <p><strong>Use this to load function libraries:</strong></p>
                    <textarea
                        className="setup-script-editor"
                        value={setupScript}
                        onChange={(e) => setSetupScript(e.target.value)}
                        placeholder={`// Example: Load Excel-like functions\nREQUIRE "cwd:../../extras/functions/excel/src/excel-functions.js"\n\n// Example: Load R statistics\nREQUIRE "cwd:../../extras/functions/r-inspired/src/r-statistics-functions.js"\n\n// Now VLOOKUP, SUMIF, MEAN, MEDIAN, etc. are available in all cells`}
                        rows={12}
                    />
                    {executeMessage && (
                        <div className={`execute-message ${executeMessage.startsWith('✅') ? 'success' : 'error'}`}>
                            {executeMessage}
                        </div>
                    )}
                </div>
                <div className="modal-footer">
                    <button className="button-secondary" onClick={onClose}>Cancel</button>
                    <button className="button-primary" onClick={handleSave}>Save & Execute</button>
                </div>
            </div>
        </div>
    );
}

/**
 * Main App Component
 */
function App() {
    const [model, setModel] = useState(null);
    const [adapter, setAdapter] = useState(null);
    const [selectedCell, setSelectedCell] = useState('A1');
    const [selectionRange, setSelectionRange] = useState(null);
    const [isSelecting, setIsSelecting] = useState(false);
    const [sheetName, setSheetName] = useState('Sheet1');
    const [updateCounter, setUpdateCounter] = useState(0);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState(null);
    const [settingsOpen, setSettingsOpen] = useState(false);
    const [viewMode, setViewMode] = useState('normal');
    const [startEditCallback, setStartEditCallback] = useState(null);
    const [contextMenu, setContextMenu] = useState(null);
    const [clipboard, setClipboard] = useState(null);
    const [currentFilePath, setCurrentFilePath] = useState(null);
    const isTransitioningToEdit = useRef(false);
    const bufferedKeys = useRef([]);
    const initializationInProgress = useRef(false);
    const hasInitialized = useRef(false);

    const visibleRows = 20;
    const visibleCols = 10;

    // Define handlers before useEffects that use them
    const handleNavigation = useCallback((key, shiftKey) => {
        if (!selectedCell) return;

        isTransitioningToEdit.current = false;  // Reset transition flag
        bufferedKeys.current = [];  // Clear buffered keys
        const { col, row } = parseCellRef(selectedCell);
        let newCol = col;
        let newRow = row;

        switch (key.toLowerCase()) {
            case 'arrowdown':
            case 'enter':
                newRow = Math.min(row + 1, visibleRows);
                break;
            case 'arrowup':
                newRow = Math.max(row - 1, 1);
                break;
            case 'arrowright':
            case 'tab':
                if (shiftKey && key.toLowerCase() === 'tab') {
                    newCol = Math.max(col - 1, 1);
                } else {
                    newCol = Math.min(col + 1, visibleCols);
                }
                break;
            case 'arrowleft':
                newCol = Math.max(col - 1, 1);
                break;
        }

        const newCellRef = SpreadsheetModel.formatCellRef(newCol, newRow);
        setSelectedCell(newCellRef);
        setSelectionRange(null);
    }, [selectedCell, visibleRows, visibleCols]);

    const handleCopy = useCallback((includeFormulas = true) => {
        if (!model) return;

        const clipboardData = {
            cells: [],
            asText: ''
        };

        if (selectionRange) {
            const { startCol, startRow, endCol, endRow } = selectionRange;
            const minCol = Math.min(startCol, endCol);
            const maxCol = Math.max(startCol, endCol);
            const minRow = Math.min(startRow, endRow);
            const maxRow = Math.max(startRow, endRow);

            let textRows = [];
            for (let row = minRow; row <= maxRow; row++) {
                const rowValues = [];
                for (let col = minCol; col <= maxCol; col++) {
                    const cellRef = SpreadsheetModel.formatCellRef(col, row);
                    const cell = model.getCell(cellRef);
                    rowValues.push(cell.value || '');

                    clipboardData.cells.push({
                        offsetCol: col - minCol,
                        offsetRow: row - minRow,
                        value: cell.value,
                        expression: includeFormulas ? cell.expression : null,
                        format: cell.format,
                        comment: cell.comment
                    });
                }
                textRows.push(rowValues.join('\t'));
            }
            clipboardData.asText = textRows.join('\n');
        } else if (selectedCell) {
            const cell = model.getCell(selectedCell);
            clipboardData.asText = cell.value || '';
            clipboardData.cells.push({
                offsetCol: 0,
                offsetRow: 0,
                value: cell.value,
                expression: includeFormulas ? cell.expression : null,
                format: cell.format,
                comment: cell.comment
            });
        }

        // Store in internal clipboard
        setClipboard(clipboardData);

        // Also copy text to system clipboard
        navigator.clipboard.writeText(clipboardData.asText).then(() => {
            console.log('Copied to clipboard');
        }).catch(err => {
            console.error('Failed to copy:', err);
        });
    }, [model, selectedCell, selectionRange]);

    const handleCut = useCallback(() => {
        if (!model || !adapter) return;

        // Copy first
        handleCopy();

        // Then clear cells
        if (selectionRange) {
            const { startCol, startRow, endCol, endRow } = selectionRange;
            const minCol = Math.min(startCol, endCol);
            const maxCol = Math.max(startCol, endCol);
            const minRow = Math.min(startRow, endRow);
            const maxRow = Math.max(startRow, endRow);

            for (let row = minRow; row <= maxRow; row++) {
                for (let col = minCol; col <= maxCol; col++) {
                    const cellRef = SpreadsheetModel.formatCellRef(col, row);
                    model.setCell(cellRef, '', adapter);
                }
            }
        } else if (selectedCell) {
            model.setCell(selectedCell, '', adapter);
        }

        setUpdateCounter(c => c + 1);
    }, [model, adapter, selectedCell, selectionRange, handleCopy]);

    const handlePaste = useCallback(async (valuesOnly = false) => {
        if (!model || !adapter || !clipboard || !selectedCell) return;

        const { col: startCol, row: startRow } = parseCellRef(selectedCell);

        for (const cellData of clipboard.cells) {
            const targetCol = startCol + cellData.offsetCol;
            const targetRow = startRow + cellData.offsetRow;
            const targetRef = SpreadsheetModel.formatCellRef(targetCol, targetRow);

            let content = cellData.value;
            if (!valuesOnly && cellData.expression) {
                content = '=' + cellData.expression;
            }

            const metadata = {
                format: valuesOnly ? '' : cellData.format,
                comment: valuesOnly ? '' : cellData.comment
            };

            await model.setCell(targetRef, content, adapter, metadata);
        }

        setUpdateCounter(c => c + 1);
    }, [model, adapter, clipboard, selectedCell]);

    const handleFormat = useCallback(async (format) => {
        if (!model || !selectedCell) return;

        const cell = model.getCell(selectedCell);
        const existingFormat = cell.format || '';

        // Parse existing format
        const formatParts = existingFormat.split(';').filter(p => p.trim());

        if (format === '') {
            // Clear all formatting
            model.setCellMetadata(selectedCell, { format: '' });
        } else {
            // Toggle or set format
            const [formatKey] = format.split(':');

            // Remove existing format of same type
            const filtered = formatParts.filter(p => {
                const [key] = p.split(':');
                return key !== formatKey && key !== format; // Handle both "bold" and "color:red"
            });

            // Add new format if not already present
            if (!formatParts.includes(format)) {
                filtered.push(format);
            }

            model.setCellMetadata(selectedCell, { format: filtered.join(';') });
        }

        setUpdateCounter(c => c + 1);
    }, [model, selectedCell]);

    const handleContextMenu = useCallback((e, cellRef) => {
        e.preventDefault();
        setContextMenu({
            x: e.clientX,
            y: e.clientY,
            cellRef: cellRef
        });
    }, []);

    const handleSave = useCallback(async () => {
        if (!model) return;

        try {
            // Check if we're in Tauri mode
            if (typeof window.__TAURI__ !== 'undefined') {
                // Tauri mode - save to file system
                const { save } = window.__TAURI__.dialog;
                const { writeTextFile } = window.__TAURI__.fs;

                let filePath = currentFilePath;

                // If no current file, show save dialog
                if (!filePath) {
                    filePath = await save({
                        filters: [{
                            name: 'Spreadsheet',
                            extensions: ['json']
                        }]
                    });
                }

                if (filePath) {
                    const data = model.toJSON();
                    const contents = JSON.stringify(data, null, 2);
                    await writeTextFile(filePath, contents);
                    setCurrentFilePath(filePath);
                    console.log('Saved to file:', filePath);
                    alert('Spreadsheet saved successfully!');
                }
            } else {
                // Web mode - save via PUT request
                const hash = window.location.hash.substring(1);
                let saveUrl = null;

                // Parse hash to find save URL
                if (hash.startsWith('load=')) {
                    saveUrl = hash.substring(5);
                } else if (currentFilePath) {
                    saveUrl = currentFilePath;
                } else {
                    // No URL specified - trigger download instead
                    const data = model.toJSON();
                    const contents = JSON.stringify(data, null, 2);
                    const blob = new Blob([contents], { type: 'application/json' });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'spreadsheet.json';
                    a.click();
                    URL.revokeObjectURL(url);
                    console.log('Downloaded as JSON file');
                    return;
                }

                if (saveUrl) {
                    const data = model.toJSON();
                    const response = await fetch(saveUrl, {
                        method: 'PUT',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify(data, null, 2)
                    });

                    if (response.ok) {
                        console.log('Saved to URL:', saveUrl);
                        alert('Spreadsheet saved successfully!');
                    } else {
                        throw new Error(`Save failed: ${response.status} ${response.statusText}`);
                    }
                }
            }
        } catch (error) {
            console.error('Save failed:', error);
            alert('Failed to save spreadsheet: ' + error.message);
        }
    }, [model, currentFilePath]);

    const loadSampleData = useCallback(async (model, adapter) => {
        const sampleData = {
            'A1': '10',
            'A2': '20',
            'A3': '30',
            'A4': '=A1 + A2',
            'A5': '=SUM_RANGE("A1:A3")',
            'B1': 'Hello',
            'B2': '=UPPER(B1)',
            'C1': '5',
            'C2': '3',
            'C3': '=C1 * C2'
        };

        for (const [ref, content] of Object.entries(sampleData)) {
            await model.setCell(ref, content, adapter);
        }

        // Note: No setUpdateCounter here - SETCELL already triggers spreadsheet-update events
    }, []); // No dependencies - function is stable

    const initializeSpreadsheet = useCallback(async (initialFilePath = null) => {
        // Prevent multiple initializations using ref
        if (hasInitialized.current || initializationInProgress.current) {
            console.log('[Init] Already initialized or in progress, skipping');
            return;
        }

        initializationInProgress.current = true;
        console.log('[Init] Starting initialization...');

        try {
            setIsLoading(true);

            // Wait for RexxJS to be available
            let attempts = 0;
            while ((typeof RexxInterpreter === 'undefined' || typeof parse === 'undefined') && attempts < 100) {
                await new Promise(resolve => setTimeout(resolve, 100));
                attempts++;
            }

            if (typeof RexxInterpreter === 'undefined' || typeof parse === 'undefined') {
                throw new Error('RexxJS interpreter failed to load. Please refresh the page.');
            }

            const newModel = new SpreadsheetModel(100, 26);
            const newAdapter = new SpreadsheetRexxAdapter(newModel);
            await newAdapter.initializeInterpreter(RexxInterpreter);
            window.spreadsheetAdapter = newAdapter;

            // Register spreadsheet control functions (SETCELL, GETCELL, etc.)
            // These are available in cell expressions AND via remote ADDRESS commands
            const controlFunctions = createSpreadsheetControlFunctions(newModel, newAdapter);
            for (const [name, func] of Object.entries(controlFunctions)) {
                newAdapter.interpreter.externalFunctions[name] = func;
            }
            console.log('[Init] Registered control functions:', Object.keys(controlFunctions));

            // Note: No ADDRESS handler needed anymore!
            // Remote commands now execute via isolated interpreters calling REXX functions directly.
            // This ensures single source of truth - same SETCELL/GETCELL used in expressions and remotely.

            // Load spreadsheet functions via script tag instead of REQUIRE
            // (REQUIRE doesn't work well with tauri:// protocol)
            const libUrl = window.location.origin + '/lib/spreadsheet-functions.js';
            const scriptLoaded = await new Promise((resolve) => {
                const script = document.createElement('script');
                script.src = libUrl;
                script.onload = () => {
                    console.log('✓ Loaded', libUrl, 'via script tag');
                    resolve(true);
                };
                script.onerror = () => {
                    console.error('Failed to load', libUrl);
                    resolve(false);
                };
                document.head.appendChild(script);
            });

            if (!scriptLoaded) {
                throw new Error(`Failed to load spreadsheet functions from ${libUrl}`);
            }

            const result = { success: true, message: 'Spreadsheet functions loaded' };

            if (!result.success) {
                throw new Error(`Failed to load spreadsheet functions: ${result.error}`);
            }

            // Load file if provided
            if (initialFilePath) {
                try {
                    console.log('Loading spreadsheet from file:', initialFilePath);
                    const data = await SpreadsheetLoader.loadFromFile(initialFilePath);
                    await SpreadsheetLoader.importIntoModel(newModel, data, newAdapter);
                    console.log('Spreadsheet loaded successfully');
                } catch (error) {
                    console.log('File loading not available or failed:', error.message);
                    // Fall through to sample data
                }
            }

            if (!initialFilePath) {
                // Load sample data if no file provided
                const setupScript = newModel.getSetupScript();
                if (setupScript) {
                    await newAdapter.executeSetupScript(setupScript);
                }
                loadSampleData(newModel, newAdapter);
            }

            setModel(newModel);
            setAdapter(newAdapter);
            setIsLoading(false);
            hasInitialized.current = true;
            initializationInProgress.current = false;
            console.log('[Init] ✅ Initialization complete');
        } catch (err) {
            console.error('Failed to initialize spreadsheet:', err);
            setError(err.message);
            setIsLoading(false);
            initializationInProgress.current = false;
        }
    }, []); // No dependencies - stable function

    // Initialize on mount and listen for file path from Tauri
    useEffect(() => {
        let unlistenFn = null;
        let timeout = null;
        let initialized = false;

        const setupTauriListener = async () => {
            try {
                // Try to listen for Tauri events (will work in desktop app)
                console.log('[Init] Setting up Tauri initial-file listener...');
                unlistenFn = await listen('initial-file', (event) => {
                    if (!initialized) {
                        initialized = true;
                        console.log('Received initial-file event:', event.payload);
                        const filePath = event.payload.file_path;
                        initializeSpreadsheet(filePath);
                    }
                });

                // Fallback: Initialize without file after timeout if no event received
                timeout = setTimeout(() => {
                    if (!initialized) {
                        initialized = true;
                        console.log('No initial file event received, initializing with defaults');
                        initializeSpreadsheet();
                    }
                }, 1000);
            } catch (error) {
                // Not in Tauri mode, or Tauri not available - initialize immediately
                if (!initialized) {
                    initialized = true;
                    console.log('[Init] Not in Tauri mode, initializing immediately');
                    initializeSpreadsheet();
                }
            }
        };

        setupTauriListener();

        return () => {
            if (unlistenFn) unlistenFn();
            if (timeout) clearTimeout(timeout);
        };
    }, []); // Empty deps - only run once on mount

    // COMET-style polling for control bus commands (Selenium-RC style)
    useEffect(() => {
        if (!model || !adapter) {
            console.log('[Control Bus] Waiting for model and adapter...');
            return;
        }

        console.log('[Control Bus] Starting COMET-style polling...');
        let stopPolling = false;
        let requestCounter = 0;

        const pollForCommands = async () => {
            while (!stopPolling) {
                try {
                    // Poll for pending commands
                    const response = await fetch('http://localhost:2410/api/poll', {
                        method: 'GET',
                        headers: {
                            'Authorization': 'Bearer dev-token-12345',
                            'Content-Type': 'application/json',
                        },
                    });

                    if (!response.ok) {
                        console.error('[Control Bus] Poll failed:', response.status);
                        await new Promise(resolve => setTimeout(resolve, 1000));
                        continue;
                    }

                    const data = await response.json();

                    if (data && data.request_id && data.command) {
                        const requestId = ++requestCounter;
                        console.log(`[Control Bus #${requestId}] 🎯 Got command!`, data);
                        const { request_id, command } = data;

                        // Create isolated interpreter for this request
                        let scopeElement = null;
                        let isolatedInterpreter = null;
                        let executionResult = null;
                        let executionError = null;

                        try {
                            // Create hidden DOM element for scoping this interpreter
                            scopeElement = document.createElement('div');
                            scopeElement.id = `rexx-scope-${requestId}`;
                            scopeElement.className = 'RexxScript';
                            scopeElement.style.display = 'none';
                            document.body.appendChild(scopeElement);

                            console.log(`[Control Bus #${requestId}] Created isolated scope`);

                            // Create isolated interpreter
                            isolatedInterpreter = RexxInterpreter.builder().build();
                            isolatedInterpreter.scopeElement = scopeElement;

                            // Register control functions
                            const controlFunctions = createSpreadsheetControlFunctions(model, adapter);
                            for (const [name, func] of Object.entries(controlFunctions)) {
                                isolatedInterpreter.externalFunctions[name] = func;
                            }

                            // Command is already valid REXX syntax - just wrap with assignment
                            const rexxCode = `result = ${command.trim()}`;

                            console.log(`[Control Bus #${requestId}] Executing:`, rexxCode);

                            // Execute in isolated interpreter
                            const commands = parse(rexxCode);
                            await isolatedInterpreter.run(commands);

                            const result = isolatedInterpreter.getVariable('result');

                            console.log(`[Control Bus #${requestId}] ✅ Done!`, { result });
                            executionResult = result;
                        } catch (error) {
                            console.error(`[Control Bus #${requestId}] ❌ Error:`, error);
                            executionError = error.message || String(error);
                        } finally {
                            // Clean up
                            if (scopeElement && scopeElement.parentNode) {
                                document.body.removeChild(scopeElement);
                            }

                            // Post result back to Rust
                            try {
                                await fetch('http://localhost:2410/api/result', {
                                    method: 'POST',
                                    headers: {
                                        'Authorization': 'Bearer dev-token-12345',
                                        'Content-Type': 'application/json',
                                    },
                                    body: JSON.stringify({
                                        request_id: request_id,
                                        result: {
                                            value: executionResult,
                                            error: executionError
                                        }
                                    })
                                });
                                console.log(`[Control Bus #${requestId}] 📤 Result posted`);
                            } catch (postError) {
                                console.error(`[Control Bus #${requestId}] ❌ Failed to post result:`, postError);
                            }
                        }
                    } else {
                        // No commands, wait a bit before polling again
                        await new Promise(resolve => setTimeout(resolve, 100));
                    }
                } catch (error) {
                    console.error('[Control Bus] Polling error:', error);
                    await new Promise(resolve => setTimeout(resolve, 1000));
                }
            }
        };

        pollForCommands();

        return () => {
            console.log('[Control Bus] Stopping polling');
            stopPolling = true;
        };
    }, [model, adapter]);

    // Listen for spreadsheet updates from ADDRESS handler
    useEffect(() => {
        const handleUpdate = () => {
            console.log('[UI] 🔄 spreadsheet-update event received, triggering re-render');
            setUpdateCounter(c => c + 1);
        };
        console.log('[UI] Setting up spreadsheet-update listener');
        window.addEventListener('spreadsheet-update', handleUpdate);
        return () => {
            console.log('[UI] Removing spreadsheet-update listener');
            window.removeEventListener('spreadsheet-update', handleUpdate);
        };
    }, []);

    // Keyboard handler for navigation and copy
    useEffect(() => {
        const handleKeyDown = (e) => {
            // Only handle if not in an input/textarea
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
                return;
            }

            // View mode shortcuts
            const key = e.key.toLowerCase();
            if (key === 'v') {
                setViewMode('values');
                return;
            } else if (key === 'e') {
                setViewMode('expressions');
                return;
            } else if (key === 'f') {
                setViewMode('formats');
                return;
            }

            // Copy selection
            if ((e.ctrlKey || e.metaKey) && key === 'c') {
                e.preventDefault();
                handleCopy();
                return;
            }

            // Navigation
            if (['arrowdown', 'arrowup', 'arrowleft', 'arrowright', 'tab', 'enter'].includes(key)) {
                e.preventDefault();
                handleNavigation(e.key, e.shiftKey);
                return;
            }

            // Start editing on any printable character
            if (key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
                e.preventDefault();
                if (!isTransitioningToEdit.current) {
                    // First character - start edit mode
                    if (startEditCallback) {
                        isTransitioningToEdit.current = true;
                        bufferedKeys.current = [e.key];
                        startEditCallback(e.key);
                    }
                } else {
                    // Subsequent characters while transitioning - buffer them
                    bufferedKeys.current.push(e.key);
                }
            }
        };

        const handleKeyUp = (e) => {
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
                return;
            }

            const key = e.key.toLowerCase();
            if (key === 'v' || key === 'e' || key === 'f') {
                setViewMode('normal');
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        window.addEventListener('keyup', handleKeyUp);
        return () => {
            window.removeEventListener('keydown', handleKeyDown);
            window.removeEventListener('keyup', handleKeyUp);
        };
    }, [selectedCell, selectionRange, model, adapter, startEditCallback, handleCopy, handleNavigation]);

    // Handle hash changes for sheet name
    useEffect(() => {
        const handleHashChange = () => {
            const hash = window.location.hash.substring(1);
            if (hash) {
                setSheetName(hash);
            }
        };

        window.addEventListener('hashchange', handleHashChange);
        handleHashChange();

        return () => window.removeEventListener('hashchange', handleHashChange);
    }, []);


    const handleEditCell = useCallback(async (cellRef, content, navigationKey, shiftKey) => {
        if (!model || !adapter) return;

        await model.setCell(cellRef, content, adapter);
        // Note: No setUpdateCounter here - model.setCell triggers spreadsheet-update event

        // Handle navigation after edit
        if (navigationKey) {
            handleNavigation(navigationKey, shiftKey);
        }
    }, [model, adapter, handleNavigation]);

    const handleSelectCell = useCallback((cellRef, shiftKey) => {
        isTransitioningToEdit.current = false;  // Reset transition flag
        bufferedKeys.current = [];  // Clear buffered keys
        if (shiftKey && selectedCell) {
            // Extend selection
            const start = parseCellRef(selectedCell);
            const end = parseCellRef(cellRef);
            setSelectionRange({
                startCol: start.col,
                startRow: start.row,
                endCol: end.col,
                endRow: end.row
            });
        } else {
            setSelectedCell(cellRef);
            setSelectionRange(null);
        }
    }, [selectedCell]);

    const handleSelectionStart = useCallback((cellRef) => {
        setIsSelecting(true);
        setSelectedCell(cellRef);
        const { col, row } = parseCellRef(cellRef);
        setSelectionRange({
            startCol: col,
            startRow: row,
            endCol: col,
            endRow: row
        });
    }, []);

    const handleSelectionMove = useCallback((cellRef) => {
        if (isSelecting && selectionRange) {
            const { col, row } = parseCellRef(cellRef);
            setSelectionRange(prev => ({
                ...prev,
                endCol: col,
                endRow: row
            }));
        }
    }, [isSelecting, selectionRange]);

    const handleSelectionEnd = useCallback(() => {
        setIsSelecting(false);
    }, []);

    const handleStartCellEdit = (callback) => {
        setStartEditCallback(() => (char) => callback(char));
        return () => setStartEditCallback(null);
    };

    if (isLoading) {
        return (
            <div className="app loading">
                <div className="loading-message">
                    <h2>Loading RexxJS Spreadsheet...</h2>
                    <p>Initializing interpreter and spreadsheet model...</p>
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div className="app error">
                <div className="error-message">
                    <h2>Error</h2>
                    <p>{error}</p>
                </div>
            </div>
        );
    }

    return (
        <div className="app">
            <div className="header">
                <h1>RexxJS Spreadsheet POC</h1>
                <div className="header-controls">
                    <button className="save-button" onClick={handleSave} title="Save spreadsheet">
                        💾 Save
                    </button>
                    <div className="view-mode-badge" title="Hold V/E/F to peek at different views">
                        View: {viewMode.toUpperCase()}
                    </div>
                    <button className="settings-button" onClick={() => setSettingsOpen(true)}>
                        ⚙️ Setup
                    </button>
                    <div className="sheet-name">Sheet: {sheetName}</div>
                </div>
            </div>

            <FormulaBar
                selectedCell={selectedCell}
                model={model}
                onEdit={handleEditCell}
            />

            <div className="main-content">
                <Grid
                    model={model}
                    selectedCell={selectedCell}
                    selectionRange={selectionRange}
                    onSelectCell={handleSelectCell}
                    onEditCell={handleEditCell}
                    onStartCellEdit={handleStartCellEdit}
                    visibleRows={visibleRows}
                    visibleCols={visibleCols}
                    viewMode={viewMode}
                    onSelectionStart={handleSelectionStart}
                    onSelectionMove={handleSelectionMove}
                    onSelectionEnd={handleSelectionEnd}
                    onContextMenu={handleContextMenu}
                    bufferedKeysRef={bufferedKeys}
                    isTransitioningRef={isTransitioningToEdit}
                />

                <InfoPanel
                    selectedCell={selectedCell}
                    selectionRange={selectionRange}
                    model={model}
                    viewMode={viewMode}
                />
            </div>

            {contextMenu && (
                <ContextMenu
                    x={contextMenu.x}
                    y={contextMenu.y}
                    cellRef={contextMenu.cellRef}
                    onClose={() => setContextMenu(null)}
                    onFormat={handleFormat}
                    onCut={() => { handleCut(); setContextMenu(null); }}
                    onCopy={() => { handleCopy(); setContextMenu(null); }}
                    onPaste={() => { handlePaste(false); setContextMenu(null); }}
                    onPasteValues={() => { handlePaste(true); setContextMenu(null); }}
                />
            )}

            <SettingsModal
                isOpen={settingsOpen}
                onClose={() => setSettingsOpen(false)}
                model={model}
                adapter={adapter}
                onScriptExecuted={() => setUpdateCounter(c => c + 1)}
            />
        </div>
    );
}

// Export the App component as default
export default App;
