/**
 * SpreadsheetApp - React components for the spreadsheet UI
 *
 * Components:
 * - Cell: Individual spreadsheet cell
 * - Grid: Spreadsheet grid
 * - FormulaBar: Formula/value editor
 * - App: Main application with state management
 */

const { useState, useEffect, useRef, useCallback } = React;

/**
 * Cell Component
 */
function Cell({ cellRef, cell, isSelected, onSelect, onEdit, viewMode, width, height, model, formattedValue }) {
    const inputRef = useRef(null);
    const [isEditing, setIsEditing] = useState(false);
    const [editValue, setEditValue] = useState('');

    useEffect(() => {
        if (isEditing && inputRef.current) {
            inputRef.current.focus();
            inputRef.current.select();
        }
    }, [isEditing]);

    const handleDoubleClick = () => {
        setIsEditing(true);
        setEditValue(cell.expression ? '=' + cell.expression : cell.value);
    };

    const handleKeyDown = (e) => {
        if (e.key === 'Enter') {
            onEdit(cellRef, editValue);
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
        // Show only literal values, blank if formula
        displayValue = cell.expression ? '' : (formattedValue || cell.value || '');
        showCell = !cell.expression || cell.value === '';
    } else if (viewMode === 'expressions') {
        // Show only formulas, hide value cells
        displayValue = cell.expression ? '=' + cell.expression : '';
        showCell = !!cell.expression;
    } else if (viewMode === 'formats') {
        // Show format strings only
        displayValue = cell.format || '';
        showCell = !!cell.format;
    } else {
        // Normal mode - show evaluated values (formatted if format is present)
        displayValue = cell.error ? cell.value : (formattedValue || cell.value || '');
    }

    const hasError = !!cell.error;
    const hasFormula = !!cell.expression;
    const hasFormat = !!cell.format;
    const hasComment = !!cell.comment;
    const hasStyle = !!cell.style || !!cell.styleExpression;

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
    if (cell.styleExpression) {
        title += (title ? '\n' : '') + '🎨 Style: ' + cell.styleExpression;
    }

    // Build inline styles from cell.style
    const inlineStyle = {
        width: `${width}px`,
        minWidth: `${width}px`,
        height: `${height}px`
    };

    if (cell.style && viewMode === 'normal') {
        if (cell.style.backgroundColor) {
            inlineStyle.backgroundColor = cell.style.backgroundColor;
        }
        if (cell.style.color) {
            inlineStyle.color = cell.style.color;
        }
        if (cell.style.fontWeight) {
            inlineStyle.fontWeight = cell.style.fontWeight;
        }
        if (cell.style.fontStyle) {
            inlineStyle.fontStyle = cell.style.fontStyle;
        }
        if (cell.style.textAlign) {
            inlineStyle.textAlign = cell.style.textAlign;
        }
        if (cell.style.border) {
            inlineStyle.border = cell.style.border;
        }
    }

    return (
        <div
            className={`cell ${isSelected ? 'selected' : ''} ${hasError ? 'error' : ''} ${hasFormula ? 'formula' : ''} ${hasFormat ? 'formatted' : ''} ${hasComment ? 'commented' : ''} ${hasStyle ? 'has-custom-style' : ''} ${viewMode !== 'normal' ? 'view-mode-' + viewMode : ''}`}
            onClick={() => onSelect(cellRef)}
            onDoubleClick={handleDoubleClick}
            title={title}
            style={inlineStyle}
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
function ColumnHeader({ col, width, onResize }) {
    const [isResizing, setIsResizing] = useState(false);
    const [startX, setStartX] = useState(0);
    const [startWidth, setStartWidth] = useState(0);

    const handleMouseDown = (e) => {
        setIsResizing(true);
        setStartX(e.clientX);
        setStartWidth(width);
        e.preventDefault();
    };

    useEffect(() => {
        if (!isResizing) return;

        const handleMouseMove = (e) => {
            const delta = e.clientX - startX;
            const newWidth = Math.max(20, startWidth + delta);
            onResize(col, newWidth);
        };

        const handleMouseUp = () => {
            setIsResizing(false);
        };

        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);

        return () => {
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', handleMouseUp);
        };
    }, [isResizing, startX, startWidth, col, onResize]);

    return (
        <div className="column-header" style={{ width: `${width}px`, minWidth: `${width}px` }}>
            <span>{SpreadsheetModel.colNumberToLetter(col)}</span>
            <div
                className="resize-handle resize-handle-col"
                onMouseDown={handleMouseDown}
            />
        </div>
    );
}

/**
 * Row Header Component
 */
function RowHeader({ row, height, onResize }) {
    const [isResizing, setIsResizing] = useState(false);
    const [startY, setStartY] = useState(0);
    const [startHeight, setStartHeight] = useState(0);

    const handleMouseDown = (e) => {
        setIsResizing(true);
        setStartY(e.clientY);
        setStartHeight(height);
        e.preventDefault();
    };

    useEffect(() => {
        if (!isResizing) return;

        const handleMouseMove = (e) => {
            const delta = e.clientY - startY;
            const newHeight = Math.max(15, startHeight + delta);
            onResize(row, newHeight);
        };

        const handleMouseUp = () => {
            setIsResizing(false);
        };

        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);

        return () => {
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', handleMouseUp);
        };
    }, [isResizing, startY, startHeight, row, onResize]);

    return (
        <div className="row-header" style={{ height: `${height}px`, minHeight: `${height}px` }}>
            <span>{row}</span>
            <div
                className="resize-handle resize-handle-row"
                onMouseDown={handleMouseDown}
            />
        </div>
    );
}

/**
 * Grid Component
 */
function Grid({ model, selectedCell, onSelectCell, onEditCell, visibleRows, visibleCols, viewMode, onColumnResize, onRowResize }) {
    const rows = [];

    // Header row with column letters
    const headerRow = (
        <div key="header" className="grid-row header-row">
            <div className="corner-cell"></div>
            {Array.from({ length: visibleCols }, (_, i) => {
                const col = i + 1;
                return (
                    <ColumnHeader
                        key={i}
                        col={col}
                        width={model.getColumnWidth(col)}
                        onResize={onColumnResize}
                    />
                );
            })}
        </div>
    );
    rows.push(headerRow);

    // Data rows
    for (let row = 1; row <= visibleRows; row++) {
        const cells = [];
        const rowHeight = model.getRowHeight(row);

        // Row header
        cells.push(
            <RowHeader
                key={`row-${row}`}
                row={row}
                height={rowHeight}
                onResize={onRowResize}
            />
        );

        // Data cells
        for (let col = 1; col <= visibleCols; col++) {
            const cellRef = SpreadsheetModel.formatCellRef(col, row);
            const cell = model.getCell(cellRef);
            const isSelected = selectedCell === cellRef;
            const colWidth = model.getColumnWidth(col);
            const formattedValue = model.formatCellValue(cellRef);

            cells.push(
                <Cell
                    key={cellRef}
                    cellRef={cellRef}
                    cell={cell}
                    isSelected={isSelected}
                    onSelect={onSelectCell}
                    onEdit={onEditCell}
                    viewMode={viewMode}
                    width={colWidth}
                    height={rowHeight}
                    model={model}
                    formattedValue={formattedValue}
                />
            );
        }

        rows.push(
            <div key={`row-${row}`} className="grid-row">
                {cells}
            </div>
        );
    }

    return <div className="grid">{rows}</div>;
}

/**
 * Formatting Toolbar Component
 */
function FormattingToolbar({ selectedCell, model, onApplyFormat, onApplyStyle }) {
    const [showFormatMenu, setShowFormatMenu] = useState(false);
    const [showStyleMenu, setShowStyleMenu] = useState(false);

    const formats = [
        { name: 'Currency', value: '$#,##0.00' },
        { name: 'Percentage', value: '0.00%' },
        { name: 'Number', value: '#,##0.00' },
        { name: 'Integer', value: '#,##0' },
        { name: 'Date (ISO)', value: 'yyyy-mm-dd' },
        { name: 'Date (US)', value: 'mm/dd/yyyy' }
    ];

    const stylePresets = [
        { name: 'Bold', style: { fontWeight: 'bold' } },
        { name: 'Italic', style: { fontStyle: 'italic' } },
        { name: 'Red Text', style: { color: '#d32f2f' } },
        { name: 'Green Text', style: { color: '#388e3c' } },
        { name: 'Blue Text', style: { color: '#1976d2' } },
        { name: 'Yellow BG', style: { backgroundColor: '#fff9c4' } },
        { name: 'Red BG', style: { backgroundColor: '#ffebee' } },
        { name: 'Green BG', style: { backgroundColor: '#e8f5e9' } },
        { name: 'Clear Style', style: null }
    ];

    return (
        <div className="formatting-toolbar">
            <div className="toolbar-group">
                <button
                    className="toolbar-button"
                    onClick={() => setShowFormatMenu(!showFormatMenu)}
                    disabled={!selectedCell}
                    title="Number Format"
                >
                    🔢 Format
                </button>
                {showFormatMenu && (
                    <div className="dropdown-menu">
                        {formats.map(fmt => (
                            <div
                                key={fmt.value}
                                className="dropdown-item"
                                onClick={() => {
                                    onApplyFormat(fmt.value);
                                    setShowFormatMenu(false);
                                }}
                            >
                                {fmt.name}
                            </div>
                        ))}
                        <div
                            className="dropdown-item"
                            onClick={() => {
                                onApplyFormat('');
                                setShowFormatMenu(false);
                            }}
                        >
                            Clear Format
                        </div>
                    </div>
                )}
            </div>
            <div className="toolbar-group">
                <button
                    className="toolbar-button"
                    onClick={() => setShowStyleMenu(!showStyleMenu)}
                    disabled={!selectedCell}
                    title="Cell Style"
                >
                    🎨 Style
                </button>
                {showStyleMenu && (
                    <div className="dropdown-menu">
                        {stylePresets.map(preset => (
                            <div
                                key={preset.name}
                                className="dropdown-item"
                                onClick={() => {
                                    onApplyStyle(preset.style);
                                    setShowStyleMenu(false);
                                }}
                            >
                                {preset.name}
                            </div>
                        ))}
                    </div>
                )}
            </div>
        </div>
    );
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
function InfoPanel({ selectedCell, model, viewMode }) {
    const cell = selectedCell && model ? model.getCell(selectedCell) : null;

    if (!selectedCell || !cell) {
        return (
            <div className="info-panel">
                <h3>Cell Details</h3>
                <p className="no-selection">No cell selected</p>
                <div className="help-section">
                    <p><strong>Hotkeys:</strong></p>
                    <ul>
                        <li><kbd>Ctrl+C</kbd> - Copy cell</li>
                        <li><kbd>Ctrl+X</kbd> - Cut cell</li>
                        <li><kbd>Ctrl+V</kbd> - Paste cell</li>
                        <li><kbd>V</kbd> - View values only</li>
                        <li><kbd>E</kbd> - View expressions only</li>
                        <li><kbd>F</kbd> - View formats only</li>
                        <li><kbd>N</kbd> - Normal view (default)</li>
                    </ul>
                    <p><strong>Named Variables:</strong></p>
                    <p>Use <strong>⚙️ Setup</strong> to define:</p>
                    <code style={{display: 'block', marginTop: '5px'}}>LET TAX_RATE = 0.07</code>
                    <p style={{marginTop: '5px'}}>Then use in cells:</p>
                    <code>=Revenue * TAX_RATE</code>
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

    return (
        <div className="info-panel">
            <h3>Cell: {selectedCell}</h3>

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

            {cell.style && (
                <div className="cell-detail-section">
                    <p><strong>Style:</strong></p>
                    <div className="style-preview" style={cell.style}>
                        {Object.entries(cell.style).map(([key, value]) => (
                            <div key={key} className="style-property">
                                <strong>{key}:</strong> {value}
                            </div>
                        ))}
                    </div>
                </div>
            )}

            {cell.styleExpression && (
                <div className="cell-detail-section">
                    <p><strong>Style Expression:</strong></p>
                    <code className="format-display">{cell.styleExpression}</code>
                </div>
            )}

            <div className="cell-detail-section view-mode-indicator">
                <p><strong>View Mode:</strong> {viewMode.toUpperCase()}</p>
                <p className="help-text">Press V/E/F/N to change view</p>
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
    const [sheetName, setSheetName] = useState('Sheet1');
    const [updateCounter, setUpdateCounter] = useState(0);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState(null);
    const [settingsOpen, setSettingsOpen] = useState(false);
    const [viewMode, setViewMode] = useState('normal'); // 'normal', 'values', 'expressions', 'formats'
    const [clipboard, setClipboard] = useState(null); // { content, sourceCol, sourceRow, isCut }

    const visibleRows = 20;
    const visibleCols = 10;

    // Initialize on mount
    useEffect(() => {
        initializeSpreadsheet();
    }, []);

    // Keyboard handler for view mode switching and clipboard operations
    useEffect(() => {
        const handleKeyDown = (e) => {
            // Only handle if not in an input/textarea
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
                return;
            }

            // Clipboard shortcuts (Ctrl+C, Ctrl+X, Ctrl+V)
            if (e.ctrlKey || e.metaKey) {
                if (e.key === 'c') {
                    e.preventDefault();
                    handleCopy();
                } else if (e.key === 'x') {
                    e.preventDefault();
                    handleCut();
                } else if (e.key === 'v') {
                    e.preventDefault();
                    handlePaste();
                }
                return;
            }

            // View mode switching (plain keys without modifiers)
            const key = e.key.toLowerCase();
            if (key === 'v') {
                setViewMode('values');
            } else if (key === 'e') {
                setViewMode('expressions');
            } else if (key === 'f') {
                setViewMode('formats');
            } else if (key === 'n') {
                setViewMode('normal');
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [selectedCell, model, clipboard]);

    // Handle hash changes for sheet name
    useEffect(() => {
        const handleHashChange = () => {
            const hash = window.location.hash.substring(1);
            if (hash) {
                setSheetName(hash);
            }
        };

        window.addEventListener('hashchange', handleHashChange);
        handleHashChange(); // Initial load

        return () => window.removeEventListener('hashchange', handleHashChange);
    }, []);

    const initializeSpreadsheet = async () => {
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

            // Create model
            const newModel = new SpreadsheetModel(100, 26);

            // Create adapter
            const newAdapter = new SpreadsheetRexxAdapter(newModel);

            // Initialize RexxJS interpreter
            await newAdapter.initializeInterpreter(RexxInterpreter);
            newAdapter.installSpreadsheetFunctions();

            // Check for data to load from hash parameter
            const hash = window.location.hash.substring(1);
            const loadSpec = SpreadsheetLoader.parseHashParameter(hash);

            let loadedData = false;
            if (loadSpec.type === 'url') {
                try {
                    console.log('Loading spreadsheet from:', loadSpec.value);
                    const data = await SpreadsheetLoader.loadFromURL(loadSpec.value);
                    await SpreadsheetLoader.importIntoModel(newModel, data, newAdapter);
                    setSheetName(data.name || 'Sheet1');
                    console.log('Spreadsheet loaded successfully');
                    loadedData = true;
                } catch (error) {
                    console.error('Failed to load spreadsheet:', error);
                    setError(`Failed to load spreadsheet: ${error.message}`);
                }
            }

            // Execute setup script if present
            const setupScript = newModel.getSetupScript();
            if (setupScript) {
                await newAdapter.executeSetupScript(setupScript);
            }

            setModel(newModel);
            setAdapter(newAdapter);
            setIsLoading(false);

            // Load sample data only if no data was loaded from URL
            if (!loadedData) {
                loadSampleData(newModel, newAdapter);
            }
        } catch (err) {
            console.error('Failed to initialize spreadsheet:', err);
            setError(err.message);
            setIsLoading(false);
        }
    };

    const loadSampleData = async (model, adapter) => {
        // Sample data to demonstrate features
        const sampleData = {
            'A1': '10',
            'A2': '20',
            'A3': '=A1 + A2',
            'B1': 'Hello',
            'B2': '=UPPER(B1)',
            'C1': '5',
            'C2': '3',
            'C3': '=C1 * C2'
        };

        for (const [ref, content] of Object.entries(sampleData)) {
            await model.setCell(ref, content, adapter);
        }

        setUpdateCounter(c => c + 1);
    };

    const handleEditCell = useCallback(async (cellRef, content) => {
        if (!model || !adapter) return;

        await model.setCell(cellRef, content, adapter);
        setUpdateCounter(c => c + 1);
    }, [model, adapter]);

    const handleSelectCell = useCallback((cellRef) => {
        setSelectedCell(cellRef);
    }, []);

    // Helper to adjust cell references in formulas
    const adjustCellReferences = (expression, offsetCol, offsetRow) => {
        // Match cell references like A1, B2, AA10 (case insensitive)
        const cellRefPattern = /\b([A-Z]+)(\d+)\b/gi;
        return expression.replace(cellRefPattern, (match, col, row) => {
            const colNum = SpreadsheetModel.colLetterToNumber(col);
            const rowNum = parseInt(row, 10);
            const newColNum = colNum + offsetCol;
            const newRowNum = rowNum + offsetRow;
            if (newColNum < 1 || newRowNum < 1) {
                return match; // Don't adjust if would go out of bounds
            }
            return SpreadsheetModel.formatCellRef(newColNum, newRowNum);
        });
    };

    // Copy handler
    const handleCopy = useCallback(() => {
        if (!selectedCell || !model) return;

        const cell = model.getCell(selectedCell);
        const content = cell.expression ? '=' + cell.expression : cell.value;
        const parsed = SpreadsheetModel.parseCellRef(selectedCell);

        setClipboard({
            content: content,
            sourceCol: SpreadsheetModel.colLetterToNumber(parsed.col),
            sourceRow: parsed.row,
            isCut: false
        });
    }, [selectedCell, model]);

    // Cut handler
    const handleCut = useCallback(() => {
        if (!selectedCell || !model) return;

        const cell = model.getCell(selectedCell);
        const content = cell.expression ? '=' + cell.expression : cell.value;
        const parsed = SpreadsheetModel.parseCellRef(selectedCell);

        setClipboard({
            content: content,
            sourceCol: SpreadsheetModel.colLetterToNumber(parsed.col),
            sourceRow: parsed.row,
            isCut: true
        });

        // Clear the source cell for cut operation
        model.setCell(selectedCell, '', adapter);
        setUpdateCounter(c => c + 1);
    }, [selectedCell, model, adapter]);

    // Paste handler with relative reference adjustment
    const handlePaste = useCallback(async () => {
        if (!selectedCell || !model || !clipboard) return;

        const parsed = SpreadsheetModel.parseCellRef(selectedCell);
        const targetCol = SpreadsheetModel.colLetterToNumber(parsed.col);
        const targetRow = parsed.row;

        let content = clipboard.content;

        // If pasting a formula, adjust cell references
        if (content.startsWith('=')) {
            const offsetCol = targetCol - clipboard.sourceCol;
            const offsetRow = targetRow - clipboard.sourceRow;
            const expression = content.substring(1);
            const adjustedExpression = adjustCellReferences(expression, offsetCol, offsetRow);
            content = '=' + adjustedExpression;
        }

        await model.setCell(selectedCell, content, adapter);
        setUpdateCounter(c => c + 1);

        // Clear clipboard if it was a cut operation
        if (clipboard.isCut) {
            setClipboard(null);
        }
    }, [selectedCell, model, adapter, clipboard]);

    // Column resize handler
    const handleColumnResize = useCallback((col, width) => {
        if (!model) return;
        model.setColumnWidth(col, width);
        setUpdateCounter(c => c + 1);
    }, [model]);

    // Row resize handler
    const handleRowResize = useCallback((row, height) => {
        if (!model) return;
        model.setRowHeight(row, height);
        setUpdateCounter(c => c + 1);
    }, [model]);

    // Apply number format to selected cell
    const handleApplyFormat = useCallback((format) => {
        if (!selectedCell || !model) return;

        model.setCellMetadata(selectedCell, { format });
        setUpdateCounter(c => c + 1);
    }, [selectedCell, model]);

    // Apply style to selected cell
    const handleApplyStyle = useCallback((style) => {
        if (!selectedCell || !model) return;

        model.setCellMetadata(selectedCell, { style });
        setUpdateCounter(c => c + 1);
    }, [selectedCell, model]);

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
                    <div className="view-mode-badge" title="Press V/E/F/N to change view">
                        View: {viewMode.toUpperCase()}
                    </div>
                    <button className="settings-button" onClick={() => setSettingsOpen(true)}>
                        ⚙️ Setup
                    </button>
                    <div className="sheet-name">Sheet: {sheetName}</div>
                </div>
            </div>

            <FormattingToolbar
                selectedCell={selectedCell}
                model={model}
                onApplyFormat={handleApplyFormat}
                onApplyStyle={handleApplyStyle}
            />

            <FormulaBar
                selectedCell={selectedCell}
                model={model}
                onEdit={handleEditCell}
            />

            <div className="main-content">
                <Grid
                    model={model}
                    selectedCell={selectedCell}
                    onSelectCell={handleSelectCell}
                    onEditCell={handleEditCell}
                    visibleRows={visibleRows}
                    visibleCols={visibleCols}
                    viewMode={viewMode}
                    onColumnResize={handleColumnResize}
                    onRowResize={handleRowResize}
                />

                <InfoPanel
                    selectedCell={selectedCell}
                    model={model}
                    viewMode={viewMode}
                />
            </div>

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

// Render the app when DOM and RexxJS are ready
if (typeof document !== 'undefined') {
    function renderApp() {
        // Check if RexxJS is available (need both RexxInterpreter and parse)
        if ((typeof RexxInterpreter !== 'undefined' && typeof parse !== 'undefined') || typeof RexxWebLoader !== 'undefined') {
            const root = ReactDOM.createRoot(document.getElementById('root'));
            root.render(<App />);
        } else {
            // RexxJS not ready yet, wait a bit and try again
            setTimeout(renderApp, 100);
        }
    }

    window.addEventListener('DOMContentLoaded', renderApp);
}
