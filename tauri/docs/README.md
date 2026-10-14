# Documentation

This directory contains visual documentation for the RexxJS Spreadsheet, including feature guides with screenshots.

## Files

- **[FEATURES.md](FEATURES.md)** - Visual feature guide with screenshots demonstrating all major functionality
- **screenshots/** - Directory containing all documentation screenshots
- **screenshots/manifest.json** - Metadata about each screenshot (generated automatically)

## Viewing the Documentation

Simply open [FEATURES.md](FEATURES.md) to see the complete visual guide with examples and screenshots.

## Generating Screenshots

The screenshots in this documentation are automatically generated from Playwright tests.

### Prerequisites

```bash
# Install dependencies (if not already installed)
npm install
```

### Generate Screenshots

```bash
# Run the screenshot generation test
npm run test:screenshots

# Or run manually with Playwright
TAKE_SCREENSHOT=1 npm run test:web -- spreadsheet-screenshots-web.spec.js
```

### What Happens

1. The test suite launches the spreadsheet in a browser
2. Tests interact with the application (entering values, applying formatting, etc.)
3. At key moments, screenshots are captured
4. Screenshots are saved to `docs/screenshots/` with descriptive names
5. A manifest file (`screenshots/manifest.json`) is generated with metadata

### Screenshot Organization

Screenshots are named using the pattern: `{category}-{number}-{step}.png`

**Categories:**
- `overview` - Overall application views
- `basic-editing` - Cell editing and selection
- `formulas` - Formula entry and calculation
- `rexxjs-functions` - RexxJS function demonstrations
- `range-functions` - Range function examples
- `formatting` - Number formatting options
- `styling` - Visual styling (colors, fonts, alignment)
- `interface` - UI components (formula bar, info panel)
- `examples` - Complete example spreadsheets

### Manifest File

The `screenshots/manifest.json` file contains metadata for each screenshot:

```json
{
  "generatedAt": "2025-11-15T10:30:00.000Z",
  "screenshots": [
    {
      "filename": "overview-01-initial-load.png",
      "category": "overview",
      "step": "initial-load",
      "description": "Spreadsheet interface on initial load with sample data...",
      "order": 1,
      "timestamp": "2025-11-15T10:30:05.123Z"
    }
  ]
}
```

## Updating Documentation

### Adding New Screenshots

1. Edit `tests/spreadsheet-screenshots-web.spec.js`
2. Add a new test case demonstrating the feature
3. Use `takeScreenshot()` at key moments
4. Regenerate screenshots: `npm run test:screenshots`
5. Update `FEATURES.md` to reference the new screenshots

### Modifying Existing Screenshots

1. Edit the corresponding test in `tests/spreadsheet-screenshots-web.spec.js`
2. Regenerate: `npm run test:screenshots`
3. Screenshots will be overwritten with updated versions

### Example: Adding a New Feature Screenshot

```javascript
test('My new feature', async ({ page }) => {
    // Setup
    const cell = page.locator('.grid-row').nth(1).locator('.cell').first();
    await cell.dblclick();
    await page.keyboard.type('=NEW_FUNCTION()');

    // Take screenshot
    await takeScreenshot(
        page,
        'my-category',           // Category for organization
        'my-feature',            // Short step name
        'Description of what this screenshot shows' // Full description
    );

    await page.keyboard.press('Enter');
    await page.waitForTimeout(500);

    // Take another screenshot showing the result
    await takeScreenshot(
        page,
        'my-category',
        'my-feature-result',
        'Result of my new feature'
    );
});
```

## Documentation Best Practices

### Screenshot Guidelines

1. **Clear State**: Ensure the application is in a clear, understandable state
2. **Wait for Rendering**: Use `page.waitForTimeout()` to ensure everything is rendered
3. **Descriptive Names**: Use clear, descriptive step names
4. **Logical Progression**: Capture screenshots in a logical sequence (setup → action → result)
5. **Consistent Styling**: Keep formatting consistent across screenshots

### Writing Documentation

1. **Show, Don't Just Tell**: Use screenshots to illustrate concepts
2. **Step-by-Step**: Break complex features into simple steps
3. **Provide Context**: Explain what the user is seeing and why it matters
4. **Include Code Examples**: Show formula examples that users can try
5. **Cross-Reference**: Link to other documentation for more details

## Troubleshooting

### Screenshots Not Generating

- **Check Environment Variable**: Ensure `TAKE_SCREENSHOT=1` is set
- **Check Permissions**: Ensure `docs/screenshots/` is writable
- **View Test Output**: Look for "📸 Screenshot saved" messages in test output

### Screenshots Look Wrong

- **Increase Timeouts**: Add more `waitForTimeout()` calls in tests
- **Check Viewport**: Playwright uses a specific viewport size
- **Clear Browser State**: Tests should start with a clean state

### Manifest Not Created

- **Check Initialization**: Ensure `initScreenshots()` is called in `beforeAll`
- **Check Finalization**: Ensure `finalizeScreenshots()` is called in `afterAll`

## Related Documentation

- [../README.md](../README.md) - Main project README
- [../FILE-LOADING.md](../FILE-LOADING.md) - File format documentation
- [../TESTING-CONTROL-BUS.md](../TESTING-CONTROL-BUS.md) - Control Bus API documentation
- [FEATURES.md](FEATURES.md) - This visual feature guide
