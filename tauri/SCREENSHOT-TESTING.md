# Screenshot Testing Mode

This document describes the screenshot testing feature that allows tests to generate documentation screenshots automatically.

## Overview

The spreadsheet test suite includes a special screenshot mode that captures images at key moments during tests. These screenshots are used in the visual documentation at [docs/FEATURES.md](docs/FEATURES.md).

## Quick Start

### Generating Screenshots

```bash
# Run tests with screenshot mode enabled
npm run test:screenshots

# Or manually with environment variable
TAKE_SCREENSHOT=1 npm run test:web -- spreadsheet-screenshots-web.spec.js
```

### Output

- **Screenshots**: Saved to `docs/screenshots/` with descriptive filenames
- **Manifest**: `docs/screenshots/manifest.json` contains metadata about each screenshot
- **Console Output**: Shows progress with "📸 Screenshot saved" messages

## How It Works

### 1. Screenshot Helper (`tests/screenshot-helper.js`)

The screenshot helper provides:
- `initScreenshots()` - Initialize screenshot system (call in `beforeAll`)
- `takeScreenshot(page, category, step, description, options)` - Capture a screenshot
- `finalizeScreenshots()` - Finalize and save manifest (call in `afterAll`)
- `isScreenshotMode()` - Check if screenshot mode is enabled

### 2. Environment Variable

Screenshots are only captured when `TAKE_SCREENSHOT=1` is set:

```bash
# Screenshots enabled
TAKE_SCREENSHOT=1 npm run test:web

# Normal test run (no screenshots)
npm run test:web
```

This allows the same tests to run normally in CI/CD without generating screenshots.

### 3. Screenshot Test File (`tests/spreadsheet-screenshots-web.spec.js`)

A dedicated test file that:
- Demonstrates all major features
- Takes screenshots at key moments
- Generates images for documentation
- Can be run independently or with other tests

## Screenshot Organization

### Naming Convention

Screenshots follow the pattern: `{category}-{number}-{step}.png`

**Example:**
- `basic-editing-01-cell-selected.png`
- `formulas-02-formula-entry.png`
- `formatting-04-currency-applied.png`

### Categories

- **overview** - General application views
- **basic-editing** - Cell selection and editing
- **formulas** - Formula entry and calculation
- **rexxjs-functions** - RexxJS function demonstrations
- **range-functions** - Range function examples
- **formatting** - Number formatting
- **styling** - Visual styling (colors, fonts, alignment)
- **interface** - UI components
- **examples** - Complete examples

### Manifest File

The manifest file (`docs/screenshots/manifest.json`) contains metadata:

```json
{
  "generatedAt": "2025-11-15T10:30:00.000Z",
  "screenshots": [
    {
      "filename": "overview-01-initial-load.png",
      "category": "overview",
      "step": "initial-load",
      "description": "Spreadsheet interface on initial load with sample data showing formulas and calculated values",
      "order": 1,
      "timestamp": "2025-11-15T10:30:05.123Z"
    }
  ]
}
```

## Writing Screenshot Tests

### Basic Example

```javascript
import { test } from '@playwright/test';
import { takeScreenshot, initScreenshots, finalizeScreenshots } from './screenshot-helper.js';

test.describe('My Feature Screenshots', () => {
    test.beforeAll(async () => {
        initScreenshots();
    });

    test.afterAll(async () => {
        finalizeScreenshots();
    });

    test.beforeEach(async ({ page }) => {
        await page.goto('/examples/spreadsheet-poc/index.html');
        await page.waitForSelector('.grid');
    });

    test('Feature demonstration', async ({ page }) => {
        // Setup
        const cell = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cell.click();

        // Take screenshot of initial state
        await takeScreenshot(
            page,
            'my-category',
            'initial-state',
            'Description of what this screenshot shows'
        );

        // Perform action
        await cell.dblclick();
        await page.keyboard.type('42');

        // Take screenshot of action
        await takeScreenshot(
            page,
            'my-category',
            'action',
            'Description of the action being performed'
        );

        // Confirm
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Take screenshot of result
        await takeScreenshot(
            page,
            'my-category',
            'result',
            'Description of the final result'
        );
    });
});
```

### Screenshot Options

```javascript
// Full page screenshot (default)
await takeScreenshot(page, 'category', 'step', 'Description');

// Element-specific screenshot
await takeScreenshot(
    page,
    'category',
    'step',
    'Description',
    { selector: '.formula-bar' }
);

// Viewport-only screenshot
await takeScreenshot(
    page,
    'category',
    'step',
    'Description',
    { fullPage: false }
);
```

## Best Practices

### 1. Clear State

Ensure the application is in a clear, understandable state:

```javascript
// ✅ Good - wait for rendering
await page.keyboard.press('Enter');
await page.waitForTimeout(500);
await takeScreenshot(...);

// ❌ Bad - might capture mid-animation
await page.keyboard.press('Enter');
await takeScreenshot(...);
```

### 2. Descriptive Names

Use clear, descriptive step names and descriptions:

```javascript
// ✅ Good
await takeScreenshot(
    page,
    'formulas',
    'formula-result',
    'Cell A3 displays the calculated result (150) from the formula'
);

// ❌ Bad
await takeScreenshot(page, 'test', 'step1', 'Screenshot');
```

### 3. Logical Progression

Capture screenshots in a logical sequence:

1. **Setup** - Show initial state
2. **Action** - Show the action being performed
3. **Result** - Show the outcome

### 4. Wait for Rendering

Always wait for visual changes to complete:

```javascript
// Apply formatting
await page.click('.context-menu-item:has-text("Bold")');
await page.waitForTimeout(200); // Wait for style to apply

await takeScreenshot(...);
```

### 5. Consistent Timing

Use consistent wait times for similar operations:

```javascript
// Cell entry waits
await page.keyboard.press('Enter');
await page.waitForTimeout(200);

// Formula calculation waits
await page.keyboard.press('Enter');
await page.waitForTimeout(500);
```

## Integrating with Documentation

### 1. Update FEATURES.md

Reference screenshots using relative paths:

```markdown
### Currency Formatting

Apply currency formatting to display values with dollar signs.

![Currency formatting applied](screenshots/formatting-04-currency-applied.png)

**Result:** 1234.56 → $1,234.56
```

### 2. Keep Documentation Synchronized

When adding new features:

1. Add test cases to `tests/spreadsheet-screenshots-web.spec.js`
2. Run `npm run test:screenshots`
3. Update `docs/FEATURES.md` with new screenshots
4. Commit both code and screenshots

## Troubleshooting

### Screenshots Not Generated

**Check environment variable:**
```bash
# Should see "📸 Screenshot mode enabled" message
TAKE_SCREENSHOT=1 npm run test:web
```

**Check directory permissions:**
```bash
ls -la docs/screenshots/
```

### Screenshots Look Wrong

**Increase wait times:**
```javascript
// Increase timeout for slow operations
await page.waitForTimeout(1000); // Instead of 200
```

**Check viewport size:**
```javascript
// Playwright uses a default viewport
// If needed, set custom size in test config
```

**Clear state between tests:**
```javascript
test.beforeEach(async ({ page }) => {
    // Navigate to fresh page each time
    await page.goto('/examples/spreadsheet-poc/index.html');
});
```

### Manifest Not Created

**Check initialization:**
```javascript
test.beforeAll(async () => {
    initScreenshots(); // Must be called
});
```

**Check finalization:**
```javascript
test.afterAll(async () => {
    finalizeScreenshots(); // Must be called
});
```

## CI/CD Integration

### Running Tests Without Screenshots

In CI/CD, run tests normally without screenshots:

```yaml
# GitHub Actions example
- name: Run tests
  run: npm run test:web
```

### Generating Screenshots in CI (Optional)

If you want to generate screenshots in CI:

```yaml
- name: Install Playwright browsers
  run: npx playwright install --with-deps

- name: Generate screenshots
  run: npm run test:screenshots

- name: Upload screenshots
  uses: actions/upload-artifact@v3
  with:
    name: screenshots
    path: docs/screenshots/
```

## System Requirements

### Browser Dependencies

The screenshot tests require a full browser environment. On Linux, install dependencies:

```bash
# Ubuntu/Debian
npx playwright install-deps

# Or manually
sudo apt-get install -y \
    libgtk-4-1 \
    libgraphene-1.0-0 \
    libwoff2dec1.0.2 \
    # ... other dependencies
```

### Headless Mode

Playwright runs in headless mode by default, which is suitable for screenshot generation.

## Performance

### Test Duration

Screenshot tests take longer due to:
- Browser startup
- Page rendering
- Screenshot capture (~100-200ms each)
- File I/O

**Typical timing:**
- Full screenshot test suite: ~30-60 seconds
- Individual test: ~2-5 seconds

### Optimization

**Run only screenshot tests:**
```bash
npm run test:screenshots
```

**Skip screenshots in development:**
```bash
npm run test:web
```

## Related Documentation

- [docs/FEATURES.md](docs/FEATURES.md) - Visual feature guide (uses these screenshots)
- [docs/README.md](docs/README.md) - Documentation overview
- [README.md](README.md) - Main project README
- [tests/spreadsheet-screenshots-web.spec.js](tests/spreadsheet-screenshots-web.spec.js) - Screenshot test implementation
- [tests/screenshot-helper.js](tests/screenshot-helper.js) - Screenshot utility functions
