/**
 * Spreadsheet Screenshot Tests
 *
 * This test suite captures screenshots of key features for documentation purposes.
 * Run with TAKE_SCREENSHOT=1 to generate screenshots in docs/screenshots/
 *
 * Usage:
 *   TAKE_SCREENSHOT=1 npm run test:web -- spreadsheet-screenshots-web.spec.js
 */

import { test, expect } from '@playwright/test';
import { takeScreenshot, initScreenshots, finalizeScreenshots } from './screenshot-helper.js';

test.describe('Spreadsheet Documentation Screenshots', () => {
    test.beforeAll(async () => {
        initScreenshots();
    });

    test.afterAll(async () => {
        finalizeScreenshots();
    });

    test.beforeEach(async ({ page }) => {
        // Navigate to the spreadsheet
        await page.goto('/examples/spreadsheet-poc/index.html');

        // Wait for React to load and render
        await page.waitForSelector('.app:not(.loading)', { timeout: 10000 });

        // Wait for grid to be visible
        await page.waitForSelector('.grid');
    });

    test('01. Initial spreadsheet with sample data', async ({ page }) => {
        // Wait for sample data to load
        await page.waitForTimeout(500);

        await takeScreenshot(
            page,
            'overview',
            'initial-load',
            'Spreadsheet interface on initial load with sample data showing formulas and calculated values'
        );

        // Verify sample data loaded
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        const a1Value = await cellA1.textContent();
        expect(a1Value.trim()).toBe('10');
    });

    test('02. Basic cell editing - entering values', async ({ page }) => {
        // Click on cell B1
        const cellB1 = page.locator('.grid-row').nth(1).locator('.cell').nth(1);
        await cellB1.click();
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'basic-editing',
            'cell-selected',
            'Cell B1 selected, showing the cell reference in the formula bar'
        );

        // Enter edit mode
        await cellB1.dblclick();
        await page.keyboard.type('42');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'basic-editing',
            'entering-value',
            'Editing cell B1 with value "42" being entered'
        );

        // Confirm entry
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'basic-editing',
            'value-entered',
            'Cell B1 now displays the value 42 after confirmation'
        );
    });

    test('03. Formulas - basic arithmetic', async ({ page }) => {
        // Set up values in A1 and A2
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.dblclick();
        await page.keyboard.type('100');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        const cellA2 = page.locator('.grid-row').nth(2).locator('.cell').first();
        await cellA2.dblclick();
        await page.keyboard.type('50');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'formulas',
            'values-setup',
            'Values entered in A1 (100) and A2 (50) ready for formula demonstration'
        );

        // Enter formula in A3
        const cellA3 = page.locator('.grid-row').nth(3).locator('.cell').first();
        await cellA3.dblclick();
        await page.keyboard.type('=A1 + A2');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'formulas',
            'formula-entry',
            'Formula "=A1 + A2" being entered in cell A3'
        );

        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Click on A3 to show formula in title
        await cellA3.click();
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'formulas',
            'formula-result',
            'Cell A3 displays the calculated result (150) from the formula'
        );
    });

    test('04. Formulas - cell dependency updates', async ({ page }) => {
        // Set up A1 = 10, A2 = A1 * 2
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.dblclick();
        await page.keyboard.type('10');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        const cellA2 = page.locator('.grid-row').nth(2).locator('.cell').first();
        await cellA2.dblclick();
        await page.keyboard.type('=A1 * 2');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        await takeScreenshot(
            page,
            'formulas',
            'dependency-initial',
            'Cell A2 shows 20 (A1 * 2) demonstrating cell dependencies'
        );

        // Change A1 value
        await cellA1.dblclick();
        await page.keyboard.press('Control+A');
        await page.keyboard.type('25');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        await takeScreenshot(
            page,
            'formulas',
            'dependency-updated',
            'Cell A2 automatically updates to 50 when A1 changes to 25'
        );
    });

    test('05. RexxJS functions - string manipulation', async ({ page }) => {
        // Test UPPER function
        const cellB1 = page.locator('.grid-row').nth(1).locator('.cell').nth(1);
        await cellB1.dblclick();
        await page.keyboard.type('=UPPER("hello world")');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Test LENGTH function
        const cellB2 = page.locator('.grid-row').nth(2).locator('.cell').nth(1);
        await cellB2.dblclick();
        await page.keyboard.type('=LENGTH("RexxJS")');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Test SUBSTR function
        const cellB3 = page.locator('.grid-row').nth(3).locator('.cell').nth(1);
        await cellB3.dblclick();
        await page.keyboard.type('=SUBSTR("Spreadsheet", 1, 6)');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        await takeScreenshot(
            page,
            'rexxjs-functions',
            'string-functions',
            'RexxJS string functions: UPPER, LENGTH, and SUBSTR demonstrated in cells B1-B3'
        );
    });

    test('06. RexxJS functions - function pipelines', async ({ page }) => {
        const cellC1 = page.locator('.grid-row').nth(1).locator('.cell').nth(2);
        await cellC1.dblclick();
        await page.keyboard.type('="hello" |> UPPER |> LENGTH');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'rexxjs-functions',
            'pipeline-entry',
            'Function pipeline being entered: "hello" |> UPPER |> LENGTH'
        );

        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        await takeScreenshot(
            page,
            'rexxjs-functions',
            'pipeline-result',
            'Pipeline result shows 5 (length of "HELLO") demonstrating chained transformations'
        );
    });

    test('07. Range functions - SUM_RANGE', async ({ page }) => {
        // Enter values A1-A5
        for (let i = 1; i <= 5; i++) {
            const cell = page.locator('.grid-row').nth(i).locator('.cell').first();
            await cell.dblclick();
            await page.keyboard.type(String(i * 10));
            await page.keyboard.press('Enter');
            await page.waitForTimeout(100);
        }

        await takeScreenshot(
            page,
            'range-functions',
            'values-setup',
            'Values 10, 20, 30, 40, 50 entered in cells A1-A5 for range function demonstration'
        );

        // Add SUM_RANGE formula in A6
        const cellA6 = page.locator('.grid-row').nth(6).locator('.cell').first();
        await cellA6.dblclick();
        await page.keyboard.type('=SUM_RANGE("A1:A5")');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        await takeScreenshot(
            page,
            'range-functions',
            'sum-result',
            'SUM_RANGE function in A6 calculates the sum (150) of cells A1-A5'
        );
    });

    test('08. Number formatting - currency', async ({ page }) => {
        // Enter a value
        const cellD1 = page.locator('.grid-row').nth(1).locator('.cell').nth(3);
        await cellD1.dblclick();
        await page.keyboard.type('1234.56');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'formatting',
            'unformatted-number',
            'Cell D1 shows raw number 1234.56 before formatting'
        );

        // Right-click and apply currency format
        await cellD1.click({ button: 'right' });
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'formatting',
            'context-menu',
            'Context menu opened on cell D1 showing formatting options'
        );

        await page.hover('.context-menu-item:has-text("Number Format")');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'formatting',
            'number-format-submenu',
            'Number Format submenu showing available format options (Currency, Percent, etc.)'
        );

        await page.click('.context-submenu .context-menu-item:has-text("Currency (USD)")');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'formatting',
            'currency-applied',
            'Cell D1 now displays as $1,234.56 with currency formatting applied'
        );
    });

    test('09. Number formatting - percentage', async ({ page }) => {
        // Enter a decimal value
        const cellD2 = page.locator('.grid-row').nth(2).locator('.cell').nth(3);
        await cellD2.dblclick();
        await page.keyboard.type('0.856');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Apply percentage format
        await cellD2.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Number Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Percent (0.0%)")');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'formatting',
            'percentage-applied',
            'Cell D2 displays 85.6% after percentage formatting is applied to 0.856'
        );
    });

    test('10. Cell styling - text color and bold', async ({ page }) => {
        // Enter a value
        const cellE1 = page.locator('.grid-row').nth(1).locator('.cell').nth(4);
        await cellE1.dblclick();
        await page.keyboard.type('Important');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Apply bold
        await cellE1.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Bold")');
        await page.waitForTimeout(200);

        // Apply red color
        await cellE1.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Text: Red")');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'styling',
            'bold-red-text',
            'Cell E1 displays "Important" in bold red text demonstrating cell styling'
        );
    });

    test('11. Cell styling - background colors', async ({ page }) => {
        // Create a header row with different background colors
        const colors = [
            { col: 0, label: 'Red', menu: 'BG: Red' },
            { col: 1, label: 'Green', menu: 'BG: Green' },
            { col: 2, label: 'Blue', menu: 'BG: Blue' },
            { col: 3, label: 'Yellow', menu: 'BG: Yellow' }
        ];

        for (const { col, label, menu } of colors) {
            const cell = page.locator('.grid-row').nth(1).locator('.cell').nth(col);
            await cell.dblclick();
            await page.keyboard.type(label);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(100);

            await cell.click({ button: 'right' });
            await page.waitForTimeout(100);
            await page.hover('.context-menu-item:has-text("Format")');
            await page.waitForTimeout(100);
            await page.click(`.context-submenu .context-menu-item:has-text("${menu}")`);
            await page.waitForTimeout(100);
        }

        await takeScreenshot(
            page,
            'styling',
            'background-colors',
            'Cells with different background colors (red, green, blue, yellow) demonstrating styling options'
        );
    });

    test('12. Cell alignment - center and right align', async ({ page }) => {
        // Left aligned (default)
        const cellF1 = page.locator('.grid-row').nth(1).locator('.cell').nth(5);
        await cellF1.dblclick();
        await page.keyboard.type('Left');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(100);

        // Center aligned
        const cellF2 = page.locator('.grid-row').nth(2).locator('.cell').nth(5);
        await cellF2.dblclick();
        await page.keyboard.type('Center');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(100);

        await cellF2.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Align")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Align Center")');
        await page.waitForTimeout(100);

        // Right aligned
        const cellF3 = page.locator('.grid-row').nth(3).locator('.cell').nth(5);
        await cellF3.dblclick();
        await page.keyboard.type('Right');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(100);

        await cellF3.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Align")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Align Right")');
        await page.waitForTimeout(100);

        await takeScreenshot(
            page,
            'styling',
            'text-alignment',
            'Cells F1-F3 demonstrating left, center, and right text alignment'
        );
    });

    test('13. Formula bar editing', async ({ page }) => {
        // Click on a cell
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.click();
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'interface',
            'formula-bar-empty',
            'Formula bar showing cell reference A1 when cell is selected'
        );

        // Type in formula bar
        const formulaInput = page.locator('.formula-input');
        await formulaInput.fill('=10 + 20 + 30');
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'interface',
            'formula-bar-editing',
            'Formula bar with formula "=10 + 20 + 30" being edited'
        );

        await formulaInput.press('Enter');
        await page.waitForTimeout(500);

        await cellA1.click();
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'interface',
            'formula-bar-result',
            'Cell A1 displays result 60, formula bar shows the formula when cell is selected'
        );
    });

    test('14. Complex spreadsheet - budget example', async ({ page }) => {
        // Create a simple budget spreadsheet
        const data = [
            { row: 1, col: 0, value: 'Item' },
            { row: 1, col: 1, value: 'Amount' },
            { row: 2, col: 0, value: 'Rent' },
            { row: 2, col: 1, value: '1200' },
            { row: 3, col: 0, value: 'Food' },
            { row: 3, col: 1, value: '400' },
            { row: 4, col: 0, value: 'Transport' },
            { row: 4, col: 1, value: '150' },
            { row: 5, col: 0, value: 'Total' },
            { row: 5, col: 1, value: '=SUM_RANGE("B2:B4")' }
        ];

        // Enter all data
        for (const { row, col, value } of data) {
            const cell = page.locator('.grid-row').nth(row).locator('.cell').nth(col);
            await cell.dblclick();
            await page.keyboard.type(value);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(50);
        }

        await page.waitForTimeout(300);

        // Format header row (bold)
        for (let col = 0; col <= 1; col++) {
            const cell = page.locator('.grid-row').nth(1).locator('.cell').nth(col);
            await cell.click({ button: 'right' });
            await page.waitForTimeout(50);
            await page.hover('.context-menu-item:has-text("Format")');
            await page.waitForTimeout(50);
            await page.click('.context-submenu .context-menu-item:has-text("Bold")');
            await page.waitForTimeout(50);
        }

        // Format currency column
        for (let row = 2; row <= 5; row++) {
            const cell = page.locator('.grid-row').nth(row).locator('.cell').nth(1);
            await cell.click({ button: 'right' });
            await page.waitForTimeout(50);
            await page.hover('.context-menu-item:has-text("Number Format")');
            await page.waitForTimeout(50);
            await page.click('.context-submenu .context-menu-item:has-text("Currency (USD)")');
            await page.waitForTimeout(50);
        }

        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'examples',
            'budget-spreadsheet',
            'Complete budget spreadsheet example with headers, values, formatting, and SUM_RANGE formula'
        );
    });

    test('15. Info panel and cell details', async ({ page }) => {
        // Create a cell with formula
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.dblclick();
        await page.keyboard.type('100');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(100);

        const cellA2 = page.locator('.grid-row').nth(2).locator('.cell').first();
        await cellA2.dblclick();
        await page.keyboard.type('=A1 * 1.5');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Click on A2 to show info
        await cellA2.click();
        await page.waitForTimeout(200);

        await takeScreenshot(
            page,
            'interface',
            'info-panel',
            'Info panel showing cell details including value, formula, and dependencies for cell A2'
        );
    });
});
