/**
 * Spreadsheet POC Playwright Tests
 *
 * Tests for the spreadsheet UI interactions and RexxJS integration.
 */

import { test, expect } from '@playwright/test';

test.describe('Spreadsheet POC', () => {
    test.beforeEach(async ({ page }) => {
        // Navigate to the spreadsheet
        await page.goto('/examples/spreadsheet-poc/index.html');

        // Wait for React to load and render
        await page.waitForSelector('.app:not(.loading)', { timeout: 10000 });

        // Wait for grid to be visible
        await page.waitForSelector('.grid');
    });

    test('should load and display the spreadsheet', async ({ page }) => {
        // Check header is visible
        const header = await page.textContent('h1');
        expect(header).toBe('RexxJS Spreadsheet POC');

        // Check grid is rendered
        const grid = await page.locator('.grid');
        await expect(grid).toBeVisible();

        // Check column headers
        const colA = await page.locator('.column-header').first();
        await expect(colA).toHaveText('A');
    });

    test('should handle sheet name in hash parameter', async ({ page }) => {
        // Navigate with custom sheet name
        await page.goto('/examples/spreadsheet-poc/index.html#TestSheet');
        await page.waitForSelector('.app:not(.loading)');

        // Check sheet name is displayed
        const sheetName = await page.textContent('.sheet-name');
        expect(sheetName).toContain('TestSheet');
    });

    test('should allow entering literal values in cells', async ({ page }) => {
        // Find cell A1 and double-click to edit
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.dblclick();

        // Type a value
        await page.keyboard.type('42');
        await page.keyboard.press('Enter');

        // Wait a moment for update
        await page.waitForTimeout(200);

        // Verify cell shows the value
        const cellValue = await cellA1.textContent();
        expect(cellValue.trim()).toBe('42');
    });

    test('should evaluate simple formulas', async ({ page }) => {
        // Set value in A1
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.dblclick();
        await page.keyboard.type('10');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Set value in A2
        const cellA2 = page.locator('.grid-row').nth(2).locator('.cell').first();
        await cellA2.dblclick();
        await page.keyboard.type('20');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Set formula in A3
        const cellA3 = page.locator('.grid-row').nth(3).locator('.cell').first();
        await cellA3.dblclick();
        await page.keyboard.type('=A1 + A2');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Verify result
        const cellValue = await cellA3.textContent();
        expect(cellValue.trim()).toBe('30');
    });

    test('should update formulas when dependencies change', async ({ page }) => {
        // Set A1 = 5
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.dblclick();
        await page.keyboard.type('5');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Set A2 = A1 * 2
        const cellA2 = page.locator('.grid-row').nth(2).locator('.cell').first();
        await cellA2.dblclick();
        await page.keyboard.type('=A1 * 2');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Verify A2 = 10
        let cellValue = await cellA2.textContent();
        expect(cellValue.trim()).toBe('10');

        // Change A1 to 7
        await cellA1.dblclick();
        await page.keyboard.press('Control+A');
        await page.keyboard.type('7');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Verify A2 updated to 14
        cellValue = await cellA2.textContent();
        expect(cellValue.trim()).toBe('14');
    });

    test('should support RexxJS string functions', async ({ page }) => {
        // Set cell with UPPER function
        const cellB1 = page.locator('.grid-row').nth(1).locator('.cell').nth(1);
        await cellB1.dblclick();
        await page.keyboard.type('=UPPER("hello")');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Verify result
        const cellValue = await cellB1.textContent();
        expect(cellValue.trim()).toBe('HELLO');
    });

    test('should support function pipelines', async ({ page }) => {
        // Set cell with pipeline
        const cellC1 = page.locator('.grid-row').nth(1).locator('.cell').nth(2);
        await cellC1.dblclick();
        await page.keyboard.type('="hello" |> UPPER |> LENGTH');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Verify result (length of "HELLO" is 5)
        const cellValue = await cellC1.textContent();
        expect(cellValue.trim()).toBe('5');
    });

    test('should use formula bar for editing', async ({ page }) => {
        // Click on cell A1
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.click();
        await page.waitForTimeout(100);

        // Check formula bar shows cell reference
        const cellRefLabel = await page.textContent('.cell-ref-label');
        expect(cellRefLabel).toContain('A1');

        // Type in formula bar
        const formulaInput = page.locator('.formula-input');
        await formulaInput.fill('99');
        await formulaInput.press('Enter');
        await page.waitForTimeout(200);

        // Verify cell updated
        const cellValue = await cellA1.textContent();
        expect(cellValue.trim()).toBe('99');
    });

    test('should show formula in title attribute', async ({ page }) => {
        // Set a formula
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.dblclick();
        await page.keyboard.type('=10 + 20');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(500);

        // Check cell has formula class
        await expect(cellA1).toHaveClass(/formula/);

        // Check title attribute shows formula
        const title = await cellA1.getAttribute('title');
        expect(title).toContain('10 + 20');
    });

    test('should handle cell selection', async ({ page }) => {
        // Click on cell B2
        const cellB2 = page.locator('.grid-row').nth(2).locator('.cell').nth(1);
        await cellB2.click();
        await page.waitForTimeout(100);

        // Check cell is selected
        await expect(cellB2).toHaveClass(/selected/);

        // Check formula bar shows B2
        const cellRefLabel = await page.textContent('.cell-ref-label');
        expect(cellRefLabel).toContain('B2');

        // Click on different cell
        const cellC3 = page.locator('.grid-row').nth(3).locator('.cell').nth(2);
        await cellC3.click();
        await page.waitForTimeout(100);

        // Check B2 is no longer selected
        await expect(cellB2).not.toHaveClass(/selected/);

        // Check C3 is selected
        await expect(cellC3).toHaveClass(/selected/);

        // Check formula bar shows C3
        const newCellRefLabel = await page.textContent('.cell-ref-label');
        expect(newCellRefLabel).toContain('C3');
    });

    test('should display sample data on load', async ({ page }) => {
        // Check that sample data is loaded
        // A1 should be 10
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        const a1Value = await cellA1.textContent();
        expect(a1Value.trim()).toBe('10');

        // A2 should be 20
        const cellA2 = page.locator('.grid-row').nth(2).locator('.cell').first();
        const a2Value = await cellA2.textContent();
        expect(a2Value.trim()).toBe('20');

        // A3 should be 30 (=A1 + A2)
        const cellA3 = page.locator('.grid-row').nth(3).locator('.cell').first();
        const a3Value = await cellA3.textContent();
        expect(a3Value.trim()).toBe('30');
    });

    test('should clear cells with empty input', async ({ page }) => {
        // Set a value
        const cellA1 = page.locator('.grid-row').nth(1).locator('.cell').first();
        await cellA1.dblclick();
        await page.keyboard.type('Test');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Verify value is set
        let cellValue = await cellA1.textContent();
        expect(cellValue.trim()).toBe('Test');

        // Clear the cell
        await cellA1.dblclick();
        await page.keyboard.press('Control+A');
        await page.keyboard.press('Backspace');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Verify cell is empty
        cellValue = await cellA1.textContent();
        expect(cellValue.trim()).toBe('');
    });

    test('should apply number formatting from context menu', async ({ page }) => {
        // Set a numeric value
        const cellD1 = page.locator('.grid-row').nth(1).locator('.cell').nth(3);
        await cellD1.dblclick();
        await page.keyboard.type('123.456789');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Right-click to open context menu
        await cellD1.click({ button: 'right' });
        await page.waitForTimeout(100);

        // Navigate to Number Format submenu
        await page.hover('.context-menu-item:has-text("Number Format")');
        await page.waitForTimeout(100);

        // Click on "2 Decimals (0.00)"
        await page.click('.context-submenu .context-menu-item:has-text("2 Decimals")');
        await page.waitForTimeout(200);

        // Verify cell displays formatted value
        const cellValue = await cellD1.textContent();
        expect(cellValue.trim()).toBe('123.46');
    });

    test('should apply currency formatting from context menu', async ({ page }) => {
        // Set a numeric value
        const cellD2 = page.locator('.grid-row').nth(2).locator('.cell').nth(3);
        await cellD2.dblclick();
        await page.keyboard.type('1234.56');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Right-click to open context menu
        await cellD2.click({ button: 'right' });
        await page.waitForTimeout(100);

        // Navigate to Number Format submenu
        await page.hover('.context-menu-item:has-text("Number Format")');
        await page.waitForTimeout(100);

        // Click on "Currency (USD)"
        await page.click('.context-submenu .context-menu-item:has-text("Currency (USD)")');
        await page.waitForTimeout(200);

        // Verify cell displays formatted value (with $ symbol)
        const cellValue = await cellD2.textContent();
        expect(cellValue.trim()).toContain('$');
        expect(cellValue.trim()).toContain('1,234.56');
    });

    test('should apply percentage formatting from context menu', async ({ page }) => {
        // Set a numeric value
        const cellD3 = page.locator('.grid-row').nth(3).locator('.cell').nth(3);
        await cellD3.dblclick();
        await page.keyboard.type('0.125');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Right-click to open context menu
        await cellD3.click({ button: 'right' });
        await page.waitForTimeout(100);

        // Navigate to Number Format submenu
        await page.hover('.context-menu-item:has-text("Number Format")');
        await page.waitForTimeout(100);

        // Click on "Percent (0.0%)"
        await page.click('.context-submenu .context-menu-item:has-text("Percent (0.0%)")');
        await page.waitForTimeout(200);

        // Verify cell displays formatted value
        const cellValue = await cellD3.textContent();
        expect(cellValue.trim()).toBe('12.5%');
    });

    test('should apply alignment formatting from context menu', async ({ page }) => {
        // Set a text value
        const cellE1 = page.locator('.grid-row').nth(1).locator('.cell').nth(4);
        await cellE1.dblclick();
        await page.keyboard.type('Center Me');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Right-click to open context menu
        await cellE1.click({ button: 'right' });
        await page.waitForTimeout(100);

        // Navigate to Align submenu
        await page.hover('.context-menu-item:has-text("Align")');
        await page.waitForTimeout(100);

        // Click on "Align Center"
        await page.click('.context-submenu .context-menu-item:has-text("Align Center")');
        await page.waitForTimeout(200);

        // Verify cell has center alignment style
        const textAlign = await cellE1.evaluate(el => window.getComputedStyle(el).textAlign);
        expect(textAlign).toBe('center');
    });

    test('should apply combined formatting (bold + color + currency)', async ({ page }) => {
        // Set a numeric value
        const cellF1 = page.locator('.grid-row').nth(1).locator('.cell').nth(5);
        await cellF1.dblclick();
        await page.keyboard.type('999.99');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Apply bold
        await cellF1.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Bold")');
        await page.waitForTimeout(200);

        // Apply red color
        await cellF1.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Text: Red")');
        await page.waitForTimeout(200);

        // Apply currency format
        await cellF1.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Number Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Currency (USD)")');
        await page.waitForTimeout(200);

        // Verify all formats are applied
        const fontWeight = await cellF1.evaluate(el => window.getComputedStyle(el).fontWeight);
        expect(parseInt(fontWeight)).toBeGreaterThanOrEqual(700); // Bold

        const color = await cellF1.evaluate(el => window.getComputedStyle(el).color);
        expect(color).toContain('255, 0, 0'); // Red (rgb format)

        const cellValue = await cellF1.textContent();
        expect(cellValue.trim()).toContain('$');
        expect(cellValue.trim()).toContain('999.99');
    });

    test('should preserve formatting when updating cell value', async ({ page }) => {
        // Set a value and apply formatting
        const cellG1 = page.locator('.grid-row').nth(1).locator('.cell').nth(6);
        await cellG1.dblclick();
        await page.keyboard.type('100');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Apply bold format
        await cellG1.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Bold")');
        await page.waitForTimeout(200);

        // Change cell value
        await cellG1.dblclick();
        await page.keyboard.press('Control+A');
        await page.keyboard.type('200');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Verify formatting is still applied
        const fontWeight = await cellG1.evaluate(el => window.getComputedStyle(el).fontWeight);
        expect(parseInt(fontWeight)).toBeGreaterThanOrEqual(700); // Still bold

        const cellValue = await cellG1.textContent();
        expect(cellValue.trim()).toBe('200');
    });

    test('should clear formatting from context menu', async ({ page }) => {
        // Set a value and apply formatting
        const cellH1 = page.locator('.grid-row').nth(1).locator('.cell').nth(7);
        await cellH1.dblclick();
        await page.keyboard.type('Test');
        await page.keyboard.press('Enter');
        await page.waitForTimeout(200);

        // Apply italic format
        await cellH1.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Italic")');
        await page.waitForTimeout(200);

        // Verify italic is applied
        let fontStyle = await cellH1.evaluate(el => window.getComputedStyle(el).fontStyle);
        expect(fontStyle).toBe('italic');

        // Clear formatting
        await cellH1.click({ button: 'right' });
        await page.waitForTimeout(100);
        await page.hover('.context-menu-item:has-text("Format")');
        await page.waitForTimeout(100);
        await page.click('.context-submenu .context-menu-item:has-text("Clear Format")');
        await page.waitForTimeout(200);

        // Verify formatting is cleared
        fontStyle = await cellH1.evaluate(el => window.getComputedStyle(el).fontStyle);
        expect(fontStyle).toBe('normal');
    });
});
