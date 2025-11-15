/**
 * Screenshot Helper for Documentation
 *
 * Provides utilities for capturing screenshots during tests when TAKE_SCREENSHOT=1
 * environment variable is set. Screenshots are saved to docs/screenshots/ and
 * metadata is tracked in manifest.json for documentation purposes.
 *
 * Usage:
 *   import { takeScreenshot, initScreenshots } from './screenshot-helper.js';
 *
 *   test('my test', async ({ page }) => {
 *     await initScreenshots();
 *     await takeScreenshot(page, 'feature-name', 'step-description', 'Detailed description');
 *   });
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Check if screenshot mode is enabled
const SCREENSHOT_MODE = process.env.TAKE_SCREENSHOT === '1';

// Screenshot directory
const SCREENSHOT_DIR = path.join(__dirname, '..', 'docs', 'screenshots');
const MANIFEST_PATH = path.join(SCREENSHOT_DIR, 'manifest.json');

// Screenshot counter for ordering
let screenshotCounter = 0;

// Manifest data
let manifest = {
  generatedAt: null,
  screenshots: []
};

/**
 * Initialize screenshot system
 * Creates screenshot directory and resets manifest
 */
export function initScreenshots() {
  if (!SCREENSHOT_MODE) return;

  // Create screenshot directory if it doesn't exist
  if (!fs.existsSync(SCREENSHOT_DIR)) {
    fs.mkdirSync(SCREENSHOT_DIR, { recursive: true });
  }

  // Reset manifest
  manifest = {
    generatedAt: new Date().toISOString(),
    screenshots: []
  };
  screenshotCounter = 0;

  console.log('📸 Screenshot mode enabled - screenshots will be saved to:', SCREENSHOT_DIR);
}

/**
 * Take a screenshot if TAKE_SCREENSHOT=1 is set
 *
 * @param {Page} page - Playwright page object
 * @param {string} category - Feature category (e.g., 'basic-editing', 'formulas', 'formatting')
 * @param {string} step - Short step description (e.g., 'enter-values', 'apply-formula')
 * @param {string} description - Full description for documentation
 * @param {object} options - Additional options (fullPage, selector)
 * @returns {Promise<string|null>} Path to screenshot or null if not taken
 */
export async function takeScreenshot(page, category, step, description, options = {}) {
  if (!SCREENSHOT_MODE) {
    return null;
  }

  screenshotCounter++;
  const counter = String(screenshotCounter).padStart(2, '0');
  const filename = `${category}-${counter}-${step}.png`;
  const filepath = path.join(SCREENSHOT_DIR, filename);

  // Default screenshot options
  const screenshotOptions = {
    path: filepath,
    fullPage: options.fullPage !== false, // Default to full page
  };

  // If a specific element is specified, screenshot only that element
  if (options.selector) {
    const element = await page.locator(options.selector);
    await element.screenshot({ path: filepath });
  } else {
    await page.screenshot(screenshotOptions);
  }

  // Add to manifest
  manifest.screenshots.push({
    filename,
    category,
    step,
    description,
    order: screenshotCounter,
    timestamp: new Date().toISOString()
  });

  // Save manifest
  fs.writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2));

  console.log(`📸 Screenshot saved: ${filename} - ${description}`);

  return filepath;
}

/**
 * Finalize screenshots and save final manifest
 */
export function finalizeScreenshots() {
  if (!SCREENSHOT_MODE) return;

  // Save final manifest
  fs.writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2));

  console.log(`📸 Screenshot capture complete: ${screenshotCounter} screenshots saved`);
}

/**
 * Check if screenshot mode is enabled
 */
export function isScreenshotMode() {
  return SCREENSHOT_MODE;
}
