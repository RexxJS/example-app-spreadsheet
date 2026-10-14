# Spreadsheet App - Standalone Setup Guide

This document describes how the spreadsheet POC has been extracted from `@RexxJS/examples/spreadsheet-poc` to be a standalone example app at `/home/paul/scm/rexxjs/example-app-spreadsheet`.

## Status

✅ **Fully Functional**
- Vite build: Working
- Unit tests: All 95 passing
- Ready for development and deployment

## What Was Changed

### 1. Directory Structure
```
BEFORE: RexxJS/examples/spreadsheet-poc/
AFTER:  /home/paul/scm/rexxjs/example-app-spreadsheet/
```

### 2. package.json Updates
- Renamed package from `spreadsheet-poc` to `example-app-spreadsheet`
- Updated `npm scripts` for standalone operation:
  - `npm run dev` - Vite dev server
  - `npm run build` - Vite production build
  - `npm run test` - Run all tests (unit + web)
  - `npm run test:unit` - Unit tests only
  - `npm run test:web` - Playwright web tests
  - `npm run tauri:dev:control-bus` - Run with control bus enabled

- Added dev dependencies:
  - Jest + babel-jest (unit testing)
  - Playwright (web testing)
  - ESLint, Prettier (code quality)

### 3. Build Configuration
- Updated `vite.config.js`:
  - Fixed __dirname for ES modules
  - Made RexxJS bundle copy optional (graceful degradation)
  - Changed strictPort to false (allows fallback ports)
  - Compatible with Node 18+

- Added `jest.config.js`:
  - Configured for ES modules
  - Proper babel transformation
  - JSDOM environment for DOM testing
  - Test path filtering (exclude web/POC tests from unit)

- Added `.babelrc`:
  - Babel preset-env
  - Babel preset-react (for JSX support)

- Added `playwright.config.js`:
  - Web test configuration
  - Local dev server integration
  - Multi-browser testing support

### 4. Import Path Fixes
- Updated test files to use ES imports:
  - `spreadsheet-functions.spec.js` - ES import syntax
  - `spreadsheet-model.spec.js` - ES import syntax
  - `spreadsheet-loader.spec.js` - ES import syntax

- Added missing import in source:
  - `src/spreadsheet-rexx-adapter.js` - Now imports SpreadsheetModel

### 5. Dependencies Compatibility
- Downgraded Vite from 7.2.0 → 4.5.0 (Node 18 compatibility)
- Downgraded @vitejs/plugin-react from 5.1.0 → 4.2.0 (Node 18 compatibility)
- All other dependencies compatible with Node 18+

## Development Workflow

### Installation
```bash
cd /home/paul/scm/rexxjs/example-app-spreadsheet
npm install
```

### Development Server
```bash
npm run dev
# Vite dev server starts on http://localhost:5173
```

### Building
```bash
npm run build
# Production build in ./dist/
# Note: RexxJS bundle copy is optional, gracefully skipped if not found
```

### Testing

#### Unit Tests (95 tests)
```bash
npm run test:unit
# Tests: spreadsheet-model, spreadsheet-functions, spreadsheet-loader
# All passing ✅
```

#### Web Tests (Playwright)
```bash
npm run test:web
# Runs Playwright tests on http://localhost:5173
# Requires dev server running or will start one automatically
```

#### All Tests
```bash
npm run test
# Runs unit tests AND web tests
```

### Control Bus (HTTP API)
```bash
npm run tauri:dev:control-bus
# Starts Tauri app with ENABLE_CONTROL_BUS=1
# HTTP API available on port 2410
```

## Test Results

### Unit Tests: 95 Passing ✅
```
Test Suites: 3 passed, 3 total
Tests:       95 passed, 95 total
```

**Test Files**:
1. `tests/spreadsheet-model.spec.js` - SpreadsheetModel core functionality
2. `tests/spreadsheet-functions.spec.js` - Range functions (SUM_RANGE, AVERAGE_RANGE, etc.)
3. `tests/spreadsheet-loader.spec.js` - File loading functionality

### Coverage Areas
- Cell storage and retrieval
- Range functions (statistical, conditional)
- Data validation
- Error handling
- File format parsing

## Scripts Reference

```json
{
  "dev": "vite",
  "build": "vite build",
  "preview": "vite preview",
  "test": "npm run test:unit && npm run test:web",
  "test:unit": "jest --testMatch='**/tests/**/*.spec.js' --testPathIgnore='web\\.spec\\.js'",
  "test:web": "playwright test",
  "test:web:debug": "playwright test --debug",
  "lint": "eslint src --ext .js,.jsx",
  "format": "prettier --write \"src/**/*.{js,jsx,css}\""
}
```

## File Structure

```
example-app-spreadsheet/
├── src/
│   ├── spreadsheet-model.js              Core spreadsheet logic
│   ├── spreadsheet-rexx-adapter.js       RexxJS integration
│   ├── spreadsheet-loader.js             File loading
│   ├── spreadsheet-address-handler.js    ADDRESS SPREADSHEET handler
│   ├── spreadsheet-control-functions.js  Control bus functions
│   └── main.jsx                          React entry point
│
├── tests/
│   ├── spreadsheet-model.spec.js         ✅ Unit tests
│   ├── spreadsheet-functions.spec.js     ✅ Range function tests
│   ├── spreadsheet-loader.spec.js        ✅ Loader tests
│   ├── spreadsheet-poc.spec.js           Playwright tests
│   └── spreadsheet-*-web.spec.js         Playwright web tests
│
├── public/
│   └── lib/spreadsheet-functions.js      Browser library
│
├── index.html                            HTML entry point
├── vite.config.js                        ✅ Updated for standalone
├── jest.config.js                        ✅ Added
├── playwright.config.js                  ✅ Added
├── .babelrc                              ✅ Added
└── package.json                          ✅ Updated
```

## Dependencies

### Production
- `@tauri-apps/api` - Tauri integration

### Development
- `react`, `react-dom` - UI framework
- `vite` - Build tool (v4.5.0 for Node 18)
- `@vitejs/plugin-react` - React plugin (v4.2.0 for Node 18)
- `jest`, `babel-jest` - Unit testing
- `@playwright/test` - Web testing
- `eslint`, `prettier` - Code quality

## Troubleshooting

### Issue: "RexxJS bundle not found"
**Status**: ✅ Expected in standalone mode
**Fix**: This is normal. The bundle copy is optional. The app works without it.

### Issue: "Cannot find module SpreadsheetModel"
**Status**: ✅ Fixed
**Fix**: Already fixed in this setup. All imports are correct.

### Issue: "Jest transform error"
**Status**: ✅ Fixed
**Fix**: Babel and Jest configs are properly set up.

### Issue: Vite crypto error (Node version)
**Status**: ✅ Fixed
**Fix**: Using Vite v4.5.0 instead of v7.2.0 for Node 18 compatibility.

## Integration with Control Bus

The spreadsheet supports RexxJS control via ADDRESS:

```rexx
/* In a RexxJS script */
ADDRESS SPREADSHEET "set-cell A1 value"
ADDRESS SPREADSHEET "get-cell A1"
SAY RESULT
```

Control bus HTTP endpoint: `http://127.0.0.1:2410`

See `/REMEDIATION_TARGETS.md` for adding to the unified control bus standard.

## Next Steps

### As a Full Example App
This spreadsheet should be brought into parity with other example apps:

1. **Add to remediation plan** (`/REMEDIATION_TARGETS.md`)
2. **Create SPREADSHEET_CONTROL_BUS.md** (architecture guide)
3. **Enhance SPREADSHEET_COMMANDS.md** (400+ lines)
4. **Add example automation scripts** (test-spreadsheet-automation.rexx)
5. **Expand test coverage** (web tests, integration tests)

### Development Work
1. Fix any remaining Playwright test issues
2. Add missing test coverage
3. Document remaining commands
4. Test with real RexxJS automation

## Related Documentation

- Original POC: `RexxJS/examples/spreadsheet-poc/`
- Control Bus Standard: `/REXXJS_CONTROL_BUS_STANDARD.md`
- Example Apps Guide: `/EXAMPLE_APPS_CONTROL_BUS_GUIDE.md`
- Remediation Plan: `/REMEDIATION_TARGETS.md`

## Status Summary

✅ **Extraction Complete**
- ✅ Copied to standalone location
- ✅ Updated package.json for standalone
- ✅ Fixed import paths
- ✅ Updated build configs
- ✅ Added test configs
- ✅ All 95 unit tests passing
- ✅ Build working

⏳ **Ready for**
- Vite development (`npm run dev`)
- Production build (`npm run build`)
- Unit test execution (`npm run test:unit`)
- Web test development (`npm run test:web`)
- Control bus work (phase 2+)

## Build Status

```bash
$ npm run build
vite v4.5.14 building for production...
✓ 42 modules transformed.
dist/index.html                   2.75 kB │ gzip:   1.18 kB
dist/assets/index-7c8b8f6f.css   12.34 kB │ gzip:   2.67 kB
dist/assets/index-7974f34f.js   712.99 kB │ gzip: 121.41 kB
✓ built in 1.56s
```

**Result**: ✅ BUILD SUCCESSFUL
