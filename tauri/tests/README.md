# Spreadsheet POC Tests

This directory contains tests for the spreadsheet proof-of-concept application.

## Test Types

### Unit Tests (Jest)
- `spreadsheet-model.spec.js` - SpreadsheetModel unit tests
- `spreadsheet-functions.spec.js` - Function library unit tests
- `spreadsheet-poc.spec.js` - Main application unit tests

Run with:
```bash
npm test
```

### Integration Tests (REXX via ADDRESS)
- `spreadsheet-functions-test.rexx` - Tests spreadsheet range functions via control bus

These tests require the spreadsheet to be running with the control bus enabled.

#### Running Integration Tests

**Terminal 1 - Start the spreadsheet:**
```bash
cd examples/spreadsheet-poc
./rexxsheet-dev --control-bus
```

Wait for the spreadsheet window to open and show "Control Bus HTTP API enabled on port 2410" in the terminal.

> **Note**: Control bus communication logs are at DEBUG level, so you won't see verbose output during normal operation. If you need to debug the control bus, use `RUST_LOG=debug ./rexxsheet-dev --control-bus`.
>
> **Port**: The control bus uses port 2410 by default (see [port-calculator](https://paul-hammant.github.io/port-calculator/#rexxsheet)). You can override this with `REXXSHEET_CONTROL_BUS_PORT=<port>` environment variable.

**Terminal 2 - Run the test:**
```bash
cd examples/spreadsheet-poc
./tests/spreadsheet-functions-test.rexx
```

#### What the Integration Test Does

The `spreadsheet-functions-test.rexx` test:
1. Connects to the running spreadsheet via `ADDRESS SPREADSHEET`
2. Sets cell values using `SETCELL()`
3. Creates formulas using range functions like `SUM_RANGE(A1:A5)`
4. Retrieves calculated results using `GETCELL()`
5. Validates results using `ADDRESS EXPECTATIONS`

This validates that:
- The control bus communication works correctly
- Formulas are evaluated properly in the browser
- Range functions (SUM_RANGE, AVERAGE_RANGE, MIN_RANGE, MAX_RANGE, MEDIAN_RANGE, SUMIF_RANGE, COUNTIF_RANGE) produce correct results

### Web Tests (Playwright)
- `spreadsheet-loader.spec.js` - Web loader tests
- `spreadsheet-loader-web.spec.js` - Browser-based loader tests
- `spreadsheet-controlbus-web.spec.js` - Control bus web tests

Run with:
```bash
npm run test:web
```

## Helper Scripts

- `spreadsheet-integration-helper.sh` - Manages HTTP server for integration tests

## Control Bus Architecture

The integration tests use a COMET-style control bus inspired by Selenium-RC (2007):

1. Node.js REXX script → POST command to `/api/spreadsheet` (default port 2410)
2. Rust/Tauri server → Queues command, waits for browser to poll
3. Browser → Polls `/api/poll` continuously (100ms interval)
4. Browser → Executes command in isolated REXX interpreter
5. Browser → POSTs result to `/api/result`
6. Rust/Tauri → Returns result to Node.js caller
7. REXX script → `RESULT` variable contains the return value

This allows Node.js REXX scripts to control the spreadsheet running in the browser, similar to how ARexx worked on Amiga systems.

The default port (2410) is registered at [paul-hammant.github.io/port-calculator](https://paul-hammant.github.io/port-calculator/#rexxsheet).
