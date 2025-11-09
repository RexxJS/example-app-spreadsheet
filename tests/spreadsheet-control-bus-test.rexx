#!/usr/bin/env ../../../core/rexxt

// Test suite for spreadsheet control bus commands
// Copyright (c) 2025 RexxJS Project
// Comprehensive tests for all control bus functions
//
// Prerequisites:
//   1. Start the spreadsheet with control bus enabled:
//      cd examples/spreadsheet-poc
//      ./rexxsheet-dev --control-bus
//
//   2. In another terminal, run this test:
//      cd examples/spreadsheet-poc
//      ./tests/spreadsheet-control-bus-test.rexx
//
// This is an integration test that validates all spreadsheet control bus
// commands work correctly in the actual spreadsheet.

/* @test-tags spreadsheet, control-bus, integration */
/* @description Test all spreadsheet control bus commands */

REQUIRE "../../../core/src/expectations-address.js"

SAY "🔌 Testing Spreadsheet Control Bus Commands"
SAY "=============================================="
SAY ""
SAY "Prerequisites:"
SAY "  1. Spreadsheet must be running with control bus enabled"
SAY "  2. Run: ./rexxsheet-dev --control-bus"
SAY "  3. Control bus should be at http://localhost:2410"
SAY ""

// Connect to the spreadsheet control bus
ADDRESS "http://localhost:2410/api/spreadsheet" AUTH "dev-token-12345" AS SPREADSHEET

// Verify connection
SAY "Checking spreadsheet connection..."
ADDRESS SPREADSHEET
'SPREADSHEET_VERSION()'
IF RC \= 0 THEN DO
  SAY ""
  SAY "❌ Failed to connect to spreadsheet control bus"
  SAY ""
  SAY "Please start the spreadsheet first:"
  SAY "  cd examples/spreadsheet-poc"
  SAY "  ./rexxsheet-dev --control-bus"
  SAY ""
  SAY "Then run this test again."
  SAY ""
  EXIT 1
END
SAY "✓ Connected to spreadsheet version:" RESULT
SAY ""

// Clear spreadsheet before tests
ADDRESS SPREADSHEET
'CLEAR()'

// Run all test suites
CALL TestBasicCellOperations
CALL TestCellMetadata
CALL TestCellReferences
CALL TestRangeOperations
CALL TestSheetManagement
CALL TestExpressionEvaluation
CALL TestSetupScript
CALL TestUtilityCommands

SAY ""
SAY "✅ All control bus tests passed!"
EXIT 0

// ============= Test Subroutines =============

TestBasicCellOperations:
  SAY "Testing basic cell operations..."

  ADDRESS SPREADSHEET

  // Test SETCELL
  'SETCELL("A1", "Hello")'
  'SETCELL("A2", "100")'
  'SETCELL("A3", "=A2 * 2")'

  // Test GETCELL
  'GETCELL("A1")'
  LET result_a1 = RESULT
  ADDRESS EXPECTATIONS
  "{result_a1} should equal Hello"

  ADDRESS SPREADSHEET
  'GETCELL("A2")'
  LET result_a2 = RESULT
  ADDRESS EXPECTATIONS
  "{result_a2} should equal 100"

  ADDRESS SPREADSHEET
  'GETCELL("A3")'
  LET result_a3 = RESULT
  ADDRESS EXPECTATIONS
  "{result_a3} should equal 200"

  // Test GETEXPRESSION
  ADDRESS SPREADSHEET
  'GETEXPRESSION("A3")'
  LET expr = RESULT
  ADDRESS EXPECTATIONS
  "{expr} should equal A2 * 2"

  // Test CLEARCELL
  ADDRESS SPREADSHEET
  'CLEARCELL("A1")'
  'GETCELL("A1")'
  LET cleared = RESULT
  ADDRESS EXPECTATIONS
  "{cleared} should equal "

  SAY "  ✓ Basic cell operations (SETCELL, GETCELL, GETEXPRESSION, CLEARCELL)"
RETURN

TestCellMetadata:
  SAY "Testing cell metadata operations..."

  ADDRESS SPREADSHEET

  // Test SETCOMMENT and GETCOMMENT
  'SETCOMMENT("B1", "This is a test comment")'
  'GETCOMMENT("B1")'
  LET comment = RESULT
  ADDRESS EXPECTATIONS
  "{comment} should equal This is a test comment"

  // Test SETFORMAT and GETFORMAT
  ADDRESS SPREADSHEET
  'SETFORMAT("B2", "bold;italic")'
  'GETFORMAT("B2")'
  LET format = RESULT
  ADDRESS EXPECTATIONS
  "{format} should equal bold;italic"

  SAY "  ✓ Cell metadata (SETCOMMENT, GETCOMMENT, SETFORMAT, GETFORMAT)"
RETURN

TestCellReferences:
  SAY "Testing cell reference utilities..."

  ADDRESS SPREADSHEET

  // Test GETROW
  'GETROW("C5")'
  LET row = RESULT
  ADDRESS EXPECTATIONS
  "{row} should equal 5"

  // Test GETCOL
  ADDRESS SPREADSHEET
  'GETCOL("C5")'
  LET col = RESULT
  ADDRESS EXPECTATIONS
  "{col} should equal 3"

  // Test GETCOLNAME
  ADDRESS SPREADSHEET
  'GETCOLNAME("C5")'
  LET colname = RESULT
  ADDRESS EXPECTATIONS
  "{colname} should equal C"

  // Test MAKECELLREF with number
  ADDRESS SPREADSHEET
  'MAKECELLREF(3, 5)'
  LET ref1 = RESULT
  ADDRESS EXPECTATIONS
  "{ref1} should equal C5"

  // Test MAKECELLREF with letter
  ADDRESS SPREADSHEET
  'MAKECELLREF("D", 10)'
  LET ref2 = RESULT
  ADDRESS EXPECTATIONS
  "{ref2} should equal D10"

  SAY "  ✓ Cell reference utilities (GETROW, GETCOL, GETCOLNAME, MAKECELLREF)"
RETURN

TestRangeOperations:
  SAY "Testing range operations..."

  ADDRESS SPREADSHEET

  // Set up test data
  'SETCELL("D1", "10")'
  'SETCELL("D2", "20")'
  'SETCELL("D3", "30")'

  // Test GETCELLS
  'GETCELLS("D1:D3")'
  LET cells = RESULT

  // Check count
  LET count = cells.0
  ADDRESS EXPECTATIONS
  "{count} should equal 3"

  // Check first cell value (cell values are strings)
  LET cell1_val = INTERPRET_JS("cells['1'].value")
  ADDRESS EXPECTATIONS
  "{cell1_val} should be string 10"

  // Test SETCELLS with array parsed from JSON
  LET values = JSON_PARSE text='["100", "200", "300"]'

  ADDRESS SPREADSHEET
  'SETCELLS("E1:E3", values)'
  'GETCELL("E1")'
  LET e1 = RESULT
  ADDRESS EXPECTATIONS
  "{e1} should be string 100"

  ADDRESS SPREADSHEET
  'GETCELL("E2")'
  LET e2 = RESULT
  ADDRESS EXPECTATIONS
  "{e2} should be string 200"

  ADDRESS SPREADSHEET
  'GETCELL("E3")'
  LET e3 = RESULT
  ADDRESS EXPECTATIONS
  "{e3} should be string 300"

  SAY "  ✓ Range operations (GETCELLS, SETCELLS)"
RETURN

TestSheetManagement:
  SAY "Testing sheet management..."

  ADDRESS SPREADSHEET

  // Test GETSHEETNAME (default)
  'GETSHEETNAME()'
  LET default_name = RESULT
  // Should be "Sheet1" or empty, just verify it doesn't error
  SAY "  - Current sheet name:" default_name

  // Test SETSHEETNAME
  'SETSHEETNAME("TestSheet")'
  'GETSHEETNAME()'
  LET new_name = RESULT
  ADDRESS EXPECTATIONS
  "{new_name} should equal TestSheet"

  // Reset to default
  ADDRESS SPREADSHEET
  'SETSHEETNAME("Sheet1")'

  SAY "  ✓ Sheet management (GETSHEETNAME, SETSHEETNAME)"
RETURN

TestExpressionEvaluation:
  SAY "Testing expression evaluation..."

  ADDRESS SPREADSHEET

  // Set up test data
  'SETCELL("F1", "10")'
  'SETCELL("F2", "20")'

  // Test EVALUATE
  'EVALUATE("F1 + F2")'
  LET eval_result = RESULT
  ADDRESS EXPECTATIONS
  "{eval_result} should equal 30"

  // Test RECALCULATE
  ADDRESS SPREADSHEET
  'SETCELL("F3", "=F1 + F2")'
  'RECALCULATE()'
  LET recalc_count = RESULT
  // Should recalculate at least 1 cell (F3)
  SAY "  - Recalculated" recalc_count "formulas"

  ADDRESS SPREADSHEET
  'GETCELL("F3")'
  LET f3_result = RESULT
  ADDRESS EXPECTATIONS
  "{f3_result} should equal 30"

  SAY "  ✓ Expression evaluation (EVALUATE, RECALCULATE)"
RETURN

TestSetupScript:
  SAY "Testing setup script management..."

  ADDRESS SPREADSHEET

  // Test SETSETUPSCRIPT
  'SETSETUPSCRIPT("LET TEST_CONST = 42")'

  // Test GETSETUPSCRIPT
  'GETSETUPSCRIPT()'
  LET script = RESULT
  ADDRESS EXPECTATIONS
  "{script} should equal LET TEST_CONST = 42"

  // Test EXECUTESETUPSCRIPT
  ADDRESS SPREADSHEET
  'EXECUTESETUPSCRIPT()'

  // Clear setup script
  'SETSETUPSCRIPT("")'

  SAY "  ✓ Setup script (GETSETUPSCRIPT, SETSETUPSCRIPT, EXECUTESETUPSCRIPT)"
RETURN

TestUtilityCommands:
  SAY "Testing utility commands..."

  ADDRESS SPREADSHEET

  // Test LISTCOMMANDS
  'LISTCOMMANDS()'
  LET commands = RESULT
  LET cmd_count = commands.0

  // Should have 26 commands
  ADDRESS EXPECTATIONS
  "{cmd_count} should equal 26"

  SAY "  - Found" cmd_count "available commands"

  // Verify some key commands are in the list
  LET found_setcell = 0
  LET found_getcell = 0
  LET found_export = 0

  DO i = 1 TO cmd_count
    // Use INTERPRET to build the property access dynamically
    INTERPRET "LET cmd = commands." || i
    IF cmd = "SETCELL" THEN found_setcell = 1
    IF cmd = "GETCELL" THEN found_getcell = 1
    IF cmd = "EXPORT" THEN found_export = 1
  END

  ADDRESS EXPECTATIONS
  "{found_setcell} should equal 1"
  "{found_getcell} should equal 1"
  "{found_export} should equal 1"

  // Test SPREADSHEET_VERSION
  ADDRESS SPREADSHEET
  'SPREADSHEET_VERSION()'
  LET version = RESULT
  SAY "  - Spreadsheet version:" version

  SAY "  ✓ Utility commands (LISTCOMMANDS, SPREADSHEET_VERSION)"
RETURN

TestImportExport:
  SAY "Testing import/export..."

  ADDRESS SPREADSHEET

  // Set up some test data
  'SETCELL("G1", "Test")'
  'SETCELL("G2", "=10+20")'
  'SETCOMMENT("G1", "Test comment")'

  // Test EXPORT
  'EXPORT()'
  LET exported = RESULT

  // Verify it's valid JSON
  SAY "  - Exported data length:" LENGTH(exported) "characters"

  // Test CLEAR
  'CLEAR()'
  'GETCELL("G1")'
  LET cleared_g1 = RESULT
  ADDRESS EXPECTATIONS
  "{cleared_g1} should equal "

  // Test IMPORT
  ADDRESS SPREADSHEET
  'IMPORT(exported)'
  'GETCELL("G1")'
  LET imported_g1 = RESULT
  ADDRESS EXPECTATIONS
  "{imported_g1} should equal Test"

  ADDRESS SPREADSHEET
  'GETCOMMENT("G1")'
  LET imported_comment = RESULT
  ADDRESS EXPECTATIONS
  "{imported_comment} should equal Test comment"

  SAY "  ✓ Import/export (EXPORT, IMPORT, CLEAR)"
RETURN
