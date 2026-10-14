// Test suite for spreadsheet-poc functions via ADDRESS
// Copyright (c) 2025 RexxJS Project
// Tests spreadsheet functions running in the browser via control bus
//
// Prerequisites:
//   1. Start the spreadsheet (../app/) with its control bus on, from the
//      repository's root:
//      ../sae/target/build/bin/sae --app app --control-bus
//
//   2. In another terminal, from the repository's root:
//      RexxJS/core/rexx rexx/spreadsheet-functions-test.rexx
//
// ADDRESS SPREADSHEET is ./spreadsheet-address.js: each command goes to the
// app sealed, through sae's sae-bus (see README.md).
//
// This is an integration test that validates spreadsheet range functions
// (SUM_RANGE, AVERAGE_RANGE, MIN_RANGE, MAX_RANGE, MEDIAN_RANGE,
//  SUMIF_RANGE, COUNTIF_RANGE) work correctly in the actual spreadsheet.

/* @test-tags spreadsheet, functions, statistical, address, integration */
/* @description Test spreadsheet range functions via ADDRESS control bus */

REQUIRE "./spreadsheet-address.js"
REQUIRE "../RexxJS/core/src/expectations-address.js"

SAY "📊 Testing Spreadsheet Range Functions via ADDRESS Control Bus"
SAY "================================================================"
SAY ""
SAY "Prerequisites:"
SAY "  1. Spreadsheet must be running with control bus enabled"
SAY "  2. Run: ../sae/target/build/bin/sae --app app --control-bus"
SAY "  3. sae-bus finds it through its courier file"
SAY ""

// Verify connection
SAY "Checking spreadsheet connection..."
ADDRESS SPREADSHEET
'SPREADSHEET_VERSION()'
IF RC \= 0 THEN DO
  SAY ""
  SAY "❌ Failed to connect to spreadsheet control bus"
  SAY ""
  SAY "Please start the spreadsheet first:"
  SAY "  ../sae/target/build/bin/sae --app app --control-bus"
  SAY "  (" ERRORTEXT ")"
  SAY ""
  SAY "Then run this test again."
  SAY ""
  EXIT 1
END
SAY "✓ Connected to spreadsheet version:" RESULT
// The Tauri app's sheets were RexxJS formulas only; the sae app's start as
// Excel's, so ask for RexxJS's (SUM_RANGE, |> and the rest).
'SETFORMULALANGUAGE("RexxJS")'
SAY ""

// Test basic statistical functions
CALL TestSumRangeTest
CALL TestAverageRangeTest
CALL TestMinMaxRangeTest

// Test formulas with statistical functions
CALL TestMedianFormulaTest
CALL TestConditionalFunctionsTest

SAY ""
SAY "✅ All spreadsheet function tests passed!"
EXIT 0

// ============= Test Subroutines =============

TestSumRangeTest:
  SAY "Testing SUM_RANGE via spreadsheet..."

  ADDRESS SPREADSHEET
  // Set up test data in cells A1-A5
  'SETCELL("A1", "10")'
  'SETCELL("A2", "20")'
  'SETCELL("A3", "30")'
  'SETCELL("A4", "40")'
  'SETCELL("A5", "50")'

  // Use SUM_RANGE function in A6
  'SETCELL("A6", "=SUM_RANGE(A1:A5)")'

  // Get the result
  'GETCELL("A6")'
  LET sum_result = RESULT

  ADDRESS EXPECTATIONS
  "{sum_result} should equal 150"

  SAY "  ✓ SUM_RANGE: 10+20+30+40+50 = 150"
RETURN

TestAverageRangeTest:
  SAY "Testing AVERAGE_RANGE via spreadsheet..."

  ADDRESS SPREADSHEET
  // Use AVERAGE_RANGE function on same data (A1-A5)
  'SETCELL("B1", "=AVERAGE_RANGE(A1:A5)")'

  // Get the result
  'GETCELL("B1")'
  LET avg_result = RESULT

  ADDRESS EXPECTATIONS
  "{avg_result} should equal 30"

  SAY "  ✓ AVERAGE_RANGE: average of [10,20,30,40,50] = 30"
RETURN

TestMinMaxRangeTest:
  SAY "Testing MIN_RANGE and MAX_RANGE via spreadsheet..."

  ADDRESS SPREADSHEET
  // Set up test data with unsorted values
  'SETCELL("C1", "50")'
  'SETCELL("C2", "5")'
  'SETCELL("C3", "30")'
  'SETCELL("C4", "2")'
  'SETCELL("C5", "10")'

  // Test MIN_RANGE
  'SETCELL("C6", "=MIN_RANGE(C1:C5)")'
  'GETCELL("C6")'
  LET min_result = RESULT

  ADDRESS EXPECTATIONS
  "{min_result} should equal 2"

  SAY "  ✓ MIN_RANGE: minimum of [50,5,30,2,10] = 2"

  ADDRESS SPREADSHEET
  // Test MAX_RANGE
  'SETCELL("C7", "=MAX_RANGE(C1:C5)")'
  'GETCELL("C7")'
  LET max_result = RESULT

  ADDRESS EXPECTATIONS
  "{max_result} should equal 50"

  SAY "  ✓ MAX_RANGE: maximum of [50,5,30,2,10] = 50"
RETURN

TestMedianFormulaTest:
  SAY "Testing MEDIAN_RANGE via spreadsheet..."

  ADDRESS SPREADSHEET
  // Test with odd count (5 values)
  'SETCELL("D1", "10")'
  'SETCELL("D2", "20")'
  'SETCELL("D3", "30")'
  'SETCELL("D4", "40")'
  'SETCELL("D5", "50")'
  'SETCELL("D6", "=MEDIAN_RANGE(D1:D5)")'
  'GETCELL("D6")'
  LET median_odd = RESULT

  ADDRESS EXPECTATIONS
  "{median_odd} should equal 30"

  SAY "  ✓ MEDIAN_RANGE: median of [10,20,30,40,50] = 30"

  ADDRESS SPREADSHEET
  // Test with even count (4 values)
  'SETCELL("E1", "10")'
  'SETCELL("E2", "20")'
  'SETCELL("E3", "30")'
  'SETCELL("E4", "40")'
  'SETCELL("E5", "=MEDIAN_RANGE(E1:E4)")'
  'GETCELL("E5")'
  LET median_even = RESULT

  ADDRESS EXPECTATIONS
  "{median_even} should equal 25"

  SAY "  ✓ MEDIAN_RANGE: median of [10,20,30,40] = 25"
RETURN

TestConditionalFunctionsTest:
  SAY "Testing SUMIF_RANGE and COUNTIF_RANGE via spreadsheet..."

  ADDRESS SPREADSHEET
  // Set up test data
  'SETCELL("F1", "5")'
  'SETCELL("F2", "10")'
  'SETCELL("F3", "15")'
  'SETCELL("F4", "20")'
  'SETCELL("F5", "25")'
  'SETCELL("F6", "30")'

  // Test SUMIF_RANGE with >15 condition
  'SETCELL("F7", "=SUMIF_RANGE(F1:F6, \">15\")")'
  'GETCELL("F7")'
  LET sumif_result = RESULT

  ADDRESS EXPECTATIONS
  "{sumif_result} should equal 75"

  SAY "  ✓ SUMIF_RANGE: sum where >15 = 20+25+30 = 75"

  ADDRESS SPREADSHEET
  // Test COUNTIF_RANGE with >15 condition
  'SETCELL("F8", "=COUNTIF_RANGE(F1:F6, \">15\")")'
  'GETCELL("F8")'
  LET countif_result = RESULT

  ADDRESS EXPECTATIONS
  "{countif_result} should equal 3"

  SAY "  ✓ COUNTIF_RANGE: count where >15 = 3 values"
RETURN
