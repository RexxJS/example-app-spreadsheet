// Test suite for row and column operations through the control bus:
// INSERTROW, DELETEROW, INSERTCOLUMN, DELETECOLUMN, on a running sheet.
//
// (The Tauri app's version of this only printed what each test should do;
// these are the same cases, run.)
//
// Prerequisites:
//   1. Start the spreadsheet (../app/) with its control bus on, from the
//      repository's root:
//      ../sae/target/build/bin/sae --app app --control-bus
//
//   2. In another terminal, from the repository's root:
//      RexxJS/core/rexx rexx/spreadsheet-row-column-ops.rexx

/* @test-tags spreadsheet, control-bus, integration */
/* @description Row and column insertion and deletion via ADDRESS SPREADSHEET */

REQUIRE "./spreadsheet-address.js"
REQUIRE "../RexxJS/core/src/expectations-address.js"

SAY "Testing row and column operations via the control bus"
SAY ""

ADDRESS SPREADSHEET
'SPREADSHEET_VERSION()'
IF RC \= 0 THEN DO
  SAY "❌ Failed to connect to the spreadsheet's control bus:" ERRORTEXT
  SAY "Start it first: ../sae/target/build/bin/sae --app app --control-bus"
  EXIT 1
END

CALL TestInsertRow
CALL TestDeleteRow
CALL TestInvalidRow
CALL TestInsertColumn
CALL TestDeleteColumn
CALL TestColumnLetter
CALL TestInvalidColumn

SAY ""
SAY "✅ All row/column operation tests passed!"
EXIT 0

// ============= Test Subroutines =============

// A1=10, A2=20, A3=30; INSERTROW(2) gives A1=10, A2 empty, A3=20, A4=30.
TestInsertRow:
  ADDRESS SPREADSHEET
  'CLEAR()'
  'SETCELLS("A1:A3", "[10, 20, 30]")'
  'INSERTROW(2)'
  'GETCELL("A2")'
  LET a2 = RESULT
  'GETCELL("A3")'
  LET a3 = RESULT
  'GETCELL("A4")'
  LET a4 = RESULT
  ADDRESS EXPECTATIONS
  "{a2} should equal "
  "{a3} should equal 20"
  "{a4} should equal 30"
  SAY "  ✓ INSERTROW inserts a row and shifts cells down"
RETURN

// Then DELETEROW(2) takes it back: A1=10, A2=20, A3=30, A4 empty.
TestDeleteRow:
  ADDRESS SPREADSHEET
  'DELETEROW(2)'
  'GETCELL("A2")'
  LET a2 = RESULT
  'GETCELL("A3")'
  LET a3 = RESULT
  'GETCELL("A4")'
  LET a4 = RESULT
  ADDRESS EXPECTATIONS
  "{a2} should equal 20"
  "{a3} should equal 30"
  "{a4} should equal "
  SAY "  ✓ DELETEROW deletes a row and shifts cells up"
RETURN

TestInvalidRow:
  ADDRESS SPREADSHEET
  'INSERTROW(0)'
  LET rc = RC
  LET why = ERRORTEXT
  ADDRESS EXPECTATIONS
  "{rc} should equal 1"
  "{why} should contain 'INSERTROW needs a row number'"
  SAY "  ✓ INSERTROW(0) is refused, saying why"
RETURN

// A1=10, B1=20, C1=30; INSERTCOLUMN(2) gives A1=10, B1 empty, C1=20, D1=30.
TestInsertColumn:
  ADDRESS SPREADSHEET
  'CLEAR()'
  'SETCELLS("A1:C1", "[10, 20, 30]")'
  'INSERTCOLUMN(2)'
  'GETCELL("B1")'
  LET b1 = RESULT
  'GETCELL("C1")'
  LET c1 = RESULT
  'GETCELL("D1")'
  LET d1 = RESULT
  ADDRESS EXPECTATIONS
  "{b1} should equal "
  "{c1} should equal 20"
  "{d1} should equal 30"
  SAY "  ✓ INSERTCOLUMN inserts a column and shifts cells right"
RETURN

// Then DELETECOLUMN(2) takes it back: B1=20, C1=30, D1 empty.
TestDeleteColumn:
  ADDRESS SPREADSHEET
  'DELETECOLUMN(2)'
  'GETCELL("B1")'
  LET b1 = RESULT
  'GETCELL("C1")'
  LET c1 = RESULT
  'GETCELL("D1")'
  LET d1 = RESULT
  ADDRESS EXPECTATIONS
  "{b1} should equal 20"
  "{c1} should equal 30"
  "{d1} should equal "
  SAY "  ✓ DELETECOLUMN deletes a column and shifts cells left"
RETURN

TestColumnLetter:
  ADDRESS SPREADSHEET
  'INSERTCOLUMN("B")'
  'GETCELL("C1")'
  LET c1 = RESULT
  ADDRESS EXPECTATIONS
  "{c1} should equal 20"
  SAY "  ✓ INSERTCOLUMN takes a column letter"
RETURN

TestInvalidColumn:
  ADDRESS SPREADSHEET
  'INSERTCOLUMN(0)'
  LET rc = RC
  LET why = ERRORTEXT
  ADDRESS EXPECTATIONS
  "{rc} should equal 1"
  "{why} should contain 'INSERTCOLUMN needs a column'"
  SAY "  ✓ INSERTCOLUMN(0) is refused, saying why"
RETURN
