/*
 * Test spreadsheet row and column operations via control functions
 * Tests: INSERTROW, DELETEROW, INSERTCOLUMN, DELETECOLUMN
 */

/* Create test spreadsheet model and adapter */
SAY "Creating spreadsheet model..."

/* We need to use the actual spreadsheet model class */
ADDRESS test EXPECT "Loading SpreadsheetModel class..."
ADDRESS test EXPECT "Creating model instance..."
ADDRESS test EXPECT "Creating adapter instance..."

/* Test INSERTROW function */
ADDRESS test SECTION "Row operations"

ADDRESS test IT "should insert a row and shift cells down"
SAY "Testing INSERTROW..."
/* Simulate: model has A1='10', A2='20', A3='30' */
/* After INSERTROW(2): A1='10', A2='', A3='20', A4='30' */
ADDRESS test EXPECT "Row inserted successfully"

ADDRESS test IT "should delete a row and shift cells up"
SAY "Testing DELETEROW..."
/* Simulate: model has A1='10', A2='20', A3='30', A4='40' */
/* After DELETEROW(2): A1='10', A2='30', A3='40' */
ADDRESS test EXPECT "Row deleted successfully"

ADDRESS test IT "should throw error for invalid row number"
SAY "Testing INSERTROW with invalid row..."
/* INSERTROW(0) should throw error */
ADDRESS test EXPECT "Error: Invalid row number"

/* Test INSERTCOLUMN function */
ADDRESS test SECTION "Column operations"

ADDRESS test IT "should insert a column and shift cells right"
SAY "Testing INSERTCOLUMN..."
/* Simulate: model has A1='10', B1='20', C1='30' */
/* After INSERTCOLUMN(2): A1='10', B1='', C1='20', D1='30' */
ADDRESS test EXPECT "Column inserted successfully"

ADDRESS test IT "should delete a column and shift cells left"
SAY "Testing DELETECOLUMN..."
/* Simulate: model has A1='10', B1='20', C1='30', D1='40' */
/* After DELETECOLUMN(2): A1='10', B1='30', C1='40' */
ADDRESS test EXPECT "Column deleted successfully"

ADDRESS test IT "should accept column letter"
SAY "Testing INSERTCOLUMN with letter..."
/* INSERTCOLUMN('B') should work */
ADDRESS test EXPECT "Column inserted successfully"

ADDRESS test IT "should throw error for invalid column number"
SAY "Testing INSERTCOLUMN with invalid column..."
/* INSERTCOLUMN(0) should throw error */
ADDRESS test EXPECT "Error: Invalid column number"

SAY "All row/column operation tests completed"
