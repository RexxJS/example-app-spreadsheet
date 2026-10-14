-- test-connection.rexx
-- Is the spreadsheet there? One command to its control bus, and what to do
-- when nothing answers.
--
-- The Tauri app's version of this POSTed JSON itself, with HTTP_POST and a
-- fixed token. The sae app's bus takes only sealed requests, with a key
-- fresh for every launch, so the check goes through ADDRESS SPREADSHEET
-- (./spreadsheet-address.js), which runs sae's sae-bus.
--
-- Usage, from the repository's root:
--   1. Start the spreadsheet (../app/) with its control bus on:
--      ../sae/target/build/bin/sae --app app --control-bus
--
--   2. Run this script in another terminal:
--      RexxJS/core/rexx rexx/test-connection.rexx

REQUIRE "./spreadsheet-address.js"

SAY "Testing connection to the spreadsheet's control bus..."
SAY ""

ADDRESS SPREADSHEET
'SPREADSHEET_VERSION()'

IF RC = 0 THEN DO
  SAY "✓ Connection successful: spreadsheet version" RESULT
  EXIT 0
END

SAY "ERROR: no answer, RC" RC
SAY "  " ERRORTEXT
SAY ""
SAY "Troubleshooting:"
SAY "1. Is the spreadsheet running? Start it with:"
SAY "     ../sae/target/build/bin/sae --app app --control-bus"
SAY "2. Without --control-bus the app does not listen at all."
SAY "3. Started with SAE_APPDATA_DIR? Set the same here, so sae-bus finds"
SAY "   the courier file."
EXIT 1
