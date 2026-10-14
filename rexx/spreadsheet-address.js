/*!
 * spreadsheet-address | MIT License (../LICENSE-PROTOCOL-MIT)
 * @rexxjs-meta=SPREADSHEET_ADDRESS_META
 */
/**
 * ADDRESS SPREADSHEET for the sae app (../app/), ARexx style:
 *
 *   REQUIRE "./spreadsheet-address.js"
 *   ADDRESS SPREADSHEET
 *   'SETCELL("A1", "100")'
 *   'GETCELL("A1")'
 *   SAY RESULT
 *
 * Each command is a string of REXX the app runs (its control functions:
 * SETCELL, GETCELL, LISTCOMMANDS and the rest); RC is 0 when it worked and
 * RESULT what it came to, else RC is non-zero and ERRORTEXT says why. A
 * command runs in the app, so it sees the app's names, not this script's
 * variables: put values in the string.
 *
 * The app must be running with its control bus on
 * (`sae --app app --control-bus`). The bus is sae's: a port chosen at
 * random and a key fresh for every launch, in a courier file only this user
 * can read, and every message sealed (sae's docs/control-bus.md). This
 * library does none of that itself: it runs sae's `sae-bus` command, which
 * reads the courier file, seals, sends and opens the answer.
 *
 *   SAE_BUS            the sae-bus to run (default: ../sae/target/sae-bus,
 *                      beside this repository's sae link)
 *   SPREADSHEET_APP    the app's name, which finds its courier file
 *                      (default "Spreadsheet", app/app.json's name)
 *   SAE_APPDATA_DIR    the app's data folder, when the app was started
 *                      with one (sae-bus reads it)
 *
 * RC is 2 when there was no app to ask (not running, the bus off, or the
 * answer did not open), with sae-bus's reason in ERRORTEXT.
 */

function SPREADSHEET_ADDRESS_META() {
  return {
    type: 'address-target',
    name: 'sae spreadsheet control bus',
    version: '1.0.0',
    description: 'ADDRESS SPREADSHEET: commands to the sae spreadsheet app over its sealed control bus',
    provides: {
      addressTarget: 'spreadsheet',
      commandSupport: true,
      methodSupport: false
    },
    dependencies: {},
    loaded: true,
    requirements: {
      environment: 'nodejs'
    }
  };
}

// One command, through sae-bus: {success, errorCode, errorMessage, output}.
function ADDRESS_SPREADSHEET_HANDLER(command, context, sourceContext) {
  const { execFile } = require('child_process');
  const bus = process.env.SAE_BUS || require('path').join(__dirname, '..', 'sae', 'target', 'sae-bus');
  const app = process.env.SPREADSHEET_APP || 'Spreadsheet';
  return new Promise((resolve) => {
    execFile(bus, [app, String(command)], { encoding: 'utf8' }, (error, stdout, stderr) => {
      if (error) {
        const why = (stderr || '').trim() || error.message;
        resolve({ success: false, errorCode: 2, errorMessage: why, output: '' });
        return;
      }
      const answer = JSON.parse(stdout);
      if (answer.success) {
        resolve({ success: true, output: answer.result === null ? '' : answer.result });
      } else {
        resolve({ success: false, errorCode: answer.errorCode || 1, errorMessage: answer.error, output: '' });
      }
    });
  });
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { SPREADSHEET_ADDRESS_META, ADDRESS_SPREADSHEET_HANDLER };
}
if (typeof global !== 'undefined') {
  global.SPREADSHEET_ADDRESS_META = SPREADSHEET_ADDRESS_META;
  global.ADDRESS_SPREADSHEET_HANDLER = ADDRESS_SPREADSHEET_HANDLER;
}
