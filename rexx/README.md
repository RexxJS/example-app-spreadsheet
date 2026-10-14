# REXX scripts that drive the spreadsheet

Scripts that control a running spreadsheet from another program, ARexx
style: `ADDRESS SPREADSHEET`, then `SETCELL`, `GETCELL` and the rest of the
control protocol (MIT; see [../LICENSING.md](../LICENSING.md)).

| Script | What it does |
|---|---|
| `test-spreadsheet-address.rexx` | The ARexx-style demonstration: sets cells and formulas, reads them back |
| `test-connection.rexx` | Checks the spreadsheet answers |
| `spreadsheet-control-bus-test.rexx` | Every control-bus command, as tests |
| `spreadsheet-functions-test.rexx` | The range functions, over the bus |
| `spreadsheet-row-column-ops.rexx` | Inserting, deleting and moving rows and columns |

## Which spreadsheet they reach

They were written for the Tauri app (`../tauri/`), whose control bus is
plain HTTP on a fixed port with a fixed token:

```rexx
ADDRESS "http://localhost:2410/api/spreadsheet" AUTH "dev-token-12345" AS SPREADSHEET
```

and they run from inside the RexxJS repository (`../../core/rexx`), where
the Tauri app used to live.

The sae app (`../app/`) answers the same commands, as the same strings,
but its control bus is sae's: switched on only at launch
(`--control-bus`), on a port chosen at random, with a key that changes on
every launch, written to a file only this user can read, and every
request and answer sealed with it (sae's `docs/control-bus.md`). A script
cannot name the port or key in its source, and RexxJS's `ADDRESS ... AUTH`
sends the token in the clear, which sae refuses.

So these scripts do not reach the sae app yet. What would make them:

- **A client that does the sealing.** For example a small `sae-bus`
  command in sae that reads the courier file, seals a command, and prints
  the answer, and an ADDRESS environment that runs it; or RexxJS learning
  to seal (Ascon-AEAD128, as sae's `services/secure_channel` does).
- **Then the scripts here changed** from the fixed URL and token to
  `ADDRESS SPREADSHEET` by name, with that client finding the rest.

`tests/specs/spec_app_sheet_bus.ae` is what drives the sae app over its bus
today: an Aether program doing what such a client would.
