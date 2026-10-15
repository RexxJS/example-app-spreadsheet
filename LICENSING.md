# Licensing

| What | License | Text |
|---|---|---|
| `app/`, `mcp/`, `tests/`, `tools/`, `docs/`, `run-tests.sh` (the sae app and its MCP server) | MIT | [LICENSE-MIT](LICENSE-MIT) |
| `vendor/rexxjs/` (RexxJS, vendored) | MIT | [vendor/rexxjs/LICENSE](vendor/rexxjs/LICENSE) |
| `rexx/` (control-protocol scripts) and the protocol itself | MIT | [LICENSE-PROTOCOL-MIT](LICENSE-PROTOCOL-MIT) |

The sae app was written in sae (MIT) and moved here on 2026-10-14. What it
takes from the Tauri app, its author relicenses as MIT.

The Tauri app itself (React + Tauri, `tauri/`) was dual-licensed, the
application AGPL v3 and its control protocol MIT. It was removed on
2026-10-15 and stays under those terms in this repository's history: its
files, `LICENSE-AGPL` and this file's account of the dual licensing are at
commit `bd1a002`.
