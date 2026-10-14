# Licensing

| What | License | Text |
|---|---|---|
| `app/`, `tests/`, `tools/`, `docs/`, `run-tests.sh` (the sae app) | MIT | [LICENSE-MIT](LICENSE-MIT) |
| `vendor/rexxjs/` (RexxJS, vendored) | MIT | [vendor/rexxjs/LICENSE](vendor/rexxjs/LICENSE) |
| `rexx/` (control-protocol scripts) and the protocol itself | MIT | [LICENSE-PROTOCOL-MIT](LICENSE-PROTOCOL-MIT) |
| `tauri/` (the legacy app), but for its protocol files | AGPL v3 | [LICENSE-AGPL](LICENSE-AGPL) |

The sae app was written in sae (MIT) and moved here on 2026-10-14. What it
takes from the Tauri app, its author relicenses as MIT: the AGPL below
covers `tauri/` as it stands, not the sae app.

The rest of this file is the Tauri app's dual licensing, as it was written
(paths now under `tauri/` and `rexx/`).

---

# RexxSheet Dual Licensing (the Tauri app)

This project uses **dual licensing** to balance open source values with ecosystem growth:

## 📋 Summary

- **Application**: AGPL v3 (strong copyleft)
- **Control Protocol**: MIT (permissive)

## 🎯 Why Dual Licensing?

### AGPL for the Application
The RexxSheet spreadsheet application itself is licensed under **GNU Affero General Public License v3** to ensure:
- Source code modifications must be shared
- Network users have access to source code
- Derived works remain open source
- Community benefits from improvements

### MIT for the Protocol
The spreadsheet control protocol/grammar is licensed under **MIT License** to enable:
- Alternative implementations in any language
- Commercial integrations without restrictions
- Proprietary client development
- Ecosystem growth and interoperability

## 📄 File-by-File Breakdown

### AGPL Licensed (Application)
See [LICENSE-AGPL](LICENSE-AGPL)

- `tauri/src/SpreadsheetApp.jsx` - Main UI component
- `tauri/src/spreadsheet-model.js` - Data model
- `tauri/src/spreadsheet-rexx-adapter.js` - REXX integration
- `tauri/src/spreadsheet-loader.js` - Application loader
- `tauri/src/main.jsx` - Entry point
- `tauri/src/spreadsheet-address-handler.js` - ADDRESS handler
- `tauri/src-tauri/` - Tauri desktop app (Rust)
- Build and distribution files
- UI components and styling

**Requirements when modifying:**
- Share source code of modifications
- Maintain AGPL license
- Provide source to network users
- Credit original authors

### MIT Licensed (Protocol)
See [LICENSE-PROTOCOL-MIT](LICENSE-PROTOCOL-MIT)

- `tauri/src/spreadsheet-control-functions.js` - **Protocol/Grammar Definition**
- `rexx/test-spreadsheet-address.rexx` - Protocol example
- Protocol documentation and specifications

**Freedom granted:**
- Use in proprietary software
- Commercial applications
- No source code sharing required
- Create alternative implementations
- Fork for any purpose

## 🔌 Protocol/Grammar Definition

The control protocol consists of these REXX ADDRESS commands:

```rexx
/* MIT Licensed Protocol - Anyone can implement */
ADDRESS SPREADSHEET

/* Set cell content (value or formula) */
CALL SETCELL("A1", "100")
CALL SETCELL("A3", "=A1+A2")

/* Get cell value */
value = GETCELL("A1")

/* Get cell formula */
formula = GETEXPRESSION("A3")

/* Clear cell */
CALL CLEARCELL("A1")

/* Get protocol version */
version = SPREADSHEET_VERSION()
```

This protocol is **intentionally simple and well-documented** to encourage:
- Python clients
- Rust clients  
- Alternative spreadsheet implementations
- Commercial integrations
- Academic research

## 💡 Use Cases

### ✅ Allowed Without AGPL Obligations

1. **Proprietary Client**: Write a Python client that controls RexxSheet
   - MIT protocol applies
   - No source sharing required
   - Can be commercial

2. **Alternative Implementation**: Create "MySheet" that implements the protocol
   - MIT protocol applies
   - Your implementation can use any license
   - Can be proprietary

3. **Integration**: Use RexxSheet as backend for commercial app
   - MIT protocol applies for API calls
   - Your app can be proprietary
   - **BUT**: If you modify RexxSheet application itself, AGPL applies

### ⚠️ AGPL Obligations Apply

1. **Fork RexxSheet**: Create modified version
   - Must share source code
   - Must use AGPL
   - Must provide source to users

2. **Network Service**: Deploy RexxSheet as web service
   - Must provide source to users
   - Including any modifications
   - Cannot relicense to proprietary

3. **Embed UI**: Include RexxSheet UI components
   - Entire application becomes AGPL
   - Must share source code
   - Cannot create proprietary version

## 🤝 Commercial Use

### Protocol (MIT)
- ✅ Use in commercial products
- ✅ No licensing fees
- ✅ No source code requirements
- ✅ Create proprietary clients

### Application (AGPL)
- ✅ Use as-is commercially (must provide source)
- ❌ Cannot make proprietary modifications
- ✅ Can charge for services
- ✅ Can offer support/hosting

## 📞 Contact

Questions about licensing?
- Protocol implementation: Use freely under MIT
- Application modifications: Must comply with AGPL
- Dual licensing queries: Open an issue on GitHub

## 🎓 Educational Use

Both licenses allow full educational use:
- Teaching programming concepts
- Research projects
- Thesis implementations  
- Classroom demonstrations

**Academic derivatives should cite:**
```
RexxSheet Spreadsheet
Copyright (c) 2024-2025 Paul Hammant and Contributors
https://github.com/RexxJS/RexxJS
```

## 🔄 Contributing

Contributions to the protocol or application are welcome:
- Protocol contributions: Will remain MIT
- Application contributions: Will be AGPL
- Contributors retain copyright
- Agree to project licensing by submitting PR

## ⚖️ Legal Notes

1. **Licenses are compatible**: MIT code can be used in AGPL projects
2. **Protocol independence**: You can implement protocol without using our code
3. **Trademark**: "RexxSheet" name may be trademarked separately
4. **Warranty**: Both licenses provide NO WARRANTY

## 📚 Full License Texts

- AGPL v3: [LICENSE-AGPL](LICENSE-AGPL) or https://www.gnu.org/licenses/agpl-3.0.html
- MIT: [LICENSE-PROTOCOL-MIT](LICENSE-PROTOCOL-MIT)

---

**TL;DR:**
- Want to build a client? → **MIT** (do whatever you want)
- Want to modify the app? → **AGPL** (share your changes)
- Want to fork? → **AGPL** applies to application, **MIT** to protocol
- Not sure? → Ask! We want to enable innovation while keeping improvements open.
