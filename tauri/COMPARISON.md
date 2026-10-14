# Spreadsheet Feature Comparison

Comprehensive comparison of RexxSheet vs. Microsoft Excel vs. Google Sheets

## Feature Matrix

| Feature Category | Feature | Excel | Google Sheets | RexxSheet |
|-----------------|---------|-------|---------------|-----------|
| **Basic Spreadsheet** |
| | Cell values and formulas | ✅ Full | ✅ Full | ✅ Full |
| | A1-style cell references | ✅ Yes | ✅ Yes | ✅ Yes |
| | Cell ranges (A1:B10) | ✅ Yes | ✅ Yes | ✅ Yes |
| | Multiple sheets/tabs | ✅ Yes | ✅ Yes | ❌ No (single sheet) |
| | Rows × Columns | ✅ 1M × 16K | ✅ 10M cells total | ⚠️ 100 × 26 (configurable) |
| | Cell comments/notes | ✅ Yes | ✅ Yes | ✅ Yes |
| | Auto-recalculation | ✅ Yes | ✅ Yes | ✅ Yes |
| | Dependency tracking | ✅ Yes | ✅ Yes | ✅ Yes |
| | Circular reference detection | ✅ Yes | ✅ Yes | ✅ Yes |
| **Formula Language** |
| | Expression language | ⚠️ Excel formulas | ⚠️ Similar to Excel | ✅ **RexxJS** (full programming language) |
| | Named variables/constants | ✅ Named ranges | ✅ Named ranges | ✅ **Setup Script** (LET statements) |
| | User-defined functions | ✅ VBA/Lambda | ✅ Apps Script | ✅ **RexxJS REQUIRE** |
| | Function pipelines | ❌ No | ❌ No | ✅ **Yes** (|> operator) |
| | Ternary/IF expressions | ✅ IF() | ✅ IF() | ✅ Native RexxJS |
| **Built-in Functions** |
| | Math functions | ✅ 50+ | ✅ 50+ | ✅ 15+ (extensible) |
| | String functions | ✅ 30+ | ✅ 30+ | ✅ 20+ (extensible) |
| | Date/Time functions | ✅ 30+ | ✅ 30+ | ✅ 10+ (extensible) |
| | Statistical functions | ✅ 100+ | ✅ 100+ | ⚠️ Via libraries |
| | Lookup functions (VLOOKUP) | ✅ Yes | ✅ Yes | ⚠️ Via Excel library |
| | Array functions | ✅ Yes | ✅ Yes | ✅ MAP, FILTER, REDUCE |
| | Financial functions | ✅ 50+ | ✅ 50+ | ❌ Not yet |
| | Database functions | ✅ Yes | ✅ Yes | ❌ No |
| | NumPy integration | ❌ No | ❌ No | ✅ **Yes** (via PyOdide) |
| **Number Formatting** |
| | Currency format | ✅ Yes | ✅ Yes | ✅ Yes ($#,##0.00) |
| | Percentage format | ✅ Yes | ✅ Yes | ✅ Yes (0.00%) |
| | Date formats | ✅ Extensive | ✅ Extensive | ✅ Basic (yyyy-mm-dd, mm/dd/yyyy) |
| | Custom number formats | ✅ Full syntax | ✅ Full syntax | ⚠️ Basic patterns |
| | Scientific notation | ✅ Yes | ✅ Yes | ❌ Not yet |
| | Fractions | ✅ Yes | ✅ Yes | ❌ No |
| **Cell Styling** |
| | Font (family, size, style) | ✅ Full control | ✅ Full control | ✅ Bold, Italic |
| | Text color | ✅ Yes | ✅ Yes | ✅ Yes (any CSS color) |
| | Background color | ✅ Yes | ✅ Yes | ✅ Yes (any CSS color) |
| | Borders | ✅ Full control | ✅ Full control | ✅ Basic (CSS border) |
| | Text alignment | ✅ Full control | ✅ Full control | ✅ Left/Center/Right |
| | Cell merge | ✅ Yes | ✅ Yes | ❌ No |
| | Wrap text | ✅ Yes | ✅ Yes | ❌ No |
| **Conditional Formatting** |
| | Rules-based formatting | ✅ GUI builder | ✅ GUI builder | ❌ No GUI |
| | Expression-based formatting | ⚠️ Limited | ⚠️ Custom formulas | ✅ **Full RexxJS expressions** |
| | Color scales | ✅ Yes | ✅ Yes | ⚠️ Manual via expressions |
| | Data bars | ✅ Yes | ✅ Yes | ❌ No |
| | Icon sets | ✅ Yes | ✅ Yes | ❌ No |
| | Programming language for conditions | ❌ No | ❌ No | ✅ **Yes** (RexxJS) |
| **Data Entry & Editing** |
| | In-cell editing | ✅ Yes | ✅ Yes | ✅ Yes |
| | Formula bar | ✅ Yes | ✅ Yes | ✅ Yes |
| | AutoFill/drag-fill | ✅ Yes | ✅ Yes | ❌ No |
| | Copy/Paste | ✅ Full support | ✅ Full support | ⚠️ Copy only (no paste yet) |
| | Undo/Redo | ✅ Yes | ✅ Yes | ❌ No |
| | Find/Replace | ✅ Yes | ✅ Yes | ❌ No |
| | Data validation | ✅ Extensive | ✅ Extensive | ❌ No |
| | Drop-down lists | ✅ Yes | ✅ Yes | ❌ No |
| **Row/Column Operations** |
| | Insert rows/columns | ✅ Yes | ✅ Yes | ✅ Yes |
| | Delete rows/columns | ✅ Yes | ✅ Yes | ✅ Yes |
| | Hide rows/columns | ✅ Yes | ✅ Yes | ❌ No |
| | Resize rows/columns | ✅ Yes | ✅ Yes | ✅ Yes (interactive) |
| | Freeze panes | ✅ Yes | ✅ Yes | ❌ No |
| | AutoFit | ✅ Yes | ✅ Yes | ❌ No |
| **Data Analysis** |
| | Sorting | ✅ Multi-level | ✅ Multi-level | ❌ No |
| | Filtering | ✅ Advanced | ✅ Advanced | ❌ No |
| | Pivot tables | ✅ Full-featured | ✅ Full-featured | ❌ No |
| | Charts/Graphs | ✅ 50+ types | ✅ 40+ types | ❌ No |
| | Sparklines | ✅ Yes | ✅ Yes | ❌ No |
| | What-if analysis | ✅ Goal Seek, Solver | ✅ Goal Seek | ❌ No |
| | Data tables | ✅ Yes | ✅ Yes | ❌ No |
| **Import/Export** |
| | Excel (.xlsx) format | ✅ Native | ✅ Import/Export | ❌ No |
| | CSV format | ✅ Yes | ✅ Yes | ❌ Not yet |
| | JSON format | ⚠️ Via Power Query | ⚠️ Via scripts | ✅ **Native format** |
| | PDF export | ✅ Yes | ✅ Yes | ❌ No |
| | Copy as image | ✅ Yes | ✅ Yes | ❌ No |
| | Load from URL | ⚠️ Power Query | ⚠️ IMPORTDATA() | ✅ **Yes** (hash parameter) |
| **Collaboration** |
| | Real-time collaboration | ⚠️ OneDrive/365 | ✅ **Best-in-class** | ❌ No |
| | Comments/discussions | ✅ Yes | ✅ Yes | ⚠️ Cell comments only |
| | Version history | ⚠️ OneDrive/365 | ✅ Yes | ❌ No |
| | Share permissions | ⚠️ OneDrive/365 | ✅ Yes | ❌ No |
| | Simultaneous editing | ⚠️ OneDrive/365 | ✅ Yes | ❌ No |
| **Automation & Scripting** |
| | Macro recording | ✅ VBA macros | ✅ Apps Script | ❌ No |
| | Scripting language | ✅ VBA | ✅ JavaScript (Apps Script) | ✅ **RexxJS** |
| | Formula language is programming language | ❌ No | ❌ No | ✅ **Yes** (RexxJS) |
| | External library loading | ⚠️ Add-ins | ⚠️ Libraries | ✅ **REQUIRE** statement |
| | Remote control API | ⚠️ COM/OLE | ✅ Apps Script API | ✅ **Control Bus** (ARexx-style) |
| | HTTP control API | ❌ No | ⚠️ Web app | ✅ **Yes** (Tauri mode) |
| | iframe postMessage API | ❌ No | ❌ No | ✅ **Yes** (web mode) |
| **View Modes** |
| | Normal view | ✅ Yes | ✅ Yes | ✅ Yes |
| | Formula view | ✅ Ctrl+` | ✅ Show formulas | ✅ **Formulas Only mode** |
| | Page layout view | ✅ Yes | ❌ No | ❌ No |
| | Custom view modes | ❌ No | ❌ No | ✅ Values/Formulas/Formats |
| **Platform Availability** |
| | Windows desktop | ✅ Yes | ⚠️ Web only | ✅ Tauri app |
| | macOS desktop | ✅ Yes | ⚠️ Web only | ✅ Tauri app |
| | Linux desktop | ❌ No | ⚠️ Web only | ✅ **Tauri app** |
| | Web browser | ⚠️ Excel Online | ✅ Primary | ✅ Yes |
| | Mobile (iOS/Android) | ✅ Apps | ✅ Apps | ⚠️ Web responsive |
| | Offline mode | ✅ Desktop | ⚠️ Limited | ✅ Desktop app |
| **File Storage** |
| | Local files | ✅ Yes | ⚠️ Download | ✅ Yes (Tauri) |
| | Cloud storage | ✅ OneDrive | ✅ **Google Drive** | ❌ No |
| | Auto-save | ✅ Yes | ✅ **Continuous** | ❌ No |
| | Manual save | ✅ Yes | ⚠️ Not needed | ❌ Not yet |
| **Unique Features** |
| | | VBA macros, Pivot tables, Power Query | Real-time collab, IMPORTDATA, QUERY | **RexxJS expressions, Function pipelines, Control Bus, NumPy** |
| **Licensing** |
| | Cost | 💰 $70-160/year | ✅ **Free** | ✅ **Free (AGPL)** |
| | Open source | ❌ Proprietary | ❌ Proprietary | ✅ **Yes (AGPL/MIT)** |
| | Self-hosted | ❌ No | ❌ No | ✅ **Yes** |
| | Commercial use | ✅ With license | ✅ Free | ✅ Yes (share changes) |
| **Performance** |
| | Large datasets (1M+ rows) | ✅ Optimized | ⚠️ Can struggle | ❌ Not designed for this |
| | Calculation speed | ✅ Fast | ✅ Fast | ⚠️ Adequate |
| | Load time | ⚠️ Heavy app | ✅ Fast | ✅ Fast (~1s) |
| | Memory usage | ⚠️ Heavy | ✅ Efficient | ✅ Light |

## Legend

- ✅ **Fully supported** - Feature is implemented and works well
- ⚠️ **Partially supported** - Feature exists but with limitations
- ❌ **Not supported** - Feature is not available
- 💰 **Paid** - Requires payment/subscription

## Key Differentiators

### Excel Strengths
1. **Most mature** - 30+ years of development
2. **Most powerful** - Pivot tables, Power Query, Power BI integration
3. **Most functions** - 400+ built-in functions
4. **Best offline** - Full desktop app with rich features
5. **Enterprise standard** - Industry-wide compatibility

### Google Sheets Strengths
1. **Best collaboration** - Real-time multi-user editing
2. **Best accessibility** - Works anywhere with a browser
3. **Best integration** - Google Workspace ecosystem
4. **Always up-to-date** - No installation or updates needed
5. **Free** - No cost for personal and small business use

### RexxSheet Strengths
1. **Full programming language** - RexxJS expressions in every cell
2. **Function pipelines** - Chain operations with `|>` operator
3. **Extensible** - Load libraries via REQUIRE (Excel functions, R stats, NumPy)
4. **Control Bus** - ARexx-style remote control and automation
5. **Open source** - AGPL license, self-hostable, fully transparent
6. **NumPy integration** - Real Python scientific computing via PyOdide
7. **Conditional formatting with code** - Full RexxJS for dynamic styling
8. **Cross-platform native** - Windows, macOS, Linux via Tauri

## Use Case Recommendations

### Choose Excel if you need:
- Industry-standard compatibility
- Advanced data analysis (pivot tables, Power Query)
- Complex financial modeling
- Offline work on Windows/Mac
- VBA automation and macros

### Choose Google Sheets if you need:
- Real-time collaboration with teams
- Access from any device/browser
- Integration with Google Workspace
- No installation required
- Free solution for most use cases

### Choose RexxSheet if you need:
- Programming language power in formulas
- Open source and self-hosted
- Custom automation via scripting
- Educational/research use with custom functions
- Cross-platform desktop app (including Linux)
- Integration with external systems via Control Bus
- Scientific computing with NumPy

## RexxSheet Roadmap

### Currently Missing (Potential Enhancements)
- **Save/Export** - JSON export works, need UI save button
- **Undo/Redo** - Not implemented yet
- **Paste** - Copy works, paste not yet implemented
- **Charts** - No visualization features yet
- **Multi-sheet** - Single sheet only
- **CSV Import/Export** - JSON only currently
- **AutoFill** - No drag-fill pattern detection
- **Filtering/Sorting** - Manual via formulas only

### Unique Capabilities Not in Others
- **RexxJS as formula language** - Full Turing-complete language in cells
- **Function pipelines** - `="hello" |> UPPER |> LENGTH` returns 5
- **REQUIRE system** - Load external function libraries dynamically
- **Control Bus** - Remote control via HTTP API or postMessage
- **Setup Scripts** - Page-level code with global variables
- **NumPy via PyOdide** - 100% accurate Python scientific computing
- **View modes** - Switch between values/formulas/formats/normal
- **Editable metadata** - All cell properties editable in UI

## Philosophy Comparison

| Aspect | Excel | Google Sheets | RexxSheet |
|--------|-------|---------------|-----------|
| **Design Goal** | Professional spreadsheet for business | Cloud-first collaborative spreadsheet | Programmable spreadsheet with RexxJS |
| **Target User** | Business professionals, analysts | Teams, educators, casual users | Developers, researchers, automation enthusiasts |
| **Primary Strength** | Power and features | Collaboration and accessibility | Programmability and extensibility |
| **Complexity** | High - many features | Medium - streamlined | Medium - code-focused |
| **Learning Curve** | Steep for advanced features | Gentle | Medium (requires programming) |
| **Extensibility** | VBA, Add-ins | Apps Script, Add-ons | **REQUIRE, native RexxJS** |

## Example: Conditional Formatting

### Excel
```
Rule: Formula: =A1<0
Format: Red text
```

### Google Sheets
```
Custom formula: =A1<0
Format: Red text
```

### RexxSheet
```json
{
  "A1": {
    "content": "-150",
    "format": "$#,##0.00",
    "styleExpression": "STYLE_IF(A1 < 0, RED_TEXT(), GREEN_TEXT())"
  }
}
```

**RexxSheet advantage:** Full programming language for conditions, can call functions, use complex logic, load custom style libraries.

## Conclusion

**Excel** remains the most powerful and feature-rich spreadsheet for professional use, especially for complex data analysis and financial modeling.

**Google Sheets** is the best choice for collaboration and accessibility, with seamless real-time editing and cloud integration.

**RexxSheet** offers a unique approach: a **programmable spreadsheet** where every cell formula is written in a full programming language (RexxJS), with function pipelines, dynamic library loading, and powerful automation capabilities. It's ideal for developers, researchers, and automation enthusiasts who want more programming power than traditional spreadsheets offer, with the added benefits of being open source and cross-platform.

All three have their place:
- **Excel** for professional/enterprise data analysis
- **Google Sheets** for team collaboration
- **RexxSheet** for programmable automation and extensibility
