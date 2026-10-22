# Architecture — afo-mcp

Technical reference for the afo-mcp server: how an Analysis for Office workbook is built, how the
server reads and writes it, how the agent bridge works, and how the operate tools drive Excel.

---

## Stack

| Component | Technology |
|---|---|
| Language | TypeScript (ES modules) |
| MCP SDK | `@modelcontextprotocol/sdk`, stdio transport |
| ZIP | `fflate` |
| XML | `fast-xml-parser` (reading), string templates (writing) |
| VBA project (CFB) | `cfb` |
| Excel automation | Windows PowerShell 5.1 + COM (operate tools only) |
| Runtime | Node.js 22+ |
| Tests | `node --test`, fixtures built from scratch |

---

## Source layout

```
src/
  index.ts           MCP server: tool definitions, argument checks, dispatch
  stdio.ts           stdio entry point
  workbook/          reading
    package.ts       the .xlsx/.xlsm package: parts, relationships, content types
    afo.ts           the AfO definition: data sources, planning objects, items, settings
    bics.ts          compressed blocks and the stored navigation state (BICS_VIEW)
    sheets.ts        sheets, defined names, SAP cell formulas
    vba.ts           VBA project: MS-OVBA decompression, modules, procedures, AfO calls
    read.ts          afo_read_workbook: one workbook, sections, issues
    scan.ts          afo_scan_workbooks, afo_where_used
  build/             writing
    build.ts         afo_build_workbook: spec check, template or minimal layout, read-back
    afo.ts           the AfO definition as XML: data sources, grids, planning objects, settings
    source.ts        taking parts over verbatim from a source workbook
    cells.ts         values and formulas into sheet XML
    vba.ts           a VBA project from scratch (MS-OVBA compression, CFB)
    vbaPatch.ts      patching an existing VBA project: add, append, edit modules
    agent.ts         the agent bridge: sheet, help, VBA
    update.ts        afo_update_workbook
  operate/           Excel on the local machine
    open.ts          afo_open_workbook: Excel start, logon, prompts, refresh, save
    quick.ts         afo_quick_report
    landscape.ts     the SAP Logon landscape: connection of a system
```

---

## The workbook format

An AfO workbook is an Office Open XML package. Everything Analysis for Office adds lives in
places Excel keeps but does not interpret.

**The definition.** A sheet named `_com.sap.ip.bi.xl.hiddensheet` with `state="veryHidden"`
carries a custom property `serializedData2`. Its part (`xl/customPropertyN.bin`) is UTF-16LE
XML, root `<Application xmlns="http://www.sap.com/ip/bi/pioneer/excel/application">`:

```
Application/com.sap.ip.bi.pioneer
  Version, AAO_Revision, SystemId, LastRefreshed
  RefreshOnOpen, PlanningModeSetToChangeMode, Cleaned, ForcePromptOnInitialRefresh,
  StorePromptsInDocument, MergeVariables, PromptsVisibility, StyleSetName
  PlanningSystemDesignTime/SystemInfo        the connection (server, SAProuter, SNC)
  DataProviders/DataProvider                 Key (alias), DataSourceInfo, SystemInfo,
                                             BICS_VIEW_HEADER (compressed navigation state)
  PlanningObjects/PlanningFunctionUsage …    Key, name, data source, variables, filter
  Items/Grid, Items/Fdd …                    crosstabs and formula items with their range
```

`Cleaned` is the setting *remove data before saving*: on, AfO empties every crosstab when the
workbook is saved.

**Compressed blocks.** Navigation states, planning variables and planning filters are stored as
`compressed="true"` elements: base64 of a 4-byte little-endian length followed by a gzip stream
of UTF-8 XML. The navigation state (`BICS_VIEW`) holds the axes, the filter per characteristic
(members, hierarchy nodes, intervals, exclusions), hierarchies, display settings, design rules,
exceptions, conditions and — when `StorePromptsInDocument` is on — the variable values.

**Cells.** Every crosstab and formula item is tied to its range by a defined name `SAP<Key>`;
AfO coordinates are zero-based and converted to A1. SAP cell formulas are stored with the add-in
prefix (`_xll.SAPGetData(…)`); a formula without it is not recognised as an AfO function.

**Cell styles.** Crosstabs are drawn with the workbook's `SAP*` cell styles (`SAPDataCell`,
`SAPMemberCell`, `SAPDimensionCell`, …). A workbook may bring its own; AfO keeps them on refresh.
That is how a corporate design reaches the crosstabs.

**VBA.** `xl/vbaProject.bin` is a compound file (CFB) holding the project per MS-OVBA: a `dir`
stream describing the modules, one stream per module with compressed source, and caches.

---

## Reading

Reading never needs Excel. `WorkbookPackage` unzips the file once; the definition is located
through the hidden sheet's relationship, decoded from UTF-16 and parsed into data sources,
planning objects and items. Compressed blocks are expanded on demand. Sheets give the SAP
formulas with parsed arguments, and the VBA project is decompressed to list procedures and the
AfO API calls in them. Issues are links that do not resolve: an alias a formula or macro names
but no data source has, a button whose macro does not exist, a crosstab without its data source.

`afo_scan_workbooks` reads every workbook of a folder through the same reader and aggregates:
usage per query, planning function and variable, recurring alias-to-query bindings, and
outliers. `afo_where_used` searches one name through every place a name can occur, VBA source
included.

---

## Building

`buildWorkbook` runs in four steps.

1. **Check** — the spec is validated as a whole and every problem is reported at once: missing
   fields, duplicate or unknown keys, references to data sources that do not exist, `from`
   without a source workbook, malformed cells and VBA names, button macros without their module.
   Nothing is written when anything is wrong.
2. **Base** — a layout template (sheets, texts, formats, button shapes), an existing AfO
   workbook (its definition is replaced, its sheet ids kept), or a generated minimal layout.
3. **Definition** — the AfO XML is generated from the spec. Entries with `from` are copied
   verbatim from the source workbook's definition (`source.ts`): only the key and the references
   to renamed keys change, and a crosstab gets its new range. Settings follow the rule *spec
   wins, then the source, then the default*. Prompt values are written into the stored
   navigation state of their data source.
4. **Write and read back** — sheets, defined names, custom property, content types and the VBA
   project are written; the result is read back with the reader and broken links are returned.

**VBA.** A new project is written from scratch with MS-OVBA compression. An existing project is
patched (`vbaPatch.ts`): modules are added to the `dir` stream, every module becomes
source-only (offset 0), the `_VBA_PROJECT` version is set to `0xFFFF` and the `__SRP_*` caches
are removed, so Excel recompiles from source on open instead of trusting stale p-code.

---

## The agent bridge

**The problem.** Claude in Excel writes cells through Office.js. Such writes raise no `Change`
event in VBA, and Analysis for Office registers input through exactly that event: a value
written by the agent stays in the cell but never reaches the plan buffer. Office.js writes do,
however, trigger a recalculation.

**The chain.**

```
agent writes command into AgentLog!B1 (Office.js)
  → formula =LEN(B1) recalculates
  → Workbook_SheetCalculate sees a new command
  → Application.OnTime schedules the runner (outside the calculation)
  → VBA executes the command: Range.Value for SET (a real Change event, AfO registers it),
    AfO API for the rest (SAPExecuteCommand, SAPExecutePlanningFunction, SAPSetVariable, …)
  → new AfO messages (SAPListOfMessages) are compared with those before the step
  → result, status and log are written back to the bridge sheet
```

A step fails when AfO reports an error or a value was rejected. `PlanDataTransfer` or `Reset`
answering *not enabled* means nothing is pending and counts as success. The bridge sheet also
holds the help text with commands and rules, live lists of data sources, planning objects and
dimensions (AfO list formulas, initialised on open), and the last 500 log lines.

**Retrofit.** `afo_update_workbook` adds the sheet, the defined names and a VBA module to an
existing workbook and hooks one call into existing `Workbook_Open` and
`Workbook_SheetCalculate` handlers instead of replacing them.

---

## Operating

The operate tools run a PowerShell script per call; the server itself never loads COM.

1. **Excel outside the server.** Excel is started through WMI (`Win32_Process.Create`). Its
   parent is then the WMI service, not the server, so it survives the server, the client and
   any job object the client puts its processes in. It stays open for the user.
2. **Attach.** The script finds the Excel window of that process and gets the `Application`
   object through `AccessibleObjectFromWindow` on the workbook window.
3. **AfO ready.** It waits until the AfO API answers (`SAPGetProperty`).
4. **Logon.** `SAPLogon(alias, client, user, password)`. The password reaches only this script,
   through its environment: with a secret runner (`AFO_SECRET_COMMAND`) the script is started
   through the runner, which puts the entry's secrets into the script's environment, so the
   server never holds one. The script clears the variable once read and never prints it.
   While Excel shows the security bar, AfO refuses the logon; the script retries for a while.
5. **Refresh, variables, refresh.** The first refresh initialises the data sources —
   `SAPSetVariable` returns 0 before it. Variables are then set with refresh behaviour off and
   applied in one refresh.
6. **Report.** Variable values, used ranges, AfO messages and visible dialogs come back as JSON.

**The connection of an ad-hoc report** comes from the SAP Logon landscape of the machine
(`SAPUILandscape.xml` and the global landscape next to it): server and system number from the
service entry, the SAProuter string through its `routerid`, the SNC name from `sncname`. The
report therefore connects exactly as SAP Logon does, including SAProuter and SNC.

---

## Testing

`npm test` runs the suites in `tests/` against fixtures generated in the tests themselves — no
customer workbook is ever part of the repository. The build tests read every written workbook
back and compare it with the spec; the round trip of the take-over path was additionally checked
element by element on real workbooks outside the repository. The operate tools are tested for
their configuration and landscape parsing; driving Excel needs a machine with Excel and AfO and
is checked there by hand.
