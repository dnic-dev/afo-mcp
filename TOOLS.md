# Tools Reference

Every tool the server exposes — 7 in total — with what it does and the parameters that matter.
Tools marked _(Read only)_ write nothing. The build tools write a new file and never touch their
input; the operate tools drive Excel on the local machine and reach the BW system through
Analysis for Office.

For what the server can do as a whole, see [What it can do](README.md#what-it-can-do) in the README.
For the file format and the inner workings, see [ARCHITECTURE.md](ARCHITECTURE.md).

---

## Contents

- [Read](#read) — 3 tools, no Excel needed
- [Build](#build) — 2 tools, no Excel needed
- [Operate](#operate) — 2 tools, Windows with Excel and Analysis for Office
- [The agent bridge](#the-agent-bridge) — commands for Claude in Excel inside a workbook
- [How the tools combine](#how-the-tools-combine)

---

## Read

### `afo_read_workbook` _(Read only)_
Describe one workbook completely, without Excel and without a logon.

- **Settings** — system, AfO version, refresh on open, change mode on open, prompt settings, style set, the connection per system.
- **Data sources** — query or InfoProvider, input readiness, and the saved navigation state: filter per characteristic (members, hierarchy nodes, exclusions), rows, columns and free characteristics, hierarchies, stored variable values. `stateStored: false` means the data source opens in the query's default view.
- **Planning objects** — planning functions and sequences with their data source, aggregation level, filter, stored variable values and the cells variables are read from.
- **Items** — crosstabs and formula items (filter components) with sheet and range.
- **Formulas** — SAP cell formulas (`SAPGetData`, `SAPGetVariable`, `SAPSetFilterComponent`, …) with parsed arguments.
- **Macros** — VBA modules with their procedures and the AfO API calls in them, buttons with the macro they run.
- **Issues** — broken links: a formula or macro naming an undefined data source alias, a button whose macro does not exist, a crosstab without its data source.

| Parameter | |
|---|---|
| `path` | Workbook (.xlsx, .xlsm) |
| `sections` | Subset of `settings`, `sheets`, `dataProviders`, `planningObjects`, `items`, `formulas`, `macros`; summary and issues always come |
| `macro_source` | Full VBA source of every module (default: procedures and AfO calls only) |
| `max_formulas_per_sheet` | Cap on listed SAP formulas per sheet (default 200); counts stay complete |

Structure members and some filter values appear as BW-internal IDs; `bw_get_query` of bw-modeling-mcp resolves them.

### `afo_scan_workbooks` _(Read only)_
Inventory a folder of workbooks in one call: a short profile per workbook, then the view across all of them that reading them one by one does not give —

- **Usage** — which query or data source is used in which workbooks; the same for planning functions and variables
- **Recurring layouts** — workbooks binding the same aliases to the same queries, i.e. copies of one design
- **Outliers** — another system or AfO version than the majority, unreadable files, files without an AfO definition

| Parameter | |
|---|---|
| `folder` | Folder to scan |
| `recursive` | Include subfolders |
| `limit` | Maximum number of workbooks (default 2000) |
| `detail` | `full` (profile per workbook, complete lists) or `compact` (no profiles, lists capped at five workbooks); default full up to 50 workbooks |

### `afo_where_used` _(Read only)_
Every place a technical name occurs across a folder: as query or InfoProvider, data source alias, planning function or sequence, aggregation level, planning filter, variable (prompt, planning variable, cell reference, `SAPGetVariable`), characteristic in filters and axes, argument of SAP formulas and formula items, and in VBA code. The question to ask before a query, variable or planning function changes. Matching is exact on the whole name, case-insensitive.

| Parameter | |
|---|---|
| `folder` | Folder to search |
| `name` | Technical name, or an alias such as `DS_1` |
| `recursive` | Include subfolders |

For workbooks stored in BW, `bw_list_afo_workbooks` of bw-modeling-mcp answers the same question from BW's own where-used list.

---

## Build

### `afo_build_workbook`
Write a new workbook from a spec — on a layout template, on an existing AfO workbook, or on a generated minimal layout. The spec is checked first and every problem is reported at once; nothing is written then. After writing, the file is read back and broken links are returned as issues. The workbook opens in Excel with Analysis for Office, which asks for the logon.

**What a spec describes**

| Spec field | |
|---|---|
| `dataProviders` | Alias, query, InfoProvider, input readiness, stored `prompts` (internal format; `member`, `literal` or `low`/`high`) `axes` (characteristics by name, a key figure structure by its runtime UID) and `filters` (per characteristic: `members`, `low`/`high` or `operator` + `value`, optionally `exclude`; internal keys) |
| `grids` | Crosstabs: data source, sheet, top-left cell. AfO does not move crosstabs apart — leave room for the rows each one grows to |
| `planningFunctions` | Planning function or sequence: alias, data source, aggregation level, its characteristics, fixed filter values, stored variables, variables read from cells |
| `formulaItems` | Filter components and other formula items; the formula is written into the cell as well |
| `cells` | Values and formulas into cells of the template, e.g. `=SAPGetVariable("DS_1","VAR","Value")` in a header |
| `system` | The connection as in the SAP Logon landscape: id, name, client, language, server, system number, `sapRouter`, `sncPartnerName` |
| `vbaModules`, `documentModules`, `shapeMacros` | Standard modules (code, file, or taken from the source), event handlers of sheets and `ThisWorkbook`, button shapes → macro |
| `agentBridge` | The bridge sheet for Claude in Excel, see [The agent bridge](#the-agent-bridge) |
| Settings | `refreshOnOpen`, `removeDataOnSave`, `changeModeOnOpen`, `storePrompts`, `mergeVariables`, `styleSetName` |

**A new layout for an existing workbook.** With `source` set to an existing AfO workbook, every entry with `from` is taken over verbatim — a data source with its whole stored state (filters incl. hierarchy nodes, display settings, design rules, exceptions, conditions, prompts), a crosstab with its settings, a planning function with variables, filter and references, a filter component with its formula. Keys may be renamed; references follow. The connection (SAProuter, SNC) and the workbook settings come from the source unless the spec sets them, VBA modules can be taken over with text replacements (`{ from, replace }`). Only the position on the new layout is new.

| Parameter | |
|---|---|
| `output` | File to write, `.xlsm` (VBA) or `.xlsx`; the folder must exist |
| `spec` / `spec_path` | The spec, inline or as a JSON file (relative paths resolve against its folder) |
| `template` | Layout workbook with sheets, texts, formats and button shapes; an AfO workbook is accepted as it is; `"source"` uses the source workbook |
| `layout` | Generated minimal layout when there is no template: sheets, cell values, defined names |
| `overwrite` | Replace an existing output |

### `afo_update_workbook`
Change an existing workbook instead of rebuilding it. Layout, cells, data sources, crosstabs and existing macros stay as they are; the VBA project is patched, not replaced, and existing `Workbook_Open` / `Workbook_SheetCalculate` handlers get one call added. Writes a new file, never the input.

| Parameter | |
|---|---|
| `input`, `output` | Existing workbook and the changed copy |
| `prompts` | Variable values by data source: `{ "DS_1": { "VAR_A": "KEY", "VAR_B": { "low": "202501", "high": "202512" } } }`. The workbook then opens in that selection without the prompt dialog. A data source with "reset data source" gets it switched off, one without stored state a minimal one |
| `storePrompts` | `StorePromptsInDocument`, default true when prompts are given |
| `refreshOnOpen` | `RefreshOnOpen`. Off lets `afo_open_workbook` log on before the first refresh (see there); users opening the workbook by hand then refresh themselves |
| `planningFunctions` | Planning functions bound to a data source already in the workbook |
| `agentBridge` | Retrofit the agent bridge |
| `vbaModules` | Additional standard modules |
| `system` | Point the workbook to another system — development to quality to production, or after a system move. `{ "systemId": "QAS" }` takes server, system number, SAProuter and SNC from the SAP Logon landscape of the machine (client and language stay unless given); `{ "from": "<workbook>" }` copies the connection of a reference workbook. Every data source and planning object gets the new connection |

Warnings name variables that still have no stored value, and remind that queries, planning objects and variable values must exist in a new target system. Signed VBA projects are refused.

---

## Operate

Both tools need Windows with Excel and Analysis for Office. Excel is started outside the server's process tree and stays open for the user. The logon uses the credentials configured for the workbook's system ([Configuration](README.md#configuration)); a workbook for a system without configured logon opens anyway, and the user logs on in Excel.

### `afo_open_workbook`
Open a workbook so the user sees the data at once: start Excel, log on, refresh once (which initialises the data sources), set variable values, refresh with them, save on request.

| Parameter | |
|---|---|
| `path` | Workbook to open |
| `prompts` | Variable values set with `SAPSetVariable`: `name`, `value`, `format` (`INTERNAL_KEY` default, or `INPUT_STRING` as typed in the prompt), `dataProvider` (default the first) |
| `logon` | Log on with the configured credentials (default true) |
| `save` | Save after the refresh — with the data when the workbook keeps data on save (`removeDataOnSave: false`) |
| `logonWaitSeconds` | How long to retry the logon while Excel shows the security bar (default 90) |
| `keepRefreshOnOpen` | Open a workbook with *refresh on open* as it is even when logging on (default false, see below) |
| `timeoutSeconds` | Overall timeout (default 900) |

Returns the Excel process id, the steps, the logon result, variable values after the refresh, used ranges per sheet and AfO messages.

**Refresh on open.** A workbook with *refresh on open* refreshes while it loads, and Analysis for Office asks for its own logon before `SAPLogon` can run. The data sources then report connected but stay inactive: variables show `#N/A`, crosstabs stay empty. When the tool logs on, it therefore opens a working copy with refresh on open off, `<name> (afo-mcp).xlsm` next to the workbook (or in the temp folder), and names it in `notes`. The original stays as it is; `file` in the result is the copy.

**Dialogs.** When Excel or AfO waits on a dialog — the prompt dialog for a mandatory variable without value, the AfO logon (with `keepRefreshOnOpen` or without configured logon), a message box, a repair or file-in-use prompt — the call stops after about 20 seconds without progress and returns `ok: false` with the title, text and buttons of the dialog. The dialog stays open in Excel for the user; nothing is clicked. The steps after it did not run.

**Security bar.** For a macro workbook outside a trusted location, Excel disables the content until the user clicks *Enable Content*, and Analysis for Office refuses the logon until then. The tool keeps retrying for `logonWaitSeconds`; tell the user beforehand, or keep such workbooks in a trusted location.

### `afo_quick_report`
An ad-hoc report for one query: a plain `.xlsx` with the query in a crosstab and the variable values stored, connected the way the SAP Logon landscape describes the system (server, system number, SAProuter, SNC name), opened logged on with the data. Without macros Excel shows no security bar, so the logon runs without the user.

| Parameter | |
|---|---|
| `query` | Technical name of the query |
| `system`, `client`, `language` | Default: the configured logon (when exactly one system is configured) |
| `prompts` | Variable values, internal format (`member`, `literal`, `low`/`high`) |
| `axes` | Optional layout: rows, columns, free characteristics |
| `filters` | Optional stored filters, as in `afo_build_workbook` |
| `folder` | Where the file goes (default `Documents\AfO Reports`) |
| `open` | `false` writes the file only |

Give every mandatory variable a value — `bw_get_query` lists the variables, `bw_get_filter_values` the valid keys — otherwise AfO shows the prompt dialog.

---

## The agent bridge

Not a tool of its own but a part of a workbook, added by `afo_build_workbook` or `afo_update_workbook` with `agentBridge`. It exists because Claude in Excel writes cells through Office.js, and Analysis for Office registers only user input: a plan value written by the agent is silently ignored.

The bridge sheet (default `AgentLog`, hidden by default) has a command cell (`B1`, defined name `AfoAgentCommand`), a result cell (`B3`, `AfoAgentResult`) and a status cell (`B4`, `AfoAgentStatus`), a help text, the AfO messages and a step log. Several commands go into one cell, separated by `;`.

| Command | |
|---|---|
| `SET Sheet!A1=value` | Write a value as user input; a range writes the same value into every cell. Follow with `TRANSFER` or a planning function |
| `TRANSFER` | Send changed values to the server and check them |
| `PLAN key` | Run a planning function or sequence by its alias |
| `PVAR key NAME=value` | Set a variable of a planning function or sequence |
| `VAR DS_1 NAME=value` | Set a query variable, external format as in the prompt |
| `FILTER DS_1 DIM=value` | Filter a characteristic, as in the filter dialog |
| `REFRESH` | Refresh all data sources, or one |
| `API SAPFunction a\|b` | Call any AfO API function; `SAPLogon` is excluded |
| `SAVE` / `RESET` | Save plan data / discard unsaved plan data |
| workbook commands | Own words that run a macro, e.g. the one behind a button |

After every step the new AfO messages are logged; an error message or a rejected value (data slice, characteristic relationship, lock) fails the step. The help text carries the rules an agent should follow: ask before `SAVE` and `RESET`, ask when a selection is ambiguous, write only input-ready cells, prefer planning functions for mass changes.

Reading needs no command. The bridge sheet explains the AfO read formulas (`SAPListOf`, `SAPGetCellInfo`, `SAPGetVariable`, …) and keeps live lists of data sources, planning objects and messages.

---

## How the tools combine

**Understand a folder of workbooks** — `afo_scan_workbooks` for the overview, `afo_where_used` for one name, `afo_read_workbook` for the details of one workbook.

**Give a workbook a new layout** — `afo_read_workbook` on the original; a layout template with the new design; `afo_build_workbook` with `source` and `from` for every data source, crosstab, planning object and filter component; `afo_read_workbook` on the result to check the links.

**Change a query, see what breaks** — `afo_where_used` on the query (or `bw_list_afo_workbooks` for workbooks stored in BW), then `afo_read_workbook` on each hit.

**Same workbook, other selection** — `afo_update_workbook` with `prompts`; or, to look at it right away, `afo_open_workbook` with `prompts`.

**Look at a query quickly** — `bw_get_query` for its variables, `bw_get_filter_values` for valid keys, then `afo_quick_report`.

**Planning workbook for Claude in Excel** — the planning model from bw-modeling-mcp (aggregation level, planning functions, sequence), then `afo_build_workbook` with `planningFunctions` and `agentBridge`; or `afo_update_workbook` with `agentBridge` on an existing workbook.
