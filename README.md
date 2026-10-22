# afo-mcp

A Model Context Protocol (MCP) server for SAP Analysis for Office (AfO) workbooks.

- **Read** workbooks without Excel and without an SAP logon: data sources and their saved state, planning objects, crosstabs, SAP formulas, VBA macros, for one workbook or a whole folder. *(available)*
- **Operate** workbooks through Excel and the AfO API on Windows: open a workbook or an ad-hoc query report logged on, with prompt values set and the data refreshed. *(available)* Filters, planning functions and saving plan data from outside Excel. *(planned; inside Excel the agent bridge does this)*
- **Build** workbooks without Excel: data sources with stored prompts and axes, crosstabs, planning functions, VBA modules and button macros, on a layout template or a generated minimal layout. *(available)*

## Tools

| Tool | What it does |
| --- | --- |
| `afo_read_workbook` | Everything about one workbook: settings, sheets, data sources with filters, axes and hierarchies, planning functions and sequences with binding, aggregation level, filter, variable values and references, crosstabs and formula items with sheet and range, SAP cell formulas with parsed arguments, VBA modules with procedures and AfO API calls, buttons and the macros they run. Broken links are listed as issues. |
| `afo_scan_workbooks` | Inventory of a folder: a profile per workbook, which query, planning function and variable is used in which workbooks, recurring layouts, and outliers (system, AfO version, unreadable files). |
| `afo_where_used` | Every place a technical name occurs across a folder: data source, alias, planning object, aggregation level, filter, variable, characteristic, SAP formulas and VBA code. |
| `afo_update_workbook` | Change an existing AfO workbook instead of rebuilding it: store prompt values (same workbook, other preset selection), add planning functions, retrofit the agent bridge, add VBA modules. Everything else stays as it is; the existing VBA project is patched, not replaced. |
| `afo_build_workbook` | Write a new AfO workbook from a JSON spec onto a layout template (.xlsx/.xlsm with sheets, texts, formats and button shapes; an AfO workbook is accepted too, `"source"` uses the source workbook) or a generated minimal layout. Checks the spec first and reports every problem at once, then reads the result back and lists broken links. The facts a spec needs (variables, aggregation level characteristics, structure UIDs) come from the BW system, e.g. through bw-modeling-mcp. |
| `afo_open_workbook` | Open a workbook in Excel on this machine for the user: logged on with the credentials configured for its system, prompt values set, data refreshed, saved on request. Excel stays open. Needs Windows, Excel and AfO. |
| `afo_quick_report` | Ad-hoc report for one query: a plain .xlsx with the query in a crosstab and the prompt values stored, connected as the SAP Logon landscape describes the system (server, SAProuter, SNC), opened logged on with the data. |

## How it reads a workbook

An `.xlsx`/`.xlsm` file is a ZIP package. AfO stores its definition as a custom property of a hidden sheet: a UTF-16 XML document listing data providers, planning objects and items. Navigation state and variable values sit inside it as compressed blocks (base64, a 4-byte length, gzip). Sheets carry an AfO id (`<code name>|<guid>`) that items point to. Macros come from `xl/vbaProject.bin` (a compound file with MS-OVBA compressed module streams), buttons from the drawing parts.

What a workbook does not contain: the query definition itself, and query variable values unless the workbook setting *StorePromptsInDocument* is on. Structure members appear as BW-internal IDs.

## Setup

Requires Node.js 22 or later.

```sh
npm install
npm run build
npm test
```

MCP client configuration (stdio):

```json
{
  "mcpServers": {
    "afo-mcp": { "command": "node", "args": ["C:/path/to/afo-mcp/dist/stdio.js"] }
  }
}
```

Logon for the operate tools (environment of the server, never tool arguments):

```json
"env": {
  "AFO_LOGON_SYSTEMS": "[{\"system\":\"SYS\",\"client\":\"100\",\"user\":\"USER\",\"language\":\"EN\",\"vault\":\"entry@sys\",\"passwordVar\":\"BW_PASSWORD\"}]",
  "AFO_VAULT_PYTHON": "C:/path/to/python.exe",
  "AFO_VAULT_STARTER": "C:/path/to/vault-starter.py"
}
```

With `vault`, the logon runs in a child process started through a vault starter (`<python> -I <starter> run <entry> -- <command>`) that puts the entry's secrets into that child's environment only. Without `vault`, the password is read from the variable `passwordVar` of the server's environment. A workbook whose content Excel has disabled (security bar) refuses the logon until the user enables it; a trusted location avoids that, and ad-hoc reports are .xlsx without macros.

## Status

Reading and building are in place and tested against synthetic fixtures; built workbooks were checked in Excel with Analysis for Office. Opening workbooks and ad-hoc reports logged on works on Windows with Excel and AfO; more operate tools are next. The plan for agentic planning with Claude in Excel and the agent bridge is in [docs/ROADMAP.md](docs/ROADMAP.md).
