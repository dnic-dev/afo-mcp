# afo-mcp

A Model Context Protocol (MCP) server for SAP Analysis for Office (AfO) workbooks.

- **Read** workbooks without Excel and without an SAP logon: data sources and their saved state, planning objects, crosstabs, SAP formulas, VBA macros, for one workbook or a whole folder. *(available)*
- **Operate** workbooks through Excel and the AfO API on Windows: variables, filters, refresh, planning functions and sequences, saving plan data. *(planned)*
- **Build** workbooks without Excel: data sources with stored prompts and axes, crosstabs, planning functions, VBA modules and button macros, on a layout template or a generated minimal layout. *(available)*

## Tools

| Tool | What it does |
| --- | --- |
| `afo_read_workbook` | Everything about one workbook: settings, sheets, data sources with filters, axes and hierarchies, planning functions and sequences with binding, aggregation level, filter, variable values and references, crosstabs and formula items with sheet and range, SAP cell formulas with parsed arguments, VBA modules with procedures and AfO API calls, buttons and the macros they run. Broken links are listed as issues. |
| `afo_scan_workbooks` | Inventory of a folder: a profile per workbook, which query, planning function and variable is used in which workbooks, recurring layouts, and outliers (system, AfO version, unreadable files). |
| `afo_where_used` | Every place a technical name occurs across a folder: data source, alias, planning object, aggregation level, filter, variable, characteristic, SAP formulas and VBA code. |
| `afo_update_workbook` | Change an existing AfO workbook instead of rebuilding it: add planning functions, retrofit the agent bridge, add VBA modules. Everything else stays as it is; the existing VBA project is patched, not replaced. |
| `afo_build_workbook` | Write a new AfO workbook from a JSON spec onto a layout template (.xlsx/.xlsm with sheets, texts, formats and button shapes) or a generated minimal layout. Checks the spec first and reports every problem at once, then reads the result back and lists broken links. The facts a spec needs (variables, aggregation level characteristics, structure UIDs) come from the BW system, e.g. through bw-modeling-mcp. |

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

## Status

Reading and building are in place and tested against synthetic fixtures; built workbooks were checked in Excel with Analysis for Office. Operating workbooks through Excel is next. The plan for agentic planning with Claude in Excel and the agent bridge is in [docs/ROADMAP.md](docs/ROADMAP.md).
