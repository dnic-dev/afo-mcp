# afo-mcp

A Model Context Protocol (MCP) server that lets AI assistants like Claude work with **SAP Analysis for Office (AfO)** workbooks — reading what a workbook contains, building and changing workbooks, and opening them in Excel logged on with fresh data.

Three parts, each usable on its own:

- **Read** — a workbook or a whole folder, **without Excel and without an SAP logon**: data sources with their saved navigation state, variables, planning functions and sequences, crosstabs, SAP formulas, VBA macros and buttons, and the links between them that are broken.
- **Build** — new workbooks from a JSON spec, or a new layout for an existing workbook that keeps every AfO feature, **without Excel**: data sources with stored prompts, crosstabs, planning functions, filter components, VBA and button macros, and the **agent bridge** that lets Claude in Excel plan inside the workbook.
- **Operate** — on a Windows machine with Excel and AfO: open a workbook or an **ad-hoc report for any query**, logged on, with variables set and the data refreshed.

**Reading and building only touch files.** Nothing reaches a BW system until a workbook is opened in Excel with Analysis for Office — by the user, or by the operate tools on the user's own machine.

---

## 🆕 What's New — v0.1.0 · First release

**📖 Read workbooks without Excel**

- `afo_read_workbook` — everything one workbook contains, with broken links listed as issues
- `afo_scan_workbooks` and `afo_where_used` — inventory a folder, and find every workbook that
  uses a query, variable or planning function before changing it

**🏗️ Build and change workbooks without Excel**

- `afo_build_workbook` — a workbook from a spec, on a layout template or a generated layout
- A **new layout for an existing workbook**: data sources, crosstabs, planning objects, filter
  components, connection and settings taken over verbatim from the source workbook (`source`
  and `from`), VBA modules included — verified element by element on 30 real workbooks
- `afo_update_workbook` — store variable values, add planning functions, retrofit the agent
  bridge, add VBA modules, keeping everything else

**🤖 Agentic planning with Claude in Excel**

- The **agent bridge**: a command cell through which Claude in Excel writes plan values that AfO
  accepts, runs planning functions, sets variables and filters, and saves — commands any agent
  can learn from the help on the bridge sheet

**🖥️ Open workbooks and ad-hoc reports logged on**

- `afo_open_workbook` — open a workbook in Excel, logged on, variables set, data refreshed
- `afo_quick_report` — an ad-hoc report for one or more queries, on a layout template if
  wanted, connected the way SAP Logon connects
- Logon per system from the server's environment or a secret runner; passwords never pass
  through tool arguments

The full history is in [CHANGELOG.md](CHANGELOG.md); later releases archive their notes in [WHATS_NEW.md](WHATS_NEW.md).

---

## What it can do

Every tool in detail — parameters, behaviour, and how the tools combine — is in the **[Tools Reference](TOOLS.md)** (7 tools).

| Area | Tool | Needs Excel and AfO |
|---|---|---|
| **Read** | `afo_read_workbook` — one workbook, completely | no |
| | `afo_scan_workbooks` — a folder: profiles, usage across workbooks, outliers | no |
| | `afo_where_used` — every place a technical name occurs across a folder | no |
| **Build** | `afo_build_workbook` — a new workbook, or a new layout for an existing one | no |
| | `afo_update_workbook` — change an existing workbook in place of rebuilding it | no |
| **Operate** | `afo_open_workbook` — open a workbook logged on, with variables and fresh data | yes (Windows) |
| | `afo_quick_report` — ad-hoc report for one or more queries, opened logged on | yes (Windows) |

### Agentic planning with Claude in Excel

Claude in Excel writes cells through Office.js, and Analysis for Office does not register such writes as user input — a plan value typed by the agent is silently ignored. The **agent bridge** closes that gap. `afo_build_workbook` and `afo_update_workbook` add a bridge sheet with a command cell; the agent writes a command such as `SET Plan!E12=1500` or `PLAN PF_1` into it, and VBA in the workbook carries it out as a user would. AfO messages are logged after every step, and a rejected value fails the step.

| Command | What it does |
|---|---|
| `SET Sheet!A1=value` | Write a value as user input |
| `SETDATA DS Key figure\|DIM=key;…\|value` | Write a plan value by member combination, wherever the crosstab shows it |
| `NEWLINES Crosstab n` · `CHANGEMODE DS` | New lines for new member combinations / switch to change mode |
| `TRANSFER` | Send changed values to the server and check them |
| `PLAN key` / `PVAR key NAME=value` | Run a planning function or sequence / set its variables |
| `VAR DS NAME=value` / `FILTER DS DIM=value` | Set a query variable / filter a characteristic (also on a planning function) |
| `MOVE DS DIM=ROWS` · `EXPAND DS node` · `DRILLLEVEL DS DIM=n` | Navigate: move characteristics, open hierarchy nodes and levels |
| `REFRESH` · `SAVE` · `RESET` | Refresh, save plan data, discard unsaved plan data |
| `API SAPFunction a\|b` | Any AfO API function (except the logon) |
| your own words | Workbook commands that run the macro behind a button |

Reading needs no command: the bridge sheet explains the AfO read formulas and keeps live lists of data sources, planning objects and messages.

**The skill for Claude in Excel.** [`skills/afo-agentic-planning`](skills/afo-agentic-planning/SKILL.md) tells Claude in Excel how to recognise an AfO workbook, find the bridge, read the workbook's queries and planning objects through AfO formulas, and when to ask before acting. Claude in Excel uses the skills of the claude.ai account it is signed in with: add the folder there as a skill.

---

## Combining with bw-modeling-mcp

**afo-mcp works best alongside [bw-modeling-mcp](https://github.com/dnic-dev/bw-modeling-mcp)**, which reads and writes the BW side. The two divide the work cleanly:

- A spec needs facts from BW — the variables of a query and their valid values, the characteristics of an aggregation level, the runtime UID of a key figure structure. bw-modeling-mcp delivers them (`bw_get_query`, `bw_get_filter_values`, `bw_get_aggregation_level`, `bw_query_data`); afo-mcp has no BW access of its own.
- Workbooks stored in BW are fetched and saved by bw-modeling-mcp (`bw_list_afo_workbooks`, `bw_get_afo_workbook`, `bw_save_afo_workbook`) and read, changed or rebuilt here in between.
- A planning model built with bw-modeling-mcp — aggregation level, planning functions, sequence — becomes a planning workbook here, with the agent bridge for Claude in Excel.

---

## Requirements

- Node.js 22 or later
- An MCP-compatible AI client (Claude Desktop, Claude Code, etc.)
- For the operate tools only: Windows with Microsoft Excel and SAP Analysis for Office, and the systems in the SAP Logon landscape of that machine

---

## Installation

```bash
git clone https://github.com/dnic-dev/afo-mcp.git
cd afo-mcp
npm install
npm run build
npm test
```

---

## Configuration

Reading and building need no configuration. The operate tools log on with credentials from the server's environment — never from tool arguments, so a password never appears in a conversation.

| Variable | Description | Required |
|---|---|---|
| `AFO_LOGON_SYSTEMS` | JSON list of logons, one per system: `system`, `client`, `user`, optional `language`, and either `vault` (an entry of a secret runner) or nothing (password from this environment). `passwordVar` names the variable that holds the password (default `BW_PASSWORD`). | for operate |
| `AFO_SECRET_COMMAND` | JSON array: command prefix of a secret runner that starts a command with the secrets of an entry in its environment; `{vault}` stands for the entry, e.g. `["op", "run", "--env-file", "{vault}.env", "--"]`. With it, the logon runs in a short-lived child process and the server itself never holds a password. | with `vault` |
| `AFO_LOGON_SYSTEM`, `AFO_LOGON_CLIENT`, `AFO_LOGON_USER`, `AFO_LOGON_LANGUAGE`, `AFO_LOGON_PASSWORD_ENV` | The same for a single system; the password is read from `AFO_LOGON_PASSWORD` or the variable `AFO_LOGON_PASSWORD_ENV` names. | alternative |

Server, system number, SAProuter and SNC name are not configured here: they come from the SAP Logon landscape of the machine (`%APPDATA%\SAP\Common\SAPUILandscape.xml`), so a workbook connects the way SAP Logon does.

### Claude Desktop / Claude Code

```json
{
  "mcpServers": {
    "afo-mcp": {
      "command": "node",
      "args": ["/path/to/afo-mcp/dist/stdio.js"],
      "env": {
        "AFO_LOGON_SYSTEMS": "[{\"system\":\"ABC\",\"client\":\"100\",\"user\":\"YOUR_USER\",\"language\":\"EN\",\"vault\":\"abc\"}]",
        "AFO_SECRET_COMMAND": "[\"op\", \"run\", \"--env-file\", \"C:/secrets/{vault}.env\", \"--\"]"
      }
    }
  }
}
```

Claude Desktop reads `claude_desktop_config.json`, Claude Code `.mcp.json` in the project or home folder. Without the `env` block the read and build tools work as they are.

---

## How it works

**An AfO workbook is an ordinary Excel file with a second, hidden document inside.** Analysis for Office keeps its definition — data sources, planning objects, crosstabs, settings — as UTF-16 XML in a custom property (`serializedData2`) of a very hidden sheet, `_com.sap.ip.bi.xl.hiddensheet`. Navigation states and variable values sit in compressed blocks inside that XML (base64 of a 4-byte length followed by gzip). Crosstabs and filter components are tied to cells through defined names, and SAP cell formulas are stored as `_xll.SAP…`.

Reading unzips the file and parses that definition next to the sheets and the VBA project. Building writes the same structures; VBA projects are written per MS-OVBA, and existing ones are patched rather than replaced, so a workbook keeps its macros. The result opens in Excel like a workbook AfO saved itself.

Operating needs Excel. The tools start Excel outside the server's process tree, attach to it through COM and drive Analysis for Office through its own API (`SAPLogon`, `SAPSetVariable`, `SAPExecuteCommand`, …) — the same calls a VBA macro would make.

See [ARCHITECTURE.md](ARCHITECTURE.md) for the file format, the build pipeline, the agent bridge and the operate layer in detail.

---

## Roadmap

The plan is in [docs/ROADMAP.md](docs/ROADMAP.md); open work is tracked as [issues](https://github.com/dnic-dev/afo-mcp/issues). Next up:

- **Describe, not only take over** — filters with hierarchy nodes, design rules and planning sequences in the spec

Ideas for later:

- **Planning parcours** — input on hierarchy nodes, data slices, characteristic relationships, sequences and new lines, tested end to end on a test model built with bw-modeling-mcp
- **The agent bridge as an Excel add-in** — installed once per user instead of built into each workbook, so any AfO workbook, including existing ones, can be operated by an agent without changing the file

---

## Contributing

Issues and feature requests are welcome. The most useful contribution is to try it on your own workbooks and report what works, what doesn't, and what's missing. Real workbooks often hold customer data — please describe them with generic names rather than attaching them.

---

## Disclaimer

afo-mcp is an independent open-source project. It is not affiliated with, endorsed by or
supported by SAP SE. SAP, SAP BW/4HANA and SAP Analysis for Office are trademarks or
registered trademarks of SAP SE in Germany and other countries.

The server reads and writes the format Analysis for Office stores in workbook files,
derived from saved workbooks, and drives Analysis for Office through its documented API.
It contains no SAP code. The format is not documented by SAP and may change with new
Analysis for Office releases.

The operate tools and the agent bridge change data in your BW system as a user would. Use
them at your own risk, and test with non-productive systems first.

---

## License

MIT
