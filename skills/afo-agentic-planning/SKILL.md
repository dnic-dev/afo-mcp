---
name: afo-agentic-planning
description: >-
  Plan in SAP Analysis for Office (AfO) workbooks from Claude in Excel: understand the workbook
  (data sources, planning functions, crosstabs), change plan values so that AfO accepts them, run
  planning functions and sequences, navigate, check results and save, through the agent bridge
  (sheet AgentLog, defined name AfoAgentCommand). Use whenever the open workbook is an AfO
  workbook or the user asks to plan, copy actuals, revalue, distribute, recalculate or save plan
  data in Excel. Combine with the BW connector (bw-modeling-mcp) for data and definitions that are
  not on the sheet.
---

# Agentic planning in Analysis for Office workbooks

Analysis for Office (AfO) connects Excel to SAP BW. In a planning workbook, crosstabs show plan
data; the planner types values, runs planning functions and saves. **Values you write through
Office.js are not seen by AfO**: the cell shows them, but recalculation and saving use the old
values. Planning therefore goes through the **agent bridge**, a command cell that a macro in the
workbook executes as if a planner typed or clicked.

## 1. Find out what you work with

- **Bridge present?** Look for the defined name `AfoAgentCommand` (sheet `AgentLog`, also when
  the sheet is hidden). `AfoAgentStatus` (B4) must say `active`; otherwise macros are disabled:
  ask the user to enable content and reopen.
- **No bridge:** you can read and explain, but not plan. Do not write into crosstabs; tell the
  user that writes would not reach SAP and that the bridge can be added to the workbook.
- **Read the bridge sheet first.** It lists the commands, the rules, the workbook's data sources,
  planning functions and crosstabs (also as live lists from column Q), AfO messages (from column F)
  and a log of every step.

## 2. Read

AfO read functions are worksheet formulas; write them into the free area of the bridge sheet
(from AD2), read the result, delete them afterwards:

| Formula | Returns |
| --- | --- |
| `=SAPListOf("DATASOURCES")`, `"PLANNINGFUNCTIONS"`, `"PLANNINGSEQUENCES"`, `"CROSSTABS"` | the workbook's objects |
| `=SAPListOfVariables("DS_1")` | variables and values of a data source |
| `=SAPListOfVariables("PF_1",,"PLAN_PARAMETER")` | variables of a planning function |
| `=SAPListOfEffectiveFilters("DS_1")` | current filters |
| `=SAPListOfDimensions("DS_1")` | dimensions and their axis |
| `=SAPListOfMembers("DS_1","DIMENSION")` | members of a dimension |

Map data cells to members by the row and column headers of the crosstab. Members repeat only
in the first row of a block: carry them down.

For anything not on the sheet (previous years, other queries, master data, the logic of a
planning function, the query definition) use the BW connector with the technical names from the
bridge sheet.

## 3. Change plan data

1. Work out the target values yourself (percentages, distributions, rounding). Plan values with
   decimals are rounded by the server to its precision.
2. **Ambiguous?** A member under several parents, several matching rows, an unclear period:
   ask before writing.
3. Write **one command** with all steps, separated by `;`, into `AfoAgentCommand`:
   `SET Planung!I56=21150; SET Planung!I90=28998; TRANSFER`
   Decimal point, no thousands separator, input-ready cells only.
4. Wait until `AfoAgentResult` (B3) no longer says `running`, then read it:
   `OK …` or `ERROR in step n (…): message`. Rejected values (data slices, characteristic
   relationships, locks) turn the step into ERROR.
5. Recalculate derived values with the workbook's command or planning function (see the bridge
   sheet), then check a few cells against your own calculation.
6. Many cells (hundreds or more): prefer a planning function with variables
   (`PVAR PF_1 NAME=value; PLAN PF_1`) over SET.

## 4. Commands

| Command | Does |
| --- | --- |
| `SET Sheet!A1=value` | input value as typed by the planner |
| `TRANSFER` | send and check changed values |
| `PLAN key` | run planning function or sequence (`PF_1`, `PS_1`, …) |
| `PVAR key NAME=value` | variable of a planning function or sequence |
| `VAR DS_1 NAME=value` | variable of a data source (external format: `01.2025`, `01.2025 - 12.2025`) |
| `FILTER DS_1 DIMENSION=value` | filter a characteristic |
| `REFRESH` | refresh (set mandatory variables with VAR first) |
| `SAVE` / `RESET` | save / discard plan data |
| `API SAPFunction arg1\|arg2` | any AfO API function, e.g. `API SAPMoveDimension DS_1\|DIM\|ROWS` |
| workbook commands | listed on the bridge sheet, e.g. `IST_UEBERNEHMEN` |

## 5. Rules

- **Never write plan cells directly**; always SET through the bridge.
- **Ask before SAVE and RESET.** Summarise what changed (cells, totals before and after).
- Say what you did and what you did not check.
- `running` for more than 30 seconds: a dialog is probably open in Excel; ask the user to look.
- Never put credentials into cells; logon is the user's job.
