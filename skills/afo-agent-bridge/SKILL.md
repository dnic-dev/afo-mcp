---
name: afo-agent-bridge
description: Operate SAP Analysis for Office (AfO) workbooks in Excel through their agent bridge - filter, navigate, set variables, enter plan values, run planning functions, save. Use whenever the open workbook has a sheet "AgentLog" or the defined name AfoAgentCommand, or shows AfO crosstabs (SAP* cell styles, data sources DS_1 ...), and the user asks to filter, show, select, change, plan, refresh or save anything in it. Always prefer the bridge over Excel features (AutoFilter, sorting, hiding, formulas) on AfO crosstabs.
---

# Operating Analysis for Office workbooks

Analysis for Office ignores values that an add-in writes into cells (Office.js raises no Change
event), and it rebuilds every crosstab on refresh. So in an AfO workbook nothing is done with
Excel features. Everything goes through the **agent bridge**: a command cell that VBA turns into
calls of the AfO API, as if the user had clicked.

## Find the bridge

- Defined names: `AfoAgentCommand` (command cell, B1 of the hidden sheet `AgentLog`),
  `AfoAgentResult` (B3), `AfoAgentStatus` (B4).
- B4 must read `active`. Otherwise macros are disabled: ask the user to click "Enable content"
  and to reopen if needed. Do not work around it with Excel.
- The bridge sheet lists the commands, the rules, the data sources, planning functions and
  crosstabs of this workbook, and logs every step with the AfO messages. Read it once.

## Run commands

1. Write the command into `AfoAgentCommand`. Several commands: **one per line** (line breaks in
   the cell). A single line may also separate commands by `;`, but never when a value contains `;`.
2. Poll `AfoAgentResult` every second until it no longer starts with `running`.
   `OK ...` = done, `ERROR ... in step n (...)` = that step failed, with the reason.
3. Read the log rows below the commands for details and AfO messages; read the crosstab to check
   the effect.

| Command | Use |
|---|---|
| `FILTER DS_1 DIM=value` | Filter a characteristic by key, interval `a - b`, several keys `a;b;c` (own line), or `@Sheet!A1:A50` for keys in a range. Empty value removes the filter. |
| `FILTERWHERE DS_1 DIM "Key figure" >= 100000` | Filter a characteristic to the members whose crosstab value meets the condition (`=`, `<>`, `<`, `<=`, `>`, `>=`). Key figure = column header text. Result is a real AfO filter on member keys and survives refresh. The crosstab must show DIM's keys. |
| `UNFILTER DS_1 DIM` | Remove the navigation filter of a characteristic. |
| `VAR DS_1 NAME=value` | Set a variable (external format, as in the prompt). |
| `REFRESH` / `REFRESH DS_1` | Refresh. Set mandatory variables first. |
| `SET Sheet!A1=value` | Enter a plan value as user input (decimal point, no thousands separator). Follow with `TRANSFER` or `PLAN`. |
| `TRANSFER` | Send entered values to the server and check them. |
| `PLAN PF_1` / `PVAR PF_1 NAME=value` | Run a planning function or sequence / set its variable. |
| `SAVE` / `RESET` | Save / discard plan data. **Ask the user first.** |
| `API SAPFunction a|b|c` | Any other AfO API call: `SAPMoveDimension`, `SAPSetFilter`, `SAPExecuteCommand`, `SAPSetFormat` ... (not `SAPLogon`). |
| workbook commands | Listed on the bridge sheet; they run the macros behind the buttons. |

## Read

AfO read functions are worksheet formulas. Write them into free cells of `AgentLog` from `AD2`,
read the result, delete them again:
`=SAPListOfEffectiveFilters("DS_1")`, `=SAPListOfDimensions("DS_1")`,
`=SAPListOfVariables("DS_1")`, `=SAPGetVariable("DS_1","NAME","VALUEASKEY")`,
`=SAPListOfMembers("DS_1","DIM")`. Values of the crosstab itself can be read directly.

## Rules

- No AutoFilter, sorting, hiding, grouping, formulas or values written into an AfO crosstab. A
  filter on values is `FILTERWHERE`, a filter on members `FILTER`, layout changes `API
  SAPMoveDimension`. If no command covers a request, say so and ask before using Excel.
- Map the user's words to technical names with `SAPListOfDimensions`: e.g. "Anlagennummer" is
  `0UCINSTALLA`. Key figures are addressed by their column header text in `FILTERWHERE`.
- Before writing plan values: only input-ready cells (yellow), numbers only; check after
  `TRANSFER` that the values are still there and `AfoAgentResult` says OK.
- Ambiguous selection (several matching rows, a member under several parents): ask first.
- After a step, report what changed in the user's terms (e.g. "25 of 604 installations remain,
  all with a projected annual consumption of at least 100,000 kWh") and how to undo it
  (`UNFILTER ...`).
- A result still `running` after 30 seconds: a dialog is probably open in Excel; ask the user.
