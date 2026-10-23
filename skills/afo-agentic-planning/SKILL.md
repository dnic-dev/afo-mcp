---
name: afo-agentic-planning
description: >-
  Work in SAP Analysis for Office (AfO) workbooks from Claude in Excel: find out which queries,
  planning functions and crosstabs a workbook uses, and plan only through the agent bridge, never
  by writing plan cells directly. Use whenever the open workbook is an AfO workbook (sheets with
  SAP crosstabs, an "Analysis" ribbon, defined names starting with SAP) or the user asks to plan,
  copy actuals, revalue, recalculate or save plan data in Excel.
---

# Analysis for Office workbooks

Analysis for Office (AfO) connects Excel to SAP BW. **Values written through Office.js are not
seen by AfO**: the cell shows them, but recalculation and saving use the old values. A workbook
can only be planned through the **agent bridge**, a command cell that a macro in the workbook
executes as if the planner typed or clicked.

## 1. Bridge or not

Look for the defined name `AfoAgentCommand`, including on hidden sheets.

- **Found:** read the sheet it points to completely before doing anything. It explains the
  commands, the rules, this workbook's objects and how to check results, and keeps a log of every
  step. Work exactly as it says. If its status cell (`AfoAgentStatus`) is not `active`, macros are
  disabled: ask the user to enable content and reopen.
- **Not found:** read and explain only. Do not write into crosstabs or plan cells, not even as a
  "preview": the workbook would look planned while SAP keeps the old values. Tell the user that
  planning from Claude needs the agent bridge in this workbook.

Rules that hold in every workbook, whatever the bridge sheet adds:

- Ask before acting when a request matches several members, rows or versions.
- Change input-ready cells only; for changes across many cells, prefer a planning function.
- After a change, recalculate and read the values back: AfO drops some input silently (text in a
  number cell, values outside a data slice or characteristic relationship).
- Never `SAVE` or `RESET` without the user's explicit yes for that step.

## 2. Understand the workbook

AfO read functions are worksheet formulas, evaluated by AfO whoever writes them. Write them into
a free cell (on the bridge sheet: its read area), read the result, delete them afterwards.

- `=SAPListOf("DATASOURCES")`: data source aliases with the query (before logon the technical
  name, after logon the description). Likewise `"PLANNINGFUNCTIONS"`, `"PLANNINGSEQUENCES"`,
  `"CROSSTABS"`.
- `=SAPListOfVariables("<alias>")`, `=SAPListOfEffectiveFilters("<alias>")`,
  `=SAPListOfDimensions("<alias>")`: variables, filters, dimensions and their axis.

Never ask the user to look up query names in the AfO menus; read them this way.

## 3. Context from SAP BW

If a BW connector is available, use it with the technical names from step 2: query definition,
planning functions and their logic, data the crosstab does not show (previous years, other
versions, master data). Check that the connector points to the same system as the workbook; if
the names are not found there, say so instead of guessing similar objects.
