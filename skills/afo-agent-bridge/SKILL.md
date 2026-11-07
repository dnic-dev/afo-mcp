---
name: afo-agent-bridge
description: Operate SAP Analysis for Office (AfO) workbooks agentically from inside Excel through their agent bridge - read and analyse data, change variables and selections, filter (by member, interval, list or key figure value), navigate (move dimensions, swap axes, expand and collapse hierarchy nodes, set a drill level), enter plan data, run planning functions and sequences, save or discard plan data, trigger workbook buttons and call any other AfO API function. Use in every Excel workbook that has a sheet "AgentLog" or the defined name AfoAgentCommand, or that shows AfO crosstabs (SAP* cell styles, data sources DS_1, DS_2 ...), for any request that touches its data or its SAP connection. Prefer the bridge over Excel features (AutoFilter, sorting, hiding, formulas) on everything Analysis for Office owns.
---

# Operating Analysis for Office workbooks through the agent bridge

Analysis for Office (AfO) owns its crosstabs. It rebuilds them on every refresh and navigation,
and it does not see values that an add-in writes into cells (writes through the Excel JavaScript
API raise no Change event). Anything done with Excel features on a crosstab is therefore lost or,
worse, silently wrong. AfO acts only on its own API.

The **agent bridge** is the way in: a command cell that VBA in the workbook turns into AfO API
calls, as if the user had clicked. This skill works for any AI that drives Excel through the
JavaScript API (Claude in Excel or others) and for any AfO workbook that carries the bridge.

## 1. Get oriented first

1. Find the bridge: defined names `AfoAgentCommand` (command, B1 of the hidden sheet `AgentLog`),
   `AfoAgentResult` (B3), `AfoAgentStatus` (B4).
   - No bridge: say so. AfO cannot be operated from the Excel JavaScript API. Offer read-only help
     (reading crosstab values) and suggest adding the bridge (afo-mcp `afo_update_workbook`
     with `agentBridge`). Do not imitate AfO with Excel features.
   - B4 not `active`: macros are disabled. Ask the user to click "Enable content". Nothing else
     works until then.
2. Read the bridge sheet once: its command list, its rules, the data sources, planning functions,
   sequences and crosstabs of this workbook, the live lists (from column Q) and the message area
   (column F).
3. Read what the request is about (section 3): dimensions with their axes, current filters,
   variable values, members. Map the user's words to technical names this way; never guess a
   technical name.

## 2. Run commands

- Write the command into `AfoAgentCommand`. Several commands: **one per line** (line breaks in
  the cell); they run in order and stop at the first error. A single line may separate commands by
  `;`, but only if no value contains `;`.
- Poll `AfoAgentResult` about once a second until it no longer starts with `running`.
  `OK …` = all steps done; `ERROR … in step n (…): reason` = that step failed, later steps did not
  run. Still `running` after 30 seconds: a dialog is open in Excel; ask the user to look.
- The log below the command list has one line per step plus every new AfO message (errors fail the
  step). Read it when something is unclear.
- Verify the effect (section 3) before reporting success.

### Command reference

| Command | Purpose |
|---|---|
| `VAR DS NAME=value` | Set a variable of a data source, value as typed in the prompt (external format, e.g. dates and periods as displayed, intervals `a - b`). Consecutive `VAR` lines are submitted together; follow with `REFRESH` if needed. |
| `REFRESH` · `REFRESH DS` | Refresh all data sources or one. Mandatory variables must have values, or AfO opens the prompt dialog. |
| `FILTER DS DIM=value` | Navigation filter on a characteristic: one key, an interval `a - b`, several keys `a;b;c` (own line), or `@Sheet!A1:A99` for the keys in a range. Empty value removes it. |
| `FILTERWHERE DS DIM "Column header" op number` | Filter a characteristic to the members whose value in the crosstab column meets the condition (`=` `<>` `<` `<=` `>` `>=`, decimal point, no thousands separator). Becomes a real AfO filter on the member keys and survives refresh. The crosstab must show the characteristic's keys in its rows. |
| `UNFILTER DS DIM` | Remove the navigation filter of a characteristic. Variables are not touched. |
| `EXPAND DS node` · `COLLAPSE DS node` | Expand or collapse a hierarchy node in the crosstab, as a click on its `[+]`/`[-]` does. The node by key or by text as shown, or by its cell (`Sheet!B9`) when the name appears more than once. Works for BW hierarchies and for the compact display of rows (several characteristics in the rows shown as one tree). |
| `DRILLLEVEL DS DIM=n` | Show the hierarchy of a characteristic in the rows down to level n (1 = top nodes only); nodes are expanded and collapsed one by one. |
| `SET Sheet!A1=value` | Enter a value as user input, e.g. plan data in an input-ready cell. Numbers with decimal point. A range sets the same value in every cell. Empty value clears. |
| `TRANSFER` | Send entered values to the server and check them (disaggregation, characteristic relationships, data slices, locks). |
| `PLAN key` | Run a planning function or sequence (keys as listed on the bridge sheet, e.g. `PF_1`, `PS_1`). |
| `PVAR key NAME=value` | Set a variable of a planning function or sequence. Variables the workbook reads from a cell: change the cell with `SET` instead. |
| `SAVE` · `RESET` | Save plan data to the server / discard unsaved plan data. **Always ask the user first.** |
| `API SAPFunction a|b|c` | Any other AfO API function with its arguments (see below). The result is logged; `0` counts as failure. `SAPLogon` is refused. |
| Workbook commands | Listed on the bridge sheet; each runs the macro behind a button of the workbook. Prefer them where they exist, since they carry the workbook's own logic. |

### More AfO through `API`

Use the AfO API by name. Commonly needed:

- `SAPMoveDimension` – move a characteristic to rows, columns or the background (free), or change
  its position; this is how drill-down and "remove from crosstab" work.
- `SAPSetFilter` – the full filter function (FILTER is the short form).
- `SAPExecuteCommand` – AfO commands such as `Refresh`, `PlanDataTransfer`, `PlanDataSave`,
  `PlanDataReset`, `ShowPrompts`, and others AfO offers in its API.
- `SAPSetRefreshBehaviour` – pause and resume refreshing while several changes are made.
- `SAPSetFormat` and further design functions where the workbook uses formatting rules.
- `SAPCallMemberSelector` – opens AfO's member selection dialog for the user.

If unsure about a function's arguments, try it once and read the result in `AfoAgentResult` and
the log; do not repeat failing calls with guessed variations, ask instead.

## 3. Read and verify

AfO read functions are worksheet formulas. Write them into free cells of `AgentLog` from `AD2`,
read the (often multi-row) result, then delete them:

| Formula | Returns |
|---|---|
| `=SAPListOfDimensions("DS")` | Dimensions: technical name, text, axis (ROWS, COLUMNS, empty = free). |
| `=SAPListOfEffectiveFilters("DS")` | Filters in effect, from variables and navigation. |
| `=SAPListOfDynamicFilters("DS")` | Navigation filters only. |
| `=SAPListOfVariables("DS")` | Variables with their current values. |
| `=SAPGetVariable("DS","NAME","VALUEASKEY")` | One variable value (`VALUEASTEXT` for the text). |
| `=SAPListOfMembers("DS","DIM")` | Members of a dimension (keys and texts). |
| `=SAPGetDisplayedMeasures("DS")` | Key figures shown. |
| `=SAPGetData("DS","measure",…)` · `=SAPGetMember(…)` · `=SAPGetDimensionInfo(…)` | Single values and member/dimension details. |
| `=SAPGetSourceInfo("DS","DataSourceName")` | Description and other properties of a data source. |
| `=SAPListOfMessages()` | Current AfO messages (also kept on the bridge sheet). |
| `=SAPListOf("DATASOURCES")`, `"CROSSTABS"`, `"PLANNINGFUNCTIONS"`, `"PLANNINGSEQUENCES"` | Objects of the workbook (also live on the bridge sheet). |

Crosstab values themselves can be read directly from the cells. Planning and input-ready cells
are recognisable by their cell style (input-ready styles, typically highlighted).

## 4. Working patterns

- **Answer a question about the data:** read the crosstab; if the data needed is not in it,
  change the navigation (`API SAPMoveDimension`, `FILTER`) rather than computing around it in
  Excel, then read again. Undo navigation the user did not ask to keep.
- **Change the selection (period, version, organisational unit …):** `VAR` lines, then `REFRESH`.
  Check with `SAPListOfVariables`.
- **Restrict what is shown:** members → `FILTER`; a condition on a key figure → `FILTERWHERE`;
  lift with `UNFILTER`. Check with `SAPListOfDynamicFilters`. Tell the user how to undo it.
- **Drill down in a hierarchy:** read the crosstab first. Nodes show `[+] KEY` (collapsed) or
  `[-] KEY` (expanded) in the cell, indented by two blanks per level; leaves have no symbol.
  Overview first (`DRILLLEVEL DS DIM=1` or `=2`), then `EXPAND` the node the question is about,
  read again. A node that is not shown (its parent is collapsed) cannot be addressed: expand the
  parent first. Collapse what the user did not ask to keep open.
- **Hierarchy not shown:** switching a hierarchy on or off and the compact display of rows have
  no AfO API function. They come from the query or the state saved with the workbook. Ask the
  user to switch it in Analysis for Office (right-click the characteristic → Hierarchy, or
  Analysis → Hierarchy → Compact Display in Rows), or to have the workbook saved that way. For
  a lasting change the workbook can be prepared with afo-mcp (`afo_update_workbook` with
  `hierarchies` or `compactRows`), outside Excel.
- **Hierarchy commands move the mouse pointer.** Analysis for Office reads the click position
  from the pointer, so the bridge puts it on the node symbol for a moment and back; if Excel is
  behind another window, the pointer visibly jumps across it. The Excel window must not be
  minimized. Tell the user before a longer `DRILLLEVEL` run, and ask them not to use the mouse
  meanwhile.
- **Plan:** confirm the target cells are input-ready and unambiguous; `SET` the values;
  `TRANSFER`; check `AfoAgentResult` and that the values are still in the cells (AfO drops text in
  data cells silently and may disaggregate totals to the members below). Use a planning function
  (`PLAN`) for changes to many cells or for derived values. Then ask before `SAVE`.
- **Several changes at once:** put them on separate lines of one command; consider
  `API SAPSetRefreshBehaviour` to pause refreshing for larger sequences.
- **Something fails:** report the step and AfO's message in plain words, propose the next step,
  do not retry blindly.

## 5. Rules

- Never use Excel features on AfO crosstabs: no AutoFilter, sorting, hiding or grouping of rows or
  columns, no formulas or values written into crosstab cells except through `SET`, no copying a
  crosstab to compute on it when AfO can deliver the result. If no command or API function covers
  a request, say so and ask before using Excel.
- Variables, filters and planning objects are addressed by technical name, found by reading
  (section 1.3), never guessed. Ambiguous selections (several matching rows, a member under several
  parents): ask before writing.
- Writes that change data on the server (`SAVE`, planning functions that save, `RESET`) need the
  user's explicit go-ahead each time.
- Do not change the bridge sheet itself, the data source definitions or the workbook's VBA.
- Report in the user's terms: what changed, how many rows or members are affected, how to undo.
