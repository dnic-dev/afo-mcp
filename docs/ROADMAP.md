# Roadmap: agentic planning with Analysis for Office

The goal: a planner works in an Analysis for Office (AfO) workbook and tells an agent what to do,
in plain language: "copy last year's actuals, raise January by 10 % for one location, recalculate",
"filter on product X and add the region to the rows", "add an info field with the last data
update". The agent operates the workbook like a planner would, explains what it did, and asks
before it saves.

This document is the plan we work through. Items are ticked when they are done and tested.

## Architecture

### Three components, three moments

| | bw-modeling-mcp | afo-mcp | Claude in Excel + agent bridge |
| --- | --- | --- | --- |
| When | design time (backend) | design time (frontend) | run time (planning) |
| Works on | the BW system | the workbook **file** (.xlsx/.xlsm), offline | the **open** workbook, in the planner's AfO session |
| Does | aDSOs, aggregation levels, filters, planning functions and sequences, queries; reads data | builds workbooks, extends existing ones, inventories workbook landscapes | enters values, runs planning functions, navigates, checks, saves, in dialogue with the planner |
| Runs | locally and on SAP BTP | locally (Claude Code) | in the planner's Excel |

bw-modeling-mcp covers the planning backend. Best practice for planning models comes as skills.

Planning itself has to run in the planner's AfO session: plan buffer, locks, characteristic
relationships, data slices and saving with the planner's authorizations live there. Running a
planning function through RFC from outside does not save and leaves locks behind.

### Three channels for AfO features

Which component offers a feature depends on whether AfO exposes it at run time:

1. **Claude in Excel directly (Office.js):** plain Excel (formats, Excel conditional formatting,
   Excel charts) and AfO **formulas**, which AfO evaluates whoever writes them: info fields
   (`SAPGetSourceInfo`, `SAPGetInfoLabel`, `SAPGetVariable`, `SAPGetDimensionEffectiveFilter`),
   object lists (`SAPListOf`, `SAPListOfDimensions`, `SAPListOfMessages`), `SAPGetData`.
2. **Agent bridge (AfO API at run time):** everything AfO offers as an API function: input values,
   planning functions and sequences, variables, filters, moving dimensions, design rules, save
   and reset. Works in the open workbook; persisted when the workbook is saved.
3. **afo-mcp (the file):** everything that exists only in the AfO definition and has no API: data
   sources, crosstabs, planning objects, AfO conditional formatting (exceptions), AfO charts,
   number of new lines, filter components. Built or inserted offline; the workbook is reopened.

While a workbook is open, AfO holds its definition in memory and writes it back on save, so
structural changes from outside always mean: close, change the file, reopen.

### Why the agent bridge exists

- Claude in Excel writes cells through Office.js. Those writes raise **no Change event**, so AfO
  never registers them as input: the cell shows the value, but recalculation and saving use the
  old one.
- Office.js cannot call VBA or the AfO API.
- Office.js writes **do** raise **Calculate**. The bridge sheet has a command cell that feeds a
  formula; `Workbook_SheetCalculate` picks the command up, the `AfoAgent` module runs it outside
  the event (`Application.OnTime`) and writes values through VBA. VBA writes raise real Change
  events, which AfO accepts as user input.
- The bridge sheet documents itself (commands, the workbook's data sources, planning objects and
  crosstabs, messages, a step log). An agent that has never seen the workbook finds and uses it.

The bridge is generic at the level of AfO, not of the business process: business logic lives in
BW (functions, sequences, data slices, relationships) and in the agent's reasoning. The bridge has
to cover what a planner can do in AfO with keyboard and mouse.

## Done

- [x] Read workbooks offline: `afo_read_workbook`, `afo_scan_workbooks`, `afo_where_used`.
- [x] Build workbooks offline: `afo_build_workbook` (data sources with stored prompts and axes,
      crosstabs, planning functions with filters, variables and cell references, VBA, button
      macros; template or minimal layout; spec validation).
- [x] Code in document modules (`documentModules`) for event handlers.
- [x] Agent bridge in the builder (`agentBridge`): bridge sheet, `AfoAgent` module, handlers;
      commands `SET`, `TRANSFER`, `PLAN`, `VAR`, `FILTER`, `REFRESH`, `SAVE`, `RESET` and workbook
      commands mapped to macros. Tested end to end: Claude in Excel copied actuals, changed 28
      input cells and recalculated; AfO accepted every value.

## Work streams

### A. Harden the bridge

- [x] **Rejected input:** *(implemented; live test with a data slice pending. Live: "not enabled"
      from PlanDataTransfer/Reset/Save means nothing pending and counts as OK; text in a data cell
      is ignored by AfO without any message, so agents must check values after recalculating)* after `SET`/`TRANSFER`/`PLAN`, read the AfO messages; an error or a
      rejected value (data slice, characteristic relationship, lock) makes the step fail. Today
      `SET` reports OK even when AfO discards the value afterwards.
- [x] **`PVAR key NAME=value`:** *(implemented; variables read from cells are rejected by AfO and
      set through their cell instead; test with free variables in the parcours)* variables of planning functions and sequences
      (`SAPSetPlanParameter`).
- [ ] **Dialog guard:** no command may hang on a dialog (prompt dialog on refresh, message
      dialog, confirmations). Prompts: require `VAR` first or detect and report. Messages: through
      `SAPListOfMessages`; check whether the `BeforeMessageDisplay` callback with
      `SAPSuppressMessage` is needed.
- [x] **Reading** needs no command: the read functions are formulas; the bridge sheet explains them and offers a free area. Cell context (`SAPGetCellInfo`: which members a cell stands for), changed and
      unsaved values (`SAPGetChangedMeasures`), variables (`SAPListOfVariables`, `SAPGetVariable`),
      filters (`SAPListOfEffectiveFilters`), dimensions and axes (`SAPListOfDimensions`).
- [x] **Rules in the help text:** ambiguous selection → ask first; always confirm before `SAVE`;
      change input-ready cells only; prefer planning functions for mass changes.
- [x] **Live help:** *(live lists confirmed in AfO)* on open, refresh the list of data sources, planning objects and crosstabs
      from `SAPListOf`, so objects added by hand appear too.
- [x] Message area with column headers; log trimmed to the last few hundred rows.
- [x] Option `hidden`: bridge sheet hidden; agents find it through the defined names. *(Test with
      Claude in Excel pending.)*

### B. Navigation and design through the bridge

- [x] Live: `API SAPMoveDimension DS|DIMENSION|ROWS` and `…|AFTER|OTHER_DIMENSION` work; inserted
      planning function PF_3 runs; retrofitted bridge runs the workbook commands.
- [x] **`API SAPFunction arg1|arg2|...`:** any AfO API function through the bridge (SAPLogon
      excluded). Covers navigation, design rules and `SAPExecuteCommand` until convenience commands
      exist. *(Parameters of SAPMoveDimension and the design rule functions to be confirmed live.)*

- [ ] **`MOVE`:** move a dimension to rows, columns or free characteristics (`SAPMoveDimension`).
- [ ] **Hierarchies:** expand, collapse, drill level; find the `SAPExecuteCommand` commands.
- [ ] **Filter on planning objects** (`SAPSetFilter` with a planning object alias); test.
- [ ] **Write by members:** `SAPSetData` (value for a member combination), independent of the
      crosstab layout; needed for formula-based planning and for cells not on screen.
- [ ] **New lines:** input into a crosstab's new-line area through `SET`; test.
- [ ] **Design rules:** `SAPSetFormat`, `SAPSetText`, `SAPInsertLine`, `SAPFixLineSize`,
      `SAPDeleteDesignRule`, `SAPListOfDesignRules`.
- [ ] Undo/redo where AfO offers it.

### C. Planning parcours (test bench)

A deliberately varied test model on our own BW system, built with bw-modeling-mcp, and a workbook
on it built with afo-mcp, to prove the bridge generic case by case:

- [ ] hierarchy with input on nodes (top-down disaggregation)
- [ ] data slice that locks part of the data
- [ ] characteristic relationship that rejects invalid combinations
- [ ] planning sequence with several steps and variables
- [ ] new lines (plan a new product)
- [ ] two data sources on two sheets
- [ ] navigation and design rules from work stream B
- [ ] automated run through COM (operate scripts), then one run with Claude in Excel
- [ ] result matrix: feature × works / gap, gaps fed back into A and B

### D. afo-mcp: change existing workbooks

- [x] **Patch instead of rebuild:** `afo_update_workbook`. a tool that changes an existing workbook and keeps everything
      else (layout, formulas, existing macros, AfO definition).
- [x] Retrofit the agent bridge into an existing workbook, merging it into the existing VBA
      project (existing handlers get one call; signed projects are refused). Tested with an
      Excel-compiled project; AfO recognises the result.
- [x] Insert a planning function (bound to an existing data source). *(Run time test with logon
      pending.)*
- [ ] Insert a planning sequence, a data source with crosstab, change prompts and axes.
- [ ] AfO conditional formatting (exceptions), AfO charts, number of new lines, filter components
      (`SAPSetFilterComponent` items); learn each format from a workbook AfO saved itself.
- [ ] **Planning toolkit:** a spec preset that binds a set of generic planning functions (revalue
      by a percentage, copy, distribute, plus an "agent" function whose FOX bw-modeling-mcp may
      rewrite). The workbook binds the slot, the behaviour lives in BW: most requests then need no
      workbook change.

- [x] **Take over from a source workbook** (`source` + `from` on data sources, planning objects,
      crosstabs, formula items): verbatim, with the whole stored state, design rules, connection
      (SAProuter/SNC) and settings; keys may be renamed, positions are new. Round trip over 30 real
      workbooks identical element by element; rebuilt workbooks recognised by AfO.
- [x] Formula items (filter components) in the spec; the builder writes their formula into the cell.
- [ ] Describe in the spec (for workbooks designed from scratch): filters incl. hierarchy nodes,
      ranges and exclusions; design rules; display per characteristic; planning function filter
      references; system connection with SAProuter/SNC.

### D2. Workbooks stored in BW (in bw-modeling-mcp, using the afo-mcp parser as a library)

AfO workbooks saved to BW live in table `RSAOOBJ` (whole file in one field); BW keeps a
where-used list in `RSAOOBJXREF` (queries by element UID, resolved through `RSZELTDIR`; planning
functions by name). The RFC module `RSAO_BICS_OPEN` returns the file with authorization check
(`IV_IGNORELOCKS = 'X'` sets no lock). Verified: a workbook fetched this way is read completely by
the afo-mcp parser. Tracked as issues in the bw-modelling-skills repository.

- [ ] `bw_list_afo_workbooks`: workbooks with where-used (queries, planning functions).
- [ ] `bw_get_afo_workbook`: fetch and describe; file locally (stdio) or as a short-lived download
      link (BTP). Open: how bw-modeling-mcp reaches the RFC module over HTTP.
- [ ] `bw_save_afo_workbook`: save a changed workbook back (`RSAO_BICS_SAVE`), confirmation
      required, new name by default.
- [ ] Publish the afo-mcp parser so bw-modeling-mcp can depend on it.

### E. Bridge as an Excel add-in

- [ ] The same bridge as an `.xlam` installed once per user, handling application-level events
      for every open AfO workbook. The agent creates the bridge sheet itself. Existing workbooks,
      including customer workbooks, need no change and need not be `.xlsm`.

### F. Skills

- [ ] Workbook skill (company design): `agentBridge` by default in planning workbooks.
- [ ] Planning skill for Claude in Excel *(draft: `skills/afo-agentic-planning/SKILL.md`, to be
      reviewed and uploaded to the claude.ai account)*: how to understand a workbook (`SAPListOf`, bridge help),
      the rules from A, and how bw-modeling-mcp adds context. Check confidentiality before customer
      knowledge goes into an account-wide skill.
- [ ] Best-practice skill for planning models on the bw-modeling-mcp side, including the toolkit
      functions from D.

### G. Operate tools for Claude Code

- [ ] A persistent Excel session behind MCP tools (`afo_open`, `afo_logon` with credentials from
      the environment, `afo_refresh`, `afo_set_variables`, `afo_read_crosstab`,
      `afo_execute_planning`, `afo_save_plan` / `afo_reset_plan`, `afo_messages`, `afo_close`),
      with dialog watcher and fresh file copies. Needed to run the parcours unattended.

### H. Central hosting (later)

- [ ] afo-mcp on SAP BTP for inventory and analysis (`afo_read_definition` on the text an agent
      reads from the hidden sheet), same pattern as bw-modeling-mcp (XSUAA, no destination).
- [ ] Structural changes from Claude in Excel through a central workbook store (SharePoint,
      OneDrive): afo-mcp changes the file there, the planner reopens it.
- [ ] For customer use, run in the customer's own subaccount.

## Order

1. A (harden) → C (parcours, first round) → B (navigation and design) → C (second round)
2. D (patch existing workbooks, toolkit) together with F (skills)
3. E (add-in), G (operate tools) as the parcours needs them
4. H when there is a concrete need

## Constraints learned the hard way

- Values entered through Office.js are invisible to AfO; never let an agent write plan cells
  directly.
- AfO formulas work as cell formulas for any writer; `SAPListOf("DATASOURCES")` names the
  queries even before logon.
- Commands from the bridge must never open a dialog: a modal dialog blocks VBA and the agent.
- `SAPSuppressMessage` outside a callback fails with "Action is currently not allowed".
- Closing an AfO workbook through COM after a session strips its data and can raise a blocking
  Excel dialog; end automated sessions by ending the process, always on file copies.
- VBA modules over 4096 bytes need real MS-OVBA compression; a crosstab needs both
  `RangeOccupied` and `RequestedRange`; without a stored state containing
  `DOCUMENT_APPLICATION_ASSIGNED`, input-ready crosstabs can make AfO drop all connections.
