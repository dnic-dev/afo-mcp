# Claude Code Instructions — afo-mcp

An MCP server for SAP Analysis for Office (AfO) workbooks. Two parts:

1. **Read** workbooks offline, without Excel or an SAP logon: data providers, planning objects,
   grids, SAP formulas, variable and filter states, VBA macros — for one workbook or a whole folder.
2. **Operate and build** workbooks through Excel COM and the AfO VBA API on a Windows machine
   with AfO installed: variables, filters, refresh, planning functions and sequences, saving plan
   data, layout, macros.

## Start here

The current plan, the spike results and the next steps are in `playbooks/handover-start.md`
(local, not versioned). Read it before doing anything else in a new session.

## Confidentiality

- Real workbooks are customer data. They live in `samples/` (gitignored) and never enter a
  commit, a test fixture, a code comment, documentation or an issue.
- Test fixtures for the repository are built from scratch against our own BW system.
- `playbooks/` is gitignored: handovers and notes may name customers and systems.
- Write examples with generic names (`QUERY_NAME`, `DS_1`, `PF_1`), never with real query,
  InfoProvider or system names.

## Language

All code, comments, tool descriptions and documentation in English, whatever language the
conversation is in.

## Git

- Commit only when David asks; push only when he asks.
- Conventional commits (`feat(scope): …`), messages safe for a public repository.
- The repository is private for now; David decides later whether it goes public.

## Conventions

- TypeScript and the MCP SDK, the same setup as `bw-modeling-mcp` (sibling project).
- Reading must stay possible without Excel: unzip the workbook and parse its parts. Anything that
  needs Excel belongs to part 2 and must say so.
