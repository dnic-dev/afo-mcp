# Security

## Reporting a vulnerability

Please do not open a public issue for a security problem. Send a description to the
maintainer (see `package.json`), with the workbook structure described in generic names rather
than attached. You will get an answer within a week.

## What runs where

- **Read and build tools** only touch files. They unzip a workbook, parse or write its parts and
  never contact a BW system. A workbook is untrusted input: the parsers are written to fail on
  malformed content, not to execute it. VBA is read and written as source text and never run
  by the server.
- **Operate tools** run only on Windows and drive Excel and Analysis for Office through COM, as
  a VBA macro would, with the rights of the Windows user running the server. Credentials come
  from the server's environment or a secret runner, never from tool arguments, so no password
  passes through a conversation or a log. Arguments reach the PowerShell child through a JSON
  file, not through string interpolation.
- **The agent bridge** is VBA inside a workbook. It runs with the rights of the user who opened
  the workbook and enabled its macros, against the BW system that user is logged on to.

## Trust boundary of the agent bridge

The bridge executes whatever is written into its command cell. Everything that can write that
cell acts as the logged-on user: Claude in Excel, any other Office.js add-in, a VBA macro, or a
person typing. This includes commands that change data in BW (`SET`, `SETDATA`, `TRANSFER`,
`PLAN`, `SAVE`) and `API`, which calls any Analysis for Office function except the logon.

Consequences:

- A prompt injection reaching the agent, for example through text in the workbook or in a
  document it reads, can lead to plan data being changed or saved. The skills in `skills/` tell
  the agent never to `SAVE` or `RESET` without the user's explicit yes and to ask when a request
  is ambiguous, but that is a convention of the agent, not a control in the bridge.
- The bridge grants no rights the user does not have: BW authorisations, data slices and locks
  apply as for input typed by hand.
- Workbooks with the bridge are `.xlsm` files. Excel's macro security (trusted locations, the
  security bar) applies; the bridge does not run until the user enables the content.

Use the bridge against non-productive systems first, and keep it out of workbooks whose users
should not be able to change plan data.
