// Agent bridge: a command cell that lets an agent operate an AfO workbook from cells alone.
//
// Why: Claude in Excel (Office.js) writes cells without raising Excel's Change event, so
// Analysis for Office never registers those values as plan input, and Office.js cannot call
// VBA or the AfO API. A write does raise Calculate. The bridge sheet has a formula on the
// command cell; Workbook_SheetCalculate picks the command up and AfoAgent runs it outside the
// event (Application.OnTime). Values written there through VBA raise real Change events, which
// AfO accepts as input (tested on BW/4HANA with AfO 2.8).
//
// Reading needs no command: AfO's read functions are worksheet formulas, which AfO evaluates
// whoever writes them. The bridge sheet keeps live lists and a free area for such formulas.

import type { AfoWorkbookSpec } from './afo.js';

export interface AgentCommandSpec {
  /** Command word, e.g. COPY_ACTUALS. Upper case letters, digits, underscores. */
  name: string;
  /** Macro the command runs, `Module.Procedure`. */
  macro: string;
  description?: string;
}

export interface AgentBridgeSpec {
  /** Name of the bridge sheet (default AgentLog). */
  sheet?: string;
  /** Hide the bridge sheet (default true); agents find it through the defined names. false shows it. */
  hidden?: boolean;
  /** Workbook-specific commands, each running a macro (e.g. the macro behind a button). */
  commands?: AgentCommandSpec[];
}

export const AGENT_MODULE = 'AfoAgent';
export const AGENT_SHEET_CODE_NAME = 'AfoAgentSheet';
export const DEFAULT_AGENT_SHEET = 'AgentLog';
export const AGENT_NAMES = { command: 'AfoAgentCommand', result: 'AfoAgentResult', status: 'AfoAgentStatus' } as const;

/** Log rows kept; older rows are dropped in blocks. */
const LOG_KEEP = 500;

/** Generic commands; workbook commands must not reuse these words. */
export const BUILTIN_COMMANDS: Array<[string, string]> = [
  ['SET Sheet!A1=value', 'Write a value as user input (Analysis for Office registers it). Decimal point, no thousands separator. Empty value clears the cell. A range writes the same value into every cell. Values are checked when they are transferred: follow with TRANSFER or a planning function. A value in double quotes is written as text ("01.2026"), as are values with a leading zero. Also for the new lines below an input-ready crosstab (see NEWLINES).'],
  ['SETDATA DS_1 Key figure|DIM=key;DIM=key|value', 'Write a plan value by member combination instead of by cell, independent of where the crosstab shows it: the key figure as its header text, then every characteristic in the rows (and columns) of DS_1 with its key, then the value (decimal point). The combination must be an input-ready cell of the current result at the lowest level, not a total. Follow with TRANSFER.'],
  ['NEWLINES Crosstab1 5', 'Show n empty lines below an input-ready crosstab (data source in change mode) for member combinations it does not show yet. Write the keys of all characteristics in the rows of a new line with SET as text ("01.2026") and the values, then TRANSFER: the line becomes a regular row. Keys in the format the crosstab shows them.'],
  ['CHANGEMODE DS_1', 'Switch an input-ready data source to change mode, so its cells take input (DISPLAYMODE DS_1: back to display mode; SAVE or RESET first, or Analysis for Office asks in a dialog whether to save). Input-ready cells have an editable cell style; read-only data cells do not.'],
  ['TRANSFER', 'Send changed input values to the server and check them, without running a planning function. "nothing pending" when there is nothing to send.'],
  ['PLAN key', 'Run a planning function or sequence by its key (PF_1, PS_1, ...).'],
  ['PVAR key NAME=value', 'Set a variable of a planning function or sequence, value as typed in the prompt. Variables the workbook reads from a cell are better set by SET on that cell: the next run reads the cell again.'],
  ['VAR DS_1 NAME=value', 'Set a variable of a data source, value in external format as typed in the prompt (01.2025, 01.2025 - 12.2025). Consecutive VAR steps are submitted together.'],
  ['FILTER DS_1 DIMENSION=value', 'Filter a characteristic, value as typed in the filter dialog: a key, an interval (a - b), several keys separated by ";" (write the command on its own line then), or @Sheet!A1:A50 for the keys in a range. An empty value removes the filter.'],
  ['FILTERWHERE DS_1 DIMENSION "Key figure" >= 100000', 'Filter a characteristic to the members whose value in the crosstab of DS_1 meets the condition (=, <>, <, <=, >, >=). The key figure is the column header text as shown. The result is a real Analysis for Office filter on the member keys; it stays when the data source is refreshed. Lift it with UNFILTER.'],
  ['UNFILTER DS_1 DIMENSION', 'Remove the navigation filter of a characteristic (variables stay as they are). FILTER and UNFILTER also take a planning function or sequence (PF_1, PS_1) in place of DS_1: that restricts the data it works on.'],
  ['MOVE DS_1 DIMENSION=ROWS', 'Move a characteristic: ROWS or COLUMNS (added at the end), FREE (out of the crosstab), AFTER OTHER or BEFORE OTHER (next to another dimension), SWAP OTHER (exchange places).'],
  ['REFRESH', 'Refresh all data sources (REFRESH DS_1: one data source). Set mandatory variables with VAR first, or the prompt dialog opens.'],
  ['EXPAND DS_1 node', 'Expand a hierarchy node in the crosstab of DS_1, as a click on its [+] does: the node by key or text as shown, or its cell (Sheet!B9) when the name appears more than once. Works for BW hierarchies and for the compact display of rows. Nodes show [+] (collapsed) or [-] (expanded) before their name.'],
  ['COLLAPSE DS_1 node', 'Collapse a hierarchy node, as a click on its [-] does; same forms as EXPAND.'],
  ['DRILLLEVEL DS_1 DIMENSION=2', 'Show the hierarchy of a characteristic in the rows down to level n (1 = top nodes only), expanding and collapsing node by node. Works for BW hierarchies and for the compact display of rows.'],
  ['API SAPFunction arg1|arg2|...', 'Call any Analysis for Office API function with its arguments, separated by "|": navigation (SAPMoveDimension), design rules (SAPSetFormat, SAPInsertLine), commands (SAPExecuteCommand). The result is logged; 0 counts as failure. SAPLogon is not allowed.'],
  ['SAVE', 'Save plan data to the server.'],
  ['RESET', 'Discard unsaved plan data.'],
];
const BUILTIN_WORDS = new Set(BUILTIN_COMMANDS.map(([c]) => c.split(' ')[0]));

const RULES = [
  'This is an Analysis for Office workbook: work through these commands. Never use Excel means on a crosstab: no AutoFilter, no sorting, no hiding rows or columns, no formulas or values written into it. They are lost on the next refresh and Analysis for Office does not see them. A filter, also one on key figure values, is FILTER or FILTERWHERE; moving characteristics is MOVE; other navigation is API (SAPSetFilter, SAPExecuteCommand). Only when no command covers a request, say so and ask before using Excel.',
  'Several commands: one per line in B1 (line breaks). A single line may also separate commands with ";" as long as no value contains ";".',
  'Ask the user before SAVE and before RESET.',
  'Selection ambiguous (a member under several parents, several matching rows)? Ask before writing.',
  'Write only input-ready cells of input-ready crosstabs. For changes to many cells prefer a planning function.',
  'After SET, run TRANSFER or a planning function and check B3: rejected values (data slices, characteristic relationships, locks) turn the step into ERROR.',
  'Data cells take numbers only. Text in a data cell is ignored by Analysis for Office without a message: after recalculating, check that the values you wrote are still there.',
  'Hierarchies: EXPAND, COLLAPSE and DRILLLEVEL work on a hierarchy that is shown. Switching a hierarchy on or off, or the compact display of rows, has no API function: it comes from the query or from the state saved with the workbook (afo-mcp can prepare that); ask the user to switch it in Analysis for Office if needed. These commands click the [+]/[-] symbol as a user does: the crosstab sheet is shown, the mouse pointer jumps there for a moment, and the Excel window must not be minimized.',
  'Navigation (moving dimensions, drilling down) makes crosstabs grow; one growing into another is cut off (a WARNING in the log). Check the layout before navigating a crosstab placed above another.',
  'While a command runs, Analysis for Office shows no message windows: its messages go to the log, and an error fails the step. B3 still says "running" after 30 seconds: another dialog is probably open in Excel; ask the user to look.',
  'B4 is not "active": macros are disabled; ask the user to enable content and reopen.',
];

const READING: Array<[string, string]> = [
  ['=SAPListOfVariables("DS_1")', 'Variables of a data source with their values (planning objects: =SAPListOfVariables("PF_1",,"PLAN_PARAMETER")).'],
  ['=SAPGetVariable("DS_1","NAME","VALUEASKEY")', 'One variable value.'],
  ['=SAPListOfEffectiveFilters("DS_1")', 'Current filters of a data source.'],
  ['=SAPListOfDimensions("DS_1")', 'Dimensions with their axis (ROWS, COLUMNS, free).'],
  ['=SAPListOfMembers("DS_1","DIMENSION")', 'Members of a dimension.'],
  ['=SAPGetSourceInfo("DS_1","DataSourceName")', 'Description of a data source.'],
];

/** Fixed cells of the bridge sheet. Messages and live lists are formulas the bridge sets once AfO is loaded. */
const LAYOUT = {
  messages: { col: 'F', headers: ['#', 'Message', 'Class', 'Message ID', 'Severity', 'Object', 'Technical name', '', '', 'System'] },
  lists: [
    { col: 'Q', title: 'Data sources (live)', formula: '=SAPListOf("DATASOURCES")' },
    { col: 'T', title: 'Planning functions (live)', formula: '=SAPListOf("PLANNINGFUNCTIONS")' },
    { col: 'W', title: 'Planning sequences (live)', formula: '=SAPListOf("PLANNINGSEQUENCES")' },
    { col: 'Z', title: 'Crosstabs (live)', formula: '=SAPListOf("CROSSTABS")' },
  ],
  readArea: { col: 'AD', title: 'Free cells for read formulas (from AD2); delete them afterwards' },
};

export function agentSheetName(spec: AfoWorkbookSpec): string {
  return spec.agentBridge?.sheet?.trim() || DEFAULT_AGENT_SHEET;
}

/** Problems in the bridge part of a spec. */
export function validateAgentBridge(spec: AfoWorkbookSpec): string[] {
  const bridge = spec.agentBridge;
  if (!bridge) return [];
  const errors: string[] = [];
  const seen = new Set<string>();
  const modules = new Set((spec.vbaModules ?? []).map((m) => m.name?.toUpperCase()));
  (bridge.commands ?? []).forEach((c, i) => {
    const what = `agentBridge.commands[${i}]`;
    if (typeof c?.name !== 'string' || !/^[A-Z][A-Z0-9_]*$/.test(c.name)) errors.push(`${what}: name must be upper case letters, digits and underscores, got ${JSON.stringify(c?.name)}.`);
    else if (BUILTIN_WORDS.has(c.name)) errors.push(`${what}: ${c.name} is a built-in command.`);
    else if (seen.has(c.name)) errors.push(`${what}: command ${c.name} appears twice.`);
    else seen.add(c.name);
    const [mod, proc] = String(c?.macro ?? '').split('.');
    if (!mod || !proc) errors.push(`${what} (${c?.name}): macro must be Module.Procedure.`);
    else if (!modules.has(mod.toUpperCase())) errors.push(`${what} (${c?.name}): module ${mod} is not among vbaModules.`);
  });
  if (modules.has(AGENT_MODULE.toUpperCase())) errors.push(`vbaModules: ${AGENT_MODULE} is the module of the agent bridge; choose another name.`);
  const own = spec.documentModules?.ThisWorkbook ?? '';
  for (const ev of ['Workbook_Open', 'Workbook_SheetCalculate']) {
    if (new RegExp(`\\bSub\\s+${ev}\\b`, 'i').test(own)) errors.push(`documentModules.ThisWorkbook defines ${ev}, which the agent bridge needs; call AfoAgent.${ev === 'Workbook_Open' ? 'MarkActive' : 'OnCalculate'} from your handler instead and drop agentBridge, or remove it.`);
  }
  return errors;
}

const vbaString = (s: string) => `"${s.replace(/"/g, '""')}"`;

/** The standard module AfoAgent. */
export function agentModuleCode(spec: AfoWorkbookSpec, logFirstRow: number): string {
  const sheet = agentSheetName(spec);
  const custom = (spec.agentBridge?.commands ?? [])
    .map((c) => `    Case ${vbaString(c.name)}: RunOne = RunMacro(${vbaString(c.macro)})`)
    .join('\r\n');
  const formulas = [
    `    PutFormula "${LAYOUT.messages.col}2", "=SAPListOfMessages(""INFORMATION"",""TRUE"")"`,
    ...LAYOUT.lists.map((l) => `    PutFormula "${l.col}2", ${vbaString(l.formula)}`),
  ].join('\r\n');
  return `Option Explicit

' Agent bridge, generated by afo-mcp. An agent writes commands into the cell ${AGENT_NAMES.command}
' (${sheet}!B1), several one per line (a single line may separate them by ";"). The bridge sheet
' lists the commands and logs every step.
' Office.js writes raise no Change event, so Analysis for Office ignores them; they do raise
' Calculate, which hands the command to this module. Values written here through VBA count as
' user input. After every AfO step new AfO messages are logged; errors fail the step.

Public Const BRIDGE_SHEET As String = ${vbaString(sheet)}
Private Const LOG_FIRST_ROW As Long = ${logFirstRow}
Private Const LOG_KEEP As Long = ${LOG_KEEP}
Public Busy As Boolean
Private pending As String
Private variablesPaused As Boolean
Private lastMessage As String
Private initTries As Long
Private stage As String
Private refreshMode As String
Private Const SETDATA_CELL As String = "AC1"

' Hierarchy nodes are expanded and collapsed by a click on their symbol: the AfO API has no
' function for it, and Analysis for Office reads the click position from the mouse pointer.
Private Type PointApi
    x As Long
    y As Long
End Type
Private Declare PtrSafe Function GetCursorPos Lib "user32" (p As PointApi) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hwnd As LongPtr, p As PointApi) As Long
Private Declare PtrSafe Function PostMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal parent As LongPtr, ByVal after As LongPtr, ByVal cls As String, ByVal title As String) As LongPtr

Public Sub MarkActive()
    SetCell "B4", "active"
    initTries = 0
    ScheduleInit
End Sub

Private Sub ScheduleInit()
    Application.OnTime Now + TimeSerial(0, 0, 3), "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!${AGENT_MODULE}.InitFormulas"
End Sub

' Messages and live lists are AfO formulas; they can only be entered once AfO is loaded.
Public Sub InitFormulas()
    initTries = initTries + 1
    If Not AfoReady() Then
        If initTries < 40 Then ScheduleInit
        Exit Sub
    End If
    EnsureFormulas
End Sub

Private Function AfoReady() As Boolean
    On Error GoTo nope
    Dim v As Variant: v = Application.Run("SAPGetProperty", "IsConnected")
    AfoReady = True
nope:
End Function

Private Sub EnsureFormulas()
${formulas}
End Sub

Private Sub PutFormula(ByVal address As String, ByVal formula As String)
    Dim cell As Range: Set cell = ThisWorkbook.Worksheets(BRIDGE_SHEET).Range(address)
    If cell.HasFormula Then Exit Sub
    Busy = True
    On Error Resume Next
    cell.Formula = formula
    On Error GoTo 0
    Busy = False
End Sub

Public Sub OnCalculate()
    If Busy Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(BRIDGE_SHEET)
    Dim cmd As String: cmd = Trim$(CStr(ws.Range("B1").Value))
    If cmd = "" Then Exit Sub
    Busy = True
    ws.Range("B1").ClearContents
    ws.Range("B3").Value = "running: " & cmd
    ws.Range("B4").Value = "active"
    Busy = False
    LogLine "COMMAND", "", cmd
    If pending = "" Then pending = cmd Else pending = pending & vbLf & cmd
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!${AGENT_MODULE}.RunPending"
End Sub

Public Sub RunPending()
    Dim cmd As String: cmd = pending: pending = ""
    If cmd = "" Then Exit Sub
    EnsureFormulas
    ' Dialog guard: AfO message windows would stop the command until someone closes them. While
    ' commands run, messages only go to the message list, which the log reads after every step.
    Quiet True
    ' Commands one per line; a single line may separate them by ";" (then no value may contain ";",
    ' which Analysis for Office uses between several filter values).
    cmd = Replace(Replace(cmd, vbCrLf, vbLf), vbCr, vbLf)
    Dim steps() As String
    ' SETDATA carries ";" in its member combination: a single SETDATA line is never split.
    If InStr(cmd, vbLf) > 0 Or UCase$(Left$(LTrim$(cmd), 8)) = "SETDATA " Then steps = Split(cmd, vbLf) Else steps = Split(cmd, ";")
    Dim i As Long, n As Long, stepText As String
    For i = 0 To UBound(steps)
        stepText = Trim$(steps(i))
        If stepText <> "" Then
            n = n + 1
            If variablesPaused And UCase$(Left$(stepText, 4)) <> "VAR " Then ResumeVariables
            If Not RunOne(stepText) Then
                ResumeVariables
                Quiet False
                SetCell "B3", "ERROR " & Format$(Now, "hh:nn:ss") & " in step " & n & " (" & stepText & "): " & lastMessage
                Exit Sub
            End If
        End If
    Next i
    ResumeVariables
    Quiet False
    SetCell "B3", "OK " & Format$(Now, "hh:nn:ss") & ", " & n & " step(s): " & Replace(cmd, vbLf, " / ")
End Sub

' Message windows of AfO off (quiet) or on again; the refresh mode stays as the agent set it.
Private Sub Quiet(ByVal quietOn As Boolean)
    If refreshMode = "" Then refreshMode = "On"
    On Error Resume Next
    Application.Run "SAPSetRefreshBehaviour", refreshMode, IIf(quietOn, "Off", "On")
End Sub

Private Function RunOne(ByVal stepText As String) As Boolean
    Dim verb As String, rest As String, sp As Long
    sp = InStr(stepText, " ")
    If sp > 0 Then
        verb = UCase$(Left$(stepText, sp - 1)): rest = Trim$(Mid$(stepText, sp + 1))
    Else
        verb = UCase$(stepText)
    End If
    lastMessage = "": stage = ""
    Dim before As Object
    If verb <> "SET" Then Set before = MessageKeys()
    On Error GoTo failed
    Select Case verb
    Case "SET": RunOne = DoSet(rest)
    Case "TRANSFER": RunOne = AfoCommand("PlanDataTransfer")
    Case "SAVE": RunOne = AfoCommand("PlanDataSave")
    Case "RESET": RunOne = AfoCommand("PlanDataReset")
    Case "REFRESH"
        If rest = "" Then RunOne = Check(Application.Run("SAPExecuteCommand", "Refresh")) Else RunOne = Check(Application.Run("SAPExecuteCommand", "Refresh", rest))
    Case "PLAN": RunOne = DoPlan(rest)
    Case "PVAR": RunOne = DoPlanVar(rest)
    Case "VAR": RunOne = DoVar(rest)
    Case "FILTER": RunOne = DoFilter(rest)
    Case "FILTERWHERE": RunOne = DoFilterWhere(rest)
    Case "UNFILTER": RunOne = DoUnfilter(rest)
    Case "API": RunOne = DoApi(rest)
    Case "EXPAND": RunOne = DoNode(rest, True)
    Case "COLLAPSE": RunOne = DoNode(rest, False)
    Case "DRILLLEVEL": RunOne = DoDrillLevel(rest)
    Case "MOVE": RunOne = DoMove(rest)
    Case "SETDATA": RunOne = DoSetData(rest)
    Case "CHANGEMODE": RunOne = DoMode(rest, "PlanDataToChangeMode")
    Case "DISPLAYMODE": RunOne = DoMode(rest, "PlanDataToDisplayMode")
    Case "NEWLINES": RunOne = DoNewLines(rest)
${custom}
    Case Else
        lastMessage = "unknown command " & verb
        RunOne = False
    End Select
    On Error GoTo 0
    If Not before Is Nothing Then
        Dim problems As String: problems = NewMessages(before)
        If problems <> "" Then
            If RunOne Or lastMessage = "" Then lastMessage = problems
            RunOne = False
        End If
    End If
    LogLine IIf(RunOne, "OK", "ERROR"), stepText, lastMessage
    Exit Function
failed:
    Busy = False
    lastMessage = Err.Description
    If stage <> "" Then lastMessage = lastMessage & " (at " & stage & ")"
    LogLine "ERROR", stepText, lastMessage
    RunOne = False
End Function

' MOVE DS_1 DIM=ROWS|COLUMNS|FREE, or AFTER|BEFORE|SWAP OTHER: SAPMoveDimension.
Private Function DoMove(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    Dim eq As Long: eq = InStr(rest, "=")
    If sp = 0 Or eq < sp Then GoTo badForm
    Dim ds As String: ds = Trim$(Left$(rest, sp - 1))
    Dim dimension As String: dimension = Trim$(Mid$(rest, sp + 1, eq - sp - 1))
    Dim target As String: target = Trim$(Mid$(rest, eq + 1))
    Dim action As String, other As String, p As Long
    p = InStr(target, " ")
    If p > 0 Then
        action = UCase$(Left$(target, p - 1)): other = Trim$(Mid$(target, p + 1))
    Else
        action = UCase$(target)
    End If
    Select Case action
    Case "ROWS", "COLUMNS"
        If other <> "" Then GoTo badForm
    Case "FREE", "FILTER"
        action = "FILTER"
        If other <> "" Then GoTo badForm
    Case "AFTER", "BEFORE", "SWAP"
        If other = "" Then GoTo badForm
    Case Else
        GoTo badForm
    End Select
    If other = "" Then
        DoMove = Check(Application.Run("SAPMoveDimension", ds, dimension, action))
    Else
        DoMove = Check(Application.Run("SAPMoveDimension", ds, dimension, action, other))
    End If
    If DoMove Then lastMessage = dimension & " moved: " & target
    Exit Function
badForm:
    lastMessage = "form: MOVE DS_1 DIMENSION=ROWS, COLUMNS, FREE, AFTER OTHER, BEFORE OTHER or SWAP OTHER"
End Function

' SETDATA DS_1 Key figure|DIM=key;DIM=key|value: the AfO formula SAPSetData, entered and
' calculated once in a free cell of the bridge sheet. The value stays as user input after the
' formula is removed; TRANSFER sends it.
Private Function DoSetData(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    If sp = 0 Then GoTo badForm
    Dim parts() As String: parts = Split(Mid$(rest, sp + 1), "|")
    If UBound(parts) <> 2 Then GoTo badForm
    Dim ds As String: ds = Trim$(Left$(rest, sp - 1))
    Dim measure As String: measure = Trim$(parts(0))
    Dim combo As String: combo = Trim$(parts(1))
    Dim v As String: v = Trim$(parts(2))
    If measure = "" Or combo = "" Or v = "" Or v = "-" Or v = "." Or v Like "*[!0-9.-]*" Then GoTo badForm
    Dim cell As Range: Set cell = ThisWorkbook.Worksheets(BRIDGE_SHEET).Range(SETDATA_CELL)
    Dim result As Variant
    stage = "SAPSetData"
    Busy = True
    cell.Formula = "=SAPSetData(" & Quoted(ds) & "," & v & ",,"  & Quoted(measure) & "," & Quoted(combo) & ")"
    cell.Calculate
    result = cell.Value
    cell.ClearContents
    Busy = False
    If IsError(result) Then
        lastMessage = "SAPSetData rejected " & measure & " (" & combo & "): the combination must be an input-ready cell of the current result of " & ds & " at the lowest level, all characteristics in the rows and columns given with their keys, and the data source in change mode"
        Exit Function
    End If
    lastMessage = measure & " (" & combo & ") = " & Txt(result)
    DoSetData = True
    Exit Function
badForm:
    lastMessage = "form: SETDATA DS_1 Key figure|DIM=key;DIM=key|value (number with decimal point)"
End Function

Private Function Quoted(ByVal s As String) As String
    Quoted = """" & Replace(s, """", """""") & """"
End Function

' CHANGEMODE / DISPLAYMODE DS_1: PlanDataToChangeMode / PlanDataToDisplayMode.
Private Function DoMode(ByVal ds As String, ByVal cmd As String) As Boolean
    If ds = "" Or InStr(ds, " ") > 0 Then lastMessage = "form: " & IIf(cmd = "PlanDataToChangeMode", "CHANGEMODE", "DISPLAYMODE") & " DS_1": Exit Function
    DoMode = Check(Application.Run("SAPExecuteCommand", cmd, ds))
    If DoMode Then lastMessage = ds & IIf(cmd = "PlanDataToChangeMode", ": change mode", ": display mode")
End Function

' NEWLINES Crosstab1 n: the number of empty lines for new member combinations below the crosstab.
Private Function DoNewLines(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    If sp = 0 Then lastMessage = "form: NEWLINES Crosstab1 5": Exit Function
    Dim n As String: n = Trim$(Mid$(rest, sp + 1))
    If n = "" Or n Like "*[!0-9]*" Then lastMessage = "form: NEWLINES Crosstab1 5": Exit Function
    DoNewLines = Check(Application.Run("SAPExecuteCommand", "NumberOfNewLines", Trim$(Left$(rest, sp - 1)), n))
    If DoNewLines Then lastMessage = Trim$(Left$(rest, sp - 1)) & ": " & n & " new line(s)"
End Function

Private Function DoSet(ByVal rest As String) As Boolean
    Dim eq As Long: eq = InStr(rest, "=")
    Dim bang As Long: bang = InStrRev(rest, "!", IIf(eq > 0, eq, Len(rest)))
    If eq = 0 Or bang = 0 Then lastMessage = "form: SET Sheet!A1=value": Exit Function
    Dim sheetName As String: sheetName = Trim$(Left$(rest, bang - 1))
    If Left$(sheetName, 1) = "'" And Right$(sheetName, 1) = "'" Then sheetName = Replace(Mid$(sheetName, 2, Len(sheetName) - 2), "''", "'")
    Dim target As Range: Set target = ThisWorkbook.Worksheets(sheetName).Range(Trim$(Mid$(rest, bang + 1, eq - bang - 1)))
    Dim v As String: v = Trim$(Mid$(rest, eq + 1))
    If v = "" Then
        target.ClearContents
    ElseIf Len(v) >= 2 And Left$(v, 1) = """" And Right$(v, 1) = """" Then
        ' "01.2026": text as typed, e.g. a member key in a new line that looks like a number.
        target.Value = "'" & Mid$(v, 2, Len(v) - 2)
        v = Mid$(v, 2, Len(v) - 2)
    ElseIf v Like "*[!0-9.-]*" Or v = "-" Or v = "." Or v Like "0[0-9]*" Then
        target.Value = "'" & v
    Else
        target.Value = Val(v)
    End If
    lastMessage = sheetName & "!" & target.Address(False, False) & " = " & v
    DoSet = True
End Function

Private Function DoPlan(ByVal key As String) As Boolean
    If key = "" Then lastMessage = "form: PLAN PF_1": Exit Function
    If IsSequence(key) Then
        DoPlan = Check(Application.Run("SAPExecutePlanningSequence", key))
    Else
        DoPlan = Check(Application.Run("SAPExecutePlanningFunction", key))
    End If
End Function

Private Function IsSequence(ByVal key As String) As Boolean
    Dim list As Variant, i As Long
    On Error GoTo none
    list = Application.Run("SAPListOf", "PLANNINGSEQUENCES")
    If IsArray(list) Then
        If ArrayRank(list) = 2 Then
            For i = LBound(list, 1) To UBound(list, 1)
                If UCase$(Txt(list(i, LBound(list, 2)))) = UCase$(key) Then IsSequence = True: Exit Function
            Next i
        Else
            If UCase$(Txt(list(LBound(list)))) = UCase$(key) Then IsSequence = True: Exit Function
        End If
    End If
none:
    If Not IsSequence Then IsSequence = (UCase$(Left$(key, 3)) = "PS_")
End Function

Private Function DoPlanVar(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    Dim eq As Long: eq = InStr(rest, "=")
    If sp = 0 Or eq < sp Then lastMessage = "form: PVAR PF_1 NAME=value": Exit Function
    ' The fourth argument names the value format; without it AfO answers "Invalid parameter".
    DoPlanVar = Check(Application.Run("SAPSetPlanParameter", Trim$(Left$(rest, sp - 1)), Trim$(Mid$(rest, sp + 1, eq - sp - 1)), Trim$(Mid$(rest, eq + 1)), "INPUT_STRING"))
End Function

Private Function DoVar(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    Dim eq As Long: eq = InStr(rest, "=")
    If sp = 0 Or eq < sp Then lastMessage = "form: VAR DS_1 NAME=value": Exit Function
    Dim ds As String: ds = Trim$(Left$(rest, sp - 1))
    Dim name As String: name = Trim$(Mid$(rest, sp + 1, eq - sp - 1))
    If Not variablesPaused Then
        Application.Run "SAPExecuteCommand", "PauseVariableSubmit", "On"
        variablesPaused = True
    End If
    DoVar = Check(Application.Run("SAPSetVariable", name, Trim$(Mid$(rest, eq + 1)), "INPUT_STRING", ds))
End Function

Private Sub ResumeVariables()
    If Not variablesPaused Then Exit Sub
    variablesPaused = False
    Application.Run "SAPExecuteCommand", "PauseVariableSubmit", "Off"
End Sub

Private Function DoFilter(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    Dim eq As Long: eq = InStr(rest, "=")
    If sp = 0 Or eq < sp Then lastMessage = "form: FILTER DS_1 DIMENSION=value": Exit Function
    Dim v As String: v = Trim$(Mid$(rest, eq + 1))
    ' @Sheet!A1:A50: the keys in that range, as one filter.
    If Left$(v, 1) = "@" Then
        v = RangeKeys(Mid$(v, 2))
        If v = "" Then lastMessage = "no keys in " & Trim$(Mid$(rest, eq + 2)): Exit Function
    End If
    DoFilter = SetFilter(Trim$(Left$(rest, sp - 1)), Trim$(Mid$(rest, sp + 1, eq - sp - 1)), v)
End Function

Private Function SetFilter(ByVal ds As String, ByVal dimension As String, ByVal members As String) As Boolean
    SetFilter = Check(Application.Run("SAPSetFilter", ds, dimension, members, "INPUT_STRING"))
    If SetFilter Then
        If members = "" Then
            lastMessage = dimension & ": filter removed"
        Else
            lastMessage = dimension & ": " & (UBound(Split(members, ";")) + 1) & " value(s)"
        End If
    End If
End Function

' The non-empty cells of a range (Sheet!A1:A50), joined by ";".
Private Function RangeKeys(ByVal ref As String) As String
    Dim bang As Long: bang = InStrRev(ref, "!")
    Dim r As Range, c As Range, s As String
    If bang > 0 Then
        Dim sheetName As String: sheetName = Trim$(Left$(ref, bang - 1))
        If Left$(sheetName, 1) = "'" And Right$(sheetName, 1) = "'" Then sheetName = Replace(Mid$(sheetName, 2, Len(sheetName) - 2), "''", "'")
        Set r = ThisWorkbook.Worksheets(sheetName).Range(Trim$(Mid$(ref, bang + 1)))
    Else
        Set r = ThisWorkbook.Worksheets(BRIDGE_SHEET).Range(Trim$(ref))
    End If
    For Each c In r.Cells
        If Trim$(Txt(c.Value)) <> "" Then
            If s <> "" Then s = s & ";"
            s = s & Trim$(Txt(c.Value))
        End If
    Next c
    RangeKeys = s
End Function

Private Function DoUnfilter(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    If sp = 0 Then lastMessage = "form: UNFILTER DS_1 DIMENSION": Exit Function
    DoUnfilter = SetFilter(Trim$(Left$(rest, sp - 1)), Trim$(Mid$(rest, sp + 1)), "")
End Function

' FILTERWHERE DS_1 DIMENSION "Key figure" >= 100000: reads the crosstab of the data source, keeps
' the members of DIMENSION whose value in the key figure column meets the condition, and filters
' DIMENSION to them. The crosstab must show the member keys of DIMENSION.
Private Function DoFilterWhere(ByVal rest As String) As Boolean
    Dim q1 As Long: q1 = InStr(rest, """")
    Dim q2 As Long: If q1 > 0 Then q2 = InStr(q1 + 1, rest, """")
    Const form As String = "form: FILTERWHERE DS_1 DIMENSION ""Key figure"" >= 100000"
    If q1 = 0 Or q2 = 0 Then lastMessage = form: Exit Function
    Dim head() As String: head = Split(Application.WorksheetFunction.Trim(Left$(rest, q1 - 1)), " ")
    If UBound(head) <> 1 Then lastMessage = form: Exit Function
    Dim ds As String: ds = head(0)
    Dim dimension As String: dimension = head(1)
    Dim measure As String: measure = Normalize(Mid$(rest, q1 + 1, q2 - q1 - 1))
    Dim cond As String: cond = Trim$(Mid$(rest, q2 + 1))
    Dim op As String, opCandidate As Variant
    For Each opCandidate In Array(">=", "<=", "<>", ">", "<", "=")
        If Left$(cond, Len(opCandidate)) = opCandidate Then op = opCandidate: Exit For
    Next opCandidate
    If op = "" Then lastMessage = form: Exit Function
    Dim limitText As String: limitText = Trim$(Mid$(cond, Len(op) + 1))
    If limitText = "" Or limitText Like "*[!0-9.-]*" Then lastMessage = "threshold must be a number with decimal point, no thousands separator: " & limitText: Exit Function
    Dim limit As Double: limit = Val(limitText)

    Dim topLeft As Range: Set topLeft = CrosstabTopLeft(ds)
    If topLeft Is Nothing Then lastMessage = "no crosstab of " & ds: Exit Function
    Dim area As Range: Set area = topLeft.CurrentRegion
    Dim dimText As String: dimText = Normalize(DimensionText(ds, dimension))
    If dimText = "" Then lastMessage = dimension & " is not a dimension of " & ds: Exit Function

    ' Header: the row holding the dimension caption; the key figure column by its header text.
    Dim r As Long, c As Long, dimRow As Long, dimCol As Long, kfCol As Long, kfPartial As Long
    Dim firstRow As Long: firstRow = area.Row
    Dim lastRow As Long: lastRow = area.Row + area.Rows.Count - 1
    Dim firstCol As Long: firstCol = area.Column
    Dim lastCol As Long: lastCol = area.Column + area.Columns.Count - 1
    Dim ws As Worksheet: Set ws = topLeft.Worksheet
    For r = firstRow To Application.Min(lastRow, firstRow + 15)
        For c = firstCol To lastCol
            Dim t As String: t = Normalize(Txt(ws.Cells(r, c).Value))
            If t <> "" Then
                If dimCol = 0 And StrComp(t, dimText, vbTextCompare) = 0 Then dimRow = r: dimCol = c
                If kfCol = 0 And StrComp(t, measure, vbTextCompare) = 0 Then kfCol = c
                If kfPartial = 0 And InStr(1, t, measure, vbTextCompare) > 0 Then kfPartial = c
            End If
        Next c
    Next r
    If kfCol = 0 Then kfCol = kfPartial
    If dimCol = 0 Then lastMessage = "the crosstab of " & ds & " does not show " & dimension & " (" & dimText & ") in its rows": Exit Function
    If kfCol = 0 Then lastMessage = "no column """ & measure & """ in the crosstab of " & ds: Exit Function

    Dim keys As Object: Set keys = CreateObject("Scripting.Dictionary")
    Dim key As String, current As String, v As Variant, ok As Boolean, scanned As Long
    For r = dimRow + 1 To lastRow
        key = Trim$(Txt(ws.Cells(r, dimCol).Value))
        ' Repeated keys of an outer dimension are shown once; the rows below belong to it.
        If key <> "" Then current = key
        v = ws.Cells(r, kfCol).Value2
        If current <> "" And Not IsError(v) And Not IsEmpty(v) Then
            If IsNumeric(v) And VarType(v) <> vbString Then
                scanned = scanned + 1
                Select Case op
                Case ">=": ok = (CDbl(v) >= limit)
                Case "<=": ok = (CDbl(v) <= limit)
                Case "<>": ok = (CDbl(v) <> limit)
                Case ">": ok = (CDbl(v) > limit)
                Case "<": ok = (CDbl(v) < limit)
                Case Else: ok = (CDbl(v) = limit)
                End Select
                If ok Then keys(current) = True
            End If
        End If
    Next r
    If scanned = 0 Then lastMessage = "no values in column """ & measure & """": Exit Function
    If keys.Count = 0 Then lastMessage = "no member of " & dimension & " meets " & measure & " " & op & " " & limitText & " (" & scanned & " rows checked)": Exit Function
    DoFilterWhere = SetFilter(ds, dimension, Join(keys.keys, ";"))
    If DoFilterWhere Then lastMessage = dimension & ": " & keys.Count & " of " & scanned & " rows meet " & measure & " " & op & " " & limitText
End Function

' Header text without "[+]"/"[-]", line breaks and doubled blanks.
Private Function Normalize(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), "[+]", "")
    s = Replace(s, "[-]", "")
    Normalize = Application.WorksheetFunction.Trim(s)
End Function

' Top left cell of the first crosstab bound to the data source (defined name SAP<crosstab>).
Private Function CrosstabTopLeft(ByVal ds As String) As Range
    Dim list As Variant, i As Long
    On Error GoTo none
    list = Application.Run("SAPListOf", "CROSSTABS")
    If Not IsArray(list) Then GoTo none
    If ArrayRank(list) = 2 Then
        For i = LBound(list, 1) To UBound(list, 1)
            If UCase$(Txt(list(i, UBound(list, 2)))) = UCase$(ds) Then Set CrosstabTopLeft = ThisWorkbook.Names("SAP" & Txt(list(i, LBound(list, 2)))).RefersToRange: Exit Function
        Next i
    ElseIf UCase$(Txt(list(UBound(list)))) = UCase$(ds) Then
        Set CrosstabTopLeft = ThisWorkbook.Names("SAP" & Txt(list(LBound(list)))).RefersToRange
    End If
none:
End Function

' Description of a dimension as the crosstab shows it in its header.
Private Function DimensionText(ByVal ds As String, ByVal dimension As String) As String
    Dim list As Variant, i As Long
    On Error GoTo none
    list = Application.Run("SAPListOfDimensions", ds)
    If Not IsArray(list) Then GoTo none
    If ArrayRank(list) = 2 Then
        For i = LBound(list, 1) To UBound(list, 1)
            If UCase$(Txt(list(i, LBound(list, 2)))) = UCase$(dimension) Then DimensionText = Txt(list(i, LBound(list, 2) + 1)): Exit Function
        Next i
    End If
none:
End Function

' EXPAND DS_1 node / COLLAPSE DS_1 node: node by key or text as shown, or its cell (Sheet!B9).
Private Function DoNode(ByVal rest As String, ByVal expand As Boolean) As Boolean
    Dim verb As String: verb = IIf(expand, "EXPAND", "COLLAPSE")
    Dim sp As Long: sp = InStr(rest, " ")
    If sp = 0 Then lastMessage = "form: " & verb & " DS_1 node (key or text as shown, or Sheet!B9)": Exit Function
    Dim ds As String: ds = Trim$(Left$(rest, sp - 1))
    Dim node As String: node = Trim$(Mid$(rest, sp + 1))
    stage = "find node"
    Dim c As Range: Set c = FindNode(ds, node)
    If c Is Nothing Then Exit Function
    stage = "node state"
    Dim state As String: state = NodeState(c)
    If state = "" Then lastMessage = node & " (" & c.Address(False, False) & ") shows no [+] or [-]: a leaf, or no hierarchy is shown": Exit Function
    If (state = "-") = expand Then
        lastMessage = node & " is already " & IIf(expand, "expanded", "collapsed")
        DoNode = True
        Exit Function
    End If
    Dim back As Worksheet: Set back = ActiveSheet
    DoNode = ToggleNode(c)
    RestoreSheet back
    If DoNode Then lastMessage = Trim$(Txt(c.Value)) & IIf(expand, " expanded", " collapsed") & " (" & c.Worksheet.Name & "!" & c.Address(False, False) & ")"
End Function

' DRILLLEVEL DS_1 DIMENSION=n: nodes above level n expanded, nodes on level n and below collapsed.
Private Function DoDrillLevel(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    Dim eq As Long: eq = InStr(rest, "=")
    Const form As String = "form: DRILLLEVEL DS_1 DIMENSION=2 (1 = top nodes only)"
    If sp = 0 Or eq < sp Then lastMessage = form: Exit Function
    Dim ds As String: ds = Trim$(Left$(rest, sp - 1))
    Dim dimension As String: dimension = Trim$(Mid$(rest, sp + 1, eq - sp - 1))
    Dim levelText As String: levelText = Trim$(Mid$(rest, eq + 1))
    If levelText = "" Or levelText Like "*[!0-9]*" Then lastMessage = form: Exit Function
    Dim level As Long: level = CLng(levelText)
    If level < 1 Then lastMessage = form: Exit Function
    Dim back As Worksheet: Set back = ActiveSheet
    Dim changed As Long, c As Range, target As Range, r As Long, col As Long, headRow As Long, area As Range, state As String
    Do
        Set area = CrosstabArea(ds)
        If area Is Nothing Then lastMessage = "no crosstab of " & ds: GoTo done
        If Not NodeColumn(ds, dimension, area, col, headRow) Then GoTo done
        Set target = Nothing
        For r = headRow + 1 To area.Row + area.Rows.Count - 1
            Set c = area.Worksheet.Cells(r, col)
            state = NodeState(c)
            If state = "+" And NodeLevel(c) + 1 < level Then Set target = c: Exit For
            If state = "-" And NodeLevel(c) + 1 >= level Then Set target = c: Exit For
        Next r
        If target Is Nothing Then Exit Do
        If Not ToggleNode(target) Then GoTo done
        changed = changed + 1
        If changed >= 200 Then lastMessage = "stopped after 200 nodes; set a lower level or filter first": GoTo done
    Loop
    lastMessage = dimension & ": levels 1 to " & level & " shown (" & changed & " node(s) changed)"
    DoDrillLevel = True
done:
    RestoreSheet back
End Function

' The crosstab area of a data source, as the crosstab shows it now.
Private Function CrosstabArea(ByVal ds As String) As Range
    Dim topLeft As Range: Set topLeft = CrosstabTopLeft(ds)
    If Not topLeft Is Nothing Then Set CrosstabArea = topLeft.CurrentRegion
End Function

' A cell AfO renders as a hierarchy node: a [+]/[-] prefix, or (a leaf) the indentation prefix in a
' hierarchy cell style. Text cells next to the node carry the style too, but no prefix.
Private Function IsNodeCell(ByVal c As Range) As Boolean
    If NodeState(c) <> "" Then IsNodeCell = True: Exit Function
    If Left$(Txt(c.NumberFormat), 2) <> """ " Then Exit Function
    On Error Resume Next
    IsNodeCell = (c.Style.Name Like "SAPHierarchy*")
End Function

' "+" collapsed, "-" expanded, "" leaf or no hierarchy. AfO writes the symbol into the number format.
Private Function NodeState(ByVal c As Range) As String
    Dim nf As String: nf = Txt(c.NumberFormat)
    If InStr(nf, "[+]") > 0 Then
        NodeState = "+"
    ElseIf InStr(nf, "[-]") > 0 Then
        NodeState = "-"
    End If
End Function

' Zero-based level of a node: AfO indents each level by two blanks before the symbol.
Private Function NodeLevel(ByVal c As Range) As Long
    Dim nf As String: nf = Txt(c.NumberFormat)
    If Left$(nf, 1) = """" Then nf = Mid$(nf, 2)
    Dim i As Long
    Do While Mid$(nf, i + 1, 1) = " "
        i = i + 1
    Loop
    NodeLevel = i \\ 2
End Function

' The node cell for a key or text as shown (or a cell reference), unique in the crosstab of ds.
Private Function FindNode(ByVal ds As String, ByVal node As String) As Range
    Dim c As Range, hit As Range, n As Long, places As String
    If node Like "*!*" Then
        Dim bang As Long: bang = InStrRev(node, "!")
        Dim sheetName As String: sheetName = Trim$(Left$(node, bang - 1))
        If Left$(sheetName, 1) = "'" And Right$(sheetName, 1) = "'" Then sheetName = Replace(Mid$(sheetName, 2, Len(sheetName) - 2), "''", "'")
        On Error Resume Next
        Set c = ThisWorkbook.Worksheets(sheetName).Range(Trim$(Mid$(node, bang + 1))).Cells(1, 1)
        On Error GoTo 0
        If c Is Nothing Then lastMessage = "no cell " & node: Exit Function
        If Not IsNodeCell(c) Then lastMessage = node & " is not a hierarchy node": Exit Function
        Set FindNode = c
        Exit Function
    End If
    stage = "crosstab"
    Dim area As Range: Set area = CrosstabArea(ds)
    If area Is Nothing Then lastMessage = "no crosstab of " & ds: Exit Function
    stage = "scan nodes"
    For Each c In area.Cells
        If IsNodeCell(c) Then
            If StrComp(Trim$(Txt(c.Value)), node, vbTextCompare) = 0 Or StrComp(Trim$(Txt(c.Offset(0, 1).Value)), node, vbTextCompare) = 0 Then
                n = n + 1
                If hit Is Nothing Then Set hit = c
                If n <= 5 Then places = places & " " & c.Address(False, False)
            End If
        End If
    Next c
    If n = 0 Then lastMessage = "no hierarchy node " & node & " in the crosstab of " & ds & " (is it shown? expand its parent first)": Exit Function
    If n > 1 Then lastMessage = node & " appears " & n & " times (" & Trim$(places) & "); give the cell, e.g. Sheet!" & hit.Address(False, False): Exit Function
    Set FindNode = hit
End Function

' Column of the crosstab holding the nodes of a dimension, found by its header; row of that header.
Private Function NodeColumn(ByVal ds As String, ByVal dimension As String, ByVal area As Range, col As Long, headRow As Long) As Boolean
    Dim dimText As String: dimText = Normalize(DimensionText(ds, dimension))
    If dimText = "" Then lastMessage = dimension & " is not a dimension of " & ds: Exit Function
    Dim r As Long, c As Long, ws As Worksheet: Set ws = area.Worksheet
    For r = area.Row To Application.Min(area.Row + area.Rows.Count - 1, area.Row + 15)
        For c = area.Column To area.Column + area.Columns.Count - 1
            If StrComp(Normalize(Txt(ws.Cells(r, c).Value)), dimText, vbTextCompare) = 0 Then
                col = c: headRow = r: NodeColumn = True
                Exit Function
            End If
        Next c
    Next r
    lastMessage = "the crosstab of " & ds & " does not show " & dimension & " (" & dimText & ") in its rows"
End Function

' Click on the symbol of a node as a user does and wait until AfO has redrawn it.
' AfO takes the position from the mouse pointer, so the pointer is moved there for the click.
Private Function ToggleNode(ByVal c As Range) As Boolean
    Dim before As String: before = NodeState(c)
    stage = "show sheet"
    c.Worksheet.Parent.Activate
    c.Worksheet.Activate
    If Application.WindowState = xlMinimized Or ActiveWindow.WindowState = xlMinimized Then lastMessage = "the Excel window is minimized; a node can only be clicked in a visible window": Exit Function
    stage = "scroll"
    If Intersect(ActiveWindow.VisibleRange, c) Is Nothing Then
        ActiveWindow.ScrollRow = Application.Max(1, c.Row - 2)
        ActiveWindow.ScrollColumn = Application.Max(1, c.Column - 1)
    End If
    stage = "pane"
    ' The pane that shows the cell (frozen panes); Panes has no enumerator, hence the index.
    Dim i As Long, p As Object: Set p = ActiveWindow.ActivePane
    For i = 1 To ActiveWindow.Panes.Count
        If Not Intersect(ActiveWindow.Panes(i).VisibleRange, c) Is Nothing Then Set p = ActiveWindow.Panes(i)
    Next i
    ' The symbol sits at the left edge of the cell: AfO counts the indentation in.
    stage = "position"
    Dim pt As PointApi
    pt.x = p.PointsToScreenPixelsX(c.Left) + 3
    pt.y = p.PointsToScreenPixelsY(c.Top + c.Height / 2)
    Dim found As Object
    On Error Resume Next
    Set found = ActiveWindow.RangeFromPoint(pt.x, pt.y)
    On Error GoTo 0
    Dim onCell As Boolean
    If Not found Is Nothing Then
        If TypeName(found) = "Range" Then onCell = (found.Address = c.Address And found.Worksheet.Name = c.Worksheet.Name)
    End If
    If Not onCell Then lastMessage = "cell " & c.Address(False, False) & " is not where Excel shows it (zoom, or another window over it?)": Exit Function
    stage = "window"
    Dim wnd As LongPtr: wnd = FindWindowExA(Application.Hwnd, 0, "XLDESK", vbNullString)
    If wnd <> 0 Then wnd = FindWindowExA(wnd, 0, "EXCEL7", vbNullString)
    If wnd = 0 Then lastMessage = "workbook window not found": Exit Function
    stage = "click"
    Dim old As PointApi: GetCursorPos old
    SetCursorPos pt.x, pt.y
    Dim client As PointApi: client = pt
    ScreenToClient wnd, client
    Dim lp As LongPtr: lp = CLngPtr(client.y) * &H10000 + (client.x And &HFFFF&)
    PostMessageA wnd, &H201, 1, lp
    PostMessageA wnd, &H202, 0, lp
    DoEvents
    SetCursorPos old.x, old.y
    stage = "wait"
    Dim started As Single: started = Timer
    Do
        DoEvents
        If NodeState(c) <> before Then ToggleNode = True: Exit Function
    Loop While Abs(Timer - started) < 20
    lastMessage = "Analysis for Office did not react to the click on " & c.Address(False, False)
End Function

Private Sub RestoreSheet(ByVal back As Worksheet)
    On Error Resume Next
    If Not ActiveSheet Is back Then back.Parent.Activate: back.Activate
End Sub

' Any AfO API function: API SAPFunction arg1|arg2|... (up to 10 arguments).
Private Function DoApi(ByVal rest As String) As Boolean
    Dim sp As Long: sp = InStr(rest, " ")
    Dim fn As String, args() As String
    If sp > 0 Then
        fn = Trim$(Left$(rest, sp - 1)): args = Split(Trim$(Mid$(rest, sp + 1)), "|")
    Else
        fn = Trim$(rest): args = Split("", "|")
    End If
    If Not fn Like "SAP*" Or fn Like "*[!A-Za-z0-9_]*" Then lastMessage = "form: API SAPFunction arg1|arg2": Exit Function
    If UCase$(fn) = "SAPLOGON" Then lastMessage = "SAPLogon is not allowed through the bridge": Exit Function
    Dim n As Long: n = UBound(args) - LBound(args) + 1
    If n > 10 Then lastMessage = "at most 10 arguments": Exit Function
    Dim a(1 To 10) As Variant, i As Long
    For i = 1 To n: a(i) = Trim$(args(LBound(args) + i - 1)): Next i
    Dim result As Variant
    Select Case n
    Case 0: result = Application.Run(fn)
    Case 1: result = Application.Run(fn, a(1))
    Case 2: result = Application.Run(fn, a(1), a(2))
    Case 3: result = Application.Run(fn, a(1), a(2), a(3))
    Case 4: result = Application.Run(fn, a(1), a(2), a(3), a(4))
    Case 5: result = Application.Run(fn, a(1), a(2), a(3), a(4), a(5))
    Case 6: result = Application.Run(fn, a(1), a(2), a(3), a(4), a(5), a(6))
    Case 7: result = Application.Run(fn, a(1), a(2), a(3), a(4), a(5), a(6), a(7))
    Case 8: result = Application.Run(fn, a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8))
    Case 9: result = Application.Run(fn, a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9))
    Case Else: result = Application.Run(fn, a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9), a(10))
    End Select
    If UCase$(fn) = "SAPSETREFRESHBEHAVIOUR" And n >= 1 Then
        If UCase$(a(1)) = "ON" Or UCase$(a(1)) = "OFF" Then refreshMode = a(1)
        Quiet True
    End If
    Dim shown As String: shown = Flatten(result)
    If IsError(result) Or shown = "0" Then
        lastMessage = Txt(Application.Run("SAPGetProperty", "LastError", "Text"))
        If lastMessage = "" Then lastMessage = fn & " returned " & IIf(IsError(result), "an error", "0")
        Exit Function
    End If
    lastMessage = fn & " = " & shown
    DoApi = True
End Function

' A result as one line: arrays row by row, fields separated by " | ", rows by " / ".
Private Function Flatten(ByVal v As Variant) As String
    Dim i As Long, j As Long, s As String
    If Not IsArray(v) Then Flatten = Txt(v): Exit Function
    If ArrayRank(v) = 2 Then
        For i = LBound(v, 1) To UBound(v, 1)
            If i > LBound(v, 1) Then s = s & " / "
            For j = LBound(v, 2) To UBound(v, 2)
                If j > LBound(v, 2) Then s = s & " | "
                s = s & Txt(v(i, j))
            Next j
        Next i
    Else
        For i = LBound(v) To UBound(v)
            If i > LBound(v) Then s = s & " | "
            s = s & Txt(v(i))
        Next i
    End If
    If Len(s) > 2000 Then s = Left$(s, 2000) & " ..."
    Flatten = s
End Function

' AfO disables PlanDataTransfer, PlanDataReset and PlanDataSave while nothing is pending: not a failure.
Private Function AfoCommand(ByVal cmd As String) As Boolean
    AfoCommand = Check(Application.Run("SAPExecuteCommand", cmd))
    If Not AfoCommand And InStr(1, lastMessage, "not enabled", vbTextCompare) > 0 Then
        lastMessage = "nothing pending (" & cmd & " not enabled)"
        AfoCommand = True
    End If
End Function

Private Function RunMacro(ByVal macro As String) As Boolean
    Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macro
    RunMacro = True
End Function

' AfO API calls return 1 on success.
Private Function Check(ByVal result As Variant) As Boolean
    Check = (Txt(result) = "1")
    If Not Check Then
        On Error Resume Next
        lastMessage = Txt(Application.Run("SAPGetProperty", "LastError", "Text"))
        If lastMessage = "" Then lastMessage = "Analysis for Office returned " & IIf(Txt(result) = "", "an error", Txt(result)) & " without a message"
    End If
End Function

' The AfO messages currently listed, one key per message (all fields joined).
Private Function MessageKeys() As Object
    Dim keys As Object: Set keys = CreateObject("Scripting.Dictionary")
    Dim row As Variant
    For Each row In MessageRows()
        keys(Join(row, "|")) = True
    Next row
    Set MessageKeys = keys
End Function

' Logs the messages that are not in before; returns the texts of new errors.
Private Function NewMessages(ByVal before As Object) As String
    Dim row As Variant, i As Long, severity As String, text As String, problems As String
    For Each row In MessageRows()
        If Not before.Exists(Join(row, "|")) Then
            severity = "": text = ""
            For i = LBound(row) To UBound(row)
                Select Case UCase$(row(i))
                Case "CRITICAL", "ERROR", "WARNING", "INFORMATION", "SUCCESS": If severity = "" Then severity = UCase$(row(i))
                End Select
            Next i
            If UBound(row) >= 1 Then text = row(1) Else text = row(0)
            LogLine "MESSAGE", severity, text
            If severity = "ERROR" Or severity = "CRITICAL" Then
                If problems <> "" Then problems = problems & " / "
                problems = problems & text
            End If
        End If
    Next row
    NewMessages = problems
End Function

' SAPListOfMessages as rows of text fields: one row for a single message, none without messages.
Private Function MessageRows() As Collection
    Dim rows As New Collection, v As Variant, i As Long, j As Long, f() As String
    On Error GoTo done
    v = Application.Run("SAPListOfMessages", "INFORMATION", "TRUE")
    If Not IsArray(v) Then GoTo done
    If ArrayRank(v) = 2 Then
        For i = LBound(v, 1) To UBound(v, 1)
            ReDim f(0 To UBound(v, 2) - LBound(v, 2))
            For j = LBound(v, 2) To UBound(v, 2): f(j - LBound(v, 2)) = Txt(v(i, j)): Next j
            rows.Add f
        Next i
    Else
        ReDim f(0 To UBound(v) - LBound(v))
        For j = LBound(v) To UBound(v): f(j - LBound(v)) = Txt(v(j)): Next j
        rows.Add f
    End If
done:
    Set MessageRows = rows
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or IsObject(v) Then Exit Function
    Txt = CStr(v)
End Function

Private Function ArrayRank(ByVal a As Variant) As Long
    Dim r As Long, x As Long
    On Error GoTo done
    Do
        r = r + 1
        x = UBound(a, r)
    Loop
done:
    ArrayRank = r - 1
End Function

Private Sub SetCell(ByVal address As String, ByVal value As Variant)
    Busy = True
    ThisWorkbook.Worksheets(BRIDGE_SHEET).Range(address).Value = value
    Busy = False
End Sub

' One log line; the oldest rows go once the log holds more than LOG_KEEP.
Private Sub LogLine(ByVal kind As String, ByVal target As String, ByVal detail As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(BRIDGE_SHEET)
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < LOG_FIRST_ROW Then r = LOG_FIRST_ROW
    Busy = True
    If r - LOG_FIRST_ROW >= LOG_KEEP Then
        ws.Range(ws.Cells(LOG_FIRST_ROW, 1), ws.Cells(LOG_FIRST_ROW + 99, 4)).Delete Shift:=xlUp
        r = r - 100
    End If
    ws.Cells(r, 1).Value = "'" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).Value = kind
    ws.Cells(r, 3).Value = "'" & target
    ws.Cells(r, 4).Value = "'" & detail
    Busy = False
End Sub
`;
}

/** Event handlers for ThisWorkbook, appended to the workbook's own document module code. */
export function agentThisWorkbookCode(own: string | undefined): string {
  const handlers = `Private Sub Workbook_Open()
    ${AGENT_MODULE}.MarkActive
End Sub

Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
    If Sh.Name = ${AGENT_MODULE}.BRIDGE_SHEET Then ${AGENT_MODULE}.OnCalculate
End Sub
`;
  if (!own?.trim()) return `Option Explicit\r\n\r\n${handlers}`;
  return `${own.replace(/\s+$/, '')}\r\n\r\n${handlers}`;
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Column letters → zero-based index. */
function colIndex(letters: string): number {
  return [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
}
function colLetters(index: number): string {
  let s = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** The bridge sheet as written into the package, and the first row of its log. */
export function agentSheetXml(spec: AfoWorkbookSpec): { xml: string; logFirstRow: number } {
  type Cell = { v?: string; f?: string };
  const rows: Array<Map<number, Cell>> = [];
  const put = (...cells: Array<string | Cell>) => {
    rows.push(new Map(cells.map((c, i) => [i, typeof c === 'string' ? { v: c } : c])));
  };
  put('Command');
  put('Trigger', { f: 'LEN(B1)' });
  put('Result', '');
  put('Bridge', 'inactive: enable macros (content) for this workbook');
  put('');
  put('Write a command into B1, several one per line. B1 empties when the bridge takes it; B3 shows the result, the log below every step and every AfO message.');
  put('Values written straight into cells by an add-in (Office.js) are not seen by Analysis for Office: write plan values with SET.');
  put('');
  put('Commands');
  for (const [cmd, text] of BUILTIN_COMMANDS) put(cmd, text);
  for (const c of spec.agentBridge?.commands ?? []) put(c.name, `${c.description ?? 'Workbook command'} (runs ${c.macro}).`);
  put('');
  put('Rules');
  for (const r of RULES) put('', r);
  put('');
  put('Reading', 'AfO read functions are formulas: write them into free cells from AD2, read the result, delete them.');
  for (const [f, text] of READING) put(f, text);
  put('');
  put('This workbook (as built; live lists from column Q)');
  for (const dp of spec.dataProviders) put(dp.key, `Data source on query ${dp.query}${dp.inputReady ?? true ? ', input-ready' : ''}.`);
  for (const pf of spec.planningFunctions ?? []) put(pf.key, `Planning function ${pf.name}${pf.description ? `: ${pf.description}` : ''} (on ${pf.dataProviderKey}).`);
  for (const g of spec.grids) put(g.key, `Crosstab of ${g.dataProviderKey}, top left ${g.sheet}!${g.cell}.`);
  put('');
  put('Time', 'Event', 'Step', 'Detail');
  const logFirstRow = rows.length + 1;

  // Headers of the message area, the live lists and the read area in row 1.
  const first = rows[0];
  LAYOUT.messages.headers.forEach((h, i) => { if (h) first.set(colIndex(LAYOUT.messages.col) + i, { v: h }); });
  for (const l of LAYOUT.lists) first.set(colIndex(l.col), { v: l.title });
  first.set(colIndex(LAYOUT.readArea.col), { v: LAYOUT.readArea.title });

  const data = rows.map((cells, r) => `<row r="${r + 1}">${[...cells.entries()].sort((a, b) => a[0] - b[0]).map(([i, c]) => {
    const ref = `${colLetters(i)}${r + 1}`;
    if (c.f) return `<c r="${ref}"><f>${esc(c.f)}</f></c>`;
    if (!c.v) return '';
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(c.v)}</t></is></c>`;
  }).join('')}</row>`).join('');
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + `<sheetPr codeName="${AGENT_SHEET_CODE_NAME}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>`
    + `<cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="2" width="22" customWidth="1"/><col min="3" max="3" width="40" customWidth="1"/><col min="4" max="4" width="60" customWidth="1"/></cols>`
    + `<sheetData>${data}</sheetData><pageMargins left="0.7" right="0.7" top="0.787401575" bottom="0.787401575" header="0.3" footer="0.3"/></worksheet>`;
  return { xml, logFirstRow };
}
