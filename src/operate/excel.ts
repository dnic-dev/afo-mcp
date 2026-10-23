// Shared side of the operate tools: a PowerShell script per call that attaches to Excel and calls
// the AfO API with retries while Excel is busy.
//
// Dialogs: a COM call into Excel does not return while Excel or AfO shows a modal dialog (prompt
// dialog, AfO logon, message box, repair or file-in-use prompt). The script reports every step on
// stdout as it goes. When no step follows for a while, the server looks from outside whether the
// Excel process shows a dialog; if so, it stops its own script and reports the dialog. Excel and
// the dialog stay as they are, for the user to answer. Nothing is clicked.

import { execFile, spawn } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

/** How long a step may run without progress before the server looks for a dialog. */
const STALL_SECONDS = 20;

const CSHARP = String.raw`
using System; using System.Text; using System.Collections.Generic;
using System.Runtime.InteropServices; using System.Windows.Automation;
public class AfoDialog { public string kind; public string title; public string text; public string[] buttons; }
public static class AfoXl {
  public delegate bool EnumProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f, IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc f, IntPtr p);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object obj);
  static string Cls(IntPtr h) { var c = new StringBuilder(256); GetClassName(h, c, 256); return c.ToString(); }
  static string Title(IntPtr h) { var t = new StringBuilder(512); GetWindowText(h, t, 512); return t.ToString(); }

  public static object Application(uint pid) {
    object app = null;
    EnumWindows((h, p) => {
      uint wp; GetWindowThreadProcessId(h, out wp);
      if (wp != pid || Cls(h) != "XLMAIN") return true;
      EnumChildWindows(h, (k, q) => {
        if (Cls(k) != "EXCEL7") return true;
        var iid = new Guid("00020400-0000-0000-C000-000000000046"); object win;
        if (AccessibleObjectFromWindow(k, 0xFFFFFFF0, ref iid, out win) == 0 && win != null) {
          app = win.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, win, null); return false;
        }
        return true;
      }, IntPtr.Zero);
      return app == null;
    }, IntPtr.Zero);
    return app;
  }

  static bool DialogClass(string c) {
    return c == "#32770" || c == "bosa_sdm_XL9" || c == "NUIDialog" || c.StartsWith("WindowsForms") || c.StartsWith("HwndWrapper");
  }
  static string Clean(string s) { return (s ?? "").Replace("&", "").Trim(); }

  /** Visible dialog windows of the process, described through UI Automation. */
  public static AfoDialog[] Dialogs(uint pid) {
    var handles = new List<IntPtr>();
    EnumWindows((h, p) => {
      uint wp; GetWindowThreadProcessId(h, out wp);
      if (wp != pid || !IsWindowVisible(h) || !DialogClass(Cls(h))) return true;
      RECT rc; if (!GetWindowRect(h, out rc) || rc.R - rc.L < 40 || rc.B - rc.T < 40) return true;
      handles.Add(h); return true;
    }, IntPtr.Zero);
    var r = new List<AfoDialog>();
    foreach (var h in handles) {
      var d = new AfoDialog(); d.title = Title(h);
      var texts = new List<string>(); var buttons = new List<string>(); bool progress = false; string uiaClass = "";
      try {
        var e = AutomationElement.FromHandle(h); uiaClass = e.Current.ClassName ?? "";
        int n = 0;
        foreach (AutomationElement c in e.FindAll(TreeScope.Descendants, Condition.TrueCondition)) {
          if (++n > 500) break;
          var ct = c.Current.ControlType; var name = Clean(c.Current.Name);
          if (ct == ControlType.ProgressBar) progress = true;
          else if (ct == ControlType.Button) { if (name.Length > 0) buttons.Add(name); }
          else if ((ct == ControlType.Text || ct == ControlType.Document) && name.Length > 0 && !texts.Contains(name) && name != d.title) texts.Add(name);
        }
      } catch {}
      var text = string.Join(" | ", texts.ToArray()); if (text.Length > 3000) text = text.Substring(0, 3000) + " ...";
      d.text = text; d.buttons = buttons.ToArray();
      var probe = (d.title + " " + uiaClass).ToLowerInvariant();
      if (progress) continue;
      if (probe.Contains("prompt") || probe.Contains("variable") || probe.Contains("eingabeaufforderung")) d.kind = "prompts";
      else if (probe.Contains("logon") || probe.Contains("log on") || probe.Contains("anmeld")) d.kind = "logon";
      else d.kind = "message";
      r.Add(d);
    }
    return r.ToArray();
  }
}
`;

/** PowerShell functions every operate script starts with. $a holds the parsed arguments. */
export const PS_PREAMBLE = String.raw`
param([string]$ArgsFile)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$a = Get-Content -Raw -Encoding UTF8 $ArgsFile | ConvertFrom-Json
Add-Type -ReferencedAssemblies UIAutomationClient, UIAutomationTypes, WindowsBase @'
${CSHARP}
'@
$out = [ordered]@{ steps = @() }
# Progress lines for the server: it looks for dialogs when they stop coming.
function Emit([string]$line) { [Console]::Out.WriteLine("@@$line"); [Console]::Out.Flush() }
function Step($m) { $script:out.steps += $m; Emit "step $m" }
function Busy($e) { $e.Exception.Message -match '0x8001010A|0x80010001|RPC_E_CALL_REJECTED|RPC_E_SERVERCALL_RETRYLATER' }
function Retry($f, [int]$seconds = 120) {
  $end = (Get-Date).AddSeconds($seconds)
  while ($true) { try { return & $f } catch { if ((Busy $_) -and (Get-Date) -lt $end) { Start-Sleep -Milliseconds 1500 } else { throw } } }
}
function Table($v) {
  if ($null -eq $v) { return ,@() }
  if ($v -is [array] -and $v.Rank -eq 2) {
    $rows = @(); for ($i = $v.GetLowerBound(0); $i -le $v.GetUpperBound(0); $i++) {
      $row = @(); for ($j = $v.GetLowerBound(1); $j -le $v.GetUpperBound(1); $j++) { $row += $v[$i, $j] }; $rows += ,$row }
    return ,$rows
  }
  return ,@(,@($v))
}
# Excel error values reach COM as these numbers.
$script:xlErrors = @{ -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'; -2146826265 = '#REF!'; -2146826259 = '#NAME?'; -2146826252 = '#NUM!'; -2146826246 = '#N/A' }
function Txt($v) {
  if ($null -eq $v) { return '' }
  if ($v -is [int] -and $script:xlErrors.ContainsKey($v)) { return $script:xlErrors[$v] }
  "$v"
}
# Rows of an AfO list as text, one array per row on the output; rows of only blanks or error values are dropped.
function Rows($v) {
  foreach ($row in (Table $v)) {
    $cells = @($row | ForEach-Object { Txt $_ })
    if (@($cells | Where-Object { $_ -and $_ -notmatch '^#' }).Count) { ,$cells }
  }
}
function Messages { try { Rows (Retry { $script:xl.Run('SAPListOfMessages', 'INFORMATION', 'TRUE') }) } catch { } }
function LastError { try { Txt (Retry { $script:xl.Run('SAPGetProperty', 'LastError', 'Text') }) } catch { '' } }
# 1 means success for the AfO API; anything else fails with AfO's last error.
function Check($result, $what) { if ((Txt $result) -ne '1') { $e = LastError; throw "$what failed$(if ($e) { ": $e" } else { " (result $(Txt $result))" })" } }
`;

const PS_EPILOGUE = String.raw`
$out | ConvertTo-Json -Depth 8 -Compress
`;

const DIALOGS_SCRIPT = String.raw`
$out.dialogs = @([AfoXl]::Dialogs([uint32]$a.pid))
`;

export interface AfoDialog {
  kind: 'prompts' | 'logon' | 'message';
  title: string;
  text: string;
  buttons: string[];
}

export interface RunOptions {
  env?: NodeJS.ProcessEnv;
  timeoutSeconds?: number;
  /** Command prefix the PowerShell child runs through, e.g. a secret runner. */
  prefix?: string[];
}

const HINTS: Record<AfoDialog['kind'], string> = {
  prompts: ' A mandatory variable has no value: give it in prompts, or store prompt values in the workbook (afo_update_workbook).',
  logon: ' Analysis for Office asks for a logon: log on in that dialog, or configure a logon for this system.',
  message: '',
};

/** One line for a dialog: kind, title, text. */
export function describeDialog(d: AfoDialog): string {
  return `${d.kind} dialog "${d.title}"${d.text ? `: ${d.text}` : ''}`;
}

function scriptFiles(body: string, args: unknown): { dir: string; script: string; argsFile: string } {
  const dir = mkdtempSync(join(tmpdir(), 'afo-operate-'));
  const script = join(dir, 'operate.ps1');
  const argsFile = join(dir, 'args.json');
  // PowerShell 5.1 reads a script without BOM as the ANSI code page.
  writeFileSync(script, '﻿' + PS_PREAMBLE + body + PS_EPILOGUE, 'utf8');
  writeFileSync(argsFile, JSON.stringify(args), 'utf8');
  return { dir, script, argsFile };
}

const psCommand = (script: string, argsFile: string) =>
  ['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-ArgsFile', argsFile];

/** The dialogs an Excel process shows right now. */
export async function excelDialogs(pid: number): Promise<AfoDialog[]> {
  const { dir, script, argsFile } = scriptFiles(DIALOGS_SCRIPT, { pid });
  try {
    const cmd = psCommand(script, argsFile);
    const stdout = await new Promise<string>((ok, fail) => {
      execFile(cmd[0], cmd.slice(1), { timeout: 60_000, windowsHide: true }, (err, out) => (err && !out ? fail(err) : ok(out)));
    });
    const json = stdout.trim().split(/\r?\n/).filter((l) => !l.startsWith('@@')).pop() ?? '{}';
    return ((JSON.parse(json) as { dialogs?: AfoDialog[] }).dialogs ?? []);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/**
 * Runs PS_PREAMBLE + body with the arguments in a JSON file and returns the parsed result. The
 * body reports the Excel process with Emit "pid <id>" and its progress with Step. When no step
 * follows for STALL_SECONDS and Excel shows a dialog, the script is stopped and the result says
 * which dialog Excel waits on; Excel itself is not touched.
 */
export async function runOperateScript(body: string, args: unknown, options: RunOptions = {}): Promise<Record<string, unknown>> {
  if (process.platform !== 'win32') throw new Error('The operate tools need Windows with Excel and Analysis for Office.');
  const { dir, script, argsFile } = scriptFiles(body, args);
  try {
    const cmd = [...(options.prefix ?? []), ...psCommand(script, argsFile)];
    return await new Promise<Record<string, unknown>>((ok, fail) => {
      const child = spawn(cmd[0], cmd.slice(1), { env: options.env ?? process.env, windowsHide: true });
      let stdout = '', pending = '', stderr = '', excelPid = 0, lastProgress = Date.now(), checking = false, done = false;
      const steps: string[] = [];
      const stopChild = () => {
        // The whole tree: with a secret runner, PowerShell is a grandchild. Excel was started
        // through WMI and is not part of it.
        if (child.pid) execFile('taskkill', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true }, () => undefined);
      };
      const finish = (result: Record<string, unknown> | Error) => {
        if (done) return;
        done = true;
        clearInterval(watch);
        clearTimeout(timer);
        if (result instanceof Error) fail(result); else ok(result);
      };
      child.stdout.setEncoding('utf8');
      child.stdout.on('data', (chunk: string) => {
        stdout += chunk;
        const lines = (pending + chunk).split(/\r?\n/);
        pending = lines.pop() ?? '';
        for (const line of lines) {
          if (!line.startsWith('@@')) continue;
          lastProgress = Date.now();
          if (line.startsWith('@@pid ')) excelPid = Number(line.slice(6));
          else if (line.startsWith('@@step ')) steps.push(line.slice(7));
        }
      });
      child.stderr.setEncoding('utf8');
      child.stderr.on('data', (chunk: string) => { stderr += chunk; });
      const watch = setInterval(async () => {
        if (checking || !excelPid || Date.now() - lastProgress < STALL_SECONDS * 1000) return;
        checking = true;
        try {
          const dialogs = await excelDialogs(excelPid);
          if (dialogs.length && !done) {
            stopChild();
            finish({
              ok: false, excelPid, steps, dialogs,
              error: `Excel waits on a ${describeDialog(dialogs[0])}.${HINTS[dialogs[0].kind]} The dialog is still open in Excel for the user; the steps after it did not run.`,
            });
          }
        } catch { /* no verdict: keep waiting */ }
        lastProgress = Date.now();
        checking = false;
      }, 2000);
      const timer = setTimeout(() => {
        stopChild();
        finish({ ok: false, excelPid: excelPid || undefined, steps, error: `Timed out after ${options.timeoutSeconds ?? 300} s without a dialog in Excel.` });
      }, (options.timeoutSeconds ?? 300) * 1000);
      child.on('error', (err) => finish(err));
      child.on('close', () => {
        const json = stdout.trim().split(/\r?\n/).filter((l) => !l.startsWith('@@')).pop() ?? '';
        try {
          finish(JSON.parse(json) as Record<string, unknown>);
        } catch (err) {
          finish(new Error(`Unexpected output from the Excel script (${(err as Error).message}): ${(stderr || stdout).slice(-800)}`));
        }
      });
    });
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
