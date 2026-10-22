// Open an AfO workbook in Excel on this machine, log on, set prompt values, refresh and
// optionally save, so the user sees the data at once. Needs Windows, Excel and Analysis for
// Office. Excel is started through WMI: it then runs outside this server's process tree and
// stays open for the user when the server stops.
//
// Logon data comes from the environment of this server, never from tool arguments:
//   AFO_LOGON_SYSTEMS   JSON list, one entry per system:
//                       [{ "system": "SYS", "client": "100", "user": "USER", "language": "EN",
//                          "vault": "entry-name", "passwordVar": "SAP_PASSWORD" }]
//   AFO_SECRET_COMMAND  JSON array: the command prefix of a secret runner that starts a command
//                       with the secrets of an entry in its environment, "{vault}" standing for
//                       the entry, e.g. ["op", "run", "--env-file", "{vault}.env", "--"]. For an
//                       entry with "vault", the logon child runs through it, so this server never
//                       holds the password. Without "vault", the password is read from this
//                       server's environment (variable passwordVar).
//   AFO_LOGON_SYSTEM, AFO_LOGON_CLIENT, AFO_LOGON_USER, AFO_LOGON_LANGUAGE, AFO_LOGON_PASSWORD_ENV:
//                       the same for a single system.
// The password reaches the PowerShell child through its environment only.

import { execFile } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { readSourceDefinition } from '../build/source.js';

export interface OpenPrompt {
  name: string;
  value: string;
  /** SAPSetVariable format: INTERNAL_KEY (default) or INPUT_STRING. */
  format?: 'INTERNAL_KEY' | 'INPUT_STRING';
  /** Data source the variable belongs to (default: the first one). */
  dataProvider?: string;
}

export interface OpenOptions {
  path: string;
  prompts?: OpenPrompt[];
  /** Log on with the configured credentials (default true). */
  logon?: boolean;
  /** Save the workbook after the refresh, with its data when the workbook keeps data on save. */
  save?: boolean;
  timeoutSeconds?: number;
  /** How long to retry the logon while the workbook's content is still disabled (default 90). */
  logonWaitSeconds?: number;
}

export interface LogonConfig {
  system: string;
  client: string;
  user: string;
  language?: string;
  /** Vault entry whose secrets the vault starter hands to the logon child. */
  vault?: string;
  /** Variable that holds the password (in the vault entry, or in this server's environment). */
  passwordVar: string;
}

/** Configured logons, one per system id. */
export function logonConfigs(env: NodeJS.ProcessEnv = process.env): LogonConfig[] {
  const list: LogonConfig[] = [];
  if (env.AFO_LOGON_SYSTEMS) {
    let parsed: unknown;
    try { parsed = JSON.parse(env.AFO_LOGON_SYSTEMS); } catch { throw new Error('AFO_LOGON_SYSTEMS is not valid JSON.'); }
    for (const e of Array.isArray(parsed) ? (parsed as Array<Record<string, string>>) : []) {
      if (!e.system || !e.client || !e.user) continue;
      list.push({ system: e.system.toUpperCase(), client: e.client, user: e.user, language: e.language || undefined, vault: e.vault || undefined, passwordVar: e.passwordVar || 'BW_PASSWORD' });
    }
  }
  const { AFO_LOGON_SYSTEM: system, AFO_LOGON_CLIENT: client, AFO_LOGON_USER: user } = env;
  const passwordVar = env.AFO_LOGON_PASSWORD_ENV || 'AFO_LOGON_PASSWORD';
  if (system && client && user && !list.some((l) => l.system === system.toUpperCase())) {
    list.push({ system: system.toUpperCase(), client, user, language: env.AFO_LOGON_LANGUAGE || undefined, passwordVar });
  }
  // An entry without a vault needs its password in this environment.
  return list.filter((l) => l.vault || env[l.passwordVar]);
}

/** The secret runner prefix for a vault entry, from AFO_SECRET_COMMAND. */
export function secretCommand(env: NodeJS.ProcessEnv, vault: string): string[] {
  let prefix: unknown;
  try { prefix = JSON.parse(env.AFO_SECRET_COMMAND ?? ''); } catch { prefix = undefined; }
  if (!Array.isArray(prefix) || !prefix.length || !prefix.every((p) => typeof p === 'string')) {
    throw new Error(`The logon uses the secret entry "${vault}", but AFO_SECRET_COMMAND is not a JSON array of strings.`);
  }
  return (prefix as string[]).map((p) => p.replaceAll('{vault}', vault));
}

/** The logon for one system; without a system, the only one configured. */
export function logonConfig(system?: string, env: NodeJS.ProcessEnv = process.env): LogonConfig | undefined {
  const all = logonConfigs(env);
  return system ? all.find((l) => l.system === system.toUpperCase()) : all.length === 1 ? all[0] : undefined;
}

/** Configured logons without passwords, for status output. */
export function logonStatus(env: NodeJS.ProcessEnv = process.env): string {
  const all = logonConfigs(env);
  return all.length ? `logon configured for ${all.map((c) => `${c.system}/${c.client}`).join(', ')}` : 'no logon configured: the user logs on in Excel';
}

const PS_SCRIPT = String.raw`
param([string]$ArgsFile)
$ErrorActionPreference = 'Stop'
$a = Get-Content -Raw -Encoding UTF8 $ArgsFile | ConvertFrom-Json
$secret = $null
if ($a.passwordVar) { $secret = [Environment]::GetEnvironmentVariable($a.passwordVar); [Environment]::SetEnvironmentVariable($a.passwordVar, $null) }
Add-Type @'
using System; using System.Text; using System.Runtime.InteropServices;
public static class AfoXl {
  public delegate bool EnumProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f, IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc f, IntPtr p);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object obj);
  static string Cls(IntPtr h) { var c = new StringBuilder(256); GetClassName(h, c, 256); return c.ToString(); }
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
  public static string[] Dialogs(uint pid) {
    var r = new System.Collections.Generic.List<string>();
    EnumWindows((h, p) => {
      uint wp; GetWindowThreadProcessId(h, out wp);
      if (wp != pid || !IsWindowVisible(h)) return true;
      var cls = Cls(h); if (cls != "#32770" && !cls.StartsWith("WindowsForms")) return true;
      var t = new StringBuilder(512); GetWindowText(h, t, 512); r.Add(t.ToString()); return true;
    }, IntPtr.Zero);
    return r.ToArray();
  }
}
'@
$out = [ordered]@{ steps = @(); dialogs = @() }
function Step($m) { $script:out.steps += $m }
function Busy($e) { $e.Exception.Message -match '0x8001010A|0x80010001|RPC_E_CALL_REJECTED|RPC_E_SERVERCALL_RETRYLATER' }
function Retry($f, [int]$seconds = 120) {
  $end = (Get-Date).AddSeconds($seconds)
  while ($true) { try { return & $f } catch { if ((Busy $_) -and (Get-Date) -lt $end) { Start-Sleep -Milliseconds 1500 } else { throw } } }
}
$procId = 0
try {
  $exe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe').'(default)'
  $r = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = ('"' + $exe + '" /x "' + $a.path + '"') }
  if ($r.ReturnValue -ne 0) { throw "Excel start failed (WMI $($r.ReturnValue))" }
  $procId = [uint32]$r.ProcessId; $out.excelPid = $procId
  $xl = $null; $end = (Get-Date).AddSeconds(120)
  while (-not $xl -and (Get-Date) -lt $end) { Start-Sleep -Seconds 1; $xl = [AfoXl]::Application($procId) }
  if (-not $xl) { throw 'Excel window not found' }
  Step 'excel started'
  $ds = $a.dataProviders[0]
  $end = (Get-Date).AddSeconds(180); $api = $false
  while (-not $api -and (Get-Date) -lt $end) { try { [void]$xl.Run('SAPGetProperty', 'IsConnected', $ds); $api = $true } catch { Start-Sleep -Seconds 2 } }
  if (-not $api) { throw 'Analysis for Office API not available (add-in not loaded, or a dialog blocks Excel)' }
  $wb = Retry { $xl.Workbooks.Item(1) }
  Step 'afo ready'
  if ($secret) {
    # AfO refuses SAPLogon while the workbook's content is disabled (security bar): keep trying
    # until the user clicks "Enable Content" or the wait is over.
    $end = (Get-Date).AddSeconds($a.logonWaitSeconds); $res = 0
    while ($true) {
      $res = Retry { $xl.Run('SAPLogon', $ds, $a.client, $a.user, $secret) }
      if ($res -eq 1 -or (Get-Date) -ge $end) { break }
      Start-Sleep -Seconds 3
    }
    $secret = $null
    $out.logon = $res
    if ($res -ne 1) { throw "SAPLogon returned $res. If Excel shows the security bar, the workbook's content is disabled and AfO refuses the logon: click Enable Content (or open the workbook from a trusted location) and call again." }
    Step 'logged on'
  }
  # The first refresh initialises the data sources; variables can only be set after it.
  [void](Retry { $xl.Run('SAPExecuteCommand', 'Refresh') } 600)
  Step 'refreshed'
  if ($a.prompts.Count) {
    [void](Retry { $xl.Run('SAPSetRefreshBehaviour', 'Off') })
    foreach ($p in $a.prompts) {
      $res = Retry { $xl.Run('SAPSetVariable', $p.name, $p.value, $p.format, $p.dataProvider) }
      Step "set $($p.name) = $($p.value): $res"
    }
    [void](Retry { $xl.Run('SAPSetRefreshBehaviour', 'On') } 600)
    Step 'refreshed with prompts'
  }
  $out.dataProviders = @(foreach ($d in $a.dataProviders) { [ordered]@{ key = $d; connected = (Retry { $xl.Run('SAPGetProperty', 'IsConnected', $d) }) } })
  $out.variables = @(foreach ($p in $a.prompts) { [ordered]@{ name = $p.name; value = "$(Retry { $xl.Run('SAPGetVariable', $p.dataProvider, $p.name, 'VALUEASTEXT') })" } })
  $out.sheets = @(foreach ($ws in @($wb.Worksheets)) { if ($ws.Visible -eq -1) { [ordered]@{ name = $ws.Name; used = $ws.UsedRange.Address(0, 0) } } })
  $msgs = Retry { $xl.Run('SAPListOfMessages', '', 'False') }
  if ($msgs -is [array]) { $out.messages = @($msgs | ForEach-Object { "$_" } | Where-Object { $_ }) } elseif ($msgs -and "$msgs" -notmatch '^-?\d+$') { $out.messages = @("$msgs") }
  if ($a.save) { Retry { $wb.Save() } | Out-Null; Step 'saved' }
  $out.ok = $true
} catch {
  $out.ok = $false; $out.error = $_.Exception.Message
} finally {
  $secret = $null
  if ($procId) { $out.dialogs = @([AfoXl]::Dialogs($procId)) }
}
$out | ConvertTo-Json -Depth 5 -Compress
`;

/** Data source keys and system id of a workbook, from its AfO definition. */
function workbookFacts(path: string): { dataProviders: string[]; systemId?: string } {
  const xml = readSourceDefinition(path);
  return {
    dataProviders: [...xml.matchAll(/<DataProvider><Key>([^<]+)<\/Key>/g)].map((m) => m[1]),
    systemId: /<SystemId>([^<]+)<\/SystemId>/.exec(xml)?.[1],
  };
}

export async function openWorkbook(options: OpenOptions, env: NodeJS.ProcessEnv = process.env): Promise<unknown> {
  if (process.platform !== 'win32') throw new Error('afo_open_workbook needs Windows with Excel and Analysis for Office.');
  const path = resolve(options.path);
  if (!existsSync(path)) throw new Error(`Workbook not found: ${path}`);
  const facts = workbookFacts(path);
  if (!facts.dataProviders.length) throw new Error('The workbook has no AfO data source.');
  const notes: string[] = [];
  let logon: LogonConfig | undefined;
  if (options.logon !== false) {
    logon = facts.systemId ? logonConfig(facts.systemId, env) : undefined;
    if (!logon) notes.push(`No logon configured for ${facts.systemId ?? 'the workbook system'}: the user logs on in Excel.`);
  }
  const prompts = (options.prompts ?? []).map((p) => ({
    name: p.name, value: p.value, format: p.format ?? 'INTERNAL_KEY', dataProvider: p.dataProvider ?? facts.dataProviders[0],
  }));
  const bad = prompts.filter((p) => !facts.dataProviders.includes(p.dataProvider));
  if (bad.length) throw new Error(`Unknown data source in prompts: ${bad.map((p) => p.dataProvider).join(', ')}. The workbook has ${facts.dataProviders.join(', ')}.`);

  const dir = mkdtempSync(join(tmpdir(), 'afo-open-'));
  try {
    const script = join(dir, 'open.ps1');
    const argsFile = join(dir, 'args.json');
    writeFileSync(script, '﻿' + PS_SCRIPT, 'utf8');
    writeFileSync(argsFile, JSON.stringify({
      path, prompts, save: options.save === true, logonWaitSeconds: options.logonWaitSeconds ?? 90, dataProviders: facts.dataProviders,
      client: logon?.client, user: logon?.user, language: logon?.language, passwordVar: logon?.passwordVar,
    }), 'utf8');
    const ps = ['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-ArgsFile', argsFile];
    let cmd = ps;
    if (logon?.vault) {
      cmd = [...secretCommand(env, logon.vault), ...ps];
    }
    const stdout = await new Promise<string>((ok, fail) => {
      execFile(cmd[0], cmd.slice(1),
        { env, timeout: (options.timeoutSeconds ?? 900) * 1000, windowsHide: true, maxBuffer: 4 * 1024 * 1024 },
        (err, out, errOut) => (err && !out ? fail(new Error(`${err.message}${errOut ? `: ${errOut.slice(0, 500)}` : ''}`)) : ok(out)));
    });
    const json = stdout.trim().split(/\r?\n/).pop() ?? '{}';
    const result = JSON.parse(json) as Record<string, unknown>;
    return { file: path, systemId: facts.systemId, loggedOnAs: logon ? `${logon.user}@${logon.system}/${logon.client}` : undefined, ...result, notes };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
