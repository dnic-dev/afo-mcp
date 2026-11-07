// Ad-hoc report: a workbook with one crosstab per query, prompt values stored, connected the way
// SAP Logon connects, then opened logged on and refreshed. Without a template it is a plain .xlsx
// without macros, so Excel shows no security bar and the logon runs without the user.

import { existsSync, mkdirSync } from 'node:fs';
import { homedir } from 'node:os';
import { extname, join, resolve } from 'node:path';
import type { DataProviderSpec, FilterSpec, GridSpec, PromptValue } from '../build/afo.js';
import { type CellSpec, buildWorkbook } from '../build/build.js';
import type { CompactRowsSpec, HierarchySpec } from '../build/hierarchy.js';
import { readWorkbook } from '../workbook/read.js';
import { landscapeSystem } from './landscape.js';
import { logonConfig, openWorkbook } from './open.js';

/** One query of the report, shown in its own crosstab. */
export interface QuickReportSheet {
  query: string;
  /** Sheet of the crosstab: a sheet of the template, or a new sheet (default: the query name). */
  sheet?: string;
  /** Top-left cell of the crosstab (default B4). */
  cell?: string;
  prompts?: PromptValue[];
  axes?: { rows?: string[]; columns?: string[]; free?: string[] };
  filters?: FilterSpec[];
  /** BW hierarchies shown on open, with the levels expanded. */
  hierarchies?: HierarchySpec[];
  /** Compact display of rows (needs axes.rows). */
  compactRows?: CompactRowsSpec;
}

export interface QuickReportOptions extends Partial<QuickReportSheet> {
  /** Further queries, each in its own crosstab; with query given, they follow it. */
  reports?: QuickReportSheet[];
  /**
   * Layout workbook (.xlsx/.xlsm) the crosstabs are placed into: company or customer design
   * lives there, not in the server. Without it a plain layout is generated.
   */
  template?: string;
  /** System id; default the one the logon is configured for. */
  system?: string;
  client?: string;
  language?: string;
  /** Folder for the workbook (default Documents\AfO Reports). */
  folder?: string;
  /** Open in Excel, logged on (default true). */
  open?: boolean;
  timeoutSeconds?: number;
}

const pad = (n: number) => String(n).padStart(2, '0');
const stamp = (d = new Date()) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
const fileSafe = (s: string) => s.replace(/[\\/:*?"<>|]+/g, '_');
/** A valid, unused sheet name from a query name: no []:*?/\, at most 31 characters. */
const sheetFor = (query: string, used: Set<string>) => {
  const base = query.replace(/[[\]:*?/\\]+/g, '_').slice(0, 31) || 'Report';
  let name = base;
  for (let i = 2; used.has(name.toLowerCase()); i++) name = `${base.slice(0, 31 - String(i).length - 1)}_${i}`;
  used.add(name.toLowerCase());
  return name;
};

/** The queries of the call: query (with its fields) first, then reports. */
export function quickReportSheets(o: QuickReportOptions): QuickReportSheet[] {
  const sheets: QuickReportSheet[] = [];
  if (o.query) sheets.push({ query: o.query, sheet: o.sheet, cell: o.cell, prompts: o.prompts, axes: o.axes, filters: o.filters, hierarchies: o.hierarchies, compactRows: o.compactRows });
  sheets.push(...(o.reports ?? []));
  if (!sheets.length) throw new Error('No query: give query or reports.');
  sheets.forEach((s, i) => {
    if (typeof s?.query !== 'string' || !s.query.trim()) throw new Error(`reports[${i - (o.query ? 1 : 0)}]: query is required.`);
  });
  return sheets;
}

export async function quickReport(o: QuickReportOptions, env: NodeJS.ProcessEnv = process.env): Promise<unknown> {
  const reports = quickReportSheets(o);
  const systemId = (o.system ?? logonConfig(undefined, env)?.system ?? '').toUpperCase();
  if (!systemId) throw new Error('No system: give system (several logons are configured, or none).');
  const logon = logonConfig(systemId, env);
  const client = o.client ?? (logon?.system === systemId ? logon.client : undefined);
  if (!client) throw new Error(`No client for ${systemId}: give client.`);
  const language = o.language ?? (logon?.system === systemId ? logon.language : undefined) ?? 'DE';
  const system = landscapeSystem(systemId, client, language);

  const template = o.template ? resolve(o.template) : undefined;
  if (template && !existsSync(template)) throw new Error(`The template does not exist: ${template}`);
  const templateSheets = template
    ? (readWorkbook(template, { sections: ['sheets'] }).sheets ?? []).filter((s) => s.state === 'visible' || !s.state).map((s) => s.name)
    : [];

  // Sheets: as given; else, on a template, its visible sheets in order; else one per query.
  const used = new Set<string>();
  const placed = reports.map((r, i) => {
    let sheet = r.sheet;
    if (!sheet && template) {
      sheet = templateSheets[i];
      if (!sheet) throw new Error(`The template has ${templateSheets.length} visible sheet(s) for ${reports.length} queries: give sheet for ${r.query}.`);
    }
    if (!sheet) sheet = reports.length === 1 ? 'Report' : sheetFor(r.query, used);
    return { ...r, sheet, cell: r.cell ?? 'B4' };
  });

  const folder = resolve(o.folder ?? join(homedir(), 'Documents', 'AfO Reports'));
  if (!existsSync(folder)) mkdirSync(folder, { recursive: true });
  const ext = template && extname(template).toLowerCase() === '.xlsm' ? '.xlsm' : '.xlsx';
  const title = reports.length === 1 ? reports[0].query : `${reports[0].query} +${reports.length - 1}`;
  // Two reports in the same minute get a counter instead of failing on the existing file.
  const base = join(folder, `${fileSafe(title)} ${systemId} ${stamp()}`);
  let output = `${base}${ext}`;
  for (let i = 2; existsSync(output); i++) output = `${base} (${i})${ext}`;

  const dataProviders: DataProviderSpec[] = placed.map((r, i) => ({
    key: `DS_${i + 1}`,
    query: r.query,
    prompts: r.prompts?.length ? r.prompts : undefined,
    axes: r.axes,
    filters: r.filters?.length ? r.filters : undefined,
    hierarchies: r.hierarchies?.length ? r.hierarchies : undefined,
    compactRows: r.compactRows,
  }));
  const grids: GridSpec[] = placed.map((r, i) => ({ key: `Crosstab${i + 1}`, dataProviderKey: `DS_${i + 1}`, sheet: r.sheet, cell: r.cell }));
  // The generated layout gets a title and the selection above each crosstab; a template keeps its own.
  const cells: CellSpec[] = template ? [] : placed.flatMap((r) => [
    { sheet: r.sheet, cell: 'B1', value: r.query },
    { sheet: r.sheet, cell: 'B2', value: `${systemId}/${client} · ${(r.prompts ?? []).map((p) => `${p.name}=${p.member ?? p.literal ?? `${p.low}..${p.high}`}`).join(', ') || 'no prompt values'}` },
  ]);

  const built = buildWorkbook(
    { system, storePrompts: true, refreshOnOpen: false, dataProviders, grids },
    template
      ? { output, template }
      : { output, layout: { sheets: [...new Set(placed.map((r) => r.sheet))], cells } },
  );

  if (o.open === false) return { built };
  // A template with macros opens with the security bar; the logon waits for the user to enable content.
  const opened = await openWorkbook({ path: output, logonWaitSeconds: ext === '.xlsm' ? 90 : 15, timeoutSeconds: o.timeoutSeconds }, env);
  return {
    file: output,
    crosstabs: built.crosstabs,
    connection: { system: system.systemName, server: system.server, sapRouter: system.sapRouter, snc: !!system.sncPartnerName },
    warnings: built.warnings,
    opened,
  };
}
