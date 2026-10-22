// Ad-hoc report: a plain .xlsx with one query in a crosstab, prompt values stored, connected the
// way SAP Logon connects, then opened logged on and refreshed. No macros, so Excel shows no
// security bar and the logon runs without the user.

import { existsSync, mkdirSync } from 'node:fs';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import type { PromptValue } from '../build/afo.js';
import { buildWorkbook } from '../build/build.js';
import { landscapeSystem } from './landscape.js';
import { logonConfig, openWorkbook } from './open.js';

export interface QuickReportOptions {
  query: string;
  /** System id; default the one the logon is configured for. */
  system?: string;
  client?: string;
  language?: string;
  prompts?: PromptValue[];
  axes?: { rows?: string[]; columns?: string[]; free?: string[] };
  /** Folder for the workbook (default Documents\AfO Reports). */
  folder?: string;
  /** Open in Excel, logged on (default true). */
  open?: boolean;
  timeoutSeconds?: number;
}

const pad = (n: number) => String(n).padStart(2, '0');
const stamp = (d = new Date()) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;

export async function quickReport(o: QuickReportOptions, env: NodeJS.ProcessEnv = process.env): Promise<unknown> {
  const logon = logonConfig(env);
  const systemId = (o.system ?? logon?.system ?? '').toUpperCase();
  if (!systemId) throw new Error('No system: give system, or configure the logon (AFO_LOGON_SYSTEM).');
  const client = o.client ?? (logon?.system === systemId ? logon.client : undefined);
  if (!client) throw new Error(`No client for ${systemId}: give client.`);
  const language = o.language ?? (logon?.system === systemId ? logon.language : undefined) ?? 'DE';
  const system = landscapeSystem(systemId, client, language);

  const folder = resolve(o.folder ?? join(homedir(), 'Documents', 'AfO Reports'));
  if (!existsSync(folder)) mkdirSync(folder, { recursive: true });
  const output = join(folder, `${o.query.replace(/[\\/:*?"<>|]+/g, '_')} ${systemId} ${stamp()}.xlsx`);
  const prompts = o.prompts ?? [];

  const built = buildWorkbook({
    system,
    storePrompts: true,
    refreshOnOpen: false,
    dataProviders: [{ key: 'DS_1', query: o.query, prompts: prompts.length ? prompts : undefined, axes: o.axes }],
    grids: [{ key: 'Crosstab1', dataProviderKey: 'DS_1', sheet: 'Report', cell: 'B4' }],
  }, {
    output,
    layout: {
      sheets: ['Report'],
      cells: [
        { sheet: 'Report', cell: 'B1', value: o.query },
        { sheet: 'Report', cell: 'B2', value: `${systemId}/${client} · ${prompts.map((p) => `${p.name}=${p.member ?? p.literal ?? `${p.low}..${p.high}`}`).join(', ') || 'no prompt values'}` },
      ],
    },
  });

  if (o.open === false) return { built };
  const opened = await openWorkbook({ path: output, logonWaitSeconds: 15, timeoutSeconds: o.timeoutSeconds }, env);
  return { file: output, connection: { system: system.systemName, server: system.server, sapRouter: system.sapRouter, snc: !!system.sncPartnerName }, warnings: built.warnings, opened };
}
