// Inventory across a folder of workbooks: a short profile per workbook, which data sources,
// planning objects and variables are used where, recurring layouts, and outliers.

import { readdirSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { type WorkbookDescription, readWorkbook } from './read.js';
import { literal } from './sheets.js';

const WORKBOOK_RE = /\.(xlsx|xlsm|xltx|xltm)$/i;

export interface FindOptions {
  recursive?: boolean;
  /** Stop after this many files. */
  limit?: number;
}

export function findWorkbooks(folder: string, opts: FindOptions = {}): { files: string[]; truncated: boolean } {
  const files: string[] = [];
  const limit = opts.limit ?? 2000;
  let truncated = false;
  const walk = (dir: string) => {
    for (const entry of readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
      if (files.length >= limit) {
        truncated = true;
        return;
      }
      const full = join(dir, entry.name);
      // Excel lock files (~$Book.xlsm) are not workbooks.
      if (entry.isFile() && WORKBOOK_RE.test(entry.name) && !entry.name.startsWith('~$')) files.push(full);
      else if (entry.isDirectory() && opts.recursive) walk(full);
    }
  };
  if (!statSync(folder).isDirectory()) throw new Error(`${folder} is not a folder`);
  walk(folder);
  return { files, truncated };
}

export interface WorkbookProfile {
  file: string;
  afo: boolean;
  systemId?: string;
  afoRevision?: string;
  lastRefreshed?: string;
  dataSources: string[];
  planningObjects: string[];
  sheets: number;
  grids: number;
  sapFormulaCells: number;
  macros: string;
  issues: string[];
  error?: string;
}

interface Usage {
  name: string;
  workbooks: string[];
}

function addUsage(map: Map<string, Set<string>>, name: string | undefined, file: string) {
  if (!name) return;
  if (!map.has(name)) map.set(name, new Set());
  map.get(name)!.add(file);
}

function usageList(map: Map<string, Set<string>>, info?: Map<string, string>): Array<Usage & { detail?: string }> {
  return [...map]
    .map(([name, files]) => ({ name, ...(info?.get(name) ? { detail: info.get(name) } : {}), workbooks: [...files].sort() }))
    .sort((a, b) => b.workbooks.length - a.workbooks.length || a.name.localeCompare(b.name));
}

function majority(values: Array<string | undefined>): string | undefined {
  const counts = new Map<string, number>();
  for (const v of values) if (v) counts.set(v, (counts.get(v) ?? 0) + 1);
  return [...counts].sort((a, b) => b[1] - a[1])[0]?.[0];
}

function countBy(values: Array<string | undefined>): Record<string, number> {
  const out: Record<string, number> = {};
  for (const v of values) if (v) out[v] = (out[v] ?? 0) + 1;
  return out;
}

/** Variables a workbook names: prompts, planning object variables, references and SAPGetVariable formulas. */
function variablesOf(d: WorkbookDescription): Set<string> {
  const vars = new Set<string>();
  for (const v of d.settings?.promptVariables ?? []) vars.add(v.name);
  for (const p of d.planningObjects ?? []) {
    for (const v of p.variables) vars.add(v.name);
    for (const r of p.references) if (r.object === 'Variable') vars.add(r.technicalName);
  }
  for (const dp of d.dataProviders ?? []) for (const v of dp.state?.variables ?? []) vars.add(v.name);
  for (const s of d.formulas ?? []) {
    for (const f of s.formulas) for (const c of f.calls) if (c.function === 'SAPGetVariable') {
      const v = literal(c.args[1]);
      if (v) vars.add(v);
    }
  }
  return vars;
}

export interface ScanOptions extends FindOptions {
  /**
   * `full`: a profile per workbook and complete usage lists. `compact`: no profiles, usage lists
   * capped at five workbooks each (with the total). Default: full up to 50 workbooks.
   */
  detail?: 'full' | 'compact';
}

export interface ScanResult {
  folder: string;
  workbooks: number;
  withAfo: number;
  truncated: boolean;
  detail: 'full' | 'compact';
  systems: Record<string, number>;
  afoRevisions: Record<string, number>;
  profiles: WorkbookProfile[];
  dataSources: Array<Usage & { detail?: string }>;
  planningObjects: Array<Usage & { detail?: string }>;
  variables: Usage[];
  patterns: Array<{ signature: string[]; workbooks: string[] }>;
  outliers: string[];
}

export function scanWorkbooks(folder: string, opts: ScanOptions = {}): ScanResult {
  const { files, truncated } = findWorkbooks(folder, opts);
  const detail = opts.detail ?? (files.length > 50 ? 'compact' : 'full');
  const profiles: WorkbookProfile[] = [];
  const descriptions: Array<{ rel: string; d: WorkbookDescription }> = [];
  const dataSources = new Map<string, Set<string>>();
  const dsNames = new Map<string, string>();
  const planning = new Map<string, Set<string>>();
  const planningInfo = new Map<string, string>();
  const variables = new Map<string, Set<string>>();

  for (const file of files) {
    const rel = relative(folder, file) || file;
    try {
      const d = readWorkbook(file, { sections: ['settings', 'dataProviders', 'planningObjects', 'formulas'], maxFormulasPerSheet: 100000 });
      descriptions.push({ rel, d });
      const s = d.summary;
      profiles.push({
        file: rel,
        afo: d.afo,
        systemId: s.systemId,
        afoRevision: s.afoRevision,
        lastRefreshed: s.lastRefreshed,
        dataSources: s.dataSources,
        planningObjects: s.planningObjects,
        sheets: s.sheets,
        grids: s.grids,
        sapFormulaCells: s.sapFormulaCells,
        macros: s.vbaModules ? `${s.procedures} procedures, ${s.afoApiCalls} AfO API calls, ${s.buttons} buttons` : s.buttons ? `${s.buttons} buttons, no VBA` : 'none',
        issues: d.issues,
      });
      for (const dp of d.dataProviders ?? []) {
        addUsage(dataSources, dp.dataSource, rel);
        if (dp.dataSource && dp.dataSourceName && !dsNames.has(dp.dataSource)) dsNames.set(dp.dataSource, dp.dataSourceName);
      }
      for (const p of d.planningObjects ?? []) {
        addUsage(planning, p.name, rel);
        if (p.name && !planningInfo.has(p.name)) {
          planningInfo.set(p.name, [p.kind, p.description, p.aggregationLevel && `on ${p.aggregationLevel}`].filter(Boolean).join(', '));
        }
      }
      for (const v of variablesOf(d)) addUsage(variables, v, rel);
    } catch (err) {
      profiles.push({ file: rel, afo: false, dataSources: [], planningObjects: [], sheets: 0, grids: 0, sapFormulaCells: 0, macros: 'unknown', issues: [], error: (err as Error).message });
    }
  }

  // Recurring layouts: the same alias bound to the same shared data source, plus the planning object keys.
  const shared = new Set([...dataSources].filter(([, f]) => f.size > 1).map(([n]) => n));
  const groups = new Map<string, string[]>();
  for (const { rel, d } of descriptions) {
    const sig = [
      ...(d.dataProviders ?? []).filter((dp) => dp.dataSource && shared.has(dp.dataSource)).map((dp) => `${dp.key}=${dp.dataSource}`),
      ...(d.planningObjects ?? []).map((p) => `${p.key} (${p.kind})`),
    ].sort();
    if (sig.length < 2) continue;
    const key = sig.join(' | ');
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(rel);
  }
  const patterns = [...groups]
    .filter(([, wbs]) => wbs.length > 1)
    .map(([k, wbs]) => ({ signature: k.split(' | '), workbooks: wbs }))
    .sort((a, b) => b.workbooks.length - a.workbooks.length);

  const afoDescriptions = descriptions.filter((x) => x.d.afo);
  const mainSystem = majority(afoDescriptions.map((x) => x.d.summary.systemId));
  const mainRevision = majority(afoDescriptions.map((x) => x.d.summary.afoRevision));
  const outliers: string[] = [];
  for (const { rel, d } of afoDescriptions) {
    if (mainSystem && d.summary.systemId !== mainSystem) outliers.push(`${rel}: saved against system ${d.summary.systemId}, most workbooks use ${mainSystem}.`);
    if (mainRevision && d.summary.afoRevision !== mainRevision) outliers.push(`${rel}: saved with AfO ${d.summary.afoRevision}, most workbooks with ${mainRevision}.`);
    if (d.issues.length) outliers.push(`${rel}: ${d.issues.length} issue(s), see its profile.`);
  }
  for (const p of profiles) {
    if (p.error) outliers.push(`${p.file}: could not be read (${p.error}).`);
    else if (!p.afo) outliers.push(`${p.file}: no AfO definition.`);
  }

  const cap = <T extends Usage>(list: T[]) =>
    detail === 'full' ? list : list.map((u) => (u.workbooks.length > 5 ? { ...u, workbooks: [...u.workbooks.slice(0, 5), `… ${u.workbooks.length - 5} more (${u.workbooks.length} in total)`] } : u));
  return {
    folder,
    workbooks: files.length,
    withAfo: afoDescriptions.length,
    truncated,
    detail,
    systems: countBy(afoDescriptions.map((x) => x.d.summary.systemId)),
    afoRevisions: countBy(afoDescriptions.map((x) => x.d.summary.afoRevision)),
    profiles: detail === 'full' ? profiles : [],
    dataSources: cap(usageList(dataSources, dsNames)),
    planningObjects: cap(usageList(planning, planningInfo)),
    variables: cap(usageList(variables)),
    patterns: detail === 'full' ? patterns : patterns.map((p) => ({ ...p, workbooks: cap([{ name: '', workbooks: p.workbooks }])[0].workbooks })),
    outliers,
  };
}

export interface WhereUsedHit {
  file: string;
  places: string[];
}

/**
 * Every place a technical name occurs across a folder: as data source, InfoProvider,
 * planning object, aggregation level, filter, variable, characteristic, in SAP formulas and in
 * VBA code. Matching is exact on the whole name, case-insensitive.
 */
export function whereUsed(folder: string, name: string, opts: FindOptions = {}): { name: string; workbooks: number; hits: WhereUsedHit[]; errors: string[] } {
  const { files } = findWorkbooks(folder, opts);
  const target = name.trim().toUpperCase();
  const token = new RegExp(`(^|[^A-Za-z0-9_/])${target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}($|[^A-Za-z0-9_])`, 'i');
  const is = (v: string | undefined) => !!v && v.toUpperCase() === target;
  const hits: WhereUsedHit[] = [];
  const errors: string[] = [];
  for (const file of files) {
    const rel = relative(folder, file) || file;
    let d: WorkbookDescription;
    try {
      d = readWorkbook(file, { macroSource: true, maxFormulasPerSheet: 100000 });
    } catch (err) {
      errors.push(`${rel}: ${(err as Error).message}`);
      continue;
    }
    const places: string[] = [];
    for (const v of d.settings?.promptVariables ?? []) if (is(v.name)) places.push('prompt variable');
    for (const dp of d.dataProviders ?? []) {
      if (is(dp.dataSource)) places.push(`data source ${dp.key} (${dp.type ?? 'Query'})`);
      if (is(dp.key)) places.push(`data source alias ${dp.key} = ${dp.dataSource}`);
      const st = dp.state;
      if (!st) continue;
      if (is(st.infoProvider) && !is(dp.dataSource)) places.push(`InfoProvider of data source ${dp.key}`);
      for (const f of st.filters) if (is(f.characteristic)) places.push(`filter on ${dp.key}: ${f.selections.join(', ')}`);
      for (const axis of ['rows', 'columns', 'free'] as const) if (st[axis]?.some(is)) places.push(`${axis} of ${dp.key}`);
      for (const h of st.hierarchies ?? []) if (is(h.characteristic) || is(h.hierarchy)) places.push(`hierarchy ${h.hierarchy} on ${h.characteristic} in ${dp.key}`);
      for (const v of st.variables ?? []) if (is(v.name)) places.push(`variable of ${dp.key}: ${v.values.join(', ')}`);
    }
    for (const p of d.planningObjects ?? []) {
      if (is(p.name)) places.push(`planning ${p.kind} ${p.key} on ${p.dataProviderKey ?? '?'}`);
      if (is(p.key)) places.push(`planning object alias ${p.key} = ${p.name}`);
      if (is(p.aggregationLevel)) places.push(`aggregation level of ${p.key}`);
      if (is(p.filterObject)) places.push(`filter of ${p.key}`);
      for (const f of p.filter) if (is(f.characteristic)) places.push(`filter of ${p.key}: ${f.selections.join(', ')}`);
      for (const v of p.variables) if (is(v.name)) places.push(`variable of ${p.key}: ${v.values.join(', ') || '(no value)'}`);
      for (const r of p.references) if (is(r.technicalName)) places.push(`${r.object.toLowerCase()} reference of ${p.key} from ${r.referenceType}${r.cell ? ` ${r.cell}` : ''}`);
    }
    for (const f of d.items?.formulaItems ?? []) if (f.formula && token.test(f.formula)) places.push(`formula item ${f.key} at ${f.sheet}!${f.address}: ${f.formula}`);
    for (const s of d.formulas ?? []) {
      const cells = s.formulas.filter((f) => f.calls.some((c) => c.args.some((a) => is(literal(a)) || is(a))));
      if (cells.length) places.push(`SAP formulas on ${s.sheet}: ${cells.slice(0, 10).map((f) => f.cell).join(', ')}${cells.length > 10 ? ` and ${cells.length - 10} more` : ''}`);
    }
    for (const m of d.macros?.modules ?? []) {
      const lines = (m.source ?? '').split('\n');
      const hitLines = lines.map((l, i) => (token.test(l) ? i + 1 : 0)).filter(Boolean);
      if (hitLines.length) places.push(`VBA ${m.name}, line ${hitLines.slice(0, 10).join(', ')}${hitLines.length > 10 ? ' …' : ''}`);
    }
    if (places.length) hits.push({ file: rel, places });
  }
  return { name, workbooks: files.length, hits, errors };
}
