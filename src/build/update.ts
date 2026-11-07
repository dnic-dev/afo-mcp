// Change an existing Analysis for Office workbook instead of rebuilding it: store prompt values,
// add planning functions to its AfO definition, retrofit the agent bridge, add VBA modules. Everything else
// (layout, formulas, data sources, crosstabs, existing macros) stays as it is.
//
// The workbook must be closed: while it is open, AfO holds the definition in memory and writes
// it back on save.

import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { extname, isAbsolute, resolve } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readWorkbook } from '../workbook/read.js';
import {
  type AfoWorkbookSpec, CT_CP, CT_MACRO_MAIN, CT_VBA, CT_WS, type PlanningFunctionSpec, type PromptValue, REL_CP, REL_VBA, REL_WS,
  addRels, addVisibility, esc, guid, planningFunction, promptProblems, promptsVisibility, quoteSheet, relXml, stateXml, unstoredPrompts, utf16,
  systemInfo, withStoredPrompts,
} from './afo.js';
import { type CompactRowsSpec, type HierarchySpec, hierarchyProblems, withStoredHierarchies } from './hierarchy.js';
import { firstElement, keyedElement, keyedElements, readSourceDefinition, sourceSystemInfo } from './source.js';
import { landscapeSystem } from '../operate/landscape.js';
import {
  AGENT_MODULE, AGENT_NAMES, AGENT_SHEET_CODE_NAME, type AgentBridgeSpec, agentModuleCode, agentSheetName, agentSheetXml, validateAgentBridge,
} from './agent.js';
import { type VbaModuleSource, writeVbaProject } from './vba.js';
import { patchVbaProject, vbaModuleNames } from './vbaPatch.js';

/** Prompt values of one data source: a list as in the build spec, or variable → member key or value object. */
export type PromptInput = PromptValue[] | Record<string, string | Omit<PromptValue, 'name'>>;

export interface UpdateOperations {
  /** Prompt values to store, by data source key; other variables of the data source keep theirs. */
  prompts?: Record<string, PromptInput>;
  /** StorePromptsInDocument; true by default when prompts are given. */
  storePrompts?: boolean;
  /**
   * RefreshOnOpen. A workbook that refreshes on open asks for its own logon while it loads,
   * before a logon through the API can run; false lets `afo_open_workbook` log on first.
   */
  refreshOnOpen?: boolean;
  /** Planning functions to add to the AfO definition, bound to an existing data source. */
  planningFunctions?: PlanningFunctionSpec[];
  /** Retrofit the agent bridge (bridge sheet, AfoAgent module, event handlers). */
  agentBridge?: AgentBridgeSpec;
  /** Standard modules to add, e.g. the macros behind new agent commands. */
  vbaModules?: Array<{ name: string; code?: string; file?: string }>;
  /** Point every data source and planning object to another system. */
  system?: SystemChange;
  /** BW hierarchies to show on open, by data source key (null as hierarchy: flat display). */
  hierarchies?: Record<string, HierarchySpec[]>;
  /** Compact display of rows on or off, by data source key. */
  compactRows?: Record<string, CompactRowsSpec>;
}

/**
 * The target system: an entry of the SAP Logon landscape of this machine (client and language
 * default to the workbook's), or the connection of a reference workbook.
 */
export type SystemChange = { systemId: string; client?: string; language?: string } | { from: string };

const SYSTEM_INFO = /<SystemInfo>.*?<\/SystemInfo>/gs;
const tag = (xml: string, name: string) => new RegExp(`<${name}>([^<]*)</${name}>`).exec(xml)?.[1];

export interface UpdateOptions {
  input: string;
  output: string;
  overwrite?: boolean;
  /** Folder that relative VBA module files resolve against. */
  baseDir?: string;
}

export interface UpdateResult {
  output: string;
  sizeBytes: number;
  changes: string[];
  sheets: string[];
  dataSources: string[];
  planningObjects: string[];
  vbaModules: string[];
  warnings: string[];
  issues: string[];
}

const HIDDEN_SHEET = '_com.sap.ip.bi.xl.hiddensheet';
const VBA_NAME = /^[A-Za-z][A-Za-z0-9_]{0,30}$/;

function relTargets(relsXml: string): Map<string, string> {
  return new Map([...relsXml.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
}

/** `xl/worksheets/sheet1.xml` + `../customProperty3.bin` → `xl/customProperty3.bin`. */
function resolvePart(owner: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = owner.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

function normalizePrompts(input: PromptInput): PromptValue[] {
  if (Array.isArray(input)) return input;
  return Object.entries(input ?? {}).map(([name, v]) => (typeof v === 'string' ? { name, member: v } : { ...v, name }));
}

/** Set a workbook setting such as StorePromptsInDocument, adding it after Cleaned when it is missing. */
function setSetting(def: string, name: string, value: boolean): string {
  const el = `<${name}>${value ? 'True' : 'False'}</${name}>`;
  const re = new RegExp(`<${name}>[^<]*</${name}>|<${name}\\s*/>`);
  if (re.test(def)) return def.replace(re, () => el);
  return def.replace(/<\/Cleaned>|<PromptsVisibility/, (m) => (m === '</Cleaned>' ? `${m}${el}` : `${el}${m}`));
}

/** Put a line right after the header of an existing event handler, or append the handler. */
function hookHandler(source: string, event: string, call: (param: string) => string, handler: string): string {
  const re = new RegExp(`^([ \\t]*(?:Private\\s+|Public\\s+)?Sub\\s+${event}\\s*\\(([^)]*)\\)[^\\r\\n]*)(\\r?\\n)`, 'im');
  const m = re.exec(source);
  if (!m) return `${source.replace(/\s+$/, '')}\r\n\r\n${handler}`;
  const param = /(\w+)\s+As\s+Object/i.exec(m[2])?.[1] ?? /(\w+)\s*$/.exec(m[2].trim())?.[1] ?? 'Sh';
  return source.slice(0, m.index) + m[1] + m[3] + `    ${call(param)}\r\n` + source.slice(m.index + m[0].length);
}

export function updateWorkbook(ops: UpdateOperations, options: UpdateOptions): UpdateResult {
  const input = resolve(options.input);
  const output = resolve(options.output);
  if (!existsSync(input)) throw new Error(`The workbook does not exist: ${input}`);
  if (input.toLowerCase() === output.toLowerCase()) throw new Error('input and output must be different files; replace the original yourself once the result is checked.');
  if (existsSync(output) && !options.overwrite) throw new Error(`The output file exists: ${output}. Pass overwrite: true to replace it.`);
  const ext = extname(output).toLowerCase();
  if (ext !== '.xlsm' && ext !== '.xlsx') throw new Error(`output must end in .xlsx or .xlsm, got ${options.output}.`);
  const prompts = Object.entries(ops.prompts ?? {}).map(([key, input]) => [key, normalizePrompts(input)] as const);
  const storePrompts = ops.storePrompts ?? (prompts.length ? true : undefined);
  const displayKeys = [...new Set([...Object.keys(ops.hierarchies ?? {}), ...Object.keys(ops.compactRows ?? {})])];
  const refreshOnOpen = ops.refreshOnOpen;
  if (!prompts.length && storePrompts === undefined && refreshOnOpen === undefined && !ops.planningFunctions?.length && !ops.agentBridge && !ops.vbaModules?.length && !ops.system && !displayKeys.length) {
    throw new Error('Nothing to do: give prompts, storePrompts, refreshOnOpen, planningFunctions, agentBridge, vbaModules, system, hierarchies or compactRows.');
  }

  const d = readWorkbook(input, { sections: ['sheets', 'dataProviders', 'planningObjects', 'items', 'macros'] });
  if (!d.afo) throw new Error('Not an Analysis for Office workbook (no AfO definition found).');
  const parts: Record<string, Uint8Array> = unzipSync(readFileSync(input));
  const text = (n: string) => strFromU8(parts[n]);
  const errors: string[] = [];
  const warnings: string[] = [];
  const changes: string[] = [];

  const hasVba = !!parts['xl/vbaProject.bin'];
  const needsVba = !!ops.agentBridge || !!ops.vbaModules?.length;
  if ((hasVba || needsVba) && ext !== '.xlsm') errors.push('The result contains macros; write it as .xlsm.');
  if (needsVba && Object.keys(parts).some((p) => /^xl\/vbaProjectSignature/i.test(p))) errors.push('The VBA project is signed; changing it would break the signature. Re-sign it after a manual change instead.');
  const existingModules = hasVba ? vbaModuleNames(parts['xl/vbaProject.bin']) : [];
  const moduleSet = new Set(existingModules.map((n) => n.toUpperCase()));

  // VBA modules to add.
  const addedModules = (ops.vbaModules ?? []).map((m, i) => {
    if (typeof m?.name !== 'string' || !VBA_NAME.test(m.name)) errors.push(`vbaModules[${i}]: name must be a VBA identifier.`);
    else if (moduleSet.has(m.name.toUpperCase())) errors.push(`vbaModules[${i}]: module ${m.name} already exists in the workbook.`);
    if ((m?.code === undefined) === (m?.file === undefined)) { errors.push(`vbaModules[${i}] (${m?.name}): give either code or file.`); return undefined; }
    if (m.code !== undefined) return { name: m.name, code: m.code };
    const file = isAbsolute(m.file!) ? m.file! : options.baseDir ? resolve(options.baseDir, m.file!) : undefined;
    if (!file || !existsSync(file)) { errors.push(`vbaModules[${i}] (${m.name}): file not found or relative without a base folder: ${m.file}`); return undefined; }
    return { name: m.name, code: readFileSync(file, 'utf8').replace(/^﻿/, '') };
  }).filter((m): m is { name: string; code: string } => !!m);

  // Planning functions.
  const keys = new Set([...(d.dataProviders ?? []).map((x) => x.key), ...(d.planningObjects ?? []).map((x) => x.key), ...(d.items?.grids ?? []).map((x) => x.key)].map((k) => k.toUpperCase()));
  const dpKeys = new Set((d.dataProviders ?? []).map((x) => x.key));
  (ops.planningFunctions ?? []).forEach((pf, i) => {
    const what = `planningFunctions[${i}]`;
    if (typeof pf?.key !== 'string' || !/^[A-Za-z][A-Za-z0-9_]*$/.test(pf.key)) errors.push(`${what}: key must be letters, digits and underscores.`);
    else if (keys.has(pf.key.toUpperCase())) errors.push(`${what}: key ${pf.key} is already used in the workbook.`);
    else keys.add(pf.key.toUpperCase());
    if (!pf?.name) errors.push(`${what}: name (the planning function) is required.`);
    if (!dpKeys.has(pf?.dataProviderKey)) errors.push(`${what} (${pf?.key}): data source ${pf?.dataProviderKey} is not in the workbook. Data sources: ${[...dpKeys].join(', ')}.`);
    if (!pf?.aggregationLevel) errors.push(`${what} (${pf?.key}): aggregationLevel is required.`);
    if (!Array.isArray(pf?.characteristics) || !pf.characteristics.length) errors.push(`${what} (${pf?.key}): characteristics must list the characteristics of the aggregation level.`);
  });

  // Prompts.
  for (const [key, list] of prompts) {
    if (!dpKeys.has(key)) errors.push(`prompts.${key}: data source ${key} is not in the workbook. Data sources: ${[...dpKeys].join(', ')}.`);
    if (!list.length) errors.push(`prompts.${key}: no values given.`);
    list.forEach((p, j) => errors.push(...promptProblems(p, `prompts.${key}[${j}]`)));
  }
  for (const key of displayKeys) {
    if (!dpKeys.has(key)) errors.push(`hierarchies/compactRows.${key}: data source ${key} is not in the workbook. Data sources: ${[...dpKeys].join(', ')}.`);
    errors.push(...hierarchyProblems(key, ops.hierarchies?.[key], ops.compactRows?.[key], undefined));
  }
  if (storePrompts === false && prompts.length) warnings.push('storePrompts is false: the prompt values are written, but AfO ignores them until StorePromptsInDocument is on.');

  // Target system.
  let newSystemInfo: string | undefined;
  if (ops.system) {
    const current = readSourceDefinition(input).match(SYSTEM_INFO)?.[0] ?? '';
    if ('from' in ops.system) {
      const ref = resolve(ops.system.from);
      if (!existsSync(ref)) errors.push(`system.from: the reference workbook does not exist: ${ref}`);
      else {
        try { newSystemInfo = sourceSystemInfo(readSourceDefinition(ref)); } catch (err) { errors.push(`system.from: ${(err as Error).message}`); }
        if (!newSystemInfo && !errors.length) errors.push(`system.from: no connection found in ${ref}.`);
      }
    } else if (typeof ops.system.systemId !== 'string' || !ops.system.systemId.trim()) {
      errors.push('system: give systemId (as in SAP Logon) or from (a reference workbook).');
    } else {
      const client = ops.system.client ?? tag(current, 'Client');
      if (!client) errors.push('system.client: the workbook has no client to keep; give one.');
      else {
        try {
          newSystemInfo = systemInfo(landscapeSystem(ops.system.systemId.trim(), client, ops.system.language ?? tag(current, 'Language') ?? 'EN'));
        } catch (err) { errors.push(`system: ${(err as Error).message}`); }
      }
    }
  }

  // Agent bridge.
  const sheetNames = (d.sheets ?? []).map((s) => s.name);
  const specLike = {
    agentBridge: ops.agentBridge,
    // The installed bridge is reported on its own below, not as a module name clash.
    vbaModules: [...existingModules, ...addedModules.map((m) => m.name)].filter((n) => n.toUpperCase() !== AGENT_MODULE.toUpperCase()).map((name) => ({ name, code: '' })),
    dataProviders: (d.dataProviders ?? []).map((dp) => ({ key: dp.key, query: dp.dataSource ?? dp.dataSourceName ?? '?', inputReady: dp.inputReady })),
    planningFunctions: [
      ...(d.planningObjects ?? []).map((po) => ({ key: po.key, name: po.name ?? '?', description: po.description, dataProviderKey: po.dataProviderKey ?? '?', aggregationLevel: '', characteristics: [] })),
      ...(ops.planningFunctions ?? []),
    ],
    grids: (d.items?.grids ?? []).map((g) => ({ key: g.key, dataProviderKey: g.dataProviderKey ?? '?', sheet: g.sheet ?? '?', cell: g.address ?? '?' })),
  } as unknown as AfoWorkbookSpec;
  if (ops.agentBridge) {
    errors.push(...validateAgentBridge(specLike));
    const name = agentSheetName(specLike);
    if (sheetNames.some((n) => n.toLowerCase() === name.toLowerCase())) errors.push(`The workbook already has a sheet ${name}${moduleSet.has(AGENT_MODULE.toUpperCase()) ? ' and the agent bridge' : ''}; set agentBridge.sheet to another name.`);
    if (moduleSet.has(AGENT_MODULE.toUpperCase())) errors.push(`The workbook already has the module ${AGENT_MODULE}: the bridge is installed.`);
    if ((d.sheets ?? []).some((s) => s.codeName?.toLowerCase() === AGENT_SHEET_CODE_NAME.toLowerCase())) errors.push(`A sheet already has the code name ${AGENT_SHEET_CODE_NAME}.`);
  }
  if (errors.length) throw new Error(`The workbook was not changed:\n- ${errors.join('\n- ')}`);

  let wb = text('xl/workbook.xml');
  let wbRels = text('xl/_rels/workbook.xml.rels');
  let ct = text('[Content_Types].xml');
  const targets = relTargets(wbRels);

  // --- The definition on the hidden sheet: prompts, settings, planning functions.
  if (ops.planningFunctions?.length || prompts.length || storePrompts !== undefined || refreshOnOpen !== undefined || newSystemInfo || displayKeys.length) {
    const hidden = [...wb.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)].find((m) => m[1] === HIDDEN_SHEET);
    if (!hidden) throw new Error('The AfO definition sheet was not found.');
    const sheetPart = resolvePart('xl/workbook.xml', targets.get(hidden[2])!);
    const sheetRels = text(sheetPart.replace(/([^/]+)$/, '_rels/$1.rels'));
    const rid = /<customPr [^>]*name="serializedData2"[^>]*r:id="([^"]+)"/.exec(text(sheetPart))?.[1];
    if (!rid) throw new Error('The AfO definition (serializedData2) was not found on the hidden sheet.');
    const defPart = resolvePart(sheetPart, relTargets(sheetRels).get(rid)!);
    const raw = Buffer.from(parts[defPart]);
    const bom = raw[0] === 0xff && raw[1] === 0xfe;
    let def = raw.subarray(bom ? 2 : 0).toString('utf16le');
    const day = new Date().toISOString().slice(0, 10).replace(/-/g, '');

    if (prompts.length) {
      const merged = firstElement(def, 'MergeVariables') === '<MergeVariables>True</MergeVariables>';
      const visibility = promptsVisibility(def);
      for (const [key, list] of prompts) {
        const dpXml = keyedElement(def, 'DataProvider', key)!;
        const dp = (d.dataProviders ?? []).find((x) => x.key === key)!;
        const r = withStoredPrompts(dpXml, list, () => stateXml({ key, query: dp.dataSource ?? dp.dataSourceName ?? '', inputReady: dp.inputReady ?? false }, day));
        def = def.replace(dpXml, () => r.xml);
        addVisibility(visibility, key, list, merged);
        warnings.push(...r.notes.map((n) => `${key}: ${n}.`));
        changes.push(`prompts of ${key}: ${list.map((p) => p.name).join(', ')}`);
      }
      const block = `<PromptsVisibility>${visibility.map((v) => v.xml).join('')}</PromptsVisibility>`;
      const current = firstElement(def, 'PromptsVisibility');
      def = current ? def.replace(current, () => block) : def.replace(/<StyleSetName|<PlanningSystemDesignTime|<DataProviders/, (m) => `${block}${m}`);
      const unstored = unstoredPrompts(keyedElements(def, 'DataProvider'), visibility);
      if (unstored) warnings.push(unstored);
    }
    for (const key of displayKeys) {
      const dpXml = keyedElement(def, 'DataProvider', key)!;
      const dp = (d.dataProviders ?? []).find((x) => x.key === key)!;
      const list = ops.hierarchies?.[key];
      const compact = ops.compactRows?.[key];
      let r: { xml: string; notes: string[] };
      try {
        r = withStoredHierarchies(dpXml, list, compact, () => stateXml({ key, query: dp.dataSource ?? dp.dataSourceName ?? '', inputReady: dp.inputReady ?? false }, day));
      } catch (err) {
        throw new Error(`${key}: ${(err as Error).message}`);
      }
      def = def.replace(dpXml, () => r.xml);
      warnings.push(...r.notes.map((n) => `${key}: ${n}.`));
      for (const h of list ?? []) changes.push(`${key}: ${h.characteristic} ${h.hierarchy === null ? 'flat' : `hierarchy ${h.hierarchy}, ${h.drillLevel ?? 2} level(s) shown`}`);
      if (compact !== undefined) changes.push(`${key}: compact display of rows ${compact === false ? 'off' : `on${typeof compact === 'object' && compact.drillLevel ? `, ${compact.drillLevel} level(s) shown` : ''}`}`);
    }
    if (storePrompts !== undefined) {
      def = setSetting(def, 'StorePromptsInDocument', storePrompts);
      changes.push(`StorePromptsInDocument ${storePrompts ? 'on' : 'off'}`);
    }
    if (refreshOnOpen !== undefined) {
      def = setSetting(def, 'RefreshOnOpen', refreshOnOpen);
      changes.push(`RefreshOnOpen ${refreshOnOpen ? 'on' : 'off'}`);
    }
    if (newSystemInfo) {
      // Every connection (data sources, planning objects, planning design time) and the workbook's system id.
      const before = [...new Set((def.match(SYSTEM_INFO) ?? []).map((s) => `${tag(s, 'SystemId')}/${tag(s, 'Client')}`))];
      const count = (def.match(SYSTEM_INFO) ?? []).length;
      const target = newSystemInfo;
      const id = tag(target, 'SystemId') ?? '';
      def = def.replace(SYSTEM_INFO, () => target).replace(/(<\/AAO_Revision>)<SystemId>[^<]*<\/SystemId>/, (_, m) => `${m}<SystemId>${esc(id)}</SystemId>`);
      changes.push(`system ${before.join(', ') || '?'} → ${id}/${tag(target, 'Client')} (${count} connections)`);
      warnings.push(`The queries, planning objects and stored variable values must exist in ${id} as well; check them there (bw-modeling-mcp) before handing the workbook out.`);
    }

    const xml = (ops.planningFunctions ?? []).map((pf) => {
      const dp = new RegExp(`<DataProvider>(?:(?!</DataProvider>).)*?<Key>${pf.dataProviderKey}</Key>(?:(?!</DataProvider>).)*?(<SystemInfo>.*?</SystemInfo>)`, 's').exec(def);
      if (!dp) throw new Error(`No system information found for data source ${pf.dataProviderKey}.`);
      changes.push(`planning function ${pf.key} = ${pf.name} on ${pf.dataProviderKey}`);
      return planningFunction(pf, dp[1], day);
    }).join('');
    if (xml) {
      if (def.includes('</PlanningObjects>')) def = def.replace('</PlanningObjects>', () => `${xml}</PlanningObjects>`);
      else if (/<PlanningObjects\s*\/>/.test(def)) def = def.replace(/<PlanningObjects\s*\/>/, () => `<PlanningObjects>${xml}</PlanningObjects>`);
      else if (def.includes('</DataProviders>')) def = def.replace('</DataProviders>', () => `</DataProviders><PlanningObjects>${xml}</PlanningObjects>`);
      else throw new Error('The AfO definition has no DataProviders section to place planning objects after.');
    }
    parts[defPart] = new Uint8Array(Buffer.concat([bom ? Buffer.from([0xff, 0xfe]) : Buffer.alloc(0), Buffer.from(def, 'utf16le')]));
  }

  // --- Agent bridge: sheet, AfO ids, defined names.
  let bridgeLogRow = 0;
  if (ops.agentBridge) {
    const name = agentSheetName(specLike);
    const { xml, logFirstRow } = agentSheetXml(specLike);
    bridgeLogRow = logFirstRow;
    const used = Object.keys(parts).map((p) => /^xl\/customProperty(\d+)\.bin$/.exec(p)?.[1]).filter(Boolean).map(Number);
    const a = Math.max(900, ...used) + 1, b = a + 1;
    parts[`xl/customProperty${a}.bin`] = utf16(`${AGENT_SHEET_CODE_NAME}|${guid()}`);
    parts[`xl/customProperty${b}.bin`] = utf16('AoWorksheetType');
    parts['xl/worksheets/afoAgent.xml'] = strToU8(xml.replace('</worksheet>', `<customProperties><customPr name="_pios_id" r:id="rIdAfo${a}"/><customPr name="CofWorksheetType" r:id="rIdAfo${b}"/></customProperties></worksheet>`));
    addRels(parts, 'xl/worksheets/_rels/afoAgent.xml.rels', relXml(`rIdAfo${a}`, REL_CP, `../customProperty${a}.bin`) + relXml(`rIdAfo${b}`, REL_CP, `../customProperty${b}.bin`));
    const ids = [...wb.matchAll(/<sheet [^>]*sheetId="(\d+)"/g)].map((m) => Number(m[1]));
    wb = wb.replace('</sheets>', () => `<sheet name="${esc(name)}" sheetId="${Math.max(0, ...ids) + 1}" r:id="rIdAfoAgent"${ops.agentBridge!.hidden ?? true ? ' state="hidden"' : ''}/></sheets>`);
    wbRels = wbRels.replace('</Relationships>', () => relXml('rIdAfoAgent', REL_WS, 'worksheets/afoAgent.xml') + '</Relationships>');
    ct = ct.replace('</Types>', () => `<Override PartName="/xl/worksheets/afoAgent.xml" ContentType="${CT_WS}"/><Override PartName="/xl/customProperty${a}.bin" ContentType="${CT_CP}"/><Override PartName="/xl/customProperty${b}.bin" ContentType="${CT_CP}"/></Types>`);
    const ref = (cell: string) => `${esc(quoteSheet(name))}!${cell}`;
    const names = `<definedName name="${AGENT_NAMES.command}">${ref('$B$1')}</definedName><definedName name="${AGENT_NAMES.result}">${ref('$B$3')}</definedName><definedName name="${AGENT_NAMES.status}">${ref('$B$4')}</definedName>`;
    wb = /<definedNames>/.test(wb) ? wb.replace('<definedNames>', () => `<definedNames>${names}`) : wb.replace('</sheets>', () => `</sheets><definedNames>${names}</definedNames>`);
    changes.push(`agent bridge on sheet ${name}${ops.agentBridge.hidden ?? true ? ' (hidden)' : ''}${ops.agentBridge.commands?.length ? `, commands ${ops.agentBridge.commands.map((c) => c.name).join(', ')}` : ''}`);
  }

  // --- VBA.
  if (needsVba) {
    const newModules: VbaModuleSource[] = addedModules.map((m) => ({ name: m.name, type: 'standard', code: m.code }));
    if (ops.agentBridge) {
      newModules.push({ name: AGENT_SHEET_CODE_NAME, type: 'document', document: 'sheet', code: '' });
      newModules.push({ name: AGENT_MODULE, type: 'standard', code: agentModuleCode(specLike, bridgeLogRow) });
    }
    const hooks = (source: string) => {
      let s = hookHandler(source, 'Workbook_Open', () => `${AGENT_MODULE}.MarkActive`, `Private Sub Workbook_Open()\r\n    ${AGENT_MODULE}.MarkActive\r\nEnd Sub\r\n`);
      s = hookHandler(s, 'Workbook_SheetCalculate', (p) => `If ${p}.Name = ${AGENT_MODULE}.BRIDGE_SHEET Then ${AGENT_MODULE}.OnCalculate`,
        `Private Sub Workbook_SheetCalculate(ByVal Sh As Object)\r\n    If Sh.Name = ${AGENT_MODULE}.BRIDGE_SHEET Then ${AGENT_MODULE}.OnCalculate\r\nEnd Sub\r\n`);
      return s;
    };
    if (hasVba) {
      const workbookModule = /<workbookPr[^>]*codeName="([^"]+)"/.exec(wb)?.[1] ?? 'ThisWorkbook';
      if (ops.agentBridge && !moduleSet.has(workbookModule.toUpperCase())) throw new Error(`The workbook module ${workbookModule} is not in the VBA project.`);
      parts['xl/vbaProject.bin'] = patchVbaProject(parts['xl/vbaProject.bin'], {
        add: newModules,
        edit: ops.agentBridge ? { [workbookModule]: hooks } : undefined,
      });
    } else {
      // No project yet: code names for the workbook and every sheet, then a new project.
      if (!/<workbookPr[^>]*codeName=/.test(wb)) {
        wb = /<workbookPr/.test(wb) ? wb.replace('<workbookPr', '<workbookPr codeName="ThisWorkbook"') : wb.replace(/<bookViews|<sheets>/, (m) => `<workbookPr codeName="ThisWorkbook"/>${m}`);
      }
      const workbookModule = /<workbookPr[^>]*codeName="([^"]+)"/.exec(wb)![1];
      const docs: VbaModuleSource[] = [{ name: workbookModule, type: 'document', document: 'workbook', code: ops.agentBridge ? hooks('Option Explicit\r\n') : '' }];
      const wbTargets = relTargets(wbRels);
      [...wb.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)].forEach(([, sheetName, rid], i) => {
        if (rid === 'rIdAfoAgent') return;
        const part = resolvePart('xl/workbook.xml', wbTargets.get(rid)!);
        let xml = text(part);
        let codeName = /<sheetPr[^>]*codeName="([^"]+)"/.exec(xml)?.[1];
        if (!codeName) {
          codeName = sheetName === HIDDEN_SHEET ? 'AfoDefinition' : `Sheet${i + 1}`;
          xml = /<sheetPr/.test(xml) ? xml.replace('<sheetPr', `<sheetPr codeName="${codeName}"`) : xml.replace(/(<worksheet[^>]*>)/, `$1<sheetPr codeName="${codeName}"/>`);
          parts[part] = strToU8(xml);
        }
        docs.push({ name: codeName, type: 'document', document: 'sheet', code: '' });
      });
      parts['xl/vbaProject.bin'] = writeVbaProject([...docs, ...newModules]);
      wbRels = wbRels.replace('</Relationships>', () => relXml('rIdVba', REL_VBA, 'vbaProject.bin') + '</Relationships>');
      if (!/vbaProject\.bin/.test(ct)) ct = ct.replace('</Types>', () => `<Override PartName="/xl/vbaProject.bin" ContentType="${CT_VBA}"/></Types>`);
    }
    if (addedModules.length) changes.push(`VBA modules ${addedModules.map((m) => m.name).join(', ')}`);
  }
  if (ext === '.xlsm') ct = ct.replace(/(PartName="\/xl\/workbook\.xml" ContentType=")[^"]+"/, `$1${CT_MACRO_MAIN}"`);

  parts['xl/workbook.xml'] = strToU8(wb);
  parts['xl/_rels/workbook.xml.rels'] = strToU8(wbRels);
  parts['[Content_Types].xml'] = strToU8(ct);
  writeFileSync(output, zipSync(parts));

  const r = readWorkbook(output, { sections: ['sheets', 'macros'] });
  return {
    output,
    sizeBytes: statSync(output).size,
    changes,
    sheets: (r.sheets ?? []).filter((s) => s.state !== 'veryHidden').map((s) => s.name),
    dataSources: r.summary.dataSources,
    planningObjects: r.summary.planningObjects,
    vbaModules: (r.macros?.modules ?? []).map((m) => m.name),
    warnings,
    issues: r.issues,
  };
}
