// Change an existing Analysis for Office workbook instead of rebuilding it: add planning
// functions to its AfO definition, retrofit the agent bridge, add VBA modules. Everything else
// (layout, formulas, data sources, crosstabs, existing macros) stays as it is.
//
// The workbook must be closed: while it is open, AfO holds the definition in memory and writes
// it back on save.

import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { extname, isAbsolute, resolve } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readWorkbook } from '../workbook/read.js';
import {
  type AfoWorkbookSpec, CT_CP, CT_MACRO_MAIN, CT_VBA, CT_WS, type PlanningFunctionSpec, REL_CP, REL_VBA, REL_WS,
  addRels, esc, guid, planningFunction, quoteSheet, relXml, utf16,
} from './afo.js';
import {
  AGENT_MODULE, AGENT_NAMES, AGENT_SHEET_CODE_NAME, type AgentBridgeSpec, agentModuleCode, agentSheetName, agentSheetXml, validateAgentBridge,
} from './agent.js';
import { type VbaModuleSource, writeVbaProject } from './vba.js';
import { patchVbaProject, vbaModuleNames } from './vbaPatch.js';

export interface UpdateOperations {
  /** Planning functions to add to the AfO definition, bound to an existing data source. */
  planningFunctions?: PlanningFunctionSpec[];
  /** Retrofit the agent bridge (bridge sheet, AfoAgent module, event handlers). */
  agentBridge?: AgentBridgeSpec;
  /** Standard modules to add, e.g. the macros behind new agent commands. */
  vbaModules?: Array<{ name: string; code?: string; file?: string }>;
}

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
  if (!ops.planningFunctions?.length && !ops.agentBridge && !ops.vbaModules?.length) throw new Error('Nothing to do: give planningFunctions, agentBridge or vbaModules.');

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

  // --- Planning functions: into the definition on the hidden sheet.
  if (ops.planningFunctions?.length) {
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
    const xml = ops.planningFunctions.map((pf) => {
      const dp = new RegExp(`<DataProvider>(?:(?!</DataProvider>).)*?<Key>${pf.dataProviderKey}</Key>(?:(?!</DataProvider>).)*?(<SystemInfo>.*?</SystemInfo>)`, 's').exec(def);
      if (!dp) throw new Error(`No system information found for data source ${pf.dataProviderKey}.`);
      changes.push(`planning function ${pf.key} = ${pf.name} on ${pf.dataProviderKey}`);
      return planningFunction(pf, dp[1], day);
    }).join('');
    if (def.includes('</PlanningObjects>')) def = def.replace('</PlanningObjects>', () => `${xml}</PlanningObjects>`);
    else if (/<PlanningObjects\s*\/>/.test(def)) def = def.replace(/<PlanningObjects\s*\/>/, () => `<PlanningObjects>${xml}</PlanningObjects>`);
    else if (def.includes('</DataProviders>')) def = def.replace('</DataProviders>', () => `</DataProviders><PlanningObjects>${xml}</PlanningObjects>`);
    else throw new Error('The AfO definition has no DataProviders section to place planning objects after.');
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
    wb = wb.replace('</sheets>', () => `<sheet name="${esc(name)}" sheetId="${Math.max(0, ...ids) + 1}" r:id="rIdAfoAgent"${ops.agentBridge!.hidden ? ' state="hidden"' : ''}/></sheets>`);
    wbRels = wbRels.replace('</Relationships>', () => relXml('rIdAfoAgent', REL_WS, 'worksheets/afoAgent.xml') + '</Relationships>');
    ct = ct.replace('</Types>', () => `<Override PartName="/xl/worksheets/afoAgent.xml" ContentType="${CT_WS}"/><Override PartName="/xl/customProperty${a}.bin" ContentType="${CT_CP}"/><Override PartName="/xl/customProperty${b}.bin" ContentType="${CT_CP}"/></Types>`);
    const ref = (cell: string) => `${esc(quoteSheet(name))}!${cell}`;
    const names = `<definedName name="${AGENT_NAMES.command}">${ref('$B$1')}</definedName><definedName name="${AGENT_NAMES.result}">${ref('$B$3')}</definedName><definedName name="${AGENT_NAMES.status}">${ref('$B$4')}</definedName>`;
    wb = /<definedNames>/.test(wb) ? wb.replace('<definedNames>', () => `<definedNames>${names}`) : wb.replace('</sheets>', () => `</sheets><definedNames>${names}</definedNames>`);
    changes.push(`agent bridge on sheet ${name}${ops.agentBridge.hidden ? ' (hidden)' : ''}${ops.agentBridge.commands?.length ? `, commands ${ops.agentBridge.commands.map((c) => c.name).join(', ')}` : ''}`);
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
