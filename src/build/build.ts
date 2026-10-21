// Build an AfO workbook file from a spec: check the spec, take a template workbook or generate a
// minimal one, write the AfO definition (afo.ts) and read the result back to prove every link
// resolves. Needs neither Excel nor an SAP logon.

import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, extname, isAbsolute, resolve } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readWorkbook } from '../workbook/read.js';
import { type AfoWorkbookSpec, fromA1, makeAfoWorkbook } from './afo.js';
import { agentSheetName, validateAgentBridge } from './agent.js';
import { applySource, readSourceDefinition } from './source.js';

/** A value written into the generated minimal layout, optionally with a defined name. */
export interface CellSpec {
  sheet: string;
  cell: string;
  value: string | number;
  /** Workbook-level defined name for the cell, e.g. "PlanVersion". */
  name?: string;
}

export interface MinimalLayout {
  /** Sheets in tab order; sheets named by crosstabs, cells or cell references are added after them. */
  sheets?: string[];
  cells?: CellSpec[];
}

/** The spec as the tool takes it: VBA modules may come from files. */
export type BuildSpec = Omit<AfoWorkbookSpec, 'vbaModules'> & {
  vbaModules?: Array<{ name: string; code?: string; file?: string }>;
};

export interface BuildOptions {
  output: string;
  /** Workbook (.xlsx/.xlsm) with the layout: sheets, texts, formats, shapes for buttons. */
  template?: string;
  /** Used when there is no template. */
  layout?: MinimalLayout;
  overwrite?: boolean;
  /** Folder that relative VBA module files resolve against. */
  baseDir?: string;
}

export interface BuildResult {
  output: string;
  sizeBytes: number;
  base: string;
  sheets: string[];
  dataSources: string[];
  planningObjects: string[];
  crosstabs: string[];
  buttons: string[];
  vbaModules: string[];
  warnings: string[];
  issues: string[];
}

const RESERVED_MODULES = /^(ThisWorkbook|AfoDefinition|Sheet\d+)$/i;
const VBA_NAME = /^[A-Za-z][A-Za-z0-9_]{0,30}$/;
const DEFINED_NAME = /^[A-Za-z_\\][A-Za-z0-9_.]*$/;
const HIDDEN_SHEET = '_com.sap.ip.bi.xl.hiddensheet';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const quoteSheet = (name: string) => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;

/** `'Plan Sheet'!$C$10` → sheet and cell; undefined when it is not a sheet-qualified cell. */
export function splitCellRef(ref: string): { sheet: string; cell: string } | undefined {
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Za-z]+\$?\d+)$/.exec(ref.trim());
  return m ? { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), cell: m[3] } : undefined;
}

function isA1(cell: string): boolean {
  try { fromA1(cell); return true; } catch { return false; }
}

function sheetNameError(name: string): string | undefined {
  if (!name.trim()) return 'sheet names must not be empty';
  if (name.length > 31) return `sheet name longer than 31 characters: ${name}`;
  if (/[[\]:*?/\\]/.test(name)) return `sheet name with a character Excel does not allow ([ ] : * ? / \\): ${name}`;
  if (name.startsWith("'") || name.endsWith("'")) return `sheet name starting or ending with an apostrophe: ${name}`;
  return undefined;
}

/** All problems of a spec at once, so a caller can fix them in one go. */
export function validateSpec(spec: BuildSpec, layout?: MinimalLayout): string[] {
  const errors: string[] = [];
  const s = spec as Partial<BuildSpec>;
  if (!s.source && (!s.system || typeof s.system !== 'object')) errors.push('system is required (systemId, systemName, client, language), unless a source workbook gives it.');
  else if (s.system && typeof s.system !== 'object') errors.push('system must be an object.');
  else if (s.system) for (const k of ['systemId', 'systemName', 'client', 'language'] as const) {
    if (typeof s.system[k] !== 'string' || !s.system[k].trim()) errors.push(`system.${k} is required.`);
  }
  const dps = Array.isArray(s.dataProviders) ? s.dataProviders : [];
  if (!dps.length) errors.push('dataProviders needs at least one data source.');
  const grids = Array.isArray(s.grids) ? s.grids : [];
  if (!Array.isArray(s.grids)) errors.push('grids is required (may be empty).');
  const pfs = s.planningFunctions ?? [];

  const keys = new Map<string, string>();
  const claim = (key: unknown, what: string) => {
    if (typeof key !== 'string' || !/^[A-Za-z][A-Za-z0-9_]*$/.test(key)) { errors.push(`${what}: key must be a name of letters, digits and underscores, got ${JSON.stringify(key)}.`); return; }
    const prev = keys.get(key.toUpperCase());
    if (prev) errors.push(`${what}: key ${key} is already used by ${prev}.`);
    else keys.set(key.toUpperCase(), what);
  };
  const dpKeys = new Set(dps.map((d) => d?.key));

  dps.forEach((dp, i) => {
    const what = `dataProviders[${i}]`;
    claim(dp?.key, what);
    if (dp?.from && !s.source) errors.push(`${what}: from needs a source workbook (source).`);
    if (!dp?.from && (typeof dp?.query !== 'string' || !dp.query.trim())) errors.push(`${what}: query is required.`);
    (dp?.prompts ?? []).forEach((p, j) => {
      const kinds = [p.member !== undefined, p.literal !== undefined, p.low !== undefined].filter(Boolean).length;
      if (!p.name) errors.push(`${what}.prompts[${j}]: name is required.`);
      if (kinds !== 1) errors.push(`${what}.prompts[${j}] (${p.name}): give exactly one of member, literal or low (with high).`);
      if (p.high !== undefined && p.low === undefined) errors.push(`${what}.prompts[${j}] (${p.name}): high without low.`);
    });
    for (const axis of ['rows', 'columns', 'free'] as const) {
      const v = dp?.axes?.[axis];
      if (v !== undefined && !Array.isArray(v)) errors.push(`${what}.axes.${axis} must be a list of names.`);
    }
  });
  grids.forEach((g, i) => {
    const what = `grids[${i}]`;
    claim(g?.key, what);
    if (!dpKeys.has(g?.dataProviderKey)) errors.push(`${what} (${g?.key}): dataProviderKey ${g?.dataProviderKey} is not a data source key.`);
    if (typeof g?.sheet !== 'string' || !g.sheet) errors.push(`${what} (${g?.key}): sheet is required.`);
    if (typeof g?.cell !== 'string' || !isA1(g.cell)) errors.push(`${what} (${g?.key}): cell must be an A1 reference such as B16, got ${JSON.stringify(g?.cell)}.`);
  });
  (s.cells ?? []).forEach((c, i) => {
    const what = `cells[${i}]`;
    if (typeof c?.sheet !== 'string' || !c.sheet) errors.push(`${what}: sheet is required.`);
    if (typeof c?.cell !== 'string' || !isA1(c.cell)) errors.push(`${what}: cell must be an A1 reference, got ${JSON.stringify(c?.cell)}.`);
    if ((c?.formula === undefined) === (c?.value === undefined)) errors.push(`${what}: give either formula or value.`);
  });
  (s.formulaItems ?? []).forEach((f, i) => {
    const what = `formulaItems[${i}]`;
    claim(f?.key, what);
    if (typeof f?.sheet !== 'string' || !f.sheet) errors.push(`${what} (${f?.key}): sheet is required.`);
    if (typeof f?.cell !== 'string' || !isA1(f.cell)) errors.push(`${what} (${f?.key}): cell must be an A1 reference, got ${JSON.stringify(f?.cell)}.`);
    if (f?.from && !s.source) errors.push(`${what}: from needs a source workbook (source).`);
    if (!f?.from && (typeof f?.formula !== 'string' || !/^=\s*SAP\w+\s*\(/i.test(f.formula))) errors.push(`${what} (${f?.key}): formula must be an SAP formula such as =SAPSetFilterComponent("DS_1", "DIMENSION", "").`);
  });
  pfs.forEach((pf, i) => {
    const what = `planningFunctions[${i}]`;
    claim(pf?.key, what);
    if (pf?.from && !s.source) errors.push(`${what}: from needs a source workbook (source).`);
    if (!pf?.from && !pf?.name) errors.push(`${what}: name (the planning function) is required.`);
    if (!dpKeys.has(pf?.dataProviderKey)) errors.push(`${what} (${pf?.key}): dataProviderKey ${pf?.dataProviderKey} is not a data source key.`);
    if (!pf?.from && !pf?.aggregationLevel) errors.push(`${what} (${pf?.key}): aggregationLevel is required.`);
    if (!pf?.from && (!Array.isArray(pf?.characteristics) || !pf.characteristics.length)) errors.push(`${what} (${pf?.key}): characteristics must list the characteristics of the aggregation level.`);
    for (const c of Object.keys(pf?.filter ?? {})) {
      if (!pf.characteristics?.includes(c)) errors.push(`${what} (${pf.key}): filter on ${c}, which is not in characteristics.`);
    }
    (pf?.variableCells ?? []).forEach((v, j) => {
      if (!splitCellRef(v.cell ?? '')) errors.push(`${what}.variableCells[${j}] (${v.name}): cell must name sheet and cell, e.g. Planning!$C$9 or 'Plan Sheet'!$C$9.`);
    });
  });

  const modules = s.vbaModules ?? [];
  const moduleNames = new Set<string>();
  modules.forEach((m, i) => {
    const what = `vbaModules[${i}]`;
    if (typeof m?.name !== 'string' || !VBA_NAME.test(m.name)) errors.push(`${what}: name must be a VBA identifier (letter first, up to 31 letters, digits, underscores), got ${JSON.stringify(m?.name)}.`);
    else if (RESERVED_MODULES.test(m.name)) errors.push(`${what}: ${m.name} is reserved for a document module.`);
    else if (moduleNames.has(m.name.toUpperCase())) errors.push(`${what}: module ${m.name} appears twice.`);
    else moduleNames.add(m.name.toUpperCase());
    if ((m?.code === undefined) === (m?.file === undefined)) errors.push(`${what} (${m?.name}): give either code or file.`);
  });
  for (const [target, code] of Object.entries(s.documentModules ?? {})) {
    if (typeof code !== 'string') errors.push(`documentModules.${target}: code must be text.`);
  }
  for (const [shape, macro] of Object.entries(s.shapeMacros ?? {})) {
    const [mod, proc] = String(macro).split('.');
    if (!mod || !proc) errors.push(`shapeMacros.${shape}: macro must be Module.Procedure, got ${JSON.stringify(macro)}.`);
    else if (modules.length && !moduleNames.has(mod.toUpperCase())) errors.push(`shapeMacros.${shape}: module ${mod} is not among vbaModules.`);
  }

  if (s.agentBridge) {
    errors.push(...validateAgentBridge(s as AfoWorkbookSpec));
    const e = sheetNameError(agentSheetName(s as AfoWorkbookSpec));
    if (e) errors.push(`agentBridge.sheet: ${e}.`);
    if ([...(layout?.sheets ?? []), ...grids.map((g) => g?.sheet)].some((n) => n?.toLowerCase() === agentSheetName(s as AfoWorkbookSpec).toLowerCase())) errors.push(`agentBridge.sheet ${agentSheetName(s as AfoWorkbookSpec)} is also a layout or crosstab sheet; the bridge sheet is generated.`);
  }
  for (const sheet of layout?.sheets ?? []) {
    const e = sheetNameError(sheet);
    if (e) errors.push(`layout.sheets: ${e}.`);
  }
  const names = new Set<string>();
  (layout?.cells ?? []).forEach((c, i) => {
    const what = `layout.cells[${i}]`;
    if (typeof c?.sheet !== 'string' || !c.sheet) errors.push(`${what}: sheet is required.`);
    if (typeof c?.cell !== 'string' || !isA1(c.cell)) errors.push(`${what}: cell must be an A1 reference, got ${JSON.stringify(c?.cell)}.`);
    if (typeof c?.value !== 'string' && typeof c?.value !== 'number') errors.push(`${what}: value must be text or a number.`);
    if (c?.name !== undefined) {
      if (!DEFINED_NAME.test(c.name) || isA1(c.name) || /^SAP/i.test(c.name)) errors.push(`${what}: ${c.name} is not usable as a defined name (no cell references, nothing starting with SAP).`);
      else if (names.has(c.name.toUpperCase())) errors.push(`${what}: defined name ${c.name} appears twice.`);
      else names.add(c.name.toUpperCase());
    }
  });
  return errors;
}

/** Sheets of the minimal layout: the listed ones, then every sheet the spec refers to. */
function layoutSheets(spec: BuildSpec, layout: MinimalLayout | undefined): string[] {
  const out: string[] = [];
  const add = (n: string | undefined) => { if (n && !out.some((o) => o.toLowerCase() === n.toLowerCase())) out.push(n); };
  (layout?.sheets ?? []).forEach(add);
  (spec.grids ?? []).forEach((g) => add(g?.sheet));
  (spec.formulaItems ?? []).forEach((f) => add(f?.sheet));
  (layout?.cells ?? []).forEach((c) => add(c?.sheet));
  (spec.planningFunctions ?? []).forEach((pf) => (pf.variableCells ?? []).forEach((v) => add(splitCellRef(v.cell)?.sheet)));
  if (!out.length) out.push('Sheet1');
  return out;
}

/** A plain workbook with the given sheets and cell values, as Excel would accept it. */
export function minimalWorkbook(sheets: string[], cells: CellSpec[] = []): Uint8Array {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
  const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${NS_R}/${type}" Target="${target}"/>`;
  const rels = (items: string) => `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  const sheetIndex = (name: string) => sheets.findIndex((s) => s.toLowerCase() === name.toLowerCase());

  const parts: Record<string, Uint8Array> = {};
  sheets.forEach((name, i) => {
    const rows = new Map<number, Array<{ col: number; xml: string }>>();
    for (const c of cells.filter((c) => sheetIndex(c.sheet) === i)) {
      const { row, col } = fromA1(c.cell);
      const ref = c.cell.replace(/\$/g, '').toUpperCase();
      // Text cells get the text number format, so values like 01.2025 stay text when edited.
      const xml = typeof c.value === 'number'
        ? `<c r="${ref}"><v>${c.value}</v></c>`
        : `<c r="${ref}" s="1" t="inlineStr"><is><t xml:space="preserve">${esc(c.value)}</t></is></c>`;
      const row_ = rows.get(row) ?? [];
      rows.set(row, [...row_.filter((x) => x.col !== col), { col, xml }]);
    }
    const data = [...rows.entries()].sort((a, b) => a[0] - b[0])
      .map(([r, cs]) => `<row r="${r + 1}">${cs.sort((a, b) => a.col - b.col).map((x) => x.xml).join('')}</row>`).join('');
    parts[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(`${head}<worksheet xmlns="${NS}" xmlns:r="${NS_R}"><sheetViews><sheetView workbookViewId="0"${i === 0 ? ' tabSelected="1"' : ''}/></sheetViews>`
      + `<sheetFormatPr defaultRowHeight="15"/><sheetData>${data}</sheetData><pageMargins left="0.7" right="0.7" top="0.787401575" bottom="0.787401575" header="0.3" footer="0.3"/></worksheet>`);
  });

  const names = cells.filter((c) => c.name).map((c) => {
    const sheet = sheets[sheetIndex(c.sheet)];
    const { row, col } = fromA1(c.cell);
    let letters = '';
    for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
    return `<definedName name="${esc(c.name!)}">${esc(quoteSheet(sheet))}!$${letters}$${row + 1}</definedName>`;
  }).join('');
  parts['xl/workbook.xml'] = strToU8(`${head}<workbook xmlns="${NS}" xmlns:r="${NS_R}"><workbookPr/><bookViews><workbookView/></bookViews><sheets>`
    + sheets.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')
    + `</sheets>${names ? `<definedNames>${names}</definedNames>` : ''}</workbook>`);
  parts['xl/_rels/workbook.xml.rels'] = strToU8(rels(sheets.map((_, i) => rel(`rId${i + 1}`, 'worksheet', `worksheets/sheet${i + 1}.xml`)).join('')
    + rel(`rId${sheets.length + 1}`, 'styles', 'styles.xml')));
  parts['xl/styles.xml'] = strToU8(`${head}<styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>`
    + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
  parts['_rels/.rels'] = strToU8(rels(rel('rId1', 'officeDocument', 'xl/workbook.xml')));
  parts['[Content_Types].xml'] = strToU8(`${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  return zipSync(parts);
}

/** What the builder needs to know about a template before it writes into it. */
function inspectTemplate(bytes: Uint8Array): { sheets: string[]; codeNames: string[]; hasVba: boolean; afo: boolean } {
  const parts = unzipSync(bytes, { filter: (f) => f.name === 'xl/workbook.xml' || f.name === 'xl/vbaProject.bin' || /^xl\/worksheets\/[^/]+\.xml$/.test(f.name) });
  const wb = parts['xl/workbook.xml'];
  if (!wb) throw new Error('The template is not an Excel workbook (xl/workbook.xml missing).');
  const wbXml = strFromU8(wb);
  const sheets = [...wbXml.matchAll(/<sheet [^>]*name="([^"]+)"/g)].map((m) => m[1].replace(/&amp;/g, '&').replace(/&apos;/g, "'").replace(/&quot;/g, '"'));
  const codeNames = Object.entries(parts).filter(([n]) => n.startsWith('xl/worksheets/'))
    .map(([, b]) => /<sheetPr[^>]*codeName="([^"]+)"/.exec(strFromU8(b))?.[1]).filter((c): c is string => !!c);
  const wbCode = /<workbookPr[^>]*codeName="([^"]+)"/.exec(wbXml)?.[1];
  return { sheets, codeNames: wbCode ? [wbCode, ...codeNames] : codeNames, hasVba: !!parts['xl/vbaProject.bin'], afo: sheets.includes(HIDDEN_SHEET) };
}

/** Check the spec, write the workbook and read it back. Throws with every problem found. */
export function buildWorkbook(input: BuildSpec, options: BuildOptions): BuildResult {
  // Entries taken over from a source workbook are completed from it; the caller's object stays as it is.
  const spec: BuildSpec = structuredClone(input);
  const output = resolve(options.output);
  const ext = extname(output).toLowerCase();
  if (ext !== '.xlsx' && ext !== '.xlsm') throw new Error(`output must end in .xlsx or .xlsm, got ${options.output}.`);
  if (!existsSync(dirname(output))) throw new Error(`The output folder does not exist: ${dirname(output)}`);
  if (existsSync(output) && !options.overwrite) throw new Error(`The output file exists: ${output}. Pass overwrite: true to replace it.`);
  if (options.template && options.layout) throw new Error('Give either template or layout, not both: layout describes the generated minimal workbook.');

  const sourceErrors: string[] = [];
  if (spec.source) {
    const file = isAbsolute(spec.source) ? spec.source : resolve(options.baseDir ?? '', spec.source);
    if (!isAbsolute(spec.source) && !options.baseDir) sourceErrors.push(`source is relative (${spec.source}); give an absolute path or pass the spec as a file.`);
    else if (!existsSync(file)) sourceErrors.push(`The source workbook does not exist: ${file}`);
    else {
      try {
        const definition = readSourceDefinition(file);
        sourceErrors.push(...applySource(spec as AfoWorkbookSpec, definition).errors);
        spec.sourceDefinition = definition;
      } catch (err) {
        sourceErrors.push((err as Error).message);
      }
    }
  }
  const errors = [...sourceErrors, ...validateSpec(spec, options.layout)];
  const hasVba = !!spec.vbaModules?.length || Object.keys(spec.documentModules ?? {}).length > 0 || !!spec.agentBridge;
  const warnings: string[] = [];
  let base: Uint8Array;
  let baseName: string;
  if (options.template) {
    const template = resolve(options.template);
    if (template.toLowerCase() === output.toLowerCase()) throw new Error('template and output must be different files.');
    if (!existsSync(template)) throw new Error(`The template does not exist: ${template}`);
    base = readFileSync(template);
    baseName = template;
    const t = inspectTemplate(base);
    if (t.afo) errors.push('The template is already an Analysis for Office workbook. Use a plain layout workbook; the builder writes the complete AfO definition.');
    const missing = [...new Set([...(spec.grids ?? []), ...(spec.formulaItems ?? []), ...(spec.cells ?? [])].map((g) => g.sheet))].filter((s) => !t.sheets.includes(s));
    for (const pf of spec.planningFunctions ?? []) for (const v of pf.variableCells ?? []) {
      const ref = splitCellRef(v.cell ?? '');
      if (ref && !t.sheets.includes(ref.sheet)) warnings.push(`${pf.key}: variable ${v.name} reads ${v.cell}, but the template has no sheet ${ref.sheet}.`);
    }
    if (spec.agentBridge && t.sheets.some((n) => n.toLowerCase() === agentSheetName(spec as AfoWorkbookSpec).toLowerCase())) errors.push(`The template already has a sheet ${agentSheetName(spec as AfoWorkbookSpec)}; set agentBridge.sheet to another name.`);
    if (missing.length) errors.push(`Sheet(s) not in the template: ${missing.join(', ')}. Template sheets: ${t.sheets.join(', ')}.`);
    const clash = (spec.vbaModules ?? []).filter((m) => t.codeNames.some((c) => c.toLowerCase() === m.name?.toLowerCase()));
    if (clash.length) errors.push(`VBA module name(s) used as code name of a sheet or the workbook in the template: ${clash.map((m) => m.name).join(', ')}.`);
    const docMissing = Object.keys(spec.documentModules ?? {}).filter((n) => n !== 'ThisWorkbook' && !t.sheets.includes(n));
    if (docMissing.length) errors.push(`documentModules for sheet(s) not in the template: ${docMissing.join(', ')}.`);
    if (t.hasVba && hasVba) warnings.push('The template has a VBA project; it is replaced by the modules in the spec. Its macros are gone unless the spec lists them again.');
    if (t.hasVba && !hasVba && ext === '.xlsx') errors.push('The template has macros; write the output as .xlsm.');
  } else {
    if (Object.keys(spec.shapeMacros ?? {}).length) errors.push('shapeMacros needs a template: the generated minimal layout has no shapes to click.');
    const sheets = layoutSheets(spec, options.layout);
    for (const sheet of sheets) {
      const e = sheetNameError(sheet);
      if (e && !options.layout?.sheets?.includes(sheet)) errors.push(`${e}.`);
    }
    base = errors.length ? new Uint8Array() : minimalWorkbook(sheets, options.layout?.cells);
    baseName = 'minimal layout';
  }
  if (hasVba && ext === '.xlsx') errors.push('VBA modules need an .xlsm output.');
  if (errors.length) throw new Error(`The workbook was not written:\n- ${errors.join('\n- ')}`);

  const vbaModules = (spec.vbaModules ?? []).map((m) => {
    if (m.code !== undefined) return { name: m.name, code: m.code };
    if (!isAbsolute(m.file!) && !options.baseDir) throw new Error(`VBA module file for ${m.name} is relative (${m.file}); give an absolute path or pass the spec as a file.`);
    const file = isAbsolute(m.file!) ? m.file! : resolve(options.baseDir!, m.file!);
    if (!existsSync(file)) throw new Error(`VBA module file not found for ${m.name}: ${file}`);
    return { name: m.name, code: readFileSync(file, 'utf8').replace(/^﻿/, '') };
  });
  const bytes = makeAfoWorkbook(base, { ...spec, vbaModules }, { macroEnabled: ext === '.xlsm' });
  writeFileSync(output, bytes);

  const d = readWorkbook(output, { sections: ['sheets', 'items', 'macros'] });
  return {
    output,
    sizeBytes: statSync(output).size,
    base: baseName,
    sheets: (d.sheets ?? []).filter((s) => s.state !== 'veryHidden').map((s) => s.name),
    dataSources: d.summary.dataSources,
    planningObjects: d.summary.planningObjects,
    crosstabs: (d.items?.grids ?? []).map((g) => `${g.key} (${g.dataProviderKey}) at ${g.sheet}!${g.address}`),
    buttons: (d.macros?.buttons ?? []).map((b) => `${b.name} → ${b.procedure ?? `${b.macro} (not found)`}`),
    vbaModules: (d.macros?.modules ?? []).filter((m) => m.type === 'standard').map((m) => m.name),
    warnings,
    issues: d.issues,
  };
}
