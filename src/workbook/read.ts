// Read one workbook into a single description: AfO definition, sheets, SAP formulas, buttons
// and macros, linked to each other, with the broken links between them listed as issues.

import { statSync } from 'node:fs';
import { basename } from 'node:path';
import { type AfoDefinition, type DataProvider, type FormulaItem, type GridItem, type OtherItem, type PlanningObject, isAfoApplication, parseAfoDefinition } from './afo.js';
import { WorkbookPackage } from './package.js';
import { type Button, type SapFormula, literal, readStructure } from './sheets.js';
import { type VbaModule, readVbaProject } from './vba.js';

export const SECTIONS = ['settings', 'sheets', 'dataProviders', 'planningObjects', 'items', 'formulas', 'macros'] as const;
export type Section = (typeof SECTIONS)[number];

export interface ReadOptions {
  sections?: readonly Section[];
  /** Include the VBA source of every module. */
  macroSource?: boolean;
  /** Cap on SAP formulas listed per sheet; the count is always complete. */
  maxFormulasPerSheet?: number;
}

/** SAP functions whose first argument is a data source alias. */
const DATA_SOURCE_FUNCTIONS = new Set([
  'SAPGetData', 'SAPGetDimensionDynamicFilter', 'SAPGetDimensionEffectiveFilter', 'SAPGetDimensionInfo',
  'SAPGetDimensionStaticFilter', 'SAPGetDisplayedMeasures', 'SAPGetMeasureFilter', 'SAPGetMember',
  'SAPGetSourceInfo', 'SAPGetUniformScaling', 'SAPGetVariable', 'SAPListOfDimensions',
  'SAPListOfDynamicFilters', 'SAPListOfEffectiveFilters', 'SAPListOfMembers', 'SAPListOfVariables',
  'SAPSetFilterComponent', 'SAPSetFilter', 'SAPCallMemberSelector',
]);
/** SAP functions whose first argument is a planning object alias. */
const PLANNING_FUNCTIONS = new Set(['SAPExecutePlanningFunction', 'SAPExecutePlanningSequence', 'SAPSetPlanParameter']);

export interface ButtonInfo extends Button {
  sheet: string;
  /** `Module.Procedure` when the macro resolves to a procedure in the VBA project. */
  procedure?: string;
}

export interface SheetInfo {
  name: string;
  codeName?: string;
  state?: string;
  grids: string[];
  formulaItems: number;
  sapFormulaCells: number;
  buttons: number;
}

export interface WorkbookDescription {
  file: string;
  name: string;
  sizeBytes: number;
  modified: string;
  afo: boolean;
  summary: {
    systemId?: string;
    afoRevision?: string;
    lastRefreshed?: string;
    sheets: number;
    dataSources: string[];
    planningObjects: string[];
    grids: number;
    formulaItems: number;
    sapFormulaCells: number;
    vbaModules: number;
    procedures: number;
    afoApiCalls: number;
    buttons: number;
    issues: number;
  };
  settings?: Omit<AfoDefinition, 'dataProviders' | 'planningObjects' | 'grids' | 'formulaItems' | 'otherItems'>;
  sheets?: SheetInfo[];
  dataProviders?: DataProvider[];
  planningObjects?: PlanningObject[];
  items?: { grids: Array<GridItem & { sheet?: string }>; formulaItems: Array<FormulaItem & { sheet?: string }>; other: Array<OtherItem & { sheet?: string }> };
  formulas?: Array<{ sheet: string; cells: number; listed: number; formulas: SapFormula[] }>;
  macros?: { modules: VbaModule[]; buttons: ButtonInfo[]; error?: string };
  externalLinks?: string[];
  issues: string[];
}

export function readWorkbook(file: string, options: ReadOptions = {}): WorkbookDescription {
  const sections = new Set<Section>(options.sections ?? SECTIONS);
  const maxFormulas = options.maxFormulasPerSheet ?? 200;
  const stat = statSync(file);
  const pkg = WorkbookPackage.open(file);
  const structure = readStructure(pkg, { formulas: true, buttons: true });
  const issues: string[] = [];

  // The AfO definition: the custom property the hidden sheet names, else any part that holds it.
  let def: AfoDefinition | undefined;
  const parts = [
    ...structure.sheets.map((s) => s.afoDefinitionPart).filter((p): p is string => !!p),
    ...pkg.names.filter((n) => /^xl\/customProperty\d+\.bin$/.test(n)),
  ];
  for (const part of new Set(parts)) {
    const xml = pkg.text(part);
    if (xml && isAfoApplication(xml)) {
      def = parseAfoDefinition(xml);
      break;
    }
  }

  const sheetName = new Map<string, string>();
  for (const s of structure.sheets) if (s.afoId) sheetName.set(s.afoId, s.name);
  const resolveSheet = (id?: string) => (id ? sheetName.get(id) ?? id.split('|')[0] : undefined);

  const vba = structure.hasVba ? readVbaProject(pkg.bytes('xl/vbaProject.bin')!, options.macroSource ?? false) : undefined;
  if (vba?.error) issues.push(`VBA: ${vba.error}`);

  // Procedures by plain name and by Module.Name, for button links.
  const procedures = new Map<string, string>();
  for (const m of vba?.modules ?? []) {
    for (const p of m.procedures) {
      const qualified = `${m.name}.${p.name}`;
      procedures.set(qualified.toLowerCase(), qualified);
      if (!procedures.has(p.name.toLowerCase())) procedures.set(p.name.toLowerCase(), qualified);
    }
  }
  const buttons: ButtonInfo[] = [];
  for (const s of structure.sheets) {
    for (const b of s.buttons) {
      const procedure = procedures.get(b.macro.toLowerCase());
      buttons.push({ sheet: s.name, ...b, procedure });
      if (!procedure && vba) issues.push(`Button ${b.name ?? b.text ?? '?'} on sheet ${s.name} calls macro ${b.macro}, which is not in the VBA project.`);
      if (!vba) issues.push(`Button ${b.name ?? b.text ?? '?'} on sheet ${s.name} calls macro ${b.macro}, but the workbook has no VBA project.`);
    }
  }

  if (def) {
    const dpKeys = new Set(def.dataProviders.map((d) => d.key));
    const poKeys = new Set(def.planningObjects.map((p) => p.key));
    for (const g of def.grids) {
      if (g.dataProviderKey && !dpKeys.has(g.dataProviderKey)) issues.push(`Crosstab ${g.key} uses data source ${g.dataProviderKey}, which is not defined.`);
    }
    for (const p of def.planningObjects) {
      if (p.dataProviderKey && !dpKeys.has(p.dataProviderKey)) issues.push(`Planning object ${p.key} is bound to data source ${p.dataProviderKey}, which is not defined.`);
      if (p.error) issues.push(`Planning object ${p.key}: ${p.error}`);
    }
    for (const d of def.dataProviders) if (d.stateError) issues.push(`Data source ${d.key}: state unreadable (${d.stateError})`);
    const checkAlias = (where: string, fn: string, alias: string | undefined) => {
      if (!alias) return;
      if (DATA_SOURCE_FUNCTIONS.has(fn) && !dpKeys.has(alias)) issues.push(`${where}: ${fn} refers to data source ${alias}, which is not defined.`);
      if (PLANNING_FUNCTIONS.has(fn) && !poKeys.has(alias)) issues.push(`${where}: ${fn} refers to planning object ${alias}, which is not defined.`);
    };
    for (const s of structure.sheets) {
      for (const f of s.sapFormulas) for (const c of f.calls) checkAlias(`${s.name}!${f.cell}`, c.function, literal(c.args[0]));
    }
    for (const f of def.formulaItems) {
      for (const m of (f.formula ?? '').matchAll(/(SAP\w+)\(\s*"([^"]*)"/g)) checkAlias(`Formula item ${f.key}`, m[1], m[2]);
    }
    for (const m of vba?.modules ?? []) {
      for (const p of m.procedures) for (const c of p.afoCalls) checkAlias(`${m.name}.${p.name} line ${c.line}`, c.function, c.firstArgument);
    }
    if (def.systems.length > 1) issues.push(`The workbook connects to more than one system: ${def.systems.map((s) => s.systemId).join(', ')}.`);
  } else if (structure.sheets.some((s) => s.sapFormulas.length)) {
    issues.push('SAP formulas found, but no AfO definition: the workbook was not saved by Analysis for Office, or the definition was removed.');
  }

  const sapCells = structure.sheets.reduce((n, s) => n + s.sapFormulaCells, 0);
  const allProcedures = (vba?.modules ?? []).flatMap((m) => m.procedures);
  const out: WorkbookDescription = {
    file,
    name: basename(file),
    sizeBytes: stat.size,
    modified: stat.mtime.toISOString(),
    afo: !!def,
    summary: {
      systemId: def?.systemId,
      afoRevision: def?.afoRevision,
      lastRefreshed: def?.lastRefreshed,
      sheets: structure.sheets.length,
      dataSources: def?.dataProviders.map((d) => `${d.key}=${d.dataSource ?? '?'}${d.type && d.type !== 'Query' ? ` (${d.type})` : ''}`) ?? [],
      planningObjects: def?.planningObjects.map((p) => `${p.key}=${p.name ?? '?'} (${p.kind}${p.dataProviderKey ? ` on ${p.dataProviderKey}` : ''})`) ?? [],
      grids: def?.grids.length ?? 0,
      formulaItems: def?.formulaItems.length ?? 0,
      sapFormulaCells: sapCells,
      vbaModules: vba?.modules.length ?? 0,
      procedures: allProcedures.length,
      afoApiCalls: allProcedures.reduce((n, p) => n + p.afoCalls.length, 0),
      buttons: buttons.length,
      issues: issues.length,
    },
    issues,
  };

  if (def && sections.has('settings')) {
    const { dataProviders: _d, planningObjects: _p, grids: _g, formulaItems: _f, otherItems: _o, ...settings } = def;
    out.settings = settings;
  }
  if (sections.has('sheets')) {
    out.sheets = structure.sheets.map((s) => ({
      name: s.name,
      codeName: s.codeName,
      state: s.state,
      grids: def?.grids.filter((g) => resolveSheet(g.sheetId) === s.name).map((g) => g.key) ?? [],
      formulaItems: def?.formulaItems.filter((f) => resolveSheet(f.sheetId) === s.name).length ?? 0,
      sapFormulaCells: s.sapFormulaCells,
      buttons: s.buttons.length,
    }));
  }
  if (def && sections.has('dataProviders')) out.dataProviders = def.dataProviders;
  if (def && sections.has('planningObjects')) out.planningObjects = def.planningObjects;
  if (def && sections.has('items')) {
    const withSheet = <T extends { sheetId?: string }>(i: T) => {
      const { sheetId, ...rest } = i;
      return { ...rest, sheet: resolveSheet(sheetId) };
    };
    out.items = { grids: def.grids.map(withSheet), formulaItems: def.formulaItems.map(withSheet), other: def.otherItems.map(withSheet) };
  }
  if (sections.has('formulas')) {
    out.formulas = structure.sheets
      .filter((s) => s.sapFormulas.length)
      .map((s) => ({ sheet: s.name, cells: s.sapFormulaCells, listed: Math.min(s.sapFormulas.length, maxFormulas), formulas: s.sapFormulas.slice(0, maxFormulas) }));
  }
  if (sections.has('macros') && (vba || buttons.length)) out.macros = { modules: vba?.modules ?? [], buttons, error: vba?.error };
  if (structure.externalLinks.length) out.externalLinks = structure.externalLinks;
  return out;
}
