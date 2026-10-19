// The AfO workbook definition: the `<Application>` document that Analysis for Office stores as
// a custom property of its hidden sheet (customPr `serializedData2`). It lists data sources,
// planning objects, crosstabs and formula items, plus workbook settings.

import { type BicsView, type Variable, expandBlock, parseBicsView, parseVariables } from './bics.js';
import { type XmlNode, child, children, parseXml, path, text } from './xml.js';

export const AFO_NAMESPACE = 'http://www.sap.com/ip/bi/pioneer/excel/application';

export interface SystemRef {
  systemId?: string;
  systemName?: string;
  systemType?: string;
  client?: string;
  language?: string;
}

export interface DataProvider {
  key: string;
  dataSource?: string;
  dataSourceName?: string;
  type?: string;
  systemId?: string;
  /** False when AfO saved no navigation state: the data source opens in the query's default view. */
  stateStored: boolean;
  state?: BicsView;
  inputReady?: boolean;
  resetOnOpen?: boolean;
  stateError?: string;
}

export interface Reference {
  technicalName: string;
  /** `Variable` or `Filter`. */
  object: string;
  /** Where the value comes from: `Filter` (the bound data source), `Cell`, … */
  referenceType?: string;
  cell?: string;
}

export interface PlanningObject {
  key: string;
  /** `function`, `sequence`, … from the element name `Planning<Kind>Usage`. */
  kind: string;
  name?: string;
  description?: string;
  planningType?: string;
  dataProviderKey?: string;
  systemId?: string;
  variables: Variable[];
  /** Aggregation level the filter is defined on. */
  aggregationLevel?: string;
  /** Planning filter name, or `!!A<aggregation level>` for an ad-hoc filter. */
  filterObject?: string;
  filter: Array<{ characteristic: string; selections: string[] }>;
  references: Reference[];
  error?: string;
}

export interface SheetRange {
  /** AfO sheet identifier `<code name>|<guid>`, resolved to the sheet name by the workbook reader. */
  sheetId?: string;
  address?: string;
}

export interface GridItem extends SheetRange {
  key: string;
  description?: string;
  dataProviderKey?: string;
}

export interface FormulaItem extends SheetRange {
  key: string;
  formula?: string;
  labelFormula?: string;
}

export interface OtherItem extends SheetRange {
  type: string;
  key?: string;
  description?: string;
  dataProviderKey?: string;
}

export interface AfoDefinition {
  version?: string;
  afoRevision?: string;
  systemId?: string;
  lastRefreshed?: string;
  settings: Record<string, string>;
  promptVariables: Array<{ name: string; visible: boolean }>;
  systems: SystemRef[];
  planningModels: Array<{ name?: string; description?: string }>;
  dataProviders: DataProvider[];
  planningObjects: PlanningObject[];
  grids: GridItem[];
  formulaItems: FormulaItem[];
  otherItems: OtherItem[];
}

export function isAfoApplication(xml: string): boolean {
  return xml.trimStart().startsWith('<Application') && xml.includes(AFO_NAMESPACE);
}

/** Zero-based row and column → `B15`. */
export function toA1(row: number, col: number): string {
  let c = col + 1;
  let letters = '';
  while (c > 0) {
    const r = (c - 1) % 26;
    letters = String.fromCharCode(65 + r) + letters;
    c = Math.floor((c - 1) / 26);
  }
  return `${letters}${row + 1}`;
}

function rangeOf(item: XmlNode): SheetRange {
  const ranges = children(item, 'Range');
  const range = ranges.find((r) => r.attrs.Name === 'RangeOccupied') ?? ranges.find((r) => r.attrs.Name === 'RequestedRange') ?? ranges[0];
  if (!range) return {};
  const cell = (name: string) => {
    const c = children(range, 'Cell').find((x) => x.attrs.Name === name);
    const row = Number(text(c, 'Row'));
    const col = Number(text(c, 'Col'));
    return Number.isFinite(row) && Number.isFinite(col) && c ? toA1(row, col) : undefined;
  };
  const tl = cell('TopLeft');
  const br = cell('BottomRight');
  return { sheetId: text(range, 'SheetName'), address: tl && br && tl !== br ? `${tl}:${br}` : tl };
}

function systemRef(info: XmlNode | undefined): SystemRef | undefined {
  if (!info) return undefined;
  return {
    systemId: text(info, 'SystemId'),
    systemName: text(info, 'SystemName'),
    systemType: text(info, 'SystemType'),
    client: text(info, 'Client'),
    language: text(info, 'Language'),
  };
}

const TOP_LEVEL_FIELDS = new Set(['Version', 'AAO_Revision', 'SystemId', 'LastRefreshed', 'Cleaned']);

function parseDataProvider(dp: XmlNode): DataProvider {
  const info = child(dp, 'DataSourceInfo');
  const out: DataProvider = {
    key: text(dp, 'Key') ?? '',
    dataSource: text(info, 'Id'),
    dataSourceName: text(info, 'Name'),
    type: text(info, 'Type'),
    systemId: text(info, 'SystemInfo', 'SystemId'),
    stateStored: false,
  };
  const entry = text(dp, 'RefreshDataEntryEnabled');
  if (entry) out.inputReady = entry === 'True';
  const reset = text(dp, 'ResetDataSource');
  if (reset) out.resetOnOpen = reset === 'True';
  const header = child(dp, 'BICS_VIEW_HEADER');
  try {
    const view = header?.attrs.compressed === 'true' ? child(expandBlock(header), 'BICS_VIEW') : child(header, 'BICS_VIEW');
    if (view) {
      out.stateStored = true;
      out.state = parseBicsView(view);
    }
  } catch (err) {
    out.stateError = (err as Error).message;
  }
  return out;
}

function parsePlanningObject(node: XmlNode): PlanningObject {
  const kind = node.name.replace(/^Planning/, '').replace(/Usage$/, '').toLowerCase();
  const info = child(node, 'PlanningObjectInfo');
  const out: PlanningObject = {
    key: text(node, 'Key') ?? '',
    kind,
    name: text(info, 'Name'),
    description: text(info, 'Description'),
    planningType: text(info, 'PlanningType'),
    dataProviderKey: text(node, 'DataProviderKey'),
    systemId: text(info, 'SystemInfo', 'SystemId'),
    variables: [],
    filter: [],
    references: [],
  };
  try {
    const vars = expandBlock(path(node, 'PlanFunctionVariables', 'Variables_Header'))
      ?? expandBlock(path(node, 'PlanSequenceVariables', 'Variables_Header'));
    out.variables = parseVariables(child(vars, 'VARIABLES'));
    const filterHeader = path(node, 'PlanFilter', 'BICS_VIEW_HEADER');
    const view = child(expandBlock(filterHeader), 'BICS_VIEW');
    if (view) {
      const state = parseBicsView(view);
      out.aggregationLevel = state.infoProvider;
      out.filterObject = state.selectionObject;
      out.filter = state.filters;
    }
  } catch (err) {
    out.error = (err as Error).message;
  }
  for (const ref of child(node, 'ReferenceConfiguration')?.children ?? []) {
    const technicalName = text(ref, 'TechnicalName');
    if (!technicalName) continue;
    const r: Reference = {
      technicalName,
      object: ref.name.replace(/RefConfig$/, ''),
      referenceType: text(ref, 'ReferenceType'),
    };
    const cell = text(ref, 'CellAdress') ?? text(ref, 'CellAddress');
    if (cell) r.cell = cell;
    out.references.push(r);
  }
  return out;
}

export function parseAfoDefinition(xml: string): AfoDefinition {
  const doc = parseXml(xml);
  const app = child(doc, 'Application');
  const root = child(app, 'com.sap.ip.bi.pioneer') ?? app;
  const def: AfoDefinition = {
    version: text(root, 'Version'),
    afoRevision: text(root, 'AAO_Revision'),
    systemId: text(root, 'SystemId'),
    lastRefreshed: text(root, 'LastRefreshed'),
    settings: {},
    promptVariables: [],
    systems: [],
    planningModels: [],
    dataProviders: [],
    planningObjects: [],
    grids: [],
    formulaItems: [],
    otherItems: [],
  };
  for (const c of root?.children ?? []) {
    if (c.children.length === 0 && !TOP_LEVEL_FIELDS.has(c.name) && c.text.trim()) def.settings[c.name] = c.text.trim();
  }
  const cellLocking = text(root, 'PlanningSystemDesignTime', 'CellLocking');
  if (cellLocking) def.settings.CellLocking = cellLocking;

  def.promptVariables = children(child(root, 'PromptsVisibility'), 'Variable').map((v) => ({
    name: v.attrs.name ?? '',
    visible: v.attrs.visible !== 'False',
  }));

  const systems = new Map<string, SystemRef>();
  const addSystem = (info: XmlNode | undefined) => {
    const s = systemRef(info);
    if (s?.systemId && !systems.has(s.systemId)) systems.set(s.systemId, s);
  };
  const design = child(root, 'PlanningSystemDesignTime');
  addSystem(child(design, 'SystemInfo'));
  for (const pm of children(design, 'PlanningModel')) {
    const info = child(pm, 'PlanningObjectInfo');
    def.planningModels.push({ name: text(info, 'Name')?.replace(/\s+/g, ' '), description: text(info, 'Description') });
  }

  for (const dp of children(child(root, 'DataProviders'), 'DataProvider')) {
    addSystem(path(dp, 'DataSourceInfo', 'SystemInfo'));
    def.dataProviders.push(parseDataProvider(dp));
  }
  for (const po of child(root, 'PlanningObjects')?.children ?? []) {
    if (!/^Planning\w+Usage$/.test(po.name)) continue;
    addSystem(path(po, 'PlanningObjectInfo', 'SystemInfo'));
    def.planningObjects.push(parsePlanningObject(po));
  }
  def.systems = [...systems.values()];

  for (const item of child(root, 'Items')?.children ?? []) {
    const key = text(item, 'Key');
    if (item.name === 'Grid') {
      def.grids.push({ key: key ?? '', description: text(item, 'Description'), dataProviderKey: text(item, 'DataProviderKey'), ...rangeOf(item) });
    } else if (item.name === 'Fdd') {
      def.formulaItems.push({ key: key ?? '', formula: text(item, 'FormulaString'), labelFormula: text(item, 'LabelRangeFormula'), ...rangeOf(item) });
    } else {
      def.otherItems.push({ type: item.name, key, description: text(item, 'Description'), dataProviderKey: text(item, 'DataProviderKey'), ...rangeOf(item) });
    }
  }
  return def;
}
