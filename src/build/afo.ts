// Turn an ordinary Excel workbook into an Analysis for Office workbook: write the AfO definition
// (data sources, crosstabs, planning functions) into a hidden sheet, give every sheet its AfO id,
// add the defined names AfO binds items to, and optionally a VBA project and button macros.
//
// AfO accepts such a workbook and rewrites the definition on its next save. Learned from
// workbooks AfO saved itself: a crosstab needs both RangeOccupied and RequestedRange, and every
// item has a defined name SAP<Key> pointing at its range.

import { randomUUID } from 'node:crypto';
import { gzipSync, strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { AFO_NAMESPACE, toA1 } from '../workbook/afo.js';
import { decodeCompressed } from '../workbook/bics.js';
import { decodeText, relsPartOf, resolveTarget } from '../workbook/package.js';
import { AGENT_MODULE, AGENT_NAMES, type AgentBridgeSpec, agentModuleCode, agentSheetName, agentSheetXml, agentThisWorkbookCode } from './agent.js';
import { writeCells } from './cells.js';
import { firstElement, keyedElement, replaceFirstKey, sourceSystemInfo } from './source.js';
import { writeVbaProject, type VbaModuleSource } from './vba.js';

export interface AfoSystem {
  systemId: string;
  /** The connection name as in the SAP Logon landscape, e.g. "XYZ - BW/4HANA". */
  systemName: string;
  client: string;
  language: string;
  /** Application server host (the landscape's server without the port). */
  server?: string;
  systemNumber?: string;
  /** The landscape entry's uuid. */
  uuid?: string;
  /** SAProuter string, e.g. /H/router.example.com/S/3299. */
  sapRouter?: string;
  /** SNC partner name; with it AfO connects through SNC. */
  sncPartnerName?: string;
  /** SNC quality of protection (default 0 with a partner name, otherwise -1 = no SNC). */
  sncQuality?: number;
}

/** A stored prompt value: a member (internal key), a literal, or an interval of members. */
export interface PromptValue {
  name: string;
  member?: string;
  literal?: string;
  low?: string;
  high?: string;
}

export interface DataProviderSpec {
  key: string;
  /**
   * Take this data source over from the source workbook (its key there), with its whole stored
   * state: filters, hierarchies, display settings, design rules, prompts. Other fields are ignored
   * except key, which may rename it, prompts and inputReady.
   */
  from?: string;
  query: string;
  /** The query's InfoProvider, written into the stored state. */
  infoProvider?: string;
  inputReady?: boolean;
  /** Variable values stored with the data source; with them AfO opens without the prompt dialog. */
  prompts?: PromptValue[];
  /**
   * Stored axis layout. Characteristics by technical name; a key figure structure by its
   * runtime UID (as the reporting metadata reports it). Omitted axes keep the query default.
   */
  axes?: { rows?: string[]; columns?: string[]; free?: string[] };
  /** Stored filters on characteristics, internal keys; several entries for one characteristic combine. */
  filters?: FilterSpec[];
  /**
   * Formatting rules (Analysis Design, Format Cells): a cell style per structure member, as AfO
   * stores them in the data source state. Keys are BW element UIDs.
   */
  designRules?: DesignRulesSpec;
}

/** Cell styles for members of one structure, applied to the member and its data cells. */
export interface DesignRulesSpec {
  /** UID of the structure the members belong to. */
  structure: string;
  /** Own styles by name; colours as #RRGGBB, font defaults to #2B373D. */
  styles?: Array<{ name: string; background: string; font?: string }>;
  /**
   * Structure member UID and its style: one of `styles`, or an AfO predefined style such as
   * EXCEL_GOOD, EXCEL_BAD, EXCEL_NEUTRAL (`predefined: true`). `leaf: true` for a member without
   * children (AfO addresses it as a leaf; a member with children by the structure).
   */
  members: Array<{ member: string; style: string; predefined?: boolean; leaf?: boolean }>;
  /**
   * Cells the rules cover, as AfO writes it: BOTH_AREAS (member and data cells, default),
   * LAST_AREA (data cells only), FIRST_AREA (member cell only).
   */
  area?: 'BOTH_AREAS' | 'LAST_AREA' | 'FIRST_AREA';
}

/** A filter on one characteristic: members, an interval, or a comparison, included or excluded. */
export interface FilterSpec {
  characteristic: string;
  members?: string[];
  low?: string;
  high?: string;
  operator?: 'GREATER' | 'GREATER_EQUAL' | 'LESS' | 'LESS_EQUAL';
  value?: string;
  exclude?: boolean;
}

export interface GridSpec {
  key: string;
  /** Take the crosstab settings over from this crosstab of the source workbook; only the position is new. */
  from?: string;
  dataProviderKey: string;
  sheet: string;
  /** Top-left cell, A1 notation. */
  cell: string;
  /**
   * AfO adjusts the column widths to the crosstab on every refresh (default true). Off keeps the
   * widths of the layout, so cards and buttons above the crosstab stay in place.
   */
  autoFit?: boolean;
}

/** An AfO formula item, e.g. a filter component (=SAPSetFilterComponent(...)) in a cell. */
export interface FormulaItemSpec {
  key: string;
  /** Take the item over from the source workbook (its key there); its formula follows renamed data sources. */
  from?: string;
  sheet: string;
  cell: string;
  /** The formula, e.g. =SAPSetFilterComponent("DS_1", "0CALMONTH", ""). Required unless from. */
  formula?: string;
}

export interface PlanningFunctionSpec {
  key: string;
  /** Take this planning function or sequence over from the source workbook, with variables, filter and references. */
  from?: string;
  name: string;
  description?: string;
  dataProviderKey: string;
  aggregationLevel: string;
  /** Characteristics of the aggregation level, listed in the ad-hoc filter. */
  characteristics: string[];
  /** Fixed filter values per characteristic (internal keys); characteristics not listed stay unrestricted. */
  filter?: Record<string, string[]>;
  /** Stored variable values, internal format. */
  variables?: Array<{ name: string; value: string }>;
  /** Variables read from a cell at execution time (`Sheet!$C$7`, value in external format). */
  variableCells?: Array<{ name: string; cell: string }>;
  /** Process changed data only (AfO: "Changed data" / delta mode): the function gets only records changed in the planning buffer. */
  deltaMode?: boolean;
  /**
   * Characteristics whose filter the function takes from its data source at execution time, so a
   * filter set on the crosstab (e.g. a few meter points) restricts the function too. They must be
   * in `characteristics`.
   */
  filterFromDataSource?: string[];
}

export interface AfoWorkbookSpec {
  /** Required unless a source workbook is given; then the source's connection (SAProuter, SNC) is used. */
  system?: AfoSystem;
  /**
   * Workbook to take data sources, planning objects, crosstab settings, the system connection and
   * the workbook settings from (entries with `from`). Settings the spec sets win.
   */
  source?: string;
  /** The source workbook's AfO definition, loaded by the build step. */
  sourceDefinition?: string;
  refreshOnOpen?: boolean;
  /** AfO "remove data before saving" (stored as Cleaned). Off keeps the last result in the saved file. */
  removeDataOnSave?: boolean;
  changeModeOnOpen?: boolean;
  storePrompts?: boolean;
  /** Share variables of the same name between data sources and planning objects (AfO default: off). */
  mergeVariables?: boolean;
  /**
   * AfO style set the workbook names (e.g. "SAP Blue"). Crosstabs are drawn with the workbook's
   * SAP* cell styles; a layout template can bring its own and name them here.
   */
  styleSetName?: string;
  afoRevision?: string;
  dataProviders: DataProviderSpec[];
  grids: GridSpec[];
  /**
   * Values and formulas to write into template cells, e.g. =SAPGetVariable(...) in a header.
   * SAP formulas are stored the way AfO stores them (`_xll.`); the cell keeps its style.
   */
  cells?: Array<{ sheet: string; cell: string; formula?: string; value?: string | number }>;
  /** Formula items (filter components); the builder writes the formula into the cell too. */
  formulaItems?: FormulaItemSpec[];
  planningFunctions?: PlanningFunctionSpec[];
  /** Macro per shape name, written as the shape's click action. */
  shapeMacros?: Record<string, string>;
  /** Standard modules of a VBA project to add; turns the workbook into .xlsm. */
  vbaModules?: Array<{ name: string; code: string }>;
  /**
   * Code of document modules, by visible sheet name or `ThisWorkbook`: event handlers such as
   * Worksheet_Change or Workbook_SheetChange. Turns the workbook into .xlsm like vbaModules.
   */
  documentModules?: Record<string, string>;
  /**
   * Command cell for agents such as Claude in Excel, whose cell writes Analysis for Office does
   * not see: a bridge sheet plus VBA that runs SET, PLAN, VAR, SAVE ... (see agent.ts).
   */
  agentBridge?: AgentBridgeSpec;
}

export const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
export const guid = () => randomUUID().replace(/-/g, '');
const tf = (b: boolean | undefined) => (b ? 'True' : 'False');

/** A1 → zero-based row and column. */
export function fromA1(cell: string): { row: number; col: number } {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(cell.toUpperCase());
  if (!m) throw new Error(`Not an A1 cell reference: ${cell}`);
  const col = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  return { row: Number(m[2]) - 1, col };
}

/** An AfO compressed block: 4-byte little-endian length, gzip, base64. */
export function compressBlock(xml: string): string {
  const raw = strToU8(xml);
  const len = Buffer.alloc(4);
  len.writeUInt32LE(raw.length, 0);
  return Buffer.concat([len, Buffer.from(gzipSync(raw))]).toString('base64');
}

export function systemInfo(s: AfoSystem): string {
  return `<SystemInfo><SystemId>${esc(s.systemId)}</SystemId><SystemName>${esc(s.systemName)}</SystemName><SystemDescription>${esc(s.systemName)}</SystemDescription>`
    + `<SystemType>BW_SYSTEM</SystemType><IsLocal>True</IsLocal><SsoEnabled>False</SsoEnabled><BoeID /><BoeID2>0</BoeID2><IsBoeDataSource>False</IsBoeDataSource>`
    + (s.server ? `<ServerGroup>${esc(s.server)}</ServerGroup>` : '')
    + `<Origin>USEREDIT</Origin>`
    + (s.systemNumber ? `<SystemNumber>${esc(s.systemNumber)}</SystemNumber>` : '')
    + `<Client>${esc(s.client)}</Client><Language>${esc(s.language)}</Language>`
    + (s.sapRouter ? `<SapRouter>${esc(s.sapRouter)}</SapRouter>` : '')
    + (s.sncPartnerName ? `<SncPartnerName>${esc(s.sncPartnerName)}</SncPartnerName>` : '')
    + `<SncQuality>${s.sncQuality ?? (s.sncPartnerName ? 0 : -1)}</SncQuality><SncNoSSO>True</SncNoSSO>`
    + (s.uuid ? `<Uuid>${esc(s.uuid)}</Uuid>` : '')
    + `</SystemInfo>`;
}

function range(name: string, sheetId: string, cell: string): string {
  const { row, col } = fromA1(cell);
  const c = `<Cell Name="TopLeft"><Col>${col}</Col><Row>${row}</Row></Cell><Cell Name="BottomRight"><Col>${col}</Col><Row>${row}</Row></Cell>`;
  return `<Range Name="${name}"><SheetName>${esc(sheetId)}</SheetName>${c}</Range>`;
}

/**
 * A taken-over item's ranges moved so the first one starts at `cell`, keeping their size: a
 * crosstab keeps the area it occupies, so AfO sees its rendered cells in the layout as its own
 * and not as foreign content it would cut off (ID-103030).
 */
export function moveRanges(itemXml: string, sheetId: string, cell: string): string {
  const { row, col } = fromA1(cell);
  const first = /<Range Name="[^"]+">[\s\S]*?<Cell Name="TopLeft">[\s\S]*?<Col>(\d+)<\/Col>[\s\S]*?<Row>(\d+)<\/Row>/.exec(itemXml);
  if (!first) return itemXml.replace(/<Range Name="([^"]+)">[\s\S]*?<\/Range>/g, (_m, name: string) => range(name, sheetId, cell));
  const dCol = col - Number(first[1]), dRow = row - Number(first[2]);
  return itemXml.replace(/<Range Name="[^"]+">[\s\S]*?<\/Range>/g, (r) => r
    .replace(/<SheetName>[^<]*<\/SheetName>/, () => `<SheetName>${esc(sheetId)}</SheetName>`)
    .replace(/<Col>(\d+)<\/Col>/g, (_m, c: string) => `<Col>${Number(c) + dCol}</Col>`)
    .replace(/<Row>(\d+)<\/Row>/g, (_m, n: string) => `<Row>${Number(n) + dRow}</Row>`));
}

/** The A1 area a crosstab occupies, from its RangeOccupied (else its first range). */
export function occupiedArea(gridXml: string): string | undefined {
  const r = /<Range Name="RangeOccupied">[\s\S]*?<\/Range>/.exec(gridXml)?.[0] ?? /<Range Name="[^"]+">[\s\S]*?<\/Range>/.exec(gridXml)?.[0];
  const corner = (name: string) => {
    const c = new RegExp(`<Cell Name="${name}">[\\s\\S]*?<Col>(\\d+)</Col>[\\s\\S]*?<Row>(\\d+)</Row>`).exec(r ?? '');
    return c ? toA1(Number(c[2]), Number(c[1])) : undefined;
  };
  const tl = corner('TopLeft'), br = corner('BottomRight');
  if (!tl) return undefined;
  const abs = (a: string) => a.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  return !br || br === tl ? abs(tl) : `${abs(tl)}:${abs(br)}`;
}

export function planningFunction(pf: PlanningFunctionSpec, system: string, today: string): string {
  const vars = `<?xml version="1.0" encoding="UTF-8" standalone="no" ?><VARIABLES>${(pf.variables ?? [])
    .map((v) => `<VARIABLE name="${esc(v.name)}" type="INFO_OBJECT_MEMBER" isStateVariable="false"><SELECTIONS><SELECTION type="SINGLE_MEMBER"><MEMBER name="${esc(v.value)}" type="MEMBER" /></SELECTION></SELECTIONS></VARIABLE>`)
    .join('')}</VARIABLES>`;
  const filter = `<?xml version="1.0" encoding="UTF-8" standalone="no" ?><BICS_VIEW version="3.1"><SERVICE_DESCRIPTION><INITIALIZE_PARENT_CONTAINER>X</INITIALIZE_PARENT_CONTAINER><DATA_AREA>DEFAULT</DATA_AREA></SERVICE_DESCRIPTION>`
    + `<INFO_PROVIDER name="${esc(pf.aggregationLevel)}" /><SELECTION_OBJECT name="!!A${esc(pf.aggregationLevel)}"><DUE_DATE>${today}</DUE_DATE></SELECTION_OBJECT>`
    + `<SELECTION_SPACE_SORTING type="DEFAULT" /><SELECTION_STATE_SORTING type="DEFAULT" /><SELECTION_STATE><SELECTION type="CARTESIAN_PRODUCT"><CHARACTERISTICS>`
    + pf.characteristics.map((c) => {
      const values = pf.filter?.[c] ?? [];
      const sel = values.map((v) => `<SELECTION type="SINGLE_MEMBER"><MEMBER name="${esc(v)}" type="MEMBER" /></SELECTION>`).join('');
      return `<CHARACTERISTIC name="${esc(c)}">${sel ? `<SELECTIONS>${sel}</SELECTIONS>` : '<SELECTIONS />'}</CHARACTERISTIC>`;
    }).join('')
    + `</CHARACTERISTICS></SELECTION></SELECTION_STATE></BICS_VIEW>`;
  const refs = (pf.filterFromDataSource ?? [])
    .map((c) => `<FilterRefConfig><TechnicalName>${esc(c)}</TechnicalName><ReferenceType>Filter</ReferenceType></FilterRefConfig>`)
    .join('')
    + (pf.variableCells ?? [])
      .map((v) => `<VariableRefConfig><TechnicalName>${esc(v.name)}</TechnicalName><ReferenceType>Cell</ReferenceType><CellAdress>${esc(v.cell)}</CellAdress></VariableRefConfig>`)
      .join('');
  return `<PlanningFunctionUsage><Key>${esc(pf.key)}</Key><PlanningObjectInfo><Name>${esc(pf.name)}</Name><Description>${esc(pf.description ?? pf.name)}</Description><PlanningType>Function</PlanningType>${system}</PlanningObjectInfo>`
    // AfO spells the element DetltaMode.
    + (pf.deltaMode ? '<ChangedData><DetltaMode>True</DetltaMode></ChangedData>' : '')
    + `<DataProviderKey>${esc(pf.dataProviderKey)}</DataProviderKey><PlanFunctionVariables><Variables_Header compressed="true">${compressBlock(vars)}</Variables_Header></PlanFunctionVariables>`
    + `<PlanFilter><BICS_VIEW_HEADER compressed="true">${compressBlock(filter)}</BICS_VIEW_HEADER></PlanFilter>`
    + (refs ? `<ReferenceConfiguration>${refs}</ReferenceConfiguration>` : '')
    + `</PlanningFunctionUsage>`;
}

function promptXml(p: PromptValue): string {
  const member = (name: string) => `<MEMBER name="${esc(name)}" type="MEMBER" />`;
  let sel: string;
  if (p.low !== undefined) sel = `<SELECTION type="INTERVAL"><LOW_VALUE>${member(p.low)}</LOW_VALUE><HIGH_VALUE>${member(p.high ?? p.low)}</HIGH_VALUE></SELECTION>`;
  else if (p.literal !== undefined) sel = `<SELECTION type="SINGLE_MEMBER"><MEMBER type="LITERAL_MEMBER"><VALUE>${esc(p.literal)}</VALUE></MEMBER></SELECTION>`;
  else sel = `<SELECTION type="SINGLE_MEMBER">${member(p.member ?? '')}</SELECTION>`;
  return `<VARIABLE name="${esc(p.name)}" type="INFO_OBJECT_MEMBER" isStateVariable="false"><SELECTIONS>${sel}</SELECTIONS></VARIABLE>`;
}

const variableName = (xml: string) => /\bname="([^"]*)"/.exec(xml)?.[1]?.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

/** Variable names with a value in a stored state (its VARIABLES block). */
export function storedVariableNames(state: string): string[] {
  const block = /<VARIABLES>([\s\S]*?)<\/VARIABLES>/.exec(state)?.[1] ?? '';
  return [...block.matchAll(/<VARIABLE [^>]*>/g)].map((m) => variableName(m[0])!).filter(Boolean);
}

/** Set prompt values in a stored state: same-named variables are replaced, the others kept. */
export function setStateVariables(state: string, prompts: PromptValue[]): string {
  const block = /<VARIABLES>([\s\S]*?)<\/VARIABLES>|<VARIABLES\s*\/>/.exec(state);
  const given = new Map(prompts.map((p) => [p.name, promptXml(p)]));
  const entries = [...(block?.[1] ?? '').matchAll(/<VARIABLE [^>]*\/>|<VARIABLE [\s\S]*?<\/VARIABLE>/g)].map((m) => {
    const name = variableName(m[0])!;
    const xml = given.get(name) ?? m[0];
    given.delete(name);
    return xml;
  });
  const xml = `<VARIABLES>${[...entries, ...given.values()].join('')}</VARIABLES>`;
  if (block) return state.slice(0, block.index) + xml + state.slice(block.index + block[0].length);
  // AfO writes VARIABLES right after the QUERY element (not the one in SERVICE_DESCRIPTION).
  const query = /<QUERY name="[^"]*"\s*\/>|<QUERY name="[^"]*">[\s\S]*?<\/QUERY>/.exec(state) ?? /<BICS_VIEW[^>]*>/.exec(state);
  if (!query) throw new Error('The stored state is not a BICS view.');
  const at = query.index + query[0].length;
  return state.slice(0, at) + xml + state.slice(at);
}

/**
 * Store prompt values in a data source of an AfO definition (`<DataProvider>` XML). A data
 * source without a stored state gets a minimal one; `ResetDataSource` is switched off, since AfO
 * discards the saved state of such a data source, prompts included. Returns notes for the caller.
 */
export function withStoredPrompts(dpXml: string, prompts: PromptValue[], minimalState: () => string): { xml: string; notes: string[] } {
  const notes: string[] = [];
  let xml = dpXml;
  const compressed = /<BICS_VIEW_HEADER compressed="true">([^<]*)<\/BICS_VIEW_HEADER>/.exec(xml);
  if (compressed) {
    const state = setStateVariables(decodeCompressed(compressed[1]), prompts);
    xml = xml.replace(compressed[0], () => `<BICS_VIEW_HEADER compressed="true">${compressBlock(state)}</BICS_VIEW_HEADER>`);
  } else {
    const plain = /<BICS_VIEW_HEADER>([\s\S]*?)<\/BICS_VIEW_HEADER>|<BICS_VIEW_HEADER\s*\/>/.exec(xml);
    const design = /<TABLE_DESIGN>[\s\S]*?<\/TABLE_DESIGN>|<TABLE_DESIGN\s*\/>/.exec(plain?.[1] ?? '')?.[0];
    let state = setStateVariables(minimalState(), prompts);
    notes.push('had no stored navigation state; a minimal one now carries the prompts, so it opens in the query default view');
    if (design && !/^<TABLE_DESIGN>\s*<STYLE_LIST\s*\/>\s*<OPERATION_LIST\s*\/>\s*<\/TABLE_DESIGN>$|^<TABLE_DESIGN\s*\/>$/.test(design)) {
      state = state.replace('<DOCUMENT_APPLICATION_ASSIGNED', () => `${design}<DOCUMENT_APPLICATION_ASSIGNED`);
      notes.push('its design rules (TABLE_DESIGN) moved into the stored state; check them after opening in AfO');
    }
    const header = `<BICS_VIEW_HEADER compressed="true">${compressBlock(state)}</BICS_VIEW_HEADER>`;
    if (plain) xml = xml.replace(plain[0], () => header);
    else if (/<DimensionSynchronization\s*\/>/.test(xml)) xml = xml.replace(/<DimensionSynchronization\s*\/>/, (m) => `${m}${header}`);
    else xml = xml.replace(/<Comments\s*\/>|<Comments>|<\/DataProvider>/, (m) => `${header}${m}`);
  }
  if (/<ResetDataSource>True<\/ResetDataSource>/.test(xml)) {
    // AfO writes the element only when the option is on.
    xml = xml.replace(/<ResetDataSource>True<\/ResetDataSource>/, '');
    notes.push('had "reset data source" on, which drops stored prompts on open; switched off');
  }
  return { xml, notes };
}

export interface VisibilityEntry { dataProvider?: string; name: string; xml: string }

/** Variable entries of the PromptsVisibility block, with the data source they belong to (none when merged). */
export function promptsVisibility(definition: string): VisibilityEntry[] {
  return [...(firstElement(definition, 'PromptsVisibility') ?? '').matchAll(/<Variable [^>]*\/>/g)].map((m) => ({
    dataProvider: /dataprovider="([^"]*)"/.exec(m[0])?.[1],
    name: variableName(m[0].replace(/dataprovider="[^"]*"/, '')) ?? '',
    xml: m[0],
  }));
}

/** Add visibility entries for prompts that have none: per data source, or once when variables are merged. */
export function addVisibility(entries: VisibilityEntry[], dataProvider: string, prompts: PromptValue[], merged: boolean): void {
  const owner = merged ? undefined : dataProvider;
  for (const p of prompts) {
    if (entries.some((v) => v.name === p.name && v.dataProvider === owner)) continue;
    entries.push({ dataProvider: owner, name: p.name, xml: `<Variable ${owner ? `dataprovider="${esc(owner)}" ` : ''}name="${esc(p.name)}" visible="True" />` });
  }
}

/**
 * Variables the workbook knows (visibility entries) but stores no value for, as a warning: AfO
 * shows the prompt dialog when one of them is mandatory. Without BW access it cannot be told which are.
 */
export function unstoredPrompts(dataProviders: Array<{ key: string; xml: string }>, entries: VisibilityEntry[]): string | undefined {
  const stored = new Map(dataProviders.map((dp) => {
    const state = /<BICS_VIEW_HEADER compressed="true">([^<]*)</.exec(dp.xml)?.[1];
    return [dp.key, new Set(state ? storedVariableNames(decodeCompressed(state)) : [])];
  }));
  const missing = entries.filter((v) => v.dataProvider === undefined
    ? ![...stored.values()].some((s) => s.has(v.name))
    : stored.has(v.dataProvider) && !stored.get(v.dataProvider)!.has(v.name));
  if (!missing.length) return undefined;
  return `No stored value for ${missing.map((v) => (v.dataProvider ? `${v.name} (${v.dataProvider})` : v.name)).join(', ')}: if a mandatory variable has no value, AfO shows the prompt dialog on open. Check the query variables (bw_get_query) and add the mandatory ones to prompts.`;
}

/** Problems of one prompt value, for spec validation. */
export function promptProblems(p: Partial<PromptValue> | undefined, what: string): string[] {
  const errors: string[] = [];
  const kinds = [p?.member !== undefined, p?.literal !== undefined, p?.low !== undefined].filter(Boolean).length;
  if (!p?.name) errors.push(`${what}: name is required.`);
  if (kinds !== 1) errors.push(`${what} (${p?.name}): give exactly one of member, literal or low (with high).`);
  if (p?.high !== undefined && p?.low === undefined) errors.push(`${what} (${p?.name}): high without low.`);
  return errors;
}

const FILTER_OPERATORS = ['GREATER', 'GREATER_EQUAL', 'LESS', 'LESS_EQUAL'];

export function filterProblems(f: Partial<FilterSpec> | undefined, what: string): string[] {
  const errors: string[] = [];
  if (!f?.characteristic) errors.push(`${what}: characteristic is required.`);
  const kinds = [f?.members !== undefined, f?.low !== undefined, f?.value !== undefined].filter(Boolean).length;
  if (kinds !== 1) errors.push(`${what} (${f?.characteristic}): give exactly one of members, low (with high) or operator with value.`);
  if (f?.members !== undefined && (!Array.isArray(f.members) || !f.members.length)) errors.push(`${what} (${f?.characteristic}): members must be a non-empty list.`);
  if (f?.high !== undefined && f?.low === undefined) errors.push(`${what} (${f?.characteristic}): high without low.`);
  if ((f?.value !== undefined) !== (f?.operator !== undefined)) errors.push(`${what} (${f?.characteristic}): operator and value go together.`);
  if (f?.operator !== undefined && !FILTER_OPERATORS.includes(f.operator)) errors.push(`${what} (${f?.characteristic}): operator must be one of ${FILTER_OPERATORS.join(', ')}.`);
  return errors;
}

/** The SELECTION_STATE block of a navigation state with the given filters, in the form AfO writes it. */
function filtersXml(filters: FilterSpec[] | undefined): string {
  if (!filters?.length) return '';
  const member = (name: string) => `<MEMBER name="${esc(name)}" type="MEMBER" />`;
  const byChar = new Map<string, string[]>();
  for (const f of filters) {
    const sign = `<SIGN>${f.exclude ? 'EXCLUDING' : 'INCLUDING'}</SIGN>`;
    let sel: string[];
    if (f.members) {
      sel = f.members.map((m) => f.exclude
        ? `<SELECTION type="RANGE" operator="EQUAL">${sign}<VALUE>${member(m)}</VALUE></SELECTION>`
        : `<SELECTION type="SINGLE_MEMBER">${member(m)}</SELECTION>`);
    } else if (f.low !== undefined) {
      sel = [`<SELECTION type="RANGE" operator="BETWEEN">${sign}<LOW_VALUE>${member(f.low)}</LOW_VALUE><HIGH_VALUE>${member(f.high ?? f.low)}</HIGH_VALUE></SELECTION>`];
    } else {
      sel = [`<SELECTION type="RANGE" operator="${f.operator}">${sign}<VALUE>${member(f.value ?? '')}</VALUE></SELECTION>`];
    }
    byChar.set(f.characteristic, [...(byChar.get(f.characteristic) ?? []), ...sel]);
  }
  return `<SELECTION_SPACE_SORTING type="DEFAULT" /><SELECTION_STATE_SORTING type="DEFAULT" /><SELECTION_STATE><SELECTION type="CARTESIAN_PRODUCT"><CHARACTERISTICS>`
    + [...byChar].map(([c, sel]) => `<CHARACTERISTIC name="${esc(c)}"><SELECTIONS>${sel.join('')}</SELECTIONS></CHARACTERISTIC>`).join('')
    + `</CHARACTERISTICS></SELECTION></SELECTION_STATE>`;
}

/** The AXES block of a navigation state, in the form AfO writes it. */
function axesXml(axes: DataProviderSpec['axes']): string {
  if (!axes) return '';
  const axis = (name: string, chars: string[] | undefined) => chars === undefined ? '' :
    `<AXIS name="${name}"><HIERARCHY><ACTIVE>false</ACTIVE><DRILL_LEVEL>0</DRILL_LEVEL></HIERARCHY><SYMMETRIC_DRILL_MODE>false</SYMMETRIC_DRILL_MODE><RESULT_POSITION>BOTTOM</RESULT_POSITION>`
    + `<ZERO_SUPPRESSION_MODE>NONE</ZERO_SUPPRESSION_MODE><ZERO_SUPPRESSION_ON_DATA_ENTRY_ENABLED_DATA_CELLS>true</ZERO_SUPPRESSION_ON_DATA_ENTRY_ENABLED_DATA_CELLS>`
    + `<CHARACTERISTICS>${chars.map((c) => `<CHARACTERISTIC>${esc(c)}</CHARACTERISTIC>`).join('')}</CHARACTERISTICS><CUSTOM_PARAMETERS /></AXIS>`;
  return `<AXES>${axis('ROWS', axes.rows)}${axis('COLUMNS', axes.columns)}${axis('FREE', axes.free)}</AXES>`;
}

/**
 * A minimal stored navigation state. DOCUMENT_APPLICATION_ASSIGNED tells AfO that no BW
 * document application is assigned; without any stored state AfO asks the server, and on a
 * system without document storage that ends in an exception once the crosstab is input-ready.
 */
export function stateXml(dp: DataProviderSpec, day: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="no" ?><BICS_VIEW version="3.1"><SERVICE_DESCRIPTION><DATA_AREA>DEFAULT</DATA_AREA><QUERY>${esc(dp.query)}</QUERY></SERVICE_DESCRIPTION>`
    + (dp.designRules ? `<MODULES>${tableDesignXml(dp.designRules)}</MODULES>` : '<MODULES />')
    + (dp.infoProvider ? `<INFO_PROVIDER name="${esc(dp.infoProvider)}" />` : '')
    + `<QUERY name="${esc(dp.query)}"><DUE_DATE>${day}</DUE_DATE></QUERY>`
    + (dp.prompts?.length ? `<VARIABLES>${dp.prompts.map(promptXml).join('')}</VARIABLES>` : '')
    + filtersXml(dp.filters)
    + (dp.inputReady ?? true ? '<IS_INPUT_ENABLED>true</IS_INPUT_ENABLED>' : '')
    + axesXml(dp.axes)
    + `<DOCUMENT_APPLICATION_ASSIGNED /></BICS_VIEW>`;
}

/** #RRGGBB as the signed 32-bit ARGB integer AfO stores colours as. */
const afoColor = (hex: string) => {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) throw new Error(`Colour ${hex}: expected #RRGGBB.`);
  return ((0xff000000 | parseInt(m[1], 16)) >> 0).toString();
};

/**
 * `<TABLE_DESIGN>` as AfO saves formatting rules in the data source state (under `<MODULES>`):
 * one CELL_STYLE operation per structure member, addressed by member and structure UID, covering
 * the member and its data cells.
 */
export function tableDesignXml(rules: DesignRulesSpec): string {
  const own = new Set((rules.styles ?? []).map((s) => s.name));
  for (const m of rules.members) if (!m.predefined && !own.has(m.style)) throw new Error(`Design rule for ${m.member}: unknown style ${m.style}.`);
  const styles = (rules.styles ?? []).map((s) => `<STYLE subType="CUSTOM" name="${esc(s.name)}"><COLOR>${afoColor(s.background)}</COLOR>`
    + `<FONT_STYLE><COLOR>${afoColor(s.font ?? '#2B373D')}</COLOR></FONT_STYLE></STYLE>`).join('');
  const st = esc(rules.structure);
  const operations = rules.members.map((m, i) => `<OPERATION isRange="true" operation_id="${randomUUID()}" type="CELL_OBJECT" name="Format_${i + 1}" active="true" subType="CELL_STYLE">`
    + `<ADDRESS super_type="POINT_ADDRESS" type="TUPLE"><ADDRESS super_type="LINE_ADDRESS" type="TUPLE_ELEMENT"><KEY>${esc(m.member)}</KEY>${m.leaf ? '<NODE_TYPE>Leaf</NODE_TYPE><IS_LEAF>true</IS_LEAF>' : `<NODE_TYPE>${st}</NODE_TYPE><IS_LEAF>false</IS_LEAF>`}`
    + `<ADDRESS type="CHARACTERISTIC"><KEY>${st}</KEY><IS_STRUCTURE>true</IS_STRUCTURE></ADDRESS></ADDRESS>`
    + `<ADDRESS super_type="LINE_ADDRESS" type="CHARACTERISTIC"><KEY>${st}</KEY><IS_STRUCTURE>true</IS_STRUCTURE></ADDRESS></ADDRESS>`
    + `<OBJECT type="CELL_STYLE" subType="${m.predefined ? 'PREDEFINED' : 'CUSTOM'}" name="${esc(m.style)}" />`
    + `<RANGE><AREA>${esc(rules.area ?? 'BOTH_AREAS')}</AREA><LINE_ADDRESS_NUMBER>PRIMARY</LINE_ADDRESS_NUMBER></RANGE></OPERATION>`).join('');
  return `<TABLE_DESIGN>${styles ? `<STYLE_LIST>${styles}</STYLE_LIST>` : '<STYLE_LIST />'}<OPERATION_LIST>${operations}</OPERATION_LIST></TABLE_DESIGN>`;
}

/** The `<Application>` document. `sheetIds` maps sheet names to their AfO ids. */
export function buildDefinition(spec: AfoWorkbookSpec, sheetIds: Map<string, string>, today = new Date(), warnings: string[] = []): string {
  const src = spec.sourceDefinition;
  const srcSystem = src ? sourceSystemInfo(src) : undefined;
  if (!spec.system && !srcSystem) throw new Error('No system: give system, or a source workbook with a connection.');
  const sys = spec.system ? systemInfo(spec.system) : srcSystem!;
  const systemId = spec.system?.systemId ?? (src && /<SystemId>([^<]*)<\/SystemId>/.exec(src)?.[1]) ?? '';
  const day = today.toISOString().slice(0, 10).replace(/-/g, '');
  const sheetId = (name: string) => {
    const id = sheetIds.get(name);
    if (!id) throw new Error(`Sheet not found: ${name}`);
    return id;
  };
  const taken = (tag: string, key: string, what: string) => {
    const xml = src && keyedElement(src, tag, key);
    if (!xml) throw new Error(`The source has no ${what} ${key}.`);
    return xml;
  };
  // Data source keys may be renamed: every reference in taken-over XML follows.
  const renames = new Map(spec.dataProviders.filter((d) => d.from).map((d) => [d.from!, d.key]));
  const renameRefs = (xml: string) => xml.replace(/<DataProviderKey>([^<]*)<\/DataProviderKey>/g, (m, k: string) => (renames.has(k) ? `<DataProviderKey>${esc(renames.get(k)!)}</DataProviderKey>` : m));

  const storePrompts = spec.storePrompts ?? (src ? firstElement(src, 'StorePromptsInDocument') === '<StorePromptsInDocument>True</StorePromptsInDocument>' : false);
  const merged = spec.mergeVariables ?? (src ? firstElement(src, 'MergeVariables') === '<MergeVariables>True</MergeVariables>' : false);
  if (!storePrompts && spec.dataProviders.some((dp) => dp.prompts?.length)) warnings.push('Prompts are given but StorePromptsInDocument is off: AfO ignores them. Set storePrompts: true.');
  const takenDataProvider = (dp: DataProviderSpec) => {
    let xml = replaceFirstKey(taken('DataProvider', dp.from!, 'data source'), dp.key, esc);
    // inputReady set in the spec wins over the source's input readiness.
    if (dp.inputReady !== undefined) {
      const entry = `<RefreshDataEntryEnabled>${tf(dp.inputReady)}</RefreshDataEntryEnabled>`;
      xml = /<RefreshDataEntryEnabled>[^<]*<\/RefreshDataEntryEnabled>/.test(xml)
        ? xml.replace(/<RefreshDataEntryEnabled>[^<]*<\/RefreshDataEntryEnabled>/, entry)
        : xml.replace(/(<\/FillableNumber>)/, `$1${entry}`);
    }
    if (!dp.prompts?.length) return xml;
    const r = withStoredPrompts(xml, dp.prompts, () => stateXml({ ...dp, prompts: undefined, axes: undefined }, day));
    warnings.push(...r.notes.map((n) => `${dp.key}: ${n}.`));
    return r.xml;
  };
  const dpXml = spec.dataProviders.map((dp) => dp.from
    ? takenDataProvider(dp)
    : `<DataProvider><Key>${esc(dp.key)}</Key><FillableNumber>1</FillableNumber><RefreshDataEntryEnabled>${tf(dp.inputReady ?? true)}</RefreshDataEntryEnabled><Cleaned>False</Cleaned>`
      + `<DataSourceInfo><Id>${esc(dp.query)}</Id><Name>${esc(dp.query)}</Name><DisplayName>${esc(dp.query)}</DisplayName><Type>Query</Type><HasChildren>False</HasChildren><IsFormulaDriven>False</IsFormulaDriven><ShowScalingFactor>False</ShowScalingFactor><ShowUnits>False</ShowUnits>${sys}</DataSourceInfo>`
      + `<DimensionSynchronization /><BICS_VIEW_HEADER compressed="true">${compressBlock(stateXml(dp, day))}</BICS_VIEW_HEADER><Comments /></DataProvider>`);
  const dps = dpXml.join('');

  // Prompt visibility: the source's entries for taken-over data sources, plus the new prompts.
  // Without a data source (merged variables) an entry applies to all; else it follows its data source.
  const visibility = (src ? promptsVisibility(src) : [])
    .filter((v) => v.dataProvider === undefined || renames.has(v.dataProvider))
    .map((v) => ({ ...v, dataProvider: v.dataProvider && renames.get(v.dataProvider), xml: v.xml.replace(/dataprovider="([^"]*)"/, (_m, k: string) => `dataprovider="${esc(renames.get(k)!)}"`) }));
  for (const dp of spec.dataProviders) addVisibility(visibility, dp.key, dp.prompts ?? [], merged);
  const prompts = visibility.map((v) => v.xml).join('');
  if (storePrompts) {
    const unstored = unstoredPrompts(spec.dataProviders.map((dp, i) => ({ key: dp.key, xml: dpXml[i] })), visibility);
    if (unstored) warnings.push(unstored);
  }

  const pfs = (spec.planningFunctions ?? []).map((pf) => pf.from
    ? renameRefs(replaceFirstKey(taken('Planning\\w+Usage', pf.from, 'planning object'), pf.key, esc))
      .replace(/<DataProviderKey>[^<]*<\/DataProviderKey>/, () => `<DataProviderKey>${esc(pf.dataProviderKey)}</DataProviderKey>`)
    : planningFunction(pf, sys, day)).join('');

  const grids = spec.grids.map((g, i) => {
    const id = sheetId(g.sheet);
    if (g.from) {
      const xml = moveRanges(replaceFirstKey(taken('Grid', g.from, 'crosstab'), g.key, esc)
        .replace(/<ShapeId>[^<]*<\/ShapeId>/, () => `<ShapeId>${i + 2}</ShapeId>`)
        .replace(/<DataProviderKey>[^<]*<\/DataProviderKey>/, () => `<DataProviderKey>${esc(g.dataProviderKey)}</DataProviderKey>`), id, g.cell);
      return g.autoFit === undefined ? xml : xml.replace(/<AutoFit>[^<]*<\/AutoFit>/, `<AutoFit>${tf(g.autoFit)}</AutoFit>`);
    }
    return `<Grid><Key>${esc(g.key)}</Key><Description>${esc(g.key)}</Description><ShapeId>${i + 2}</ShapeId><DataProviderKey>${esc(g.dataProviderKey)}</DataProviderKey>`
      + range('RangeOccupied', id, g.cell) + range('RequestedRange', id, g.cell)
      + `<ApplyDefaultFormats>True</ApplyDefaultFormats><ApplyNumberFormats>True</ApplyNumberFormats><AutoFit>${tf(g.autoFit ?? true)}</AutoFit><DisplayRepeatedTexts>False</DisplayRepeatedTexts><DisplayRepeatedTitles>False</DisplayRepeatedTitles>`
      + `<NumberOfNewLines>5</NumberOfNewLines><NewLinesCheckMode>NONE</NewLinesCheckMode><DisplayScalingUnitsInHeader>False</DisplayScalingUnitsInHeader><HierarchySymbolType>1</HierarchySymbolType><WrapHeaderTextWithLinebreak>True</WrapHeaderTextWithLinebreak></Grid>`;
  }).join('');

  // Settings: the spec wins, then the source, then the builder default.
  const setting = (name: string, value: boolean | undefined, fallback: boolean) =>
    value !== undefined ? `<${name}>${tf(value)}</${name}>` : (src && firstElement(src, name)) ?? `<${name}>${tf(fallback)}</${name}>`;
  const fromSource = (name: string, fallback: string) => (src && firstElement(src, name)) ?? fallback;
  const ownDesign = `<PlanningSystemDesignTime>${sys}<CellLocking>None</CellLocking></PlanningSystemDesignTime>`;
  const design = spec.system || !src ? ownDesign : fromSource('PlanningSystemDesignTime', ownDesign);
  const revision = spec.afoRevision ?? (src && /<AAO_Revision>([^<]*)<\/AAO_Revision>/.exec(src)?.[1]) ?? '2.8.2301.1605';
  return `<Application xmlns="${AFO_NAMESPACE}"><com.sap.ip.bi.pioneer><Version>12</Version><AAO_Revision>${esc(revision)}</AAO_Revision><SystemId>${esc(systemId)}</SystemId>`
    + setting('RefreshOnOpen', spec.refreshOnOpen, false) + setting('PlanningModeSetToChangeMode', spec.changeModeOnOpen, false) + setting('Cleaned', spec.removeDataOnSave, true)
    + fromSource('ForcePromptOnInitialRefresh', '<ForcePromptOnInitialRefresh>False</ForcePromptOnInitialRefresh>')
    + setting('StorePromptsInDocument', spec.storePrompts, false) + setting('MergeVariables', spec.mergeVariables, false)
    + fromSource('RefreshPlanningObjectsOnRefreshAll', '<RefreshPlanningObjectsOnRefreshAll>True</RefreshPlanningObjectsOnRefreshAll>')
    + (prompts ? `<PromptsVisibility>${prompts}</PromptsVisibility>` : '<PromptsVisibility />')
    + (spec.styleSetName ? `<StyleSetName>${esc(spec.styleSetName)}</StyleSetName>` : fromSource('StyleSetName', '<StyleSetName>SAP Blue</StyleSetName>')) + design + fromSource('BExProperties', '<BExProperties />')
    + `<DataProviders>${dps}</DataProviders>` + (pfs ? `<PlanningObjects>${pfs}</PlanningObjects>` : '') + `<Items>${grids}${formulaItemsXml(spec, sheetId, renames)}</Items></com.sap.ip.bi.pioneer></Application>`;
}

/** A formula item's formula with renamed data source aliases (first argument). */
export function formulaItemFormula(item: FormulaItemSpec, spec: AfoWorkbookSpec, renames: Map<string, string>): string {
  let formula = item.formula;
  if (formula === undefined && item.from && spec.sourceDefinition) {
    const raw = keyedElement(spec.sourceDefinition, 'Fdd', item.from);
    formula = raw && /<FormulaString>([\s\S]*?)<\/FormulaString>/.exec(raw)?.[1]?.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  }
  if (formula === undefined) throw new Error(`Formula item ${item.key}: no formula.`);
  return formula.replace(/^(=\s*\w+\s*\(\s*")([^"]*)(")/, (m, a: string, alias: string, b: string) => (renames.has(alias) ? `${a}${renames.get(alias)}${b}` : m));
}

function formulaItemsXml(spec: AfoWorkbookSpec, sheetId: (name: string) => string, renames: Map<string, string>): string {
  return (spec.formulaItems ?? []).map((item) => {
    const id = sheetId(item.sheet);
    // Element text: AfO writes quotes as they are; only &, < and > need escaping.
    const formula = formulaItemFormula(item, spec, renames).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    if (item.from && spec.sourceDefinition) {
      const raw = keyedElement(spec.sourceDefinition, 'Fdd', item.from);
      if (!raw) throw new Error(`The source has no formula item ${item.from}.`);
      return replaceFirstKey(raw, item.key, esc)
        .replace(/<Range Name="([^"]+)">[\s\S]*?<\/Range>/g, (_m, name: string) => range(name, id, item.cell))
        .replace(/<FormulaString>[\s\S]*?<\/FormulaString>/, () => `<FormulaString>${formula}</FormulaString>`);
    }
    return `<Fdd><Key>${esc(item.key)}</Key>${range('RequestedRange', id, item.cell)}<FormulaString>${formula}</FormulaString></Fdd>`;
  }).join('');
}

export const REL_CP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty';
export const REL_WS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
export const REL_VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
export const CT_WS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
export const CT_CP = 'application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty';
export const CT_MACRO_MAIN = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
export const CT_VBA = 'application/vnd.ms-office.vbaProject';

export const utf16 = (s: string) => new Uint8Array(Buffer.from(s, 'utf16le'));
export const relXml = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
export const relsDoc = (items: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;

export function addRels(parts: Record<string, Uint8Array>, relsPath: string, items: string) {
  parts[relsPath] = strToU8(parts[relsPath] ? strFromU8(parts[relsPath]).replace('</Relationships>', `${items}</Relationships>`) : relsDoc(items));
}

export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

export function absolute(cell: string): string {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(cell.toUpperCase());
  return m ? `$${m[1]}$${m[2]}` : cell;
}

export interface MakeOptions {
  /** Declare the workbook macro-enabled (.xlsm) even without VBA modules in the spec. */
  macroEnabled?: boolean;
  /** Collects what the caller should check, e.g. variables without a stored value. */
  warnings?: string[];
}

const HIDDEN_SHEET = '_com.sap.ip.bi.xl.hiddensheet';

/** A sheet's custom properties: name → the part holding the value. */
function sheetProperties(parts: Record<string, Uint8Array>, part: string, xml: string): Map<string, string> {
  const relsPath = relsPartOf(part);
  const targets = new Map(parts[relsPath] ? [...strFromU8(parts[relsPath]).matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]) : []);
  const props = new Map<string, string>();
  for (const m of xml.matchAll(/<customPr [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
    const target = targets.get(m[2]);
    if (target) props.set(m[1], resolveTarget(part, target));
  }
  return props;
}

/** Make an existing .xlsx/.xlsm an AfO workbook according to the spec. */
export function makeAfoWorkbook(source: Uint8Array, spec: AfoWorkbookSpec, options: MakeOptions = {}): Uint8Array {
  const parts: Record<string, Uint8Array> = unzipSync(source);
  const text = (n: string) => strFromU8(parts[n]);
  let wb = text('xl/workbook.xml');
  let wbRels = text('xl/_rels/workbook.xml.rels');

  // Bridge sheet, added before the sheets are given their AfO ids so it gets one like any other.
  const bridge = spec.agentBridge ? { name: agentSheetName(spec), ...agentSheetXml(spec) } : undefined;
  if (bridge) {
    const existing = [...wb.matchAll(/<sheet [^>]*name="([^"]+)"/g)].map((m) => m[1].replace(/&amp;/g, '&'));
    if (existing.some((n) => n.toLowerCase() === bridge.name.toLowerCase())) throw new Error(`The workbook already has a sheet ${bridge.name}; choose another agentBridge.sheet.`);
    const ids = [...wb.matchAll(/<sheet [^>]*sheetId="(\d+)"/g)].map((m) => Number(m[1])).filter((n) => n < 900);
    parts['xl/worksheets/afoAgent.xml'] = strToU8(bridge.xml);
    wb = wb.replace('</sheets>', () => `<sheet name="${esc(bridge.name)}" sheetId="${Math.max(0, ...ids) + 1}" r:id="rIdAfoAgent"${spec.agentBridge?.hidden ?? true ? ' state="hidden"' : ''}/></sheets>`);
    wbRels = wbRels.replace('</Relationships>', () => relXml('rIdAfoAgent', REL_WS, 'worksheets/afoAgent.xml') + '</Relationships>');
  }
  const relTarget = new Map([...wbRels.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));

  // Code names: the workbook and every sheet need one for AfO ids and VBA document modules.
  if (!/<workbookPr[^>]*codeName=/.test(wb)) {
    // workbookPr precedes bookViews and sheets in the schema order.
    wb = /<workbookPr/.test(wb)
      ? wb.replace('<workbookPr', '<workbookPr codeName="ThisWorkbook"')
      : wb.replace(/<bookViews|<sheets>/, (m) => `<workbookPr codeName="ThisWorkbook"/>${m}`);
  }
  const workbookCodeName = /<workbookPr[^>]*codeName="([^"]+)"/.exec(wb)![1];

  // New custom property parts are numbered after the base's own.
  let cpIndex = Math.max(900, ...Object.keys(parts).map((p) => Number(/^xl\/customProperty(\d+)\.bin$/.exec(p)?.[1] ?? 0)));
  const addedProperties: string[] = [];
  const newProperty = (value: string) => {
    const n = ++cpIndex;
    parts[`xl/customProperty${n}.bin`] = utf16(value);
    addedProperties.push(`xl/customProperty${n}.bin`);
    return n;
  };
  const sheetIds = new Map<string, string>();
  const sheetCodeNames: string[] = [];
  const codeNameBySheet = new Map<string, string>();
  let definitionCodeName = 'AfoDefinition';
  // The definition sheet of a base that is already an AfO workbook: kept, only its definition is replaced.
  let definitionSheet: { part: string; props: Map<string, string> } | undefined;
  const sheets = [...wb.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)];
  sheets.forEach(([, name, rid], i) => {
    const part = 'xl/' + relTarget.get(rid)!.replace(/^\/?xl\//, '');
    // Cells the spec does not write stay byte for byte as they are (formulas, value metadata).
    const original = text(part);
    let xml = original;
    const sheetName = name.replace(/&amp;/g, '&');
    const props = sheetProperties(parts, part, xml);
    let codeName = /<sheetPr[^>]*codeName="([^"]+)"/.exec(xml)?.[1];
    if (!codeName) {
      codeName = sheetName === HIDDEN_SHEET ? 'AfoDefinition' : `Sheet${i + 1}`;
      xml = /<sheetPr/.test(xml) ? xml.replace('<sheetPr', `<sheetPr codeName="${codeName}"`) : xml.replace(/(<worksheet[^>]*>)/, `$1<sheetPr codeName="${codeName}"/>`);
    }
    if (sheetName === HIDDEN_SHEET) {
      definitionSheet = { part, props };
      definitionCodeName = codeName;
      if (xml !== original) parts[part] = strToU8(xml);
      return;
    }
    const renames = new Map(spec.dataProviders.filter((d) => d.from).map((d) => [d.from!, d.key]));
    const writes = [
      ...(spec.cells ?? []).filter((c) => c.sheet === sheetName).map((c) => ({ ref: c.cell, formula: c.formula, value: c.value })),
      ...(spec.formulaItems ?? []).filter((f) => f.sheet === sheetName).map((f) => ({ ref: f.cell, formula: formulaItemFormula(f, spec, renames) })),
    ];
    if (writes.length) xml = writeCells(xml, writes);
    sheetCodeNames.push(codeName);
    codeNameBySheet.set(sheetName, codeName);
    // A sheet of an AfO or EPM workbook keeps its id; cells and items refer to it.
    const ownId = props.get('_pios_id');
    const afoId = ownId && parts[ownId] ? decodeText(parts[ownId]).replace(/\0+$/, '') : `${codeName}|${guid()}`;
    sheetIds.set(sheetName, afoId);
    const entries: string[] = [], rels: string[] = [];
    if (!ownId || !parts[ownId]) {
      // A _pios_id whose part is missing is replaced, never doubled.
      xml = xml.replace(/<customPr [^>]*name="_pios_id"[^>]*\/>/, '');
      const a = newProperty(afoId);
      entries.push(`<customPr name="_pios_id" r:id="rIdAfo${a}"/>`);
      rels.push(relXml(`rIdAfo${a}`, REL_CP, `../customProperty${a}.bin`));
    }
    if (!props.has('CofWorksheetType')) {
      const b = newProperty('AoWorksheetType');
      entries.push(`<customPr name="CofWorksheetType" r:id="rIdAfo${b}"/>`);
      rels.push(relXml(`rIdAfo${b}`, REL_CP, `../customProperty${b}.bin`));
    }
    if (entries.length) {
      // One customProperties per sheet: entries go into an existing one (AfO, EPM).
      const cps = entries.join('');
      if (/<customProperties>/.test(xml)) xml = xml.replace('</customProperties>', () => `${cps}</customProperties>`);
      else if (/<customProperties\s*\/>/.test(xml)) xml = xml.replace(/<customProperties\s*\/>/, () => `<customProperties>${cps}</customProperties>`);
      // customProperties comes after pageMargins/pageSetup/headerFooter and before drawing.
      else if (/<drawing |<legacyDrawing /.test(xml)) xml = xml.replace(/(<drawing |<legacyDrawing )/, `<customProperties>${cps}</customProperties>$1`);
      else xml = xml.replace('</worksheet>', `<customProperties>${cps}</customProperties></worksheet>`);
      if (!/xmlns:r=/.test(xml)) xml = xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
      addRels(parts, relsPartOf(part), rels.join(''));
    }
    if (xml !== original) parts[part] = strToU8(xml);
  });

  // Hidden sheet with the definition.
  const definition = buildDefinition(spec, sheetIds, undefined, options.warnings);
  let rels = wbRels;
  let oldItemKeys: string[] = [];
  if (definitionSheet) {
    const target = definitionSheet.props.get('serializedData2');
    if (!target || !parts[target]) throw new Error('The AfO definition sheet of the base workbook has no definition (serializedData2).');
    const old = parts[target];
    oldItemKeys = [...(firstElement(decodeText(old), 'Items') ?? '').matchAll(/<Key>([^<]*)<\/Key>/g)].map((m) => m[1]);
    const bom = old[0] === 0xff && old[1] === 0xfe;
    parts[target] = bom ? new Uint8Array([0xff, 0xfe, ...utf16(definition)]) : utf16(definition);
  } else {
    const h1 = newProperty(`${HIDDEN_SHEET}|${guid()}`), h2 = newProperty('AoWorksheetType'), h3 = newProperty(definition);
    parts['xl/worksheets/sheet900.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr codeName="AfoDefinition"/><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.787401575" bottom="0.787401575" header="0.3" footer="0.3"/><customProperties><customPr name="_pios_id" r:id="rId1"/><customPr name="CofWorksheetType" r:id="rId2"/><customPr name="serializedData2" r:id="rId3"/></customProperties></worksheet>`);
    parts['xl/worksheets/_rels/sheet900.xml.rels'] = strToU8(relsDoc(relXml('rId1', REL_CP, `../customProperty${h1}.bin`) + relXml('rId2', REL_CP, `../customProperty${h2}.bin`) + relXml('rId3', REL_CP, `../customProperty${h3}.bin`)));
    wb = wb.replace('</sheets>', `<sheet name="${HIDDEN_SHEET}" sheetId="900" state="veryHidden" r:id="rIdAfoSheet"/></sheets>`);
    rels = rels.replace('</Relationships>', relXml('rIdAfoSheet', REL_WS, 'worksheets/sheet900.xml') + '</Relationships>');
  }

  // Defined names SAP<Key> for every item, a crosstab over the area it occupies. Names of the
  // base's old items and of the new ones go first, so none is defined twice.
  const itemKeys = new Set([...oldItemKeys, ...spec.grids.map((g) => g.key), ...(spec.formulaItems ?? []).map((f) => f.key)].map((k) => `SAP${k}`.toUpperCase()));
  wb = wb.replace(/<definedName [^>]*name="([^"]+)"[^>]*>[^<]*<\/definedName>/g, (m, n: string) => (itemKeys.has(n.toUpperCase()) && !/localSheetId=/.test(m) ? '' : m))
    .replace(/<definedNames>\s*<\/definedNames>/, '');
  const area = (g: GridSpec) => {
    const xml = keyedElement(definition, 'Grid', g.key);
    return (xml && occupiedArea(xml)) ?? absolute(g.cell);
  };
  let names = [
    ...spec.grids.map((g) => `<definedName name="SAP${esc(g.key)}">${esc(quoteSheet(g.sheet))}!${area(g)}</definedName>`),
    ...(spec.formulaItems ?? []).map((f) => `<definedName name="SAP${esc(f.key)}">${esc(quoteSheet(f.sheet))}!${absolute(f.cell)}</definedName>`),
  ].join('');
  if (bridge) {
    const ref = (cell: string) => `${esc(quoteSheet(bridge.name))}!${cell}`;
    names += `<definedName name="${AGENT_NAMES.command}">${ref('$B$1')}</definedName><definedName name="${AGENT_NAMES.result}">${ref('$B$3')}</definedName><definedName name="${AGENT_NAMES.status}">${ref('$B$4')}</definedName>`;
  }
  // Replacer functions throughout: names like $B$16 would otherwise be read as $1 back-references.
  if (names) wb = /<definedNames>/.test(wb) ? wb.replace('<definedNames>', () => `<definedNames>${names}`) : wb.replace('</sheets>', () => `</sheets><definedNames>${names}</definedNames>`);

  // One override per custom property part: a workbook may already map ".bin" to another type
  // (printer settings), and Excel drops parts whose type does not match.
  let ct = text('[Content_Types].xml');
  const cpOverrides = addedProperties.map((n) => `<Override PartName="/${n}" ContentType="${CT_CP}"/>`).join('');
  ct = ct.replace('</Types>', () => cpOverrides + (definitionSheet ? '' : `<Override PartName="/xl/worksheets/sheet900.xml" ContentType="${CT_WS}"/>`)
    + (bridge ? `<Override PartName="/xl/worksheets/afoAgent.xml" ContentType="${CT_WS}"/>` : '') + '</Types>');

  // VBA project: document modules for the workbook and each sheet, plus the standard modules.
  const docCode = new Map<string, string>();
  const documentModules = { ...spec.documentModules };
  if (bridge) documentModules.ThisWorkbook = agentThisWorkbookCode(documentModules.ThisWorkbook);
  for (const [target, code] of Object.entries(documentModules)) {
    const codeName = target === 'ThisWorkbook' ? workbookCodeName : codeNameBySheet.get(target);
    if (!codeName) throw new Error(`Sheet not found for document module code: ${target}`);
    docCode.set(codeName, code);
  }
  const standardModules = [...(spec.vbaModules ?? [])];
  if (bridge) standardModules.push({ name: AGENT_MODULE, code: agentModuleCode(spec, bridge.logFirstRow) });
  if (standardModules.length || docCode.size) {
    const modules: VbaModuleSource[] = [
      { name: workbookCodeName, type: 'document', document: 'workbook', code: docCode.get(workbookCodeName) ?? '' },
      ...sheetCodeNames.map((n): VbaModuleSource => ({ name: n, type: 'document', document: 'sheet', code: docCode.get(n) ?? '' })),
      { name: definitionCodeName, type: 'document', document: 'sheet', code: '' },
      ...standardModules.map((m): VbaModuleSource => ({ name: m.name, type: 'standard', code: m.code })),
    ];
    parts['xl/vbaProject.bin'] = writeVbaProject(modules);
    if (!/vbaProject\.bin/.test(rels)) rels = rels.replace('</Relationships>', relXml('rIdVba', REL_VBA, 'vbaProject.bin') + '</Relationships>');
    ct = ct.replace(/(PartName="\/xl\/workbook\.xml" ContentType=")[^"]+"/, `$1${CT_MACRO_MAIN}"`);
    if (!/vbaProject\.bin/.test(ct)) ct = ct.replace('</Types>', `<Override PartName="/xl/vbaProject.bin" ContentType="${CT_VBA}"/></Types>`);
  } else if (options.macroEnabled) {
    ct = ct.replace(/(PartName="\/xl\/workbook\.xml" ContentType=")[^"]+"/, `$1${CT_MACRO_MAIN}"`);
  }

  // Click actions on shapes, by shape name.
  for (const [shape, macro] of Object.entries(spec.shapeMacros ?? {})) {
    let found = false;
    for (const n of Object.keys(parts).filter((p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p))) {
      const xml = text(n);
      // <xdr:sp macro=""> … <xdr:cNvPr id=".." name="shape">
      const re = new RegExp(`<xdr:sp(?: macro="[^"]*")?((?:(?!<xdr:sp[ >]).)*?<xdr:cNvPr [^>]*name="${shape.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}")`, 's');
      if (re.test(xml)) {
        parts[n] = strToU8(xml.replace(re, (_m, rest: string) => `<xdr:sp macro="[0]!${esc(macro)}"${rest}`));
        found = true;
      }
    }
    if (!found) throw new Error(`Shape not found for macro ${macro}: ${shape}`);
  }

  parts['xl/workbook.xml'] = strToU8(wb);
  parts['xl/_rels/workbook.xml.rels'] = strToU8(rels);
  parts['[Content_Types].xml'] = strToU8(ct);
  return zipSync(parts);
}
