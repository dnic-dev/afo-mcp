// Turn an ordinary Excel workbook into an Analysis for Office workbook: write the AfO definition
// (data sources, crosstabs, planning functions) into a hidden sheet, give every sheet its AfO id,
// add the defined names AfO binds items to, and optionally a VBA project and button macros.
//
// AfO accepts such a workbook and rewrites the definition on its next save. Learned from
// workbooks AfO saved itself: a crosstab needs both RangeOccupied and RequestedRange, and every
// item has a defined name SAP<Key> pointing at its range.

import { randomUUID } from 'node:crypto';
import { gzipSync, strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { AFO_NAMESPACE } from '../workbook/afo.js';
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
  query: string;
  /** The query's InfoProvider, written into the stored state. */
  infoProvider?: string;
  inputReady?: boolean;
  /** Variable values stored with the data source; with them AfO opens without the prompt dialog. */
  prompts?: PromptValue[];
}

export interface GridSpec {
  key: string;
  dataProviderKey: string;
  sheet: string;
  /** Top-left cell, A1 notation. */
  cell: string;
}

export interface PlanningFunctionSpec {
  key: string;
  name: string;
  description?: string;
  dataProviderKey: string;
  aggregationLevel: string;
  /** Characteristics of the aggregation level, listed unrestricted in the ad-hoc filter. */
  characteristics: string[];
  /** Stored variable values, internal format. */
  variables?: Array<{ name: string; value: string }>;
  /** Variables read from a cell at execution time (`Sheet!$C$7`, value in external format). */
  variableCells?: Array<{ name: string; cell: string }>;
}

export interface AfoWorkbookSpec {
  system: AfoSystem;
  refreshOnOpen?: boolean;
  changeModeOnOpen?: boolean;
  storePrompts?: boolean;
  /** Share variables of the same name between data sources and planning objects (AfO default: off). */
  mergeVariables?: boolean;
  afoRevision?: string;
  dataProviders: DataProviderSpec[];
  grids: GridSpec[];
  planningFunctions?: PlanningFunctionSpec[];
  /** Macro per shape name, written as the shape's click action. */
  shapeMacros?: Record<string, string>;
  /** Standard modules of a VBA project to add; turns the workbook into .xlsm. */
  vbaModules?: Array<{ name: string; code: string }>;
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const guid = () => randomUUID().replace(/-/g, '');
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

function systemInfo(s: AfoSystem): string {
  return `<SystemInfo><SystemId>${esc(s.systemId)}</SystemId><SystemName>${esc(s.systemName)}</SystemName><SystemDescription>${esc(s.systemName)}</SystemDescription>`
    + `<SystemType>BW_SYSTEM</SystemType><IsLocal>True</IsLocal><SsoEnabled>False</SsoEnabled><BoeID /><BoeID2>0</BoeID2><IsBoeDataSource>False</IsBoeDataSource>`
    + (s.server ? `<ServerGroup>${esc(s.server)}</ServerGroup>` : '')
    + `<Origin>USEREDIT</Origin>`
    + (s.systemNumber ? `<SystemNumber>${esc(s.systemNumber)}</SystemNumber>` : '')
    + `<Client>${esc(s.client)}</Client><Language>${esc(s.language)}</Language><SncQuality>-1</SncQuality><SncNoSSO>True</SncNoSSO>`
    + (s.uuid ? `<Uuid>${esc(s.uuid)}</Uuid>` : '')
    + `</SystemInfo>`;
}

function range(name: string, sheetId: string, cell: string): string {
  const { row, col } = fromA1(cell);
  const c = `<Cell Name="TopLeft"><Col>${col}</Col><Row>${row}</Row></Cell><Cell Name="BottomRight"><Col>${col}</Col><Row>${row}</Row></Cell>`;
  return `<Range Name="${name}"><SheetName>${esc(sheetId)}</SheetName>${c}</Range>`;
}

function planningFunction(pf: PlanningFunctionSpec, system: string, today: string): string {
  const vars = `<?xml version="1.0" encoding="UTF-8" standalone="no" ?><VARIABLES>${(pf.variables ?? [])
    .map((v) => `<VARIABLE name="${esc(v.name)}" type="INFO_OBJECT_MEMBER" isStateVariable="false"><SELECTIONS><SELECTION type="SINGLE_MEMBER"><MEMBER name="${esc(v.value)}" type="MEMBER" /></SELECTION></SELECTIONS></VARIABLE>`)
    .join('')}</VARIABLES>`;
  const filter = `<?xml version="1.0" encoding="UTF-8" standalone="no" ?><BICS_VIEW version="3.1"><SERVICE_DESCRIPTION><INITIALIZE_PARENT_CONTAINER>X</INITIALIZE_PARENT_CONTAINER><DATA_AREA>DEFAULT</DATA_AREA></SERVICE_DESCRIPTION>`
    + `<INFO_PROVIDER name="${esc(pf.aggregationLevel)}" /><SELECTION_OBJECT name="!!A${esc(pf.aggregationLevel)}"><DUE_DATE>${today}</DUE_DATE></SELECTION_OBJECT>`
    + `<SELECTION_SPACE_SORTING type="DEFAULT" /><SELECTION_STATE_SORTING type="DEFAULT" /><SELECTION_STATE><SELECTION type="CARTESIAN_PRODUCT"><CHARACTERISTICS>`
    + pf.characteristics.map((c) => `<CHARACTERISTIC name="${esc(c)}"><SELECTIONS /></CHARACTERISTIC>`).join('')
    + `</CHARACTERISTICS></SELECTION></SELECTION_STATE></BICS_VIEW>`;
  const refs = (pf.variableCells ?? [])
    .map((v) => `<VariableRefConfig><TechnicalName>${esc(v.name)}</TechnicalName><ReferenceType>Cell</ReferenceType><CellAdress>${esc(v.cell)}</CellAdress></VariableRefConfig>`)
    .join('');
  return `<PlanningFunctionUsage><Key>${esc(pf.key)}</Key><PlanningObjectInfo><Name>${esc(pf.name)}</Name><Description>${esc(pf.description ?? pf.name)}</Description><PlanningType>Function</PlanningType>${system}</PlanningObjectInfo>`
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

/**
 * A minimal stored navigation state. DOCUMENT_APPLICATION_ASSIGNED tells AfO that no BW
 * document application is assigned; without any stored state AfO asks the server, and on a
 * system without document storage that ends in an exception once the crosstab is input-ready.
 */
function stateXml(dp: DataProviderSpec, day: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="no" ?><BICS_VIEW version="3.1"><SERVICE_DESCRIPTION><DATA_AREA>DEFAULT</DATA_AREA><QUERY>${esc(dp.query)}</QUERY></SERVICE_DESCRIPTION><MODULES />`
    + (dp.infoProvider ? `<INFO_PROVIDER name="${esc(dp.infoProvider)}" />` : '')
    + `<QUERY name="${esc(dp.query)}"><DUE_DATE>${day}</DUE_DATE></QUERY>`
    + (dp.prompts?.length ? `<VARIABLES>${dp.prompts.map(promptXml).join('')}</VARIABLES>` : '')
    + (dp.inputReady ?? true ? '<IS_INPUT_ENABLED>true</IS_INPUT_ENABLED>' : '')
    + `<DOCUMENT_APPLICATION_ASSIGNED /></BICS_VIEW>`;
}

/** The `<Application>` document. `sheetIds` maps sheet names to their AfO ids. */
export function buildDefinition(spec: AfoWorkbookSpec, sheetIds: Map<string, string>, today = new Date()): string {
  const sys = systemInfo(spec.system);
  const day = today.toISOString().slice(0, 10).replace(/-/g, '');
  const sheetId = (name: string) => {
    const id = sheetIds.get(name);
    if (!id) throw new Error(`Sheet not found: ${name}`);
    return id;
  };
  const dps = spec.dataProviders.map((dp) =>
    `<DataProvider><Key>${esc(dp.key)}</Key><FillableNumber>1</FillableNumber><RefreshDataEntryEnabled>${tf(dp.inputReady ?? true)}</RefreshDataEntryEnabled><Cleaned>False</Cleaned>`
    + `<DataSourceInfo><Id>${esc(dp.query)}</Id><Name>${esc(dp.query)}</Name><DisplayName>${esc(dp.query)}</DisplayName><Type>Query</Type><HasChildren>False</HasChildren><IsFormulaDriven>False</IsFormulaDriven><ShowScalingFactor>False</ShowScalingFactor><ShowUnits>False</ShowUnits>${sys}</DataSourceInfo>`
    + `<DimensionSynchronization /><BICS_VIEW_HEADER compressed="true">${compressBlock(stateXml(dp, day))}</BICS_VIEW_HEADER><Comments /></DataProvider>`).join('');
  const prompts = spec.dataProviders.flatMap((dp) => (dp.prompts ?? []).map((p) => `<Variable dataprovider="${esc(dp.key)}" name="${esc(p.name)}" visible="True" />`)).join('');
  const pfs = (spec.planningFunctions ?? []).map((pf) => planningFunction(pf, sys, day)).join('');
  const grids = spec.grids.map((g, i) => {
    const id = sheetId(g.sheet);
    return `<Grid><Key>${esc(g.key)}</Key><Description>${esc(g.key)}</Description><ShapeId>${i + 2}</ShapeId><DataProviderKey>${esc(g.dataProviderKey)}</DataProviderKey>`
      + range('RangeOccupied', id, g.cell) + range('RequestedRange', id, g.cell)
      + `<ApplyDefaultFormats>True</ApplyDefaultFormats><ApplyNumberFormats>True</ApplyNumberFormats><AutoFit>True</AutoFit><DisplayRepeatedTexts>False</DisplayRepeatedTexts><DisplayRepeatedTitles>False</DisplayRepeatedTitles>`
      + `<NumberOfNewLines>5</NumberOfNewLines><NewLinesCheckMode>NONE</NewLinesCheckMode><DisplayScalingUnitsInHeader>False</DisplayScalingUnitsInHeader><HierarchySymbolType>1</HierarchySymbolType><WrapHeaderTextWithLinebreak>True</WrapHeaderTextWithLinebreak></Grid>`;
  }).join('');
  return `<Application xmlns="${AFO_NAMESPACE}"><com.sap.ip.bi.pioneer><Version>12</Version><AAO_Revision>${esc(spec.afoRevision ?? '2.8.2301.1605')}</AAO_Revision><SystemId>${esc(spec.system.systemId)}</SystemId>`
    + `<RefreshOnOpen>${tf(spec.refreshOnOpen)}</RefreshOnOpen><PlanningModeSetToChangeMode>${tf(spec.changeModeOnOpen)}</PlanningModeSetToChangeMode><Cleaned>True</Cleaned>`
    + `<ForcePromptOnInitialRefresh>False</ForcePromptOnInitialRefresh><StorePromptsInDocument>${tf(spec.storePrompts)}</StorePromptsInDocument><MergeVariables>${tf(spec.mergeVariables)}</MergeVariables><RefreshPlanningObjectsOnRefreshAll>True</RefreshPlanningObjectsOnRefreshAll>`
    + (prompts ? `<PromptsVisibility>${prompts}</PromptsVisibility>` : '<PromptsVisibility />') + `<StyleSetName>SAP Blue</StyleSetName><PlanningSystemDesignTime>${sys}<CellLocking>None</CellLocking></PlanningSystemDesignTime><BExProperties />`
    + `<DataProviders>${dps}</DataProviders>` + (pfs ? `<PlanningObjects>${pfs}</PlanningObjects>` : '') + `<Items>${grids}</Items></com.sap.ip.bi.pioneer></Application>`;
}

const REL_CP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty';
const REL_WS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const REL_VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
const CT_WS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CT_CP = 'application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty';
const CT_MACRO_MAIN = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
const CT_VBA = 'application/vnd.ms-office.vbaProject';

const utf16 = (s: string) => new Uint8Array(Buffer.from(s, 'utf16le'));
const relXml = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
const relsDoc = (items: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;

function addRels(parts: Record<string, Uint8Array>, relsPath: string, items: string) {
  parts[relsPath] = strToU8(parts[relsPath] ? strFromU8(parts[relsPath]).replace('</Relationships>', `${items}</Relationships>`) : relsDoc(items));
}

function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function absolute(cell: string): string {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(cell.toUpperCase());
  return m ? `$${m[1]}$${m[2]}` : cell;
}

/** Make an existing .xlsx/.xlsm an AfO workbook according to the spec. */
export function makeAfoWorkbook(source: Uint8Array, spec: AfoWorkbookSpec): Uint8Array {
  const parts: Record<string, Uint8Array> = unzipSync(source);
  const text = (n: string) => strFromU8(parts[n]);
  let wb = text('xl/workbook.xml');
  const wbRels = text('xl/_rels/workbook.xml.rels');
  const relTarget = new Map([...wbRels.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));

  // Code names: the workbook and every sheet need one for AfO ids and VBA document modules.
  if (!/<workbookPr[^>]*codeName=/.test(wb)) {
    // workbookPr precedes bookViews and sheets in the schema order.
    wb = /<workbookPr/.test(wb)
      ? wb.replace('<workbookPr', '<workbookPr codeName="ThisWorkbook"')
      : wb.replace(/<bookViews|<sheets>/, (m) => `<workbookPr codeName="ThisWorkbook"/>${m}`);
  }
  const workbookCodeName = /<workbookPr[^>]*codeName="([^"]+)"/.exec(wb)![1];

  let cpIndex = 900;
  const sheetIds = new Map<string, string>();
  const sheetCodeNames: string[] = [];
  const sheets = [...wb.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)];
  sheets.forEach(([, name, rid], i) => {
    const part = 'xl/' + relTarget.get(rid)!.replace(/^\/?xl\//, '');
    let xml = text(part);
    let codeName = /<sheetPr[^>]*codeName="([^"]+)"/.exec(xml)?.[1];
    if (!codeName) {
      codeName = `Sheet${i + 1}`;
      xml = /<sheetPr/.test(xml) ? xml.replace('<sheetPr', `<sheetPr codeName="${codeName}"`) : xml.replace(/(<worksheet[^>]*>)/, `$1<sheetPr codeName="${codeName}"/>`);
    }
    sheetCodeNames.push(codeName);
    const afoId = `${codeName}|${guid()}`;
    sheetIds.set(name.replace(/&amp;/g, '&'), afoId);
    const a = ++cpIndex, b = ++cpIndex;
    parts[`xl/customProperty${a}.bin`] = utf16(afoId);
    parts[`xl/customProperty${b}.bin`] = utf16('AoWorksheetType');
    const cps = `<customProperties><customPr name="_pios_id" r:id="rIdAfo${a}"/><customPr name="CofWorksheetType" r:id="rIdAfo${b}"/></customProperties>`;
    // customProperties comes after pageMargins/pageSetup/headerFooter and before drawing.
    if (/<drawing |<legacyDrawing /.test(xml)) xml = xml.replace(/(<drawing |<legacyDrawing )/, `${cps}$1`);
    else xml = xml.replace('</worksheet>', `${cps}</worksheet>`);
    if (!/xmlns:r=/.test(xml)) xml = xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    parts[part] = strToU8(xml);
    const relsPath = part.replace(/([^/]+)$/, '_rels/$1.rels');
    addRels(parts, relsPath, relXml(`rIdAfo${a}`, REL_CP, `../customProperty${a}.bin`) + relXml(`rIdAfo${b}`, REL_CP, `../customProperty${b}.bin`));
  });

  // Hidden sheet with the definition.
  const definition = buildDefinition(spec, sheetIds);
  const h1 = ++cpIndex, h2 = ++cpIndex, h3 = ++cpIndex;
  parts[`xl/customProperty${h1}.bin`] = utf16(`_com.sap.ip.bi.xl.hiddensheet|${guid()}`);
  parts[`xl/customProperty${h2}.bin`] = utf16('AoWorksheetType');
  parts[`xl/customProperty${h3}.bin`] = utf16(definition);
  parts['xl/worksheets/sheet900.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr codeName="AfoDefinition"/><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.787401575" bottom="0.787401575" header="0.3" footer="0.3"/><customProperties><customPr name="_pios_id" r:id="rId1"/><customPr name="CofWorksheetType" r:id="rId2"/><customPr name="serializedData2" r:id="rId3"/></customProperties></worksheet>`);
  parts['xl/worksheets/_rels/sheet900.xml.rels'] = strToU8(relsDoc(relXml('rId1', REL_CP, `../customProperty${h1}.bin`) + relXml('rId2', REL_CP, `../customProperty${h2}.bin`) + relXml('rId3', REL_CP, `../customProperty${h3}.bin`)));
  wb = wb.replace('</sheets>', `<sheet name="_com.sap.ip.bi.xl.hiddensheet" sheetId="900" state="veryHidden" r:id="rIdAfoSheet"/></sheets>`);
  let rels = wbRels.replace('</Relationships>', relXml('rIdAfoSheet', REL_WS, 'worksheets/sheet900.xml') + '</Relationships>');

  // Defined names SAP<Key> for every item.
  const names = spec.grids.map((g) => `<definedName name="SAP${esc(g.key)}">${esc(quoteSheet(g.sheet))}!${absolute(g.cell)}</definedName>`).join('');
  // Replacer functions throughout: names like $B$16 would otherwise be read as $1 back-references.
  if (names) wb = /<definedNames>/.test(wb) ? wb.replace('<definedNames>', () => `<definedNames>${names}`) : wb.replace('</sheets>', () => `</sheets><definedNames>${names}</definedNames>`);

  // One override per custom property part: a workbook may already map ".bin" to another type
  // (printer settings), and Excel drops parts whose type does not match.
  let ct = text('[Content_Types].xml');
  const cpOverrides = Object.keys(parts)
    .filter((n) => /^xl\/customProperty9\d\d\.bin$/.test(n))
    .map((n) => `<Override PartName="/${n}" ContentType="${CT_CP}"/>`)
    .join('');
  ct = ct.replace('</Types>', `${cpOverrides}<Override PartName="/xl/worksheets/sheet900.xml" ContentType="${CT_WS}"/></Types>`);

  // VBA project: document modules for the workbook and each sheet, plus the standard modules.
  if (spec.vbaModules?.length) {
    const modules: VbaModuleSource[] = [
      { name: workbookCodeName, type: 'document', document: 'workbook', code: '' },
      ...sheetCodeNames.map((n): VbaModuleSource => ({ name: n, type: 'document', document: 'sheet', code: '' })),
      { name: 'AfoDefinition', type: 'document', document: 'sheet', code: '' },
      ...spec.vbaModules.map((m): VbaModuleSource => ({ name: m.name, type: 'standard', code: m.code })),
    ];
    parts['xl/vbaProject.bin'] = writeVbaProject(modules);
    if (!/vbaProject\.bin/.test(rels)) rels = rels.replace('</Relationships>', relXml('rIdVba', REL_VBA, 'vbaProject.bin') + '</Relationships>');
    ct = ct.replace(/(PartName="\/xl\/workbook\.xml" ContentType=")[^"]+"/, `$1${CT_MACRO_MAIN}"`);
    if (!/vbaProject\.bin/.test(ct)) ct = ct.replace('</Types>', `<Override PartName="/xl/vbaProject.bin" ContentType="${CT_VBA}"/></Types>`);
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
