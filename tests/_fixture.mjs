// Builds small synthetic AfO workbooks for the tests. Generic names only: no real workbook,
// query or system ever goes into a fixture.

import { zipSync, gzipSync, strToU8 } from 'fflate';
import CFB from 'cfb';

export function utf16(s) {
  return new Uint8Array(Buffer.from(s, 'utf16le'));
}

/** An AfO compressed block: 4-byte little-endian length, then gzip, all base64. */
export function compressBlock(xml) {
  const raw = strToU8(xml);
  const len = Buffer.alloc(4);
  len.writeUInt32LE(raw.length, 0);
  return Buffer.concat([len, Buffer.from(gzipSync(raw))]).toString('base64');
}

/**
 * A valid MS-OVBA container without real compression: full 4096-byte chunks stored raw, the
 * last partial chunk as literal tokens (which fit only up to 3640 bytes).
 */
export function compressVbaLiteral(bytes) {
  const out = [0x01];
  for (let start = 0; start < bytes.length; start += 4096) {
    const chunk = bytes.subarray(start, start + 4096);
    if (chunk.length === 4096) {
      out.push(0xff, 0x3f, ...chunk);
      continue;
    }
    if (chunk.length > 3640) throw new Error('last chunk too long for a literal-only chunk');
    const body = [];
    for (let i = 0; i < chunk.length; i += 8) {
      body.push(0x00);
      for (let j = i; j < Math.min(i + 8, chunk.length); j++) body.push(chunk[j]);
    }
    const size = body.length + 2;
    const header = 0xb000 | (size - 3);
    out.push(header & 0xff, header >> 8, ...body);
  }
  return Uint8Array.from(out);
}

function record(id, body) {
  const b = Buffer.alloc(6);
  b.writeUInt16LE(id, 0);
  b.writeUInt32LE(body.length, 2);
  return Buffer.concat([b, Buffer.from(body)]);
}

export function vbaProject(modules) {
  const parts = [record(0x0003, Buffer.from([0xe4, 0x04]))]; // code page 1252
  // PROJECTVERSION declares 4 bytes but carries 6.
  const version = Buffer.alloc(12);
  version.writeUInt16LE(0x0009, 0);
  version.writeUInt32LE(4, 2);
  parts.push(version);
  for (const m of modules) {
    parts.push(record(0x0019, Buffer.from(m.name, 'latin1')));
    parts.push(record(0x001a, Buffer.from(m.name, 'latin1')));
    parts.push(record(0x0031, Buffer.from([0, 0, 0, 0])));
    parts.push(record(m.document ? 0x0022 : 0x0021, Buffer.alloc(0)));
    parts.push(record(0x002b, Buffer.alloc(0)));
  }
  const cfb = CFB.utils.cfb_new();
  CFB.utils.cfb_add(cfb, '/VBA/dir', Buffer.from(compressVbaLiteral(Buffer.concat(parts))));
  for (const m of modules) {
    CFB.utils.cfb_add(cfb, `/VBA/${m.name}`, Buffer.from(compressVbaLiteral(Buffer.from(m.source, 'latin1'))));
  }
  return new Uint8Array(CFB.write(cfb, { type: 'buffer' }));
}

const SYSTEM = (id) => `<SystemInfo><SystemId>${id}</SystemId><SystemName>${id} test</SystemName><SystemType>BW_SYSTEM</SystemType><Client>001</Client><Language>EN</Language></SystemInfo>`;

export const DS1_STATE = `<BICS_VIEW version="3.1"><SERVICE_DESCRIPTION><QUERY>QUERY_NAME</QUERY></SERVICE_DESCRIPTION>
<INFO_PROVIDER name="PROVIDER_1" /><QUERY name="QUERY_NAME"><DUE_DATE>20250101</DUE_DATE></QUERY>
<SELECTION_STATE><SELECTION type="CARTESIAN_PRODUCT"><CHARACTERISTICS>
<CHARACTERISTIC name="CHAR_VERSION"><SELECTIONS><SELECTION type="SINGLE_MEMBER"><MEMBER name="PLAN1" type="MEMBER" /></SELECTION></SELECTIONS></CHARACTERISTIC>
<CHARACTERISTIC name="CHAR_EMPTY"><SELECTIONS /></CHARACTERISTIC>
<CHARACTERISTIC name="CHAR_MONTH"><SELECTIONS><SELECTION type="RANGE" operator="BETWEEN"><LOW_VALUE><MEMBER name="202501" type="MEMBER" /></LOW_VALUE><HIGH_VALUE><MEMBER name="202512" type="MEMBER" /></HIGH_VALUE></SELECTION></SELECTIONS></CHARACTERISTIC>
<CHARACTERISTIC name="CHAR_REGION"><SELECTIONS><SELECTION type="RANGE" operator="EQUAL"><SIGN>EXCLUDING</SIGN><VALUE><MEMBER name="R9" type="MEMBER" /></VALUE></SELECTION><SELECTION type="SINGLE_MEMBER"><MEMBER name="NODE_1" type="HIERARCHY_NODE"><NODE_TYPE>X</NODE_TYPE></MEMBER></SELECTION></SELECTIONS></CHARACTERISTIC>
<CHARACTERISTIC name="CHAR_YEAR"><SELECTIONS><SELECTION type="RANGE" operator="GREATER_EQUAL"><SIGN>INCLUDING</SIGN><VALUE><MEMBER name="2025" type="MEMBER" /></VALUE></SELECTION></SELECTIONS></CHARACTERISTIC>
</CHARACTERISTICS></SELECTION></SELECTION_STATE>
<CHARACTERISTICS><CHARACTERISTIC name="CHAR_REGION"><HIERARCHY name="HIER_1" dueDate="00000000" /></CHARACTERISTIC><CHARACTERISTIC name="CHAR_VERSION"><HIERARCHY name="" dueDate="00000000" /></CHARACTERISTIC></CHARACTERISTICS>
<AXES><AXIS name="ROWS"><CHARACTERISTICS><CHARACTERISTIC>CHAR_REGION</CHARACTERISTIC></CHARACTERISTICS></AXIS><AXIS name="COLUMNS"><CHARACTERISTICS><CHARACTERISTIC>KEY_FIGURES</CHARACTERISTIC></CHARACTERISTICS></AXIS><AXIS name="FREE"><CHARACTERISTICS /></AXIS></AXES>
</BICS_VIEW>`;

export const PF1_VARIABLES = `<VARIABLES><VARIABLE name="VAR_VERSION" type="INFO_OBJECT_MEMBER" isStateVariable="false"><SELECTIONS><SELECTION type="SINGLE_MEMBER"><MEMBER name="PLAN1" type="MEMBER" /></SELECTION></SELECTIONS></VARIABLE>
<VARIABLE name="VAR_FACTOR" type="INFO_OBJECT_MEMBER" isStateVariable="false"><SELECTIONS><SELECTION type="SINGLE_MEMBER"><MEMBER type="LITERAL_MEMBER"><PRESENTATION>KEY</PRESENTATION><VALUE>1.5</VALUE></MEMBER></SELECTION></SELECTIONS></VARIABLE></VARIABLES>`;

const PF1_FILTER = `<BICS_VIEW version="3.1"><INFO_PROVIDER name="ALVL_1" /><SELECTION_OBJECT name="!!AALVL_1"><DUE_DATE>20250101</DUE_DATE></SELECTION_OBJECT>
<SELECTION_STATE><SELECTION type="CARTESIAN_PRODUCT"><CHARACTERISTICS><CHARACTERISTIC name="CHAR_VERSION"><SELECTIONS><SELECTION type="SINGLE_MEMBER"><MEMBER name="PLAN1" type="MEMBER" /></SELECTION></SELECTIONS></CHARACTERISTIC></CHARACTERISTICS></SELECTION></SELECTION_STATE></BICS_VIEW>`;

export function afoDefinition({ systemId = 'SYS', query = 'QUERY_NAME' } = {}) {
  return `<Application xmlns="http://www.sap.com/ip/bi/pioneer/excel/application"><com.sap.ip.bi.pioneer>
<Version>12</Version><AAO_Revision>2.8.0.0</AAO_Revision><LastRefreshed>1/1/2025 10:00:00</LastRefreshed><SystemId>${systemId}</SystemId>
<RefreshOnOpen>True</RefreshOnOpen><PlanningModeSetToChangeMode>True</PlanningModeSetToChangeMode><Cleaned>True</Cleaned><StorePromptsInDocument>False</StorePromptsInDocument>
<PromptsVisibility><Variable name="VAR_VERSION" visible="True" /></PromptsVisibility>
<PlanningSystemDesignTime>${SYSTEM(systemId)}<CellLocking>None</CellLocking></PlanningSystemDesignTime>
<DataProviders>
<DataProvider><Key>DS_1</Key><RefreshDataEntryEnabled>True</RefreshDataEntryEnabled><DataSourceInfo><Id>${query}</Id><Name>Input query</Name><Type>Query</Type>${SYSTEM(systemId)}</DataSourceInfo>
<BICS_VIEW_HEADER compressed="true">${compressBlock(DS1_STATE)}</BICS_VIEW_HEADER></DataProvider>
<DataProvider><Key>DS_2</Key><DataSourceInfo><Id>SHARED_QUERY</Id><Name>Shared</Name><Type>Query</Type>${SYSTEM(systemId)}</DataSourceInfo>
<BICS_VIEW_HEADER><TABLE_DESIGN><STYLE_LIST /><OPERATION_LIST /></TABLE_DESIGN></BICS_VIEW_HEADER></DataProvider>
</DataProviders>
<PlanningObjects><PlanningFunctionUsage><Key>PF_1</Key><PlanningObjectInfo><Name>PLAN_FUNC_1</Name><Description>Copy</Description><PlanningType>Function</PlanningType>${SYSTEM(systemId)}</PlanningObjectInfo>
<DataProviderKey>DS_1</DataProviderKey><PlanFunctionVariables><Variables_Header compressed="true">${compressBlock(PF1_VARIABLES)}</Variables_Header></PlanFunctionVariables>
<PlanFilter><BICS_VIEW_HEADER compressed="true">${compressBlock(PF1_FILTER)}</BICS_VIEW_HEADER></PlanFilter>
<ReferenceConfiguration><FilterRefConfig><TechnicalName>CHAR_VERSION</TechnicalName><ReferenceType>Filter</ReferenceType></FilterRefConfig>
<VariableRefConfig><TechnicalName>VAR_FACTOR</TechnicalName><ReferenceType>Cell</ReferenceType><CellAdress>Input!$B$3</CellAdress></VariableRefConfig></ReferenceConfiguration>
</PlanningFunctionUsage></PlanningObjects>
<Items>
<Fdd><Key>Fdd1</Key><Range Name="RequestedRange"><SheetName>Sheet2|guid-input</SheetName><Cell Name="TopLeft"><Col>1</Col><Row>4</Row></Cell><Cell Name="BottomRight"><Col>1</Col><Row>4</Row></Cell></Range><FormulaString>=SAPSetFilterComponent("DS_1", "CHAR_MONTH", "")</FormulaString></Fdd>
<Grid><Key>Crosstab1</Key><Description>Crosstab 1</Description><DataProviderKey>DS_1</DataProviderKey>
<Range Name="RangeOccupied"><SheetName>Sheet2|guid-input</SheetName><Cell Name="TopLeft"><Col>0</Col><Row>9</Row></Cell><Cell Name="BottomRight"><Col>27</Col><Row>20</Row></Cell></Range></Grid>
<Grid><Key>Crosstab2</Key><DataProviderKey>DS_9</DataProviderKey><Range Name="RequestedRange"><SheetName>Sheet2|guid-input</SheetName><Cell Name="TopLeft"><Col>0</Col><Row>30</Row></Cell><Cell Name="BottomRight"><Col>0</Col><Row>30</Row></Cell></Range></Grid>
</Items></com.sap.ip.bi.pioneer></Application>`;
}

const MODULE_SOURCE = [
  'Attribute VB_Name = "Module1"',
  'Option Explicit',
  '',
  'Public Sub RunPlanning()',
  '    Dim r As Long',
  '    r = Application.Run("SAPSetPlanParameter", "PF_1", "VAR_VERSION", "PLAN1")',
  '    r = Application.Run("SAPExecutePlanningFunction", _',
  '        "PF_1")',
  '    \' r = Application.Run("SAPExecuteCommand", "Commented")',
  '    Range("SAPCrosstab1").Select',
  '    SaveData',
  'End Sub',
  '',
  'Private Sub SaveData()',
  '    Call Application.Run("SAPExecuteCommand", "PlanDataSave")',
  '    Debug.Print Application.Run("SapGetVariable", "DS_7", "VAR_VERSION", "Value")',
  'End Sub',
].join('\r\n');

const SHEET_SOURCE = ['Attribute VB_Name = "Sheet2"', 'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"', 'Private Sub Worksheet_Activate()', 'End Sub'].join('\r\n');

const REL = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
const RELS = (...r) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${r.join('')}</Relationships>`;

export function buildWorkbook(opts = {}) {
  const files = {
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="_com.sap.ip.bi.xl.hiddensheet" sheetId="2" state="veryHidden" r:id="rId1"/><sheet name="Input" sheetId="1" r:id="rId2"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(RELS(REL('rId1', 'worksheet', 'worksheets/sheet1.xml'), REL('rId2', 'worksheet', 'worksheets/sheet2.xml'), '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>')),
    'xl/worksheets/sheet1.xml': strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr codeName="Sheet1"/><sheetData/><customProperties><customPr name="_pios_id" r:id="rId1"/><customPr name="serializedData2" r:id="rId2"/></customProperties></worksheet>`),
    'xl/worksheets/_rels/sheet1.xml.rels': strToU8(RELS(REL('rId1', 'customProperty', '../customProperty1.bin'), REL('rId2', 'customProperty', '../customProperty2.bin'))),
    'xl/customProperty1.bin': utf16('_com.sap.ip.bi.xl.hiddensheet|guid-hidden'),
    'xl/customProperty2.bin': utf16(afoDefinition(opts)),
    'xl/worksheets/sheet2.xml': strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr codeName="Sheet2"/><sheetData>
<row r="3"><c r="B3"><f>_xll.SAPGetVariable("DS_1","VAR_VERSION","Value")</f><v>PLAN1</v></c><c r="C3"><f>SUM(A1:A2)</f><v>0</v></c></row>
<row r="4"><c r="B4"><f t="shared" ref="B4:B6" si="0">_xll.SAPGetData("DS_1","KF_1",CONCATENATE("CHAR_REGION=",A4))</f><v>1</v></c></row>
<row r="5"><c r="B5"><f t="shared" si="0"/><v>2</v></c><c r="C5" t="str"><f>_xll.SAPSetFilterComponent("DS_1", "CHAR_MONTH", "")</f><v></v></c></row>
<row r="6"><c r="B6"><f t="shared" si="0"/><v>3</v></c><c r="C6"><f>IF(_xll.SAPGetMember("DS_UNKNOWN","CHAR_REGION")=&quot;A&quot;,1,2)</f><v>2</v></c></row>
</sheetData><drawing r:id="rId2"/><legacyDrawing r:id="rId3"/><customProperties><customPr name="_pios_id" r:id="rId1"/></customProperties></worksheet>`),
    'xl/worksheets/_rels/sheet2.xml.rels': strToU8(RELS(REL('rId1', 'customProperty', '../customProperty3.bin'), REL('rId2', 'drawing', '../drawings/drawing1.xml'), REL('rId3', 'vmlDrawing', '../drawings/vmlDrawing1.vml'))),
    'xl/customProperty3.bin': utf16('Sheet2|guid-input'),
    'xl/drawings/drawing1.xml': strToU8(`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:row>2</xdr:row></xdr:to><xdr:sp macro="[0]!RunPlanning"><xdr:nvSpPr><xdr:cNvPr id="2" name="Rectangle 1"/></xdr:nvSpPr><xdr:txBody><a:p><a:r><a:t>Run planning</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor>
<xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:row>1</xdr:row></xdr:from><xdr:pic macro=""><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 1"/></xdr:nvPicPr></xdr:pic></xdr:oneCellAnchor></xdr:wsDr>`),
    'xl/drawings/vmlDrawing1.vml': strToU8(`<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025" type="#_x0000_t201"><v:textbox><div><font>Old<br>button</font></div></v:textbox><x:ClientData ObjectType="Button"><x:Anchor>1, 0, 7, 0, 3, 0, 8, 0</x:Anchor><x:FmlaMacro>[0]!MissingMacro</x:FmlaMacro></x:ClientData></v:shape></xml>`),
    'xl/vbaProject.bin': vbaProject([
      { name: 'Module1', source: MODULE_SOURCE },
      { name: 'Sheet2', source: SHEET_SOURCE, document: true },
    ]),
  };
  return zipSync(files);
}
