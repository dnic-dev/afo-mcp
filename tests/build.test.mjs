import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { writeVbaProject, compressVba, encodeCp1252 } from '../dist/build/vba.js';
import { makeAfoWorkbook, buildDefinition, fromA1 } from '../dist/build/afo.js';
import { readVbaProject, decompressVba } from '../dist/workbook/vba.js';
import { readWorkbook } from '../dist/workbook/read.js';

test('compressVba round-trips through the decompressor, across chunk sizes', () => {
  for (const n of [0, 1, 3640, 3641, 4096, 4097, 5000, 12000]) {
    const input = Buffer.from(Array.from({ length: n }, (_, i) => 32 + (i % 90)));
    assert.deepEqual(Buffer.from(decompressVba(compressVba(input))), input, `length ${n}`);
  }
  // Text without repetitions and random bytes exercise literal-heavy and raw chunks.
  const random = Buffer.from(Array.from({ length: 9000 }, (_, i) => (i * 7919 + (i >> 3) * 104729) & 0xff));
  assert.deepEqual(Buffer.from(decompressVba(compressVba(random))), random);
});

test('compressVba compresses source text with copy tokens and no raw chunks', () => {
  const source = Buffer.from(Array.from({ length: 60 }, (_, i) => `    Application.Run "SAPSetVariable", "VAR_${i}", "X", "INPUT_STRING", "DS_1"\r\n`).join(''), 'latin1');
  const packed = compressVba(source);
  assert.ok(packed.length < source.length / 2, `expected real compression, got ${packed.length} of ${source.length}`);
  // Every chunk header carries the compressed flag (bit 15).
  let pos = 1;
  while (pos < packed.length) {
    const header = packed[pos] | (packed[pos + 1] << 8);
    assert.ok(header & 0x8000, 'chunk must be compressed');
    pos += (header & 0x0fff) + 3;
  }
  assert.deepEqual(Buffer.from(decompressVba(packed)), source);
});

test('compressVba matches the MS-OVBA specification example in size', () => {
  // Byte equality is not required: among equally long matches the spec example picks a farther
  // one, this compressor the nearest. Both decompress to the same text.
  const input = Buffer.from('#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa', 'latin1');
  const packed = compressVba(input);
  assert.deepEqual(Buffer.from(decompressVba(packed)), input);
  assert.equal(packed.length, [
    0x01, 0x2f, 0xb0, 0x00, 0x23, 0x61, 0x61, 0x61, 0x62, 0x63, 0x64, 0x65, 0x82, 0x66, 0x00, 0x70,
    0x61, 0x67, 0x68, 0x69, 0x6a, 0x01, 0x38, 0x08, 0x61, 0x6b, 0x6c, 0x00, 0x30, 0x6d, 0x6e, 0x6f,
    0x70, 0x06, 0x71, 0x02, 0x70, 0x04, 0x10, 0x72, 0x73, 0x74, 0x75, 0x76, 0x10, 0x77, 0x78, 0x79,
    0x7a, 0x00, 0x3c,
  ].length);
});

test('encodeCp1252 keeps umlauts and maps typographic characters', () => {
  assert.deepEqual([...encodeCp1252('ÄöüßÖ„“–€')], [0xc4, 0xf6, 0xfc, 0xdf, 0xd6, 0x84, 0x93, 0x96, 0x80]);
});

test('writeVbaProject produces a project the reader understands', () => {
  const bin = writeVbaProject([
    { name: 'ThisWorkbook', type: 'document', document: 'workbook', code: '' },
    { name: 'Sheet1', type: 'document', document: 'sheet', code: '' },
    { name: 'Module1', type: 'standard', code: 'Public Sub Run()\n    x = Application.Run("SAPExecutePlanningFunction", "PF_1") \' Größe\nEnd Sub\n' },
  ]);
  const project = readVbaProject(bin, true);
  assert.equal(project.error, undefined);
  assert.deepEqual(project.modules.map((m) => [m.name, m.type]), [['ThisWorkbook', 'document'], ['Sheet1', 'document'], ['Module1', 'standard']]);
  const [run] = project.modules[2].procedures;
  assert.equal(run.name, 'Run');
  assert.deepEqual(run.afoCalls.map((c) => `${c.function}:${c.firstArgument}`), ['SAPExecutePlanningFunction:PF_1']);
  assert.match(project.modules[2].source, /' Größe/);
});

test('fromA1 converts to zero-based coordinates', () => {
  assert.deepEqual(fromA1('B16'), { row: 15, col: 1 });
  assert.deepEqual(fromA1('$AB$10'), { row: 9, col: 27 });
  assert.throws(() => fromA1('R1C1'));
});

const SYSTEM = { systemId: 'SYS', systemName: 'SYS - Test', client: '001', language: 'EN', server: 'host', systemNumber: '00' };

const SPEC = {
  system: SYSTEM,
  refreshOnOpen: true,
  changeModeOnOpen: true,
  storePrompts: true,
  dataProviders: [{ key: 'DS_1', query: 'QUERY_NAME', infoProvider: 'ALVL_1', prompts: [{ name: 'VAR_VERSION', literal: 'PLAN1' }, { name: 'VAR_PERIOD', low: '202501', high: '202512' }] }],
  grids: [{ key: 'Crosstab1', dataProviderKey: 'DS_1', sheet: 'Plan Sheet', cell: 'B16' }],
  planningFunctions: [{
    key: 'PF_1', name: 'PLAN_FUNC_1', dataProviderKey: 'DS_1', aggregationLevel: 'ALVL_1', characteristics: ['CHAR_A', 'CHAR_B'],
    variables: [{ name: 'VAR_SOURCE', value: '202401' }],
    variableCells: [{ name: 'VAR_SOURCE', cell: "'Plan Sheet'!$C$10" }],
  }],
  shapeMacros: { btnRun: 'Module1.Run' },
  vbaModules: [{ name: 'Module1', code: 'Public Sub Run()\n    Application.Run "SAPExecutePlanningFunction", "PF_1"\nEnd Sub\n' }],
};

/** A small workbook as Excel writes it: one sheet with a drawing, ".bin" mapped to printer settings. */
function baseWorkbook() {
  const rel = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
  const rels = (...r) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${r.join('')}</Relationships>`;
  return zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'),
    '_rels/.rels': strToU8(rels('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>')),
    'xl/workbook.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr defaultThemeVersion="166925"/><bookViews><workbookView/></bookViews><sheets><sheet name="Plan Sheet" sheetId="1" r:id="rId1"/></sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': strToU8(rels(rel('rId1', 'worksheet', 'worksheets/sheet1.xml'))),
    'xl/worksheets/sheet1.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></worksheet>'),
    'xl/worksheets/_rels/sheet1.xml.rels': strToU8(rels(rel('rId1', 'drawing', '../drawings/drawing1.xml'))),
    'xl/drawings/drawing1.xml': strToU8('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:row>8</xdr:row></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:row>9</xdr:row></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="btnRun"/></xdr:nvSpPr><xdr:txBody><a:p><a:r><a:t>Ausführen</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>'),
  });
}

test('makeAfoWorkbook writes a workbook the reader sees as a complete AfO workbook', () => {
  const dir = mkdtempSync(join(tmpdir(), 'afo-build-'));
  try {
    const file = join(dir, 'built.xlsm');
    const bytes = makeAfoWorkbook(baseWorkbook(), SPEC);
    writeFileSync(file, bytes);
    const d = readWorkbook(file, { macroSource: true });
    assert.equal(d.afo, true);
    assert.deepEqual(d.issues, []);
    assert.equal(d.settings.settings.RefreshOnOpen, 'True');
    assert.equal(d.settings.settings.PlanningModeSetToChangeMode, 'True');
    assert.equal(d.settings.settings.MergeVariables, 'False');

    const [ds] = d.dataProviders;
    assert.equal(ds.stateStored, true);
    assert.equal(ds.state.infoProvider, 'ALVL_1');
    assert.deepEqual(ds.state.variables.map((v) => [v.name, v.values]), [['VAR_VERSION', ['"PLAN1"']], ['VAR_PERIOD', ['202501 .. 202512']]]);

    const [pf] = d.planningObjects;
    assert.equal(pf.aggregationLevel, 'ALVL_1');
    assert.equal(pf.filterObject, '!!AALVL_1');
    assert.deepEqual(pf.variables.map((v) => [v.name, v.values]), [['VAR_SOURCE', ['202401']]]);
    assert.deepEqual(pf.references, [{ technicalName: 'VAR_SOURCE', object: 'Variable', referenceType: 'Cell', cell: "'Plan Sheet'!$C$10" }]);

    assert.deepEqual(d.items.grids.map((g) => [g.key, g.sheet, g.address]), [['Crosstab1', 'Plan Sheet', 'B16']]);
    assert.deepEqual(d.macros.buttons.map((b) => [b.name, b.procedure]), [['btnRun', 'Module1.Run']]);
    assert.deepEqual(d.macros.modules.map((m) => m.name), ['ThisWorkbook', 'Sheet1', 'AfoDefinition', 'Module1']);

    const parts = unzipSync(bytes);
    const ct = strFromU8(parts['[Content_Types].xml']);
    assert.match(ct, /sheet\.macroEnabled\.main\+xml/);
    assert.match(ct, /PartName="\/xl\/vbaProject\.bin"/);
    // Every custom property part has its own override; the ".bin" default stays printer settings.
    for (const n of Object.keys(parts).filter((p) => /customProperty\d+\.bin$/.test(p))) assert.ok(ct.includes(`PartName="/${n}"`), n);
    const wb = strFromU8(parts['xl/workbook.xml']);
    assert.match(wb, /<definedName name="SAPCrosstab1">'Plan Sheet'!\$B\$16<\/definedName>/);
    assert.match(wb, /<workbookPr codeName="ThisWorkbook"/);
    assert.ok(wb.indexOf('<workbookPr') < wb.indexOf('<bookViews'));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('buildDefinition and makeAfoWorkbook reject unknown sheets and shapes', () => {
  assert.throws(() => buildDefinition({ ...SPEC, grids: [{ key: 'X', dataProviderKey: 'DS_1', sheet: 'Nope', cell: 'A1' }] }, new Map()), /Sheet not found: Nope/);
  assert.throws(() => makeAfoWorkbook(baseWorkbook(), { ...SPEC, shapeMacros: { missing: 'M.Run' } }), /Shape not found/);
});

const { buildWorkbook, validateSpec, splitCellRef } = await import('../dist/build/build.js');

function withDir(fn) {
  const dir = mkdtempSync(join(tmpdir(), 'afo-build-'));
  try { return fn(dir); } finally { rmSync(dir, { recursive: true, force: true }); }
}

const MINIMAL_SPEC = { ...SPEC, shapeMacros: undefined };

test('buildWorkbook writes a minimal layout with sheets, cells and defined names', () => withDir((dir) => {
  const output = join(dir, 'minimal.xlsm');
  const result = buildWorkbook(MINIMAL_SPEC, {
    output,
    layout: { sheets: ['Overview'], cells: [{ sheet: 'Plan Sheet', cell: 'C10', value: '01.2024', name: 'SourceMonth' }, { sheet: 'Plan Sheet', cell: 'B10', value: 'Source month' }, { sheet: 'Plan Sheet', cell: 'C11', value: 12 }] },
  });
  assert.deepEqual(result.issues, []);
  assert.deepEqual(result.warnings, []);
  assert.deepEqual(result.sheets, ['Overview', 'Plan Sheet']);
  assert.deepEqual(result.dataSources, ['DS_1=QUERY_NAME']);
  assert.deepEqual(result.crosstabs, ['Crosstab1 (DS_1) at Plan Sheet!B16']);
  assert.deepEqual(result.vbaModules, ['Module1']);

  const parts = unzipSync(readFileSync(output));
  const sheet = strFromU8(parts['xl/worksheets/sheet2.xml']);
  assert.match(sheet, /<row r="10"><c r="B10" s="1" t="inlineStr"><is><t xml:space="preserve">Source month<\/t><\/is><\/c><c r="C10" s="1" t="inlineStr"><is><t xml:space="preserve">01\.2024<\/t>/);
  assert.match(sheet, /<c r="C11"><v>12<\/v><\/c>/);
  const wb = strFromU8(parts['xl/workbook.xml']);
  assert.match(wb, /<definedName name="SourceMonth">'Plan Sheet'!\$C\$10<\/definedName>/);
  assert.match(wb, /<definedName name="SAPCrosstab1">'Plan Sheet'!\$B\$16<\/definedName>/);
}));

test('buildWorkbook marks an .xlsm without VBA as macro-enabled and an .xlsx as plain', () => withDir((dir) => {
  const spec = { ...MINIMAL_SPEC, vbaModules: undefined };
  buildWorkbook(spec, { output: join(dir, 'a.xlsm') });
  buildWorkbook(spec, { output: join(dir, 'b.xlsx') });
  const ct = (f) => strFromU8(unzipSync(readFileSync(join(dir, f)))['[Content_Types].xml']);
  assert.match(ct('a.xlsm'), /sheet\.macroEnabled\.main\+xml/);
  assert.doesNotMatch(ct('b.xlsx'), /macroEnabled/);
  assert.throws(() => buildWorkbook(MINIMAL_SPEC, { output: join(dir, 'c.xlsx') }), /VBA modules need an \.xlsm output/);
}));

test('buildWorkbook uses a template and finds the button shapes in it', () => withDir((dir) => {
  const template = join(dir, 'layout.xlsx');
  writeFileSync(template, baseWorkbook());
  const result = buildWorkbook(SPEC, { output: join(dir, 'built.xlsm'), template });
  assert.deepEqual(result.issues, []);
  assert.deepEqual(result.buttons, ['btnRun → Module1.Run']);
  assert.equal(result.base, template);

  // A template that is already an AfO workbook is taken with a warning; a missing sheet is refused.
  const again = buildWorkbook(SPEC, { output: join(dir, 'again.xlsm'), template: join(dir, 'built.xlsm') });
  assert.ok(again.warnings.some((w) => /AfO definition is replaced/.test(w)), again.warnings.join(' | '));
  assert.deepEqual(again.issues, []);
  assert.throws(() => buildWorkbook({ ...SPEC, grids: [{ key: 'Crosstab1', dataProviderKey: 'DS_1', sheet: 'Other', cell: 'A1' }] }, { output: join(dir, 'x.xlsm'), template }), /Sheet\(s\) not in the template: Other\. Template sheets: Plan Sheet\./);
}));

test('buildWorkbook refuses to overwrite unless asked and reads VBA from files', () => withDir((dir) => {
  const output = join(dir, 'out.xlsm');
  writeFileSync(join(dir, 'Module1.bas'), '﻿Public Sub Run()\r\n    Application.Run "SAPExecutePlanningFunction", "PF_1"\r\nEnd Sub\r\n');
  const spec = { ...MINIMAL_SPEC, vbaModules: [{ name: 'Module1', file: 'Module1.bas' }] };
  assert.throws(() => buildWorkbook(spec, { output }), /relative/);
  buildWorkbook(spec, { output, baseDir: dir });
  assert.throws(() => buildWorkbook(spec, { output, baseDir: dir }), /exists/);
  const result = buildWorkbook(spec, { output, baseDir: dir, overwrite: true });
  assert.deepEqual(result.issues, []);
  assert.deepEqual(result.vbaModules, ['Module1']);
}));

test('validateSpec reports every problem at once', () => {
  const errors = validateSpec({
    system: { systemId: 'SYS', systemName: '', client: '001', language: 'EN' },
    dataProviders: [{ key: 'DS_1', query: 'Q', prompts: [{ name: 'V1', member: 'A', literal: 'B' }] }, { key: 'ds_1', query: 'Q2' }],
    grids: [{ key: 'Crosstab1', dataProviderKey: 'DS_9', sheet: 'S', cell: 'R1C1' }],
    planningFunctions: [{ key: 'PF_1', name: 'F', dataProviderKey: 'DS_1', aggregationLevel: 'A', characteristics: ['C1'], filter: { C2: ['X'] }, variableCells: [{ name: 'V', cell: '$C$9' }] }],
    shapeMacros: { btn: 'Missing.Run' },
    vbaModules: [{ name: 'Sheet1', code: '' }, { name: 'Module1' }],
  }, { cells: [{ sheet: 'S', cell: 'A1', value: 'x', name: 'B2' }] });
  for (const re of [/system\.systemName/, /exactly one of member, literal or low/, /key ds_1 is already used/, /DS_9 is not a data source key/, /cell must be an A1 reference/, /filter on C2/, /must name sheet and cell/, /module Missing is not among vbaModules/, /Sheet1 is reserved/, /give one of code, file or from/, /B2 is not usable as a defined name/]) {
    assert.ok(errors.some((e) => re.test(e)), `missing ${re}: ${errors.join(' | ')}`);
  }
});

test('splitCellRef handles quoted sheet names', () => {
  assert.deepEqual(splitCellRef("'Plan ''A'''!$C$9"), { sheet: "Plan 'A'", cell: '$C$9' });
  assert.deepEqual(splitCellRef('Planning!C9'), { sheet: 'Planning', cell: 'C9' });
  assert.equal(splitCellRef('$C$9'), undefined);
});

test('agentBridge adds the bridge sheet, defined names, the AfoAgent module and the event handlers', () => withDir((dir) => {
  const output = join(dir, 'agent.xlsm');
  const spec = {
    ...MINIMAL_SPEC,
    documentModules: { ThisWorkbook: 'Option Explicit\r\n\r\nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)\r\nEnd Sub\r\n' },
    agentBridge: { commands: [{ name: 'RUN_IT', macro: 'Module1.Run', description: 'Run the planning function' }] },
  };
  const result = buildWorkbook(spec, { output });
  assert.deepEqual(result.issues, []);
  assert.deepEqual(result.sheets, ['Plan Sheet', 'AgentLog']);
  assert.deepEqual(result.vbaModules, ['Module1', 'AfoAgent']);

  const parts = unzipSync(readFileSync(output));
  const wb = strFromU8(parts['xl/workbook.xml']);
  for (const [name, cell] of [['AfoAgentCommand', '$B$1'], ['AfoAgentResult', '$B$3'], ['AfoAgentStatus', '$B$4']]) {
    assert.ok(wb.includes(`<definedName name="${name}">AgentLog!${cell}</definedName>`), name);
  }
  const sheet = strFromU8(parts['xl/worksheets/afoAgent.xml']);
  assert.match(sheet, /<c r="B2"><f>LEN\(B1\)<\/f><\/c>/);
  assert.match(sheet, /RUN_IT/);
  assert.match(sheet, /Data source on query QUERY_NAME/);
  assert.match(strFromU8(parts['[Content_Types].xml']), /PartName="\/xl\/worksheets\/afoAgent\.xml"/);

  const d = readWorkbook(output, { macroSource: true });
  const thisWorkbook = d.macros.modules.find((m) => m.name === 'ThisWorkbook').source;
  assert.match(thisWorkbook, /Workbook_BeforeClose[\s\S]*Workbook_Open\(\)\s+AfoAgent\.MarkActive[\s\S]*Workbook_SheetCalculate/);
  assert.equal((thisWorkbook.match(/Option Explicit/g) ?? []).length, 1);
  const agent = d.macros.modules.find((m) => m.name === 'AfoAgent').source;
  assert.match(agent, /Case "RUN_IT": RunOne = RunMacro\("Module1\.Run"\)/);
  assert.match(agent, /Private Const LOG_FIRST_ROW As Long = \d+/);
  assert.match(agent, /Case "PVAR": RunOne = DoPlanVar\(rest\)/);
  assert.match(agent, /Case "API": RunOne = DoApi\(rest\)/);
  assert.match(agent, /SAPLogon is not allowed/);
  assert.match(sheet, /API SAPFunction arg1\|arg2/);
  assert.match(agent, /PutFormula "Q2", "=SAPListOf\(""DATASOURCES""\)"/);
  assert.match(sheet, /Ask the user before SAVE/);
  assert.match(sheet, /SAPListOfVariables/);
  assert.match(sheet, /<c r="Q1" t="inlineStr"><is><t xml:space="preserve">Data sources \(live\)/);
  assert.match(wb, /<sheet name="AgentLog" sheetId="2" r:id="rIdAfoAgent" state="hidden"\/>/);
  assert.ok(d.macros.modules.some((m) => m.name === 'AfoAgentSheet'));
}));

test('agentBridge validation names conflicts with built-ins, modules and handlers', () => withDir((dir) => {
  const spec = {
    ...MINIMAL_SPEC,
    documentModules: { ThisWorkbook: 'Private Sub Workbook_Open()\r\nEnd Sub\r\n' },
    agentBridge: { sheet: 'Plan Sheet', commands: [{ name: 'SAVE', macro: 'Module1.Run' }, { name: 'go', macro: 'Missing.Run' }] },
  };
  try {
    buildWorkbook(spec, { output: join(dir, 'x.xlsm') });
    assert.fail('expected an error');
  } catch (err) {
    for (const re of [/SAVE is a built-in command/, /name must be upper case/, /module Missing is not among vbaModules/, /defines Workbook_Open/, /also a layout or crosstab sheet/]) assert.match(err.message, re);
  }
  assert.throws(() => buildWorkbook({ ...MINIMAL_SPEC, vbaModules: undefined, agentBridge: {} }, { output: join(dir, 'y.xlsx') }), /\.xlsm/);
  buildWorkbook({ ...MINIMAL_SPEC, agentBridge: { hidden: false } }, { output: join(dir, 'h.xlsm') });
  assert.match(strFromU8(unzipSync(readFileSync(join(dir, 'h.xlsm')))['xl/workbook.xml']), /<sheet name="AgentLog" sheetId="2" r:id="rIdAfoAgent"\/>/);
}));

const { updateWorkbook } = await import('../dist/build/update.js');
const { patchVbaProject, vbaModuleNames } = await import('../dist/build/vbaPatch.js');

test('patchVbaProject adds modules and edits existing ones, keeping the rest', () => {
  const bin = writeVbaProject([
    { name: 'ThisWorkbook', type: 'document', document: 'workbook', code: 'Option Explicit\r\n\r\nPrivate Sub Workbook_Open()\r\n    x = 1\r\nEnd Sub\r\n' },
    { name: 'Sheet1', type: 'document', document: 'sheet', code: '' },
    { name: 'Module1', type: 'standard', code: 'Public Sub Run()\r\nEnd Sub\r\n' },
  ]);
  const out = patchVbaProject(bin, {
    add: [{ name: 'Module2', type: 'standard', code: 'Public Sub Two()\r\nEnd Sub\r\n' }],
    append: { thisworkbook: 'Option Explicit\r\nPublic Function Marker()\r\nEnd Function\r\n' },
  });
  const p = readVbaProject(out, true);
  assert.equal(p.error, undefined);
  assert.deepEqual(vbaModuleNames(out), ['ThisWorkbook', 'Sheet1', 'Module1', 'Module2']);
  const tw = p.modules.find((m) => m.name === 'ThisWorkbook').source;
  assert.match(tw, /Workbook_Open[\s\S]*Function Marker/);
  assert.equal((tw.match(/Option Explicit/g) ?? []).length, 1);
  assert.equal(p.modules.find((m) => m.name === 'Module1').procedures[0].name, 'Run');
  assert.throws(() => patchVbaProject(bin, { add: [{ name: 'module1', type: 'standard', code: '' }] }), /already in the project/);
  assert.throws(() => patchVbaProject(bin, { append: { Nope: 'x' } }), /not found: Nope/);
});

test('updateWorkbook adds a planning function and retrofits the bridge into an existing workbook', () => withDir((dir) => {
  const input = join(dir, 'existing.xlsm');
  buildWorkbook({ ...SPEC, documentModules: { ThisWorkbook: 'Option Explicit\r\n\r\nPrivate Sub Workbook_Open()\r\n    Module1.Run\r\nEnd Sub\r\n' } },
    { output: input, template: (writeFileSync(join(dir, 'layout.xlsx'), baseWorkbook()), join(dir, 'layout.xlsx')) });
  const output = join(dir, 'updated.xlsm');
  const r = updateWorkbook({
    planningFunctions: [{ key: 'PF_2', name: 'PLAN_FUNC_2', dataProviderKey: 'DS_1', aggregationLevel: 'ALVL_1', characteristics: ['CHAR_A'] }],
    agentBridge: { commands: [{ name: 'RUN_IT', macro: 'Module1.Run' }] },
    vbaModules: [{ name: 'Extra', code: 'Public Sub E()\r\nEnd Sub\r\n' }],
  }, { input, output });
  assert.deepEqual(r.issues, []);
  assert.deepEqual(r.planningObjects, ['PF_1=PLAN_FUNC_1 (function on DS_1)', 'PF_2=PLAN_FUNC_2 (function on DS_1)']);
  assert.deepEqual(r.sheets, ['Plan Sheet', 'AgentLog']);
  for (const m of ['Module1', 'Extra', 'AfoAgentSheet', 'AfoAgent']) assert.ok(r.vbaModules.includes(m), m);

  const d = readWorkbook(output, { macroSource: true });
  const tw = d.macros.modules.find((m) => m.name === 'ThisWorkbook').source;
  // The existing handler gets the bridge call as its first line; the missing one is appended.
  assert.match(tw, /Sub Workbook_Open\(\)\r?\n    AfoAgent\.MarkActive\r?\n    Module1\.Run/);
  assert.match(tw, /Sub Workbook_SheetCalculate\(ByVal Sh As Object\)\r?\n    If Sh\.Name = AfoAgent\.BRIDGE_SHEET/);
  assert.deepEqual(d.macros.buttons.map((b) => b.procedure), ['Module1.Run']);
  const help = strFromU8(unzipSync(readFileSync(output))['xl/worksheets/afoAgent.xml']);
  assert.match(help, /PF_2/);
  assert.match(help, /RUN_IT/);

  // A second retrofit, a missing data source and a key in use are refused without writing.
  assert.throws(() => updateWorkbook({ agentBridge: {} }, { input: output, output: join(dir, 'again.xlsm') }), /bridge is installed/);
  assert.throws(() => updateWorkbook({ planningFunctions: [{ key: 'PF_1', name: 'X', dataProviderKey: 'DS_9', aggregationLevel: 'A', characteristics: ['C'] }] },
    { input, output: join(dir, 'bad.xlsm') }), (e) => /already used/.test(e.message) && /DS_9 is not in the workbook/.test(e.message));
}));

test('updateWorkbook creates a VBA project for a workbook without macros', () => withDir((dir) => {
  const input = join(dir, 'plain.xlsx');
  buildWorkbook({ ...MINIMAL_SPEC, vbaModules: undefined }, { output: input });
  assert.throws(() => updateWorkbook({ agentBridge: {} }, { input, output: join(dir, 'x.xlsx') }), /\.xlsm/);
  const r = updateWorkbook({ agentBridge: {} }, { input, output: join(dir, 'bridged.xlsm') });
  assert.deepEqual(r.issues, []);
  assert.deepEqual(r.vbaModules, ['ThisWorkbook', 'Sheet1', 'AfoDefinition', 'AfoAgentSheet', 'AfoAgent']);
  assert.match(strFromU8(unzipSync(readFileSync(join(dir, 'bridged.xlsm')))['[Content_Types].xml']), /macroEnabled\.main/);
}));

const { writeCells, fileFormula } = await import('../dist/build/cells.js');
const { readSourceDefinition, keyedElement, firstElement } = await import('../dist/build/source.js');

test('writeCells inserts cells in row and column order and keeps styles', () => {
  const sheet = '<worksheet><sheetData><row r="2" spans="1:3"><c r="A2" s="4"><v>1</v></c><c r="C2"><v>3</v></c></row><row r="5"/></sheetData></worksheet>';
  const out = writeCells(sheet, [
    { ref: 'B2', formula: '=SAPGetVariable("DS_1", "V", "Value")' },
    { ref: 'A2', value: 'x' },
    { ref: 'D4', value: 7 },
    { ref: 'A1', formula: 'SUM(B1:B2)' },
  ]);
  assert.equal(out, '<worksheet><sheetData><row r="1"><c r="A1"><f>SUM(B1:B2)</f></c></row><row r="2"><c r="A2" s="4" t="inlineStr"><is><t xml:space="preserve">x</t></is></c><c r="B2"><f>_xll.SAPGetVariable("DS_1", "V", "Value")</f></c><c r="C2"><v>3</v></c></row><row r="4"><c r="D4"><v>7</v></c></row><row r="5"/></sheetData></worksheet>');
  assert.equal(fileFormula('=IF(A1,SAPGetData("DS_1"),_xll.SAPGetMember("DS_1"))'), 'IF(A1,_xll.SAPGetData("DS_1"),_xll.SAPGetMember("DS_1"))');
});

test('buildWorkbook takes data sources, planning objects, crosstabs and formula items over from a source workbook', () => withDir((dir) => {
  const source = join(dir, 'source.xlsm');
  buildWorkbook({ ...MINIMAL_SPEC, formulaItems: [{ key: 'Fdd1', sheet: 'Plan Sheet', cell: 'C5', formula: '=SAPSetFilterComponent("DS_1", "CHAR_A", "")' }] }, { output: source });
  const output = join(dir, 'relayout.xlsm');
  const r = buildWorkbook({
    source,
    dataProviders: [{ key: 'DS_A', from: 'DS_1' }],
    planningFunctions: [{ key: 'PF_1', from: 'PF_1' }],
    grids: [{ key: 'Crosstab1', from: 'Crosstab1', sheet: 'New Sheet', cell: 'D20' }],
    formulaItems: [{ key: 'Fdd1', from: 'Fdd1', sheet: 'New Sheet', cell: 'B3' }],
  }, { output, layout: { sheets: ['New Sheet'] } });
  assert.deepEqual(r.issues, []);
  assert.deepEqual(r.dataSources, ['DS_A=QUERY_NAME']);
  assert.deepEqual(r.planningObjects, ['PF_1=PLAN_FUNC_1 (function on DS_A)']);
  assert.deepEqual(r.crosstabs, ['Crosstab1 (DS_A) at New Sheet!D20']);

  const a = readSourceDefinition(source), b = readSourceDefinition(output);
  // The data source is taken over verbatim, only its key changes.
  assert.equal(keyedElement(b, 'DataProvider', 'DS_A'), keyedElement(a, 'DataProvider', 'DS_1').replace('<Key>DS_1</Key>', '<Key>DS_A</Key>'));
  // Connection and prompt settings come from the source; prompt entries follow the rename.
  assert.equal(firstElement(b, 'PlanningSystemDesignTime'), firstElement(a, 'PlanningSystemDesignTime'));
  assert.match(firstElement(b, 'PromptsVisibility'), /dataprovider="DS_A" name="VAR_VERSION"/);
  assert.match(keyedElement(b, 'Fdd', 'Fdd1'), /<FormulaString>=SAPSetFilterComponent\("DS_A", "CHAR_A", ""\)<\/FormulaString>/);

  const parts = unzipSync(readFileSync(output));
  assert.match(strFromU8(parts['xl/worksheets/sheet1.xml']), /<c r="B3"><f>_xll\.SAPSetFilterComponent\("DS_A", "CHAR_A", ""\)<\/f><\/c>/);
  assert.match(strFromU8(parts['xl/workbook.xml']), /<definedName name="SAPFdd1">'New Sheet'!\$B\$3<\/definedName>/);

  // Unknown keys in the source are reported together, nothing is written.
  assert.throws(() => buildWorkbook({ source, dataProviders: [{ key: 'X', from: 'DS_9' }], grids: [{ key: 'G', from: 'Nope', sheet: 'S', cell: 'A1' }] },
    { output: join(dir, 'bad.xlsx') }), (e) => /no data source DS_9/.test(e.message) && /no crosstab Nope/.test(e.message));
}));

test('cells write values and SAP formulas into template cells; styleSetName is written', () => withDir((dir) => {
  const output = join(dir, 'cells.xlsx');
  buildWorkbook({ ...MINIMAL_SPEC, vbaModules: undefined, styleSetName: 'Custom Set', cells: [
    { sheet: 'Plan Sheet', cell: 'C9', formula: '=SAPGetVariable("DS_1", "VAR_VERSION", "Value")' },
    { sheet: 'Plan Sheet', cell: 'B9', value: 'Version' },
  ] }, { output });
  const parts = unzipSync(readFileSync(output));
  const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']);
  assert.match(sheet, /<c r="C9"><f>_xll\.SAPGetVariable\("DS_1", "VAR_VERSION", "Value"\)<\/f><\/c>/);
  assert.match(sheet, /<c r="B9"[^>]*t="inlineStr"><is><t xml:space="preserve">Version<\/t>/);
  assert.match(readSourceDefinition(output), /<StyleSetName>Custom Set<\/StyleSetName>/);
  assert.throws(() => buildWorkbook({ ...MINIMAL_SPEC, cells: [{ sheet: 'Plan Sheet', cell: 'X', value: 1, formula: '=1' }] }, { output: join(dir, 'b.xlsm') }), /cell must be an A1 reference[\s\S]*either formula or value/);
}));

test('VBA modules and sheet module code are taken over from the source workbook, with replacements', () => withDir((dir) => {
  const source = join(dir, 'source.xlsm');
  buildWorkbook({ ...MINIMAL_SPEC, documentModules: { 'Plan Sheet': 'Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n    If Target.Address = "$B$22" Then Module1.Run\r\nEnd Sub\r\n' } }, { output: source });
  const output = join(dir, 'relayout.xlsm');
  const r = buildWorkbook({
    source,
    dataProviders: [{ key: 'DS_1', from: 'DS_1' }],
    planningFunctions: [{ key: 'PF_9', from: 'PF_1' }],
    grids: [{ key: 'Crosstab1', from: 'Crosstab1', sheet: 'New', cell: 'B26' }],
    vbaModules: [{ name: 'Module1', from: 'Module1', replace: [['"PF_1"', '"PF_9"']] }],
    documentModules: { New: { from: 'Plan Sheet', replace: [['$B$22', '$C$22']] } },
  }, { output, layout: { sheets: ['New'] } });
  assert.deepEqual(r.issues, []);
  const d = readWorkbook(output, { macroSource: true });
  assert.match(d.macros.modules.find((m) => m.name === 'Module1').source, /SAPExecutePlanningFunction", "PF_9"/);
  assert.match(d.macros.modules.find((m) => m.name === 'Sheet1').source, /Target\.Address = "\$C\$22"/);
  assert.throws(() => buildWorkbook({ source, dataProviders: [{ key: 'DS_1', from: 'DS_1' }], grids: [], vbaModules: [{ name: 'X', from: 'Nope' }] },
    { output: join(dir, 'bad.xlsm') }), /no standard module Nope/);
}));

const fixture = await import('./_fixture.mjs');
const { withStoredPrompts, moveRanges } = await import('../dist/build/afo.js');
const { decodeCompressed } = await import('../dist/workbook/bics.js');

const count = (s, re) => (s.match(re) ?? []).length;
const sheetData = (xml) => /<sheetData[\s\S]*?<\/sheetData>|<sheetData\s*\/>/.exec(xml)[0];

test('an AfO workbook is its own template: definition replaced in place, sheet ids, cells and VBA kept, prompts set', () => withDir((dir) => {
  const source = join(dir, 'source.xlsm');
  writeFileSync(source, fixture.buildWorkbook());
  const output = join(dir, 'rebuilt.xlsm');
  const r = buildWorkbook({
    source,
    storePrompts: true,
    dataProviders: [{ key: 'DS_1', from: 'DS_1', prompts: [{ name: 'VAR_PERIOD', low: '202501', high: '202512' }] }],
    grids: [{ key: 'Crosstab1', from: 'Crosstab1', sheet: 'Input', cell: 'A10' }],
  }, { output, template: 'source' });
  assert.ok(r.warnings.some((w) => /AfO definition is replaced/.test(w)), r.warnings.join(' | '));
  // The source knows VAR_VERSION (visibility entry) but no data source stores a value for it.
  assert.ok(r.warnings.some((w) => /No stored value for VAR_VERSION/.test(w)), r.warnings.join(' | '));
  assert.deepEqual(r.crosstabs, ['Crosstab1 (DS_1) at Input!A10:AB21']);

  const a = unzipSync(readFileSync(source)), b = unzipSync(readFileSync(output));
  const wb = strFromU8(b['xl/workbook.xml']);
  assert.equal(count(wb, /_com\.sap\.ip\.bi\.xl\.hiddensheet/g), 1);
  assert.equal(b['xl/worksheets/sheet900.xml'], undefined);
  // Exactly one customProperties per sheet, the sheet's own id kept, the missing type added.
  const sheet = strFromU8(b['xl/worksheets/sheet2.xml']);
  assert.equal(count(sheet, /<customProperties>/g), 1);
  assert.equal(count(sheet, /name="_pios_id"/g), 1);
  assert.equal(count(sheet, /name="CofWorksheetType"/g), 1);
  assert.deepEqual(b['xl/customProperty3.bin'], a['xl/customProperty3.bin']);
  assert.equal(sheetData(sheet), sheetData(strFromU8(a['xl/worksheets/sheet2.xml'])));
  assert.deepEqual(b['xl/vbaProject.bin'], a['xl/vbaProject.bin']);

  // The crosstab keeps the area it occupied (A10:AB21), in the definition and in its defined name.
  const def = readSourceDefinition(output);
  assert.match(keyedElement(def, 'Grid', 'Crosstab1'), /<SheetName>Sheet2\|guid-input<\/SheetName><Cell Name="TopLeft"><Col>0<\/Col><Row>9<\/Row><\/Cell><Cell Name="BottomRight"><Col>27<\/Col><Row>20<\/Row>/);
  assert.equal(count(wb, /name="SAPCrosstab1"/g), 1);
  assert.match(wb, /<definedName name="SAPCrosstab1">Input!\$A\$10:\$AB\$21<\/definedName>/);

  // Prompts of a taken-over data source land in its stored state; the rest of the state stays.
  const d = readWorkbook(output);
  const ds = d.dataProviders.find((x) => x.key === 'DS_1');
  assert.deepEqual(ds.state.variables.map((v) => [v.name, v.values]), [['VAR_PERIOD', ['202501 .. 202512']]]);
  assert.equal(ds.state.infoProvider, 'PROVIDER_1');
  assert.ok(ds.state.filters.some((f) => f.characteristic === 'CHAR_REGION'));
  assert.match(firstElement(def, 'PromptsVisibility'), /<Variable dataprovider="DS_1" name="VAR_PERIOD" visible="True" \/>/);
  assert.match(def, /<StorePromptsInDocument>True<\/StorePromptsInDocument>/);
}));

test('moveRanges keeps the size of every range and moves it to the new top-left cell', () => {
  const grid = '<Grid><Range Name="RangeOccupied"><SheetName>S|1</SheetName><Cell Name="TopLeft"><Col>0</Col><Row>21</Row></Cell><Cell Name="BottomRight"><Col>5</Col><Row>86</Row></Cell></Range>'
    + '<Range Name="RequestedRange"><SheetName>S|1</SheetName><Cell Name="TopLeft"><Col>0</Col><Row>21</Row></Cell><Cell Name="BottomRight"><Col>0</Col><Row>21</Row></Cell></Range></Grid>';
  assert.equal(moveRanges(grid, 'T|2', 'C30'), grid.replace(/S\|1/g, 'T|2')
    .replace('<Col>0</Col><Row>21</Row></Cell><Cell Name="BottomRight"><Col>5</Col><Row>86</Row>', '<Col>2</Col><Row>29</Row></Cell><Cell Name="BottomRight"><Col>7</Col><Row>94</Row>')
    .replace('<Col>0</Col><Row>21</Row></Cell><Cell Name="BottomRight"><Col>0</Col><Row>21</Row>', '<Col>2</Col><Row>29</Row></Cell><Cell Name="BottomRight"><Col>2</Col><Row>29</Row>'));
});

test('a template sheet with EPM custom properties gets one merged customProperties; its cells stay byte for byte', () => withDir((dir) => {
  const parts = unzipSync(baseWorkbook());
  // As EPM writes it, plus a cell whose value metadata (vm) points into xl/metadata.xml.
  const cells = '<sheetData><row r="14"><c r="B14" s="0" t="e" vm="2"><f ca="1">IFERROR(C14,"")</f><v>#VALUE!</v></c><c r="C14" t="str"><f>_xll.SAPGetMember("DS_1","CHAR_A")</f><v/></c></row></sheetData>';
  parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<sheetData/>', cells)
    .replace('<drawing ', '<customProperties><customPr name="_pios_id" r:id="rId2"/><customPr name="EpmWorksheetKeyString_GUID" r:id="rId3"/></customProperties><drawing '));
  parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']).replace('</Relationships>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" Target="../customProperty1.bin"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" Target="../customProperty2.bin"/></Relationships>'));
  parts['xl/customProperty1.bin'] = new Uint8Array(Buffer.from('Sheet1|epm-guid', 'utf16le'));
  parts['xl/customProperty2.bin'] = new Uint8Array(Buffer.from('epm-key', 'utf16le'));
  const template = join(dir, 'epm.xlsx');
  writeFileSync(template, zipSync(parts));

  const output = join(dir, 'built.xlsm');
  const r = buildWorkbook(SPEC, { output, template });
  assert.deepEqual(r.issues, []);
  const out = unzipSync(readFileSync(output));
  const sheet = strFromU8(out['xl/worksheets/sheet1.xml']);
  assert.equal(count(sheet, /<customProperties>/g), 1);
  assert.equal(count(sheet, /name="_pios_id"/g), 1);
  assert.match(sheet, /<customPr name="_pios_id" r:id="rId2"\/><customPr name="EpmWorksheetKeyString_GUID" r:id="rId3"\/><customPr name="CofWorksheetType" r:id="rIdAfo\d+"\/><\/customProperties><drawing /);
  assert.equal(sheetData(sheet), cells);
  // Items refer to the sheet by the id it already had.
  assert.match(readSourceDefinition(output), /<SheetName>Sheet1\|epm-guid<\/SheetName>/);
  assert.deepEqual(out['xl/customProperty2.bin'], parts['xl/customProperty2.bin']);
}));

test('withStoredPrompts merges into a stored state, or creates one and switches "reset data source" off', () => {
  const minimal = '<?xml version="1.0"?><BICS_VIEW version="3.1"><QUERY name="Q"><DUE_DATE>20260101</DUE_DATE></QUERY><DOCUMENT_APPLICATION_ASSIGNED /></BICS_VIEW>';
  const design = '<TABLE_DESIGN><STYLE_LIST><STYLE subType="CUSTOM" name="S1"><COLOR>-1</COLOR></STYLE></STYLE_LIST><OPERATION_LIST /></TABLE_DESIGN>';
  const plain = `<DataProvider><Key>DS_1</Key><Cleaned>True</Cleaned><ResetDataSource>True</ResetDataSource><DimensionSynchronization /><BICS_VIEW_HEADER>${design}</BICS_VIEW_HEADER><Comments /></DataProvider>`;
  const r = withStoredPrompts(plain, [{ name: 'V1', member: 'A' }], () => minimal);
  assert.doesNotMatch(r.xml, /ResetDataSource/);
  assert.equal(r.notes.length, 3);
  const state = decodeCompressed(/compressed="true">([^<]*)</.exec(r.xml)[1]);
  assert.equal(state, minimal.replace('</QUERY>', '</QUERY><VARIABLES><VARIABLE name="V1" type="INFO_OBJECT_MEMBER" isStateVariable="false"><SELECTIONS><SELECTION type="SINGLE_MEMBER"><MEMBER name="A" type="MEMBER" /></SELECTION></SELECTIONS></VARIABLE></VARIABLES>')
    .replace('<DOCUMENT_APPLICATION_ASSIGNED', `${design}<DOCUMENT_APPLICATION_ASSIGNED`));

  // A second call replaces V1 in place; a new variable is appended.
  const again = withStoredPrompts(r.xml, [{ name: 'V2', literal: 'X' }, { name: 'V1', low: '1', high: '2' }], () => { throw new Error('not needed'); });
  assert.deepEqual(again.notes, []);
  const vars = decodeCompressed(/compressed="true">([^<]*)</.exec(again.xml)[1]);
  assert.match(vars, /<VARIABLE name="V1"[^>]*><SELECTIONS><SELECTION type="INTERVAL">[\s\S]*<VARIABLE name="V2"[^>]*><SELECTIONS><SELECTION type="SINGLE_MEMBER"><MEMBER type="LITERAL_MEMBER"><VALUE>X</);
});

test('updateWorkbook stores prompts in an existing workbook and changes nothing else', () => withDir((dir) => {
  const input = join(dir, 'existing.xlsx');
  buildWorkbook({ ...MINIMAL_SPEC, vbaModules: undefined, storePrompts: false }, { output: input });
  const output = join(dir, 'preset.xlsx');
  const r = updateWorkbook({ prompts: { DS_1: { VAR_PERIOD: { low: '202601', high: '202612' }, VAR_NEW: 'M1' } } }, { input, output });
  assert.deepEqual(r.issues, []);
  assert.deepEqual(r.changes, ['prompts of DS_1: VAR_PERIOD, VAR_NEW', 'StorePromptsInDocument on']);

  const d = readWorkbook(output);
  assert.equal(d.settings.settings.StorePromptsInDocument, 'True');
  assert.deepEqual(d.dataProviders[0].state.variables.map((v) => [v.name, v.values]), [['VAR_VERSION', ['"PLAN1"']], ['VAR_PERIOD', ['202601 .. 202612']], ['VAR_NEW', ['M1']]]);
  assert.match(firstElement(readSourceDefinition(output), 'PromptsVisibility'), /dataprovider="DS_1" name="VAR_NEW"/);
  // Every part but the definition is unchanged.
  const a = unzipSync(readFileSync(input)), b = unzipSync(readFileSync(output));
  assert.deepEqual(Object.keys(b).filter((n) => Buffer.compare(Buffer.from(a[n] ?? []), Buffer.from(b[n])) !== 0).length, 1);

  assert.throws(() => updateWorkbook({ prompts: { DS_9: { V: 'X' }, DS_1: [{ name: 'V', member: 'A', literal: 'B' }] } }, { input, output: join(dir, 'bad.xlsx') }),
    (e) => /DS_9 is not in the workbook/.test(e.message) && /exactly one of member, literal or low/.test(e.message));
}));

test('updateWorkbook switches refresh on open and nothing else', () => withDir((dir) => {
  const input = join(dir, 'refreshing.xlsx');
  buildWorkbook({ ...MINIMAL_SPEC, vbaModules: undefined, refreshOnOpen: true }, { output: input });
  assert.equal(readWorkbook(input).settings.settings.RefreshOnOpen, 'True');
  const output = join(dir, 'quiet.xlsx');
  const r = updateWorkbook({ refreshOnOpen: false }, { input, output });
  assert.deepEqual(r.changes, ['RefreshOnOpen off']);
  const d = readWorkbook(output);
  assert.equal(d.settings.settings.RefreshOnOpen, 'False');
  assert.equal(d.settings.settings.StorePromptsInDocument, readWorkbook(input).settings.settings.StorePromptsInDocument);
  const a = unzipSync(readFileSync(input)), b = unzipSync(readFileSync(output));
  assert.deepEqual(Object.keys(b).filter((n) => Buffer.compare(Buffer.from(a[n] ?? []), Buffer.from(b[n])) !== 0).length, 1);
}));

test('updateWorkbook points every connection to another system, from the landscape or a reference workbook', () => withDir((dir) => {
  const input = join(dir, 'dev.xlsx');
  buildWorkbook({ ...MINIMAL_SPEC, vbaModules: undefined }, { output: input });
  const before = readSourceDefinition(input);
  const n = count(before, /<SystemInfo>/g);
  assert.ok(n >= 2, 'data source and planning function carry a connection');

  mkdirSync(join(dir, 'SAP', 'Common'), { recursive: true });
  writeFileSync(join(dir, 'SAP', 'Common', 'SAPUILandscape.xml'), `<Landscape><Services>
    <Service type="SAPGUI" uuid="u-q" name="QAS - Quality" systemid="QAS" server="qas.example.com:3205" routerid="r-1"/>
  </Services><Routers><Router uuid="r-1" name="R" router="/H/router.example.com/S/3299"/></Routers></Landscape>`);
  const appData = process.env.APPDATA;
  process.env.APPDATA = dir;
  let r;
  try {
    r = updateWorkbook({ system: { systemId: 'qas' } }, { input, output: join(dir, 'qas.xlsx') });
    assert.throws(() => updateWorkbook({ system: { systemId: 'NONE' } }, { input, output: join(dir, 'none.xlsx') }), /not in the SAP Logon landscape/);
  } finally {
    process.env.APPDATA = appData;
  }
  assert.match(r.changes.join(), /system SYS\/001 → QAS\/001 \(\d+ connections\)/);
  const def = readSourceDefinition(r.output);
  assert.equal(count(def, /<SystemInfo>/g), n);
  assert.equal(count(def, /<SystemId>SYS<\/SystemId>/g), 0);
  assert.equal(count(def, /<SapRouter>\/H\/router\.example\.com\/S\/3299<\/SapRouter>/g), n);
  assert.match(def, /<\/AAO_Revision><SystemId>QAS<\/SystemId>/);
  assert.match(def, /<Client>001<\/Client><Language>EN<\/Language>/, 'client and language kept');
  assert.deepEqual(readWorkbook(r.output).issues, []);

  // Back again, by copying the connection of the original workbook.
  const back = updateWorkbook({ system: { from: input } }, { input: r.output, output: join(dir, 'back.xlsx') });
  assert.equal(readSourceDefinition(back.output).match(/<SystemInfo>.*?<\/SystemInfo>/s)[0], before.match(/<SystemInfo>.*?<\/SystemInfo>/s)[0]);
  assert.throws(() => updateWorkbook({ system: { from: join(dir, 'missing.xlsx') } }, { input, output: join(dir, 'x.xlsx') }), /does not exist/);
}));

test('stored filters: members, intervals, comparisons and exclusions read back as written', () => withDir((dir) => {
  const output = join(dir, 'filtered.xlsx');
  buildWorkbook({
    ...MINIMAL_SPEC, vbaModules: undefined, planningFunctions: undefined,
    dataProviders: [{
      key: 'DS_1', query: 'QUERY_NAME',
      filters: [
        { characteristic: 'CHAR_A', members: ['A1', 'A2'] },
        { characteristic: 'CHAR_A', members: ['A9'], exclude: true },
        { characteristic: 'CHAR_MONTH', low: '202601', high: '202606' },
        { characteristic: 'CHAR_YEAR', operator: 'GREATER_EQUAL', value: '2025' },
      ],
    }],
  }, { output });
  const dp = readWorkbook(output).dataProviders[0];
  assert.deepEqual(dp.state.filters, [
    { characteristic: 'CHAR_A', selections: ['A1', 'A2', 'not A9'] },
    { characteristic: 'CHAR_MONTH', selections: ['202601 .. 202606'] },
    { characteristic: 'CHAR_YEAR', selections: ['>= 2025'] },
  ]);
  assert.throws(() => validateSpecOrThrow({ ...MINIMAL_SPEC, dataProviders: [{ key: 'DS_1', query: 'Q', filters: [{ characteristic: 'C', members: ['X'], low: '1' }, { characteristic: 'D', operator: 'GREATER' }] }] }),
    (e) => /exactly one of members/.test(e.message) && /operator and value go together/.test(e.message));
}));

function validateSpecOrThrow(spec) {
  const errors = validateSpec(spec);
  if (errors.length) throw new Error(errors.join('\n'));
}
