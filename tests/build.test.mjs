import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, readFileSync, rmSync } from 'node:fs';
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

  // A template that is already an AfO workbook is refused, so is a missing sheet.
  assert.throws(() => buildWorkbook(SPEC, { output: join(dir, 'again.xlsm'), template: join(dir, 'built.xlsm') }), /already an Analysis for Office workbook/);
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
  for (const re of [/system\.systemName/, /exactly one of member, literal or low/, /key ds_1 is already used/, /DS_9 is not a data source key/, /cell must be an A1 reference/, /filter on C2/, /must name sheet and cell/, /module Missing is not among vbaModules/, /Sheet1 is reserved/, /give either code or file/, /B2 is not usable as a defined name/]) {
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
  assert.match(agent, /PutFormula "Q2", "=SAPListOf\(""DATASOURCES""\)"/);
  assert.match(sheet, /Ask the user before SAVE/);
  assert.match(sheet, /SAPListOfVariables/);
  assert.match(sheet, /<c r="Q1" t="inlineStr"><is><t xml:space="preserve">Data sources \(live\)/);
  assert.doesNotMatch(wb, /state="hidden"/);
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
  buildWorkbook({ ...MINIMAL_SPEC, agentBridge: { hidden: true } }, { output: join(dir, 'h.xlsm') });
  assert.match(strFromU8(unzipSync(readFileSync(join(dir, 'h.xlsm')))['xl/workbook.xml']), /<sheet name="AgentLog" sheetId="2" r:id="rIdAfoAgent" state="hidden"\/>/);
}));
