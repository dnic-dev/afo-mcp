import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs';
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
