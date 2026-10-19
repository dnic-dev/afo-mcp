import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseXml, child, text } from '../dist/workbook/xml.js';
import { decodeCompressed, parseBicsView, parseVariables } from '../dist/workbook/bics.js';
import { decompressVba, analyseSource } from '../dist/workbook/vba.js';
import { parseSapCalls, scanSapFormulas, literal } from '../dist/workbook/sheets.js';
import { toA1 } from '../dist/workbook/afo.js';
import { compressBlock, compressVbaLiteral, DS1_STATE, PF1_VARIABLES } from './_fixture.mjs';

test('parseXml decodes entities, strips prefixes and survives unclosed elements', () => {
  const doc = parseXml('<?xml version="1.0"?><x:a k="1 &amp; 2"><b>x &lt; y</b><c/><div>one<br>two</div><d>ok</d></x:a>');
  const a = child(doc, 'a');
  assert.equal(a.attrs.k, '1 & 2');
  assert.equal(text(a, 'b'), 'x < y');
  assert.ok(child(a, 'c'));
  assert.equal(text(a, 'd'), 'ok');           // still a child of <a>, despite the unclosed <br>
});

test('a compressed block round-trips through length prefix and gzip', () => {
  assert.equal(decodeCompressed(compressBlock('<X>ä</X>')), '<X>ä</X>');
});

test('parseBicsView reads provider, query, filters, axes and hierarchies', () => {
  const view = parseBicsView(child(parseXml(DS1_STATE), 'BICS_VIEW'));
  assert.equal(view.infoProvider, 'PROVIDER_1');
  assert.equal(view.query, 'QUERY_NAME');
  assert.equal(view.dueDate, '20250101');
  assert.deepEqual(view.filters, [
    { characteristic: 'CHAR_VERSION', selections: ['PLAN1'] },
    { characteristic: 'CHAR_MONTH', selections: ['202501 .. 202512'] },
    { characteristic: 'CHAR_REGION', selections: ['not R9', 'node NODE_1'] },
    { characteristic: 'CHAR_YEAR', selections: ['>= 2025'] },
  ]);
  assert.deepEqual(view.rows, ['CHAR_REGION']);
  assert.deepEqual(view.columns, ['KEY_FIGURES']);
  assert.equal(view.free, undefined);
  assert.deepEqual(view.hierarchies, [{ characteristic: 'CHAR_REGION', hierarchy: 'HIER_1' }]);
});

test('parseVariables shows members and literals', () => {
  assert.deepEqual(parseVariables(child(parseXml(PF1_VARIABLES), 'VARIABLES')), [
    { name: 'VAR_VERSION', type: 'INFO_OBJECT_MEMBER', values: ['PLAN1'] },
    { name: 'VAR_FACTOR', type: 'INFO_OBJECT_MEMBER', values: ['"1.5"'] },
  ]);
});

test('decompressVba decodes the MS-OVBA specification example', () => {
  // MS-OVBA 3.2.3, "Example: Compressed Container".
  const compressed = Uint8Array.from([
    0x01, 0x2f, 0xb0, 0x00, 0x23, 0x61, 0x61, 0x61, 0x62, 0x63, 0x64, 0x65, 0x82, 0x66, 0x00, 0x70,
    0x61, 0x67, 0x68, 0x69, 0x6a, 0x01, 0x38, 0x08, 0x61, 0x6b, 0x6c, 0x00, 0x30, 0x6d, 0x6e, 0x6f,
    0x70, 0x06, 0x71, 0x02, 0x70, 0x04, 0x10, 0x72, 0x73, 0x74, 0x75, 0x76, 0x10, 0x77, 0x78, 0x79,
    0x7a, 0x00, 0x3c,
  ]);
  assert.equal(Buffer.from(decompressVba(compressed)).toString('latin1'), '#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa');
});

test('decompressVba handles literal-only chunks across the 4096-byte boundary', () => {
  const input = Buffer.from('x'.repeat(5000) + 'end', 'latin1');
  assert.equal(Buffer.from(decompressVba(compressVbaLiteral(input))).toString('latin1'), input.toString('latin1'));
});

test('analyseSource finds procedures, AfO calls and calls between procedures', () => {
  const src = [
    'Sub Main()',
    '  x = Application.Run("SAPExecutePlanningFunction", _',
    '      "PF_1")',
    "  ' Application.Run(\"SAPExecuteCommand\", \"Ignored\")",
    '  Range("SAPCrosstab1").Select',
    '  Helper',
    'End Sub',
    'Private Function Helper() As Boolean',
    '  Helper = Application.Run("SapGetVariable", "DS_1", "VAR_1", "Value") <> ""',
    'End Function',
  ].join('\n');
  const procs = analyseSource(src);
  assert.deepEqual(procs.map((p) => [p.name, p.kind, p.line]), [['Main', 'Sub', 1], ['Helper', 'Function', 8]]);
  assert.deepEqual(procs[0].afoCalls.map((c) => [c.function, c.firstArgument, c.line]), [['SAPExecutePlanningFunction', 'PF_1', 2]]);
  assert.deepEqual(procs[0].calls, ['Helper']);
  assert.deepEqual(procs[1].afoCalls.map((c) => [c.function, c.firstArgument]), [['SAPGetVariable', 'DS_1']]);
});

test('parseSapCalls keeps literals quoted and splits nested arguments', () => {
  const calls = parseSapCalls('IF(SAPGetData("DS_1","KF_1",CONCATENATE("A=",B2))>0,SAPGetMember("DS_2", "CHAR"),"x, y")');
  assert.deepEqual(calls, [
    { function: 'SAPGetData', args: ['"DS_1"', '"KF_1"', 'CONCATENATE("A=",B2)'] },
    { function: 'SAPGetMember', args: ['"DS_2"', '"CHAR"'] },
  ]);
  assert.equal(literal('"DS_1"'), 'DS_1');
  assert.equal(literal('B2'), undefined);
});

test('scanSapFormulas counts shared formula cells and strips the add-in prefix', () => {
  const xml = '<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">_xll.SAPGetData("DS_1","KF")</f></c></row>'
    + '<row r="2"><c r="A2"><f t="shared" si="0"/></c><c r="B2"><f>SUM(A1)</f></c></row><row r="3"><c r="A3" s="1"><f t="shared" si="0"/></c></row></sheetData>';
  const { formulas, cells } = scanSapFormulas(xml);
  assert.equal(cells, 3);
  assert.equal(formulas.length, 1);
  assert.equal(formulas[0].cell, 'A1');
  assert.equal(formulas[0].range, 'A1:A3');
  assert.equal(formulas[0].formula, '=SAPGetData("DS_1","KF")');
});

test('toA1 converts zero-based coordinates', () => {
  assert.equal(toA1(0, 0), 'A1');
  assert.equal(toA1(14, 1), 'B15');
  assert.equal(toA1(9, 27), 'AB10');
  assert.equal(toA1(0, 701), 'ZZ1');
  assert.equal(toA1(0, 702), 'AAA1');
});
