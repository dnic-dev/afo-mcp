import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readWorkbook } from '../dist/workbook/read.js';
import { scanWorkbooks, whereUsed } from '../dist/workbook/scan.js';
import { buildWorkbook } from './_fixture.mjs';

let dir;
let file;

before(() => {
  dir = mkdtempSync(join(tmpdir(), 'afo-mcp-test-'));
  file = join(dir, 'Planning A.xlsm');
  writeFileSync(file, buildWorkbook());
  writeFileSync(join(dir, 'Planning B.xlsm'), buildWorkbook({ query: 'OTHER_QUERY' }));
  writeFileSync(join(dir, 'Planning C.xlsm'), buildWorkbook({ systemId: 'QAS', query: 'THIRD_QUERY' }));
  writeFileSync(join(dir, 'broken.xlsx'), 'not a zip');
  writeFileSync(join(dir, '~$Planning A.xlsm'), 'lock file');
});

after(() => rmSync(dir, { recursive: true, force: true }));

test('readWorkbook summarises the AfO definition', () => {
  const d = readWorkbook(file);
  assert.equal(d.afo, true);
  assert.equal(d.summary.systemId, 'SYS');
  assert.equal(d.summary.afoRevision, '2.8.0.0');
  assert.deepEqual(d.summary.dataSources, ['DS_1=QUERY_NAME', 'DS_2=SHARED_QUERY']);
  assert.deepEqual(d.summary.planningObjects, ['PF_1=PLAN_FUNC_1 (function on DS_1)']);
  assert.equal(d.settings.settings.RefreshOnOpen, 'True');
  assert.equal(d.settings.settings.CellLocking, 'None');
  assert.deepEqual(d.settings.promptVariables, [{ name: 'VAR_VERSION', visible: true }]);
});

test('data sources carry their saved state, or say that none was saved', () => {
  const [ds1, ds2] = readWorkbook(file).dataProviders;
  assert.equal(ds1.stateStored, true);
  assert.equal(ds1.inputReady, true);
  assert.equal(ds1.state.filters.length, 4);
  assert.equal(ds2.stateStored, false);
  assert.equal(ds2.state, undefined);
});

test('planning objects carry binding, aggregation level, filter, variables and references', () => {
  const [pf] = readWorkbook(file).planningObjects;
  assert.equal(pf.kind, 'function');
  assert.equal(pf.dataProviderKey, 'DS_1');
  assert.equal(pf.aggregationLevel, 'ALVL_1');
  assert.equal(pf.filterObject, '!!AALVL_1');
  assert.deepEqual(pf.filter, [{ characteristic: 'CHAR_VERSION', selections: ['PLAN1'] }]);
  assert.deepEqual(pf.variables.map((v) => [v.name, v.values]), [['VAR_VERSION', ['PLAN1']], ['VAR_FACTOR', ['"1.5"']]]);
  assert.deepEqual(pf.references, [
    { technicalName: 'CHAR_VERSION', object: 'Filter', referenceType: 'Filter' },
    { technicalName: 'VAR_FACTOR', object: 'Variable', referenceType: 'Cell', cell: 'Input!$B$3' },
  ]);
});

test('items resolve the AfO sheet id to the sheet name and use A1 addresses', () => {
  const d = readWorkbook(file);
  assert.deepEqual(d.items.grids[0], { key: 'Crosstab1', description: 'Crosstab 1', dataProviderKey: 'DS_1', address: 'A10:AB21', sheet: 'Input' });
  assert.equal(d.items.formulaItems[0].address, 'B5');
  assert.equal(d.items.formulaItems[0].sheet, 'Input');
  const input = d.sheets.find((s) => s.name === 'Input');
  assert.deepEqual(input.grids, ['Crosstab1', 'Crosstab2']);
  assert.equal(d.sheets[0].state, 'veryHidden');
});

test('SAP formulas are listed per sheet with complete cell counts', () => {
  const [sheet] = readWorkbook(file).formulas;
  assert.equal(sheet.sheet, 'Input');
  assert.equal(sheet.cells, 6);                // B3, B4:B6 shared, C5, C6
  assert.deepEqual(sheet.formulas.map((f) => f.cell), ['B3', 'B4', 'C5', 'C6']);
  assert.equal(sheet.formulas[1].range, 'B4:B6');
  assert.equal(sheet.formulas[3].formula, '=IF(SAPGetMember("DS_UNKNOWN","CHAR_REGION")="A",1,2)');
});

test('macros list modules, procedures and AfO calls; buttons link to procedures', () => {
  const { macros } = readWorkbook(file);
  assert.deepEqual(macros.modules.map((m) => [m.name, m.type]), [['Module1', 'standard'], ['Sheet2', 'document']]);
  const [run, save] = macros.modules[0].procedures;
  assert.equal(run.name, 'RunPlanning');
  assert.deepEqual(run.afoCalls.map((c) => `${c.function}:${c.firstArgument}`), ['SAPSetPlanParameter:PF_1', 'SAPExecutePlanningFunction:PF_1']);
  assert.deepEqual(run.calls, ['SaveData']);
  assert.deepEqual(save.afoCalls.map((c) => `${c.function}:${c.firstArgument}`), ['SAPExecuteCommand:PlanDataSave', 'SAPGetVariable:DS_7']);
  assert.equal(macros.modules[0].source, undefined);
  assert.deepEqual(macros.buttons.map((b) => [b.macro, b.procedure, b.text, b.anchor]), [
    ['RunPlanning', 'Module1.RunPlanning', 'Run planning', 'D2'],
    ['MissingMacro', undefined, 'Old button', 'B8'],
  ]);
  assert.match(readWorkbook(file, { macroSource: true }).macros.modules[0].source, /Public Sub RunPlanning/);
});

test('issues name every broken link', () => {
  const { issues } = readWorkbook(file);
  assert.deepEqual(issues.sort(), [
    'Button _x0000_s1025 on sheet Input calls macro MissingMacro, which is not in the VBA project.',
    'Crosstab Crosstab2 uses data source DS_9, which is not defined.',
    'Input!C6: SAPGetMember refers to data source DS_UNKNOWN, which is not defined.',
    'Module1.SaveData line 15: SAPGetVariable refers to data source DS_7, which is not defined.',
  ].sort());
});

test('sections limit the output; summary and issues stay', () => {
  const d = readWorkbook(file, { sections: ['planningObjects'] });
  assert.ok(d.planningObjects);
  assert.equal(d.dataProviders, undefined);
  assert.equal(d.macros, undefined);
  assert.equal(d.summary.procedures, 3);       // two in Module1, one in the sheet module
  assert.equal(d.issues.length, 4);
});

test('scanWorkbooks profiles a folder, finds shared queries, patterns and outliers', () => {
  const r = scanWorkbooks(dir);
  assert.equal(r.workbooks, 4);                // the lock file is skipped
  assert.equal(r.withAfo, 3);
  assert.deepEqual(r.systems, { SYS: 2, QAS: 1 });
  assert.deepEqual(r.dataSources[0], { name: 'SHARED_QUERY', detail: 'Shared', workbooks: ['Planning A.xlsm', 'Planning B.xlsm', 'Planning C.xlsm'] });
  assert.deepEqual(r.patterns, [{ signature: ['DS_2=SHARED_QUERY', 'PF_1 (function)'], workbooks: ['Planning A.xlsm', 'Planning B.xlsm', 'Planning C.xlsm'] }]);
  assert.ok(r.outliers.includes('Planning C.xlsm: saved against system QAS, most workbooks use SYS.'));
  assert.ok(r.outliers.some((o) => o.startsWith('broken.xlsx: could not be read')));
  assert.equal(r.variables.find((v) => v.name === 'VAR_FACTOR').workbooks.length, 3);
});

test('compact scans drop profiles and keep totals', () => {
  const r = scanWorkbooks(dir, { detail: 'compact' });
  assert.equal(r.detail, 'compact');
  assert.deepEqual(r.profiles, []);
  assert.equal(r.dataSources[0].workbooks.length, 3);   // under the cap of five: unchanged
  assert.equal(scanWorkbooks(dir).detail, 'full');      // small folders default to full
});

test('whereUsed finds a query, a variable and a characteristic everywhere it occurs', () => {
  const q = whereUsed(dir, 'other_query');
  assert.deepEqual(q.hits, [{ file: 'Planning B.xlsm', places: ['data source DS_1 (Query)'] }]);

  const v = whereUsed(dir, 'VAR_VERSION').hits.find((h) => h.file === 'Planning A.xlsm');
  assert.deepEqual(v.places, [
    'prompt variable',
    'variable of PF_1: PLAN1',
    'SAP formulas on Input: B3',
    'VBA Module1, line 5, 15',
  ]);

  const c = whereUsed(dir, 'CHAR_MONTH').hits.find((h) => h.file === 'Planning A.xlsm');
  assert.deepEqual(c.places, [
    'filter on DS_1: 202501 .. 202512',
    'formula item Fdd1 at Input!B5: =SAPSetFilterComponent("DS_1", "CHAR_MONTH", "")',
    'SAP formulas on Input: C5',
  ]);
  assert.equal(whereUsed(dir, 'CHAR').hits.length, 0);   // whole names only
});
