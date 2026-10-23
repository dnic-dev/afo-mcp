import { test } from 'node:test';
import assert from 'node:assert/strict';
import { logonConfig, logonConfigs, logonStatus, secretCommand } from '../dist/operate/open.js';
import { landscapeSystem } from '../dist/operate/landscape.js';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const SYSTEMS = JSON.stringify([
  { system: 'sys', client: '100', user: 'USER_A', language: 'EN', vault: 'entry@sys', passwordVar: 'SAP_PASSWORD' },
  { system: 'OTH', client: '200', user: 'USER_B' },
]);

test('logon entries: vault entries count without a password in this environment', () => {
  const all = logonConfigs({ AFO_LOGON_SYSTEMS: SYSTEMS });
  assert.deepEqual(all.map((l) => l.system), ['SYS']);
  assert.equal(all[0].vault, 'entry@sys');
  assert.equal(logonConfig('sys', { AFO_LOGON_SYSTEMS: SYSTEMS })?.client, '100');
});

test('logon entries: an entry without vault needs its password variable', () => {
  const env = { AFO_LOGON_SYSTEMS: SYSTEMS, BW_PASSWORD: 'x' };
  assert.deepEqual(logonConfigs(env).map((l) => l.system), ['SYS', 'OTH']);
  assert.equal(logonConfig(undefined, env), undefined, 'several systems: no default');
  assert.match(logonStatus(env), /SYS\/100, OTH\/200/);
  assert.doesNotMatch(logonStatus(env), /x\b/);
});

test('single-system variables', () => {
  const env = { AFO_LOGON_SYSTEM: 'abc', AFO_LOGON_CLIENT: '001', AFO_LOGON_USER: 'U', AFO_LOGON_PASSWORD: 'p' };
  assert.equal(logonConfig(undefined, env)?.system, 'ABC');
  assert.equal(logonConfigs({ ...env, AFO_LOGON_PASSWORD: '' }).length, 0);
});

test('secret command: placeholder replaced, malformed setting refused', () => {
  const env = { AFO_SECRET_COMMAND: '["runner", "run", "{vault}", "--"]' };
  assert.deepEqual(secretCommand(env, 'entry@sys'), ['runner', 'run', 'entry@sys', '--']);
  assert.throws(() => secretCommand({}, 'e'), /AFO_SECRET_COMMAND/);
  assert.throws(() => secretCommand({ AFO_SECRET_COMMAND: '"runner"' }, 'e'), /AFO_SECRET_COMMAND/);
});

test('landscape: server, system number, SAProuter and SNC name', () => {
  const dir = mkdtempSync(join(tmpdir(), 'afo-landscape-'));
  const file = join(dir, 'SAPUILandscape.xml');
  writeFileSync(file, `<Landscape><Services>
    <Service type="SAPGUI" uuid="u-1" name="SYS - Test" systemid="SYS" server="host.example.com:3201" routerid="r-1" sncname="p:CN=SYS" sncnosso="1"/>
    <Service type="SAPGUI" uuid="u-2" name="OTH" systemid="OTH" server="10.0.0.1:3200"/>
  </Services><Routers><Router uuid="r-1" name="R" router="/H/router.example.com/S/3299"/></Routers></Landscape>`);
  const s = landscapeSystem('sys', '100', 'EN', [file]);
  assert.equal(s.systemName, 'SYS - Test');
  assert.equal(s.server, 'host.example.com');
  assert.equal(s.systemNumber, '01');
  assert.equal(s.sapRouter, '/H/router.example.com/S/3299');
  assert.equal(s.sncPartnerName, 'p:CN=SYS');
  const o = landscapeSystem('OTH', '200', 'DE', [file]);
  assert.equal(o.sapRouter, undefined);
  assert.equal(o.sncPartnerName, undefined);
  assert.throws(() => landscapeSystem('NONE', '1', 'EN', [file]), /not in the SAP Logon landscape/);
});

test('operate script: steps are reported as progress lines for the dialog check', async () => {
  const { PS_PREAMBLE, describeDialog } = await import('../dist/operate/excel.js');
  assert.match(PS_PREAMBLE, /function Step/);
  assert.match(PS_PREAMBLE, /Emit "step/);
  assert.doesNotMatch(PS_PREAMBLE, /Invoke\(\)|WM_CLOSE|PostMessage/, 'nothing is clicked or closed');
  assert.equal(describeDialog({ kind: 'prompts', title: 'Prompts', text: 'Version', buttons: [] }), 'prompts dialog "Prompts": Version');
});
