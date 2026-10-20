// Change an existing VBA project (`xl/vbaProject.bin`) without rebuilding it: add modules and
// append code to existing ones, keeping references, forms, project settings and every other
// module as they are.
//
// Changed or not, every module is rewritten as source only (offset 0, no performance cache) and
// `_VBA_PROJECT` declares version 0xFFFF, so Excel compiles the whole project from source on the
// next open instead of trusting stale compiled code. Signed projects are refused by the caller:
// any change breaks the signature.

import CFB from 'cfb';
import { decompressVba } from '../workbook/vba.js';
import { compressVba, encodeCp1252, moduleRecords, moduleSource, record, u16, u32, type VbaModuleSource } from './vba.js';

export interface VbaPatch {
  /** New modules (standard, class or document modules for new sheets). */
  add?: VbaModuleSource[];
  /** Code to append to existing modules, by module name (case-insensitive). */
  append?: Record<string, string>;
  /** Edit the source of existing modules, by module name: gets the source, returns the new one. */
  edit?: Record<string, (source: string) => string>;
}

interface DirRecord { id: number; buf: Buffer }

/** dir-stream records in order; PROJECTVERSION (0x0009) is the one record whose size field lies. */
function readRecords(dir: Buffer): { records: DirRecord[]; tail: Buffer } {
  const records: DirRecord[] = [];
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const id = dir.readUInt16LE(pos);
    const len = id === 0x0009 ? 12 : 6 + dir.readUInt32LE(pos + 2);
    records.push({ id, buf: dir.subarray(pos, pos + len) });
    pos += len;
    if (id === 0x0010) break;
  }
  return { records, tail: dir.subarray(pos) };
}

const data = (r: DirRecord) => r.buf.subarray(6);

function findEntry(cfb: CFB.CFB$Container, path: string): CFB.CFB$Entry | undefined {
  const i = cfb.FullPaths.findIndex((p) => p.replace(/^Root Entry/i, '').toLowerCase() === path.toLowerCase());
  return i >= 0 ? cfb.FileIndex[i] : undefined;
}

function setContent(entry: CFB.CFB$Entry, content: Buffer) {
  entry.content = content;
  entry.size = content.length;
}

/** Module names of a project, as the dir stream lists them. */
export function vbaModuleNames(bin: Uint8Array): string[] {
  const cfb = CFB.read(Buffer.from(bin), { type: 'buffer' });
  const dirEntry = findEntry(cfb, '/VBA/dir');
  if (!dirEntry) throw new Error('Not a VBA project: /VBA/dir is missing.');
  const { records } = readRecords(Buffer.from(decompressVba(dirEntry.content as Uint8Array)));
  const latin = new TextDecoder('latin1');
  return records.filter((r) => r.id === 0x0019).map((r) => latin.decode(data(r)));
}

export function patchVbaProject(bin: Uint8Array, patch: VbaPatch): Uint8Array {
  const cfb = CFB.read(Buffer.from(bin), { type: 'buffer' });
  const dirEntry = findEntry(cfb, '/VBA/dir');
  if (!dirEntry) throw new Error('Not a VBA project: /VBA/dir is missing.');
  const { records, tail } = readRecords(Buffer.from(decompressVba(dirEntry.content as Uint8Array)));

  const codePageRec = records.find((r) => r.id === 0x0003);
  const codePage = codePageRec ? data(codePageRec).readUInt16LE(0) : 1252;
  if (codePage !== 1252) throw new Error(`VBA project code page ${codePage} is not supported (only 1252).`);
  const decode = (b: Uint8Array) => new TextDecoder('windows-1252').decode(b);
  const latin = new TextDecoder('latin1');

  // Existing modules: name, stream and the record holding the source offset.
  type Mod = { name: string; stream?: string; offsetIndex?: number };
  const mods: Mod[] = [];
  let current: Mod | undefined;
  records.forEach((r, i) => {
    if (r.id === 0x0019) { current = { name: latin.decode(data(r)) }; mods.push(current); }
    else if (current && r.id === 0x001a) current.stream = latin.decode(data(r));
    else if (current && r.id === 0x0031) current.offsetIndex = i;
    else if (r.id === 0x002b) current = undefined;
  });
  const byName = new Map(mods.map((m) => [m.name.toLowerCase(), m]));
  for (const name of [...Object.keys(patch.append ?? {}), ...Object.keys(patch.edit ?? {})]) {
    if (!byName.has(name.toLowerCase())) throw new Error(`VBA module not found: ${name}. Modules: ${mods.map((m) => m.name).join(', ')}.`);
  }
  const clash = (patch.add ?? []).filter((m) => byName.has(m.name.toLowerCase()));
  if (clash.length) throw new Error(`VBA module(s) already in the project: ${clash.map((m) => m.name).join(', ')}.`);

  // Every existing module becomes source only.
  for (const m of mods) {
    if (!m.stream || m.offsetIndex === undefined) throw new Error(`VBA module ${m.name}: stream or offset missing in the dir stream.`);
    const entry = findEntry(cfb, `/VBA/${m.stream}`);
    if (!entry) throw new Error(`VBA module stream missing: ${m.stream}`);
    const offset = data(records[m.offsetIndex]).readUInt32LE(0);
    let source = decode(decompressVba(entry.content as Uint8Array, offset));
    const key = Object.keys(patch.edit ?? {}).find((k) => k.toLowerCase() === m.name.toLowerCase());
    if (key) source = patch.edit![key](source);
    const add = Object.entries(patch.append ?? {}).find(([k]) => k.toLowerCase() === m.name.toLowerCase())?.[1];
    if (add) {
      const extra = /^\s*Option Explicit\b/im.test(source) ? add.replace(/^\s*Option Explicit\s*$/gim, '') : add;
      source = `${source.replace(/\s+$/, '')}\r\n\r\n${extra.replace(/^(\r?\n)+/, '').replace(/\r\n|\r|\n/g, '\r\n').replace(/(\r\n)*$/, '')}\r\n`;
    }
    setContent(entry, Buffer.from(compressVba(encodeCp1252(source))));
    records[m.offsetIndex] = { id: 0x0031, buf: record(0x0031, u32(0)) };
  }

  // New modules: records before the terminator, a stream each, and the module count.
  const added = patch.add ?? [];
  if (added.length) {
    const countIndex = records.findIndex((r) => r.id === 0x000f);
    if (countIndex < 0) throw new Error('dir stream without PROJECTMODULES record.');
    records[countIndex] = { id: 0x000f, buf: record(0x000f, u16(data(records[countIndex]).readUInt16LE(0) + added.length)) };
    const end = records.findIndex((r) => r.id === 0x0010);
    records.splice(end, 0, ...added.map((m) => ({ id: 0x0019, buf: moduleRecords(m) })));
    for (const m of added) CFB.utils.cfb_add(cfb, `/VBA/${m.name}`, Buffer.from(compressVba(encodeCp1252(moduleSource(m)))));
  }
  setContent(dirEntry, Buffer.from(compressVba(Buffer.concat([...records.map((r) => r.buf), tail]))));

  // No compiled code: version 0xFFFF; performance caches go.
  const vbaProject = findEntry(cfb, '/VBA/_VBA_PROJECT');
  if (vbaProject) setContent(vbaProject, Buffer.from([0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00]));
  const caches = cfb.FullPaths.filter((p) => /\/VBA\/__SRP_[0-9a-f]+$/i.test(p));
  for (const p of caches) CFB.utils.cfb_del(cfb, p.replace(/^Root Entry/i, ''));
  // Deleting leaves the directory tree inconsistent until it is rebuilt.
  if (caches.length) CFB.utils.cfb_gc(cfb);

  if (added.length) {
    // PROJECT stream: module lines next to the existing ones, workspace entries at the end.
    const projectEntry = findEntry(cfb, '/PROJECT');
    if (projectEntry) {
      const lines = decode(projectEntry.content as Uint8Array).split(/\r\n/);
      const last = lines.reduce((acc, l, i) => (/^(Module|Document|Class|Package|BaseClass)=/i.test(l) ? i : acc), -1);
      const decl = added.map((m) => (m.type === 'document' ? `Document=${m.name}/&H00000000` : m.type === 'class' ? `Class=${m.name}` : `Module=${m.name}`));
      lines.splice(last + 1, 0, ...decl);
      const ws = lines.findIndex((l) => l.trim() === '[Workspace]');
      if (ws >= 0) {
        let insertAt = lines.length;
        while (insertAt > ws + 1 && lines[insertAt - 1] === '') insertAt--;
        lines.splice(insertAt, 0, ...added.map((m) => `${m.name}=0, 0, 0, 0, C`));
      }
      setContent(projectEntry, encodeCp1252(lines.join('\r\n')));
    }
    // PROJECTwm: name pairs, terminated by two zero bytes.
    const wm = findEntry(cfb, '/PROJECTwm');
    if (wm) {
      const content = Buffer.from(wm.content as Uint8Array);
      const entries = added.map((m) => Buffer.concat([encodeCp1252(m.name), Buffer.from([0]), Buffer.from(m.name, 'utf16le'), Buffer.from([0, 0])]));
      setContent(wm, Buffer.concat([content.subarray(0, content.length - 2), ...entries, Buffer.from([0, 0])]));
    }
  }
  return new Uint8Array(CFB.write(cfb, { type: 'buffer' }) as Buffer);
}
