// VBA macros from `xl/vbaProject.bin`: a compound file (CFB) whose `VBA/dir` stream lists the
// modules, each module stream holding its source compressed with the MS-OVBA algorithm.

import CFB from 'cfb';

export interface AfoApiCall {
  /** The AfO function: `SAPExecuteCommand`, `SAPSetVariable`, `SAPExecutePlanningFunction`, … */
  function: string;
  /** The first argument when it is a string literal: the command or the data source alias. */
  firstArgument?: string;
  line: number;
  code: string;
}

export interface Procedure {
  name: string;
  kind: string;
  line: number;
  afoCalls: AfoApiCall[];
  /** Other procedures of the project this one calls. */
  calls: string[];
}

export interface VbaModule {
  name: string;
  /** `standard`, `class` or `document` (sheet and workbook modules). */
  type: string;
  lines: number;
  procedures: Procedure[];
  source?: string;
  error?: string;
}

export interface VbaProject {
  modules: VbaModule[];
  error?: string;
}

/** MS-OVBA 2.4.1 decompression. */
export function decompressVba(data: Uint8Array, offset = 0): Uint8Array {
  if (data[offset] !== 0x01) throw new Error('not a compressed VBA container (signature byte missing)');
  const out: number[] = [];
  let pos = offset + 1;
  while (pos < data.length - 1) {
    const header = data[pos] | (data[pos + 1] << 8);
    const chunkSize = (header & 0x0fff) + 3;
    const compressed = (header & 0x8000) !== 0;
    const chunkEnd = Math.min(pos + chunkSize, data.length);
    pos += 2;
    const chunkStart = out.length;
    if (!compressed) {
      for (let i = 0; i < 4096 && pos < data.length; i++) out.push(data[pos++]);
      continue;
    }
    while (pos < chunkEnd) {
      const flags = data[pos++];
      for (let bit = 0; bit < 8 && pos < chunkEnd; bit++) {
        if ((flags & (1 << bit)) === 0) {
          out.push(data[pos++]);
          continue;
        }
        const token = data[pos] | (data[pos + 1] << 8);
        pos += 2;
        const decompressed = out.length - chunkStart;
        let bitCount = 4;
        while (1 << bitCount < decompressed) bitCount++;
        bitCount = Math.max(bitCount, 4);
        const lengthMask = 0xffff >> bitCount;
        const length = (token & lengthMask) + 3;
        const distance = (token >> (16 - bitCount)) + 1;
        const from = out.length - distance;
        if (from < 0) throw new Error('corrupt VBA copy token');
        for (let i = 0; i < length; i++) out.push(out[from + i]);
      }
    }
  }
  return Uint8Array.from(out);
}

interface DirModule {
  name: string;
  streamName: string;
  offset: number;
  type: string;
}

function parseDir(dir: Uint8Array): { codePage: number; modules: DirModule[] } {
  const view = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  let pos = 0;
  let codePage = 1252;
  const modules: DirModule[] = [];
  let cur: Partial<DirModule> | undefined;
  const latin = new TextDecoder('latin1');
  while (pos + 6 <= dir.length) {
    const id = view.getUint16(pos, true);
    let size = view.getUint32(pos + 2, true);
    pos += 6;
    // PROJECTVERSION declares 4 bytes but carries 6.
    if (id === 0x0009) size = 6;
    const body = dir.subarray(pos, pos + size);
    pos += size;
    switch (id) {
      case 0x0003:
        codePage = body[0] | (body[1] << 8);
        break;
      case 0x0019:
        cur = { name: latin.decode(body), type: 'standard' };
        break;
      case 0x0047:
        if (cur) cur.name = Buffer.from(body).toString('utf16le');
        break;
      case 0x001a:
        if (cur) cur.streamName = latin.decode(body);
        break;
      case 0x0032:
        if (cur) cur.streamName = Buffer.from(body).toString('utf16le');
        break;
      case 0x0031:
        if (cur) cur.offset = new DataView(body.buffer, body.byteOffset, 4).getUint32(0, true);
        break;
      case 0x0022:
        if (cur) cur.type = 'document';
        break;
      case 0x002b:
        if (cur?.name) modules.push({ name: cur.name, streamName: cur.streamName ?? cur.name, offset: cur.offset ?? 0, type: cur.type ?? 'standard' });
        cur = undefined;
        break;
    }
  }
  return { codePage, modules };
}

function decoderFor(codePage: number): TextDecoder {
  const label = codePage === 65001 ? 'utf-8' : `windows-${codePage}`;
  try {
    return new TextDecoder(label);
  } catch {
    return new TextDecoder('windows-1252');
  }
}

const PROC_RE = /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)/i;
const END_RE = /^\s*End\s+(Sub|Function|Property)\b/i;
const AFO_FUNCTION_RE = /\bSAP[A-Z]\w*/i;

/** VBA is case-insensitive: `SapGetVariable` and `SAPGetVariable` are the same call. */
function canonical(fn: string): string {
  return fn.replace(/^sap/i, 'SAP');
}
const RUN_RE = /\bRun\s*\(?\s*"(SAP\w+)"\s*(?:,\s*"([^"]*)")?/i;

/** Join VBA line continuations (` _` at the end) so each statement sits on one line. */
function logicalLines(source: string): Array<{ line: number; code: string }> {
  const out: Array<{ line: number; code: string }> = [];
  const physical = source.split(/\r\n|\r|\n/);
  for (let i = 0; i < physical.length; i++) {
    const start = i;
    let code = physical[i];
    while (/\s_\s*$/.test(code) && i + 1 < physical.length) code = code.replace(/\s_\s*$/, ' ') + physical[++i].trim();
    out.push({ line: start + 1, code });
  }
  return out;
}

function stripComment(code: string): string {
  let inString = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') inString = !inString;
    else if (ch === "'" && !inString) return code.slice(0, i);
  }
  return /^\s*Rem\s/i.test(code) ? '' : code;
}

export function analyseSource(source: string): Procedure[] {
  const procedures: Procedure[] = [];
  let cur: Procedure | undefined;
  for (const { line, code: raw } of logicalLines(source)) {
    const code = stripComment(raw);
    if (!code.trim()) continue;
    const head = PROC_RE.exec(code);
    if (head) {
      cur = { name: head[2], kind: head[1].replace(/\s+/g, ' '), line, afoCalls: [], calls: [] };
      procedures.push(cur);
      continue;
    }
    if (END_RE.test(code)) {
      cur = undefined;
      continue;
    }
    if (!cur) continue;
    const run = RUN_RE.exec(code);
    if (run) {
      cur.afoCalls.push({ function: canonical(run[1]), firstArgument: run[2], line, code: code.trim() });
      continue;
    }
    // Direct calls through an object (e.g. `AfO.SAPExecuteCommand`). Names inside string
    // literals, such as the range name in Range("SAPCrosstab1"), do not count.
    const direct = AFO_FUNCTION_RE.exec(code.replace(/"[^"]*"/g, (s) => '"' + ' '.repeat(s.length - 2) + '"'));
    if (direct) {
      const arg = /^\s*\(?\s*"([^"]*)"/.exec(code.slice(direct.index + direct[0].length))?.[1];
      cur.afoCalls.push({ function: canonical(direct[0]), firstArgument: arg, line, code: code.trim() });
    }
  }
  // Calls between procedures of the project: plain identifiers that name a known procedure.
  const names = new Map(procedures.map((p) => [p.name.toLowerCase(), p.name]));
  const lines = logicalLines(source);
  for (const p of procedures) {
    const seen = new Set<string>();
    let inside = false;
    for (const { line, code: raw } of lines) {
      if (line === p.line) {
        inside = true;
        continue;
      }
      if (!inside) continue;
      const code = stripComment(raw).replace(/"[^"]*"/g, '""');
      if (END_RE.test(code)) break;
      for (const w of code.match(/\b[A-Za-z_]\w*\b/g) ?? []) {
        const hit = names.get(w.toLowerCase());
        if (hit && hit !== p.name) seen.add(hit);
      }
    }
    p.calls = [...seen];
  }
  return procedures;
}

export function readVbaProject(bin: Uint8Array, includeSource: boolean): VbaProject {
  let cfb: CFB.CFB$Container;
  try {
    cfb = CFB.read(Buffer.from(bin), { type: 'buffer' });
  } catch (err) {
    return { modules: [], error: `vbaProject.bin is not a readable compound file: ${(err as Error).message}` };
  }
  // Full paths start with the root entry's name (`Root Entry/VBA/dir`); match on the rest.
  const index = new Map<string, number>();
  cfb.FullPaths.forEach((p, i) => index.set(p.slice(p.indexOf('/') + 1).toLowerCase(), i));
  const stream = (name: string): Uint8Array | undefined => {
    const i = index.get(name.toLowerCase());
    const content = i === undefined ? undefined : cfb.FileIndex[i]?.content;
    return content && content.length ? Uint8Array.from(content as ArrayLike<number>) : undefined;
  };
  const dirRaw = stream('VBA/dir');
  if (!dirRaw) return { modules: [], error: 'VBA/dir stream not found' };
  let dir: ReturnType<typeof parseDir>;
  try {
    dir = parseDir(decompressVba(dirRaw));
  } catch (err) {
    return { modules: [], error: `VBA/dir stream unreadable: ${(err as Error).message}` };
  }
  const decoder = decoderFor(dir.codePage);
  const modules: VbaModule[] = [];
  for (const m of dir.modules) {
    const mod: VbaModule = { name: m.name, type: m.type, lines: 0, procedures: [] };
    try {
      const raw = stream(`VBA/${m.streamName}`);
      if (!raw) throw new Error('module stream not found');
      const full = decoder.decode(decompressVba(raw, m.offset));
      // The dir stream gives sheet, workbook, class and form modules the same type code.
      if (m.type === 'document') {
        if (/^Attribute VB_Base = "0\{000208/m.test(full)) mod.type = 'document';
        else if (/^Attribute VB_Base = "0\{/m.test(full)) mod.type = 'form';
        else mod.type = 'class';
      }
      const source = full
        .split(/\r\n|\r|\n/)
        .filter((l) => !/^Attribute VB_/.test(l))
        .join('\n')
        .replace(/\n+$/, '');
      mod.lines = source ? source.split('\n').length : 0;
      mod.procedures = analyseSource(source);
      if (includeSource) mod.source = source;
    } catch (err) {
      mod.error = (err as Error).message;
    }
    modules.push(mod);
  }
  return { modules };
}
