// Write a VBA project (`xl/vbaProject.bin`) from source code only.
//
// The project carries no compiled code: `_VBA_PROJECT` declares version 0xFFFF, which tells the
// VBA host that no performance cache exists, so Excel compiles the modules from source when the
// workbook opens. Streams follow MS-OVBA: PROJECT (text), PROJECTwm, VBA/dir and one
// stream per module, each compressed with the MS-OVBA container format.

import CFB from 'cfb';
import { randomUUID } from 'node:crypto';

export interface VbaModuleSource {
  name: string;
  /** `standard` module, `document` module (workbook or sheet, by code name), or `class`. */
  type: 'standard' | 'document' | 'class';
  /** For document modules: `workbook` or `sheet`. */
  document?: 'workbook' | 'sheet';
  /** VBA code without the `Attribute VB_…` header lines; they are generated. */
  code: string;
}

const WORKBOOK_BASE = '0{00020819-0000-0000-C000-000000000046}';
const SHEET_BASE = '0{00020820-0000-0000-C000-000000000046}';

/** Windows-1252 encoding; characters outside the code page become `?`. */
export function encodeCp1252(s: string): Buffer {
  const special: Record<number, number> = {
    0x20ac: 0x80, 0x201a: 0x82, 0x0192: 0x83, 0x201e: 0x84, 0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87,
    0x02c6: 0x88, 0x2030: 0x89, 0x0160: 0x8a, 0x2039: 0x8b, 0x0152: 0x8c, 0x017d: 0x8e, 0x2018: 0x91,
    0x2019: 0x92, 0x201c: 0x93, 0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97, 0x02dc: 0x98,
    0x2122: 0x99, 0x0161: 0x9a, 0x203a: 0x9b, 0x0153: 0x9c, 0x017e: 0x9e, 0x0178: 0x9f,
  };
  const out = Buffer.alloc(s.length);
  let n = 0;
  for (const ch of s) {
    const c = ch.codePointAt(0)!;
    out[n++] = c < 0x80 || (c >= 0xa0 && c <= 0xff) ? c : special[c] ?? 0x3f;
  }
  return out.subarray(0, n);
}

/** One chunk (at most 4096 bytes) with MS-OVBA 2.4.1.3 tokens: literals and copy tokens. */
function compressChunk(chunk: Uint8Array): number[] {
  const body: number[] = [];
  let pos = 0;
  while (pos < chunk.length) {
    const flagIndex = body.length;
    body.push(0);
    let flags = 0;
    for (let bit = 0; bit < 8 && pos < chunk.length; bit++) {
      // The split between offset and length bits depends on how far the chunk has got.
      let bitCount = 4;
      while (1 << bitCount < pos) bitCount++;
      const maxLength = (0xffff >> bitCount) + 3;
      let bestLength = 0;
      let bestOffset = 0;
      for (let candidate = pos - 1; candidate >= 0; candidate--) {
        let length = 0;
        while (length < maxLength && pos + length < chunk.length && chunk[candidate + length] === chunk[pos + length]) length++;
        if (length > bestLength) {
          bestLength = length;
          bestOffset = pos - candidate;
          if (length === maxLength) break;
        }
      }
      if (bestLength >= 3) {
        const token = ((bestOffset - 1) << (16 - bitCount)) | (bestLength - 3);
        body.push(token & 0xff, token >> 8);
        flags |= 1 << bit;
        pos += bestLength;
      } else {
        body.push(chunk[pos++]);
      }
    }
    body[flagIndex] = flags;
  }
  return body;
}

/**
 * MS-OVBA compressed container (2.4.1.3), as Office writes it: every chunk compressed with
 * copy tokens; a full chunk that would not shrink is stored raw. Excel refuses module streams
 * that carry raw chunks where it expects compressed ones.
 */
export function compressVba(data: Uint8Array): Uint8Array {
  const out: number[] = [0x01];
  for (let start = 0; start < data.length; start += 4096) {
    const chunk = data.subarray(start, start + 4096);
    const body = compressChunk(chunk);
    if (body.length > 4096 && chunk.length === 4096) {
      out.push(0xff, 0x3f);
      for (const b of chunk) out.push(b);
      continue;
    }
    const header = 0xb000 | (body.length + 2 - 3);
    out.push(header & 0xff, header >> 8);
    for (const b of body) out.push(b);
  }
  return Uint8Array.from(out);
}

function record(id: number, body: Uint8Array | number[]): Buffer {
  const head = Buffer.alloc(6);
  head.writeUInt16LE(id, 0);
  head.writeUInt32LE(body.length, 2);
  return Buffer.concat([head, Buffer.from(body)]);
}

const u16 = (v: number) => { const b = Buffer.alloc(2); b.writeUInt16LE(v); return b; };
const u32 = (v: number) => { const b = Buffer.alloc(4); b.writeUInt32LE(v); return b; };
const utf16 = (s: string) => Buffer.from(s, 'utf16le');

function dirStream(projectName: string, modules: VbaModuleSource[]): Buffer {
  const parts: Buffer[] = [
    record(0x0001, u32(1)), // SYSKIND: Win32
    record(0x0002, u32(0x409)), // LCID
    record(0x0014, u32(0x409)), // LCIDINVOKE
    record(0x0003, u16(1252)), // CODEPAGE
    record(0x0004, encodeCp1252(projectName)),
    record(0x0005, []), record(0x0040, []), // DOCSTRING
    record(0x0006, []), record(0x003d, []), // HELPFILEPATH
    record(0x0007, u32(0)), // HELPCONTEXT
    record(0x0008, u32(0)), // LIBFLAGS
    Buffer.concat([u16(0x0009), u32(4), u32(1), u16(0)]), // VERSION: Reserved=4, major, minor
    record(0x000c, []), record(0x003c, []), // CONSTANTS
    record(0x000f, u16(modules.length)), // MODULES count
    record(0x0013, u16(0xffff)), // PROJECTCOOKIE
  ];
  for (const m of modules) {
    parts.push(
      record(0x0019, encodeCp1252(m.name)),
      record(0x0047, utf16(m.name)),
      record(0x001a, encodeCp1252(m.name)),
      record(0x0032, utf16(m.name)),
      record(0x001c, []), record(0x0048, []),
      record(0x0031, u32(0)), // source starts at offset 0: no performance cache
      record(0x001e, u32(0)),
      record(0x002c, u16(0xffff)),
      record(m.type === 'standard' ? 0x0021 : 0x0022, []),
      record(0x002b, []),
    );
  }
  parts.push(record(0x0010, []));
  return Buffer.concat(parts);
}

function moduleSource(m: VbaModuleSource): string {
  const lines = [`Attribute VB_Name = "${m.name}"`];
  if (m.type === 'document') {
    lines.push(
      `Attribute VB_Base = "${m.document === 'workbook' ? WORKBOOK_BASE : SHEET_BASE}"`,
      'Attribute VB_GlobalNameSpace = False',
      'Attribute VB_Creatable = False',
      'Attribute VB_PredeclaredId = True',
      'Attribute VB_Exposed = True',
      'Attribute VB_TemplateDerived = False',
      'Attribute VB_Customizable = True',
    );
  } else if (m.type === 'class') {
    lines.unshift('VERSION 1.0 CLASS', 'BEGIN', '  MultiUse = -1  \'True', 'END');
    lines.push('Attribute VB_GlobalNameSpace = False', 'Attribute VB_Creatable = False', 'Attribute VB_PredeclaredId = False', 'Attribute VB_Exposed = False');
  }
  const code = m.code.replace(/\r\n|\r|\n/g, '\r\n').replace(/(\r\n)*$/, '');
  return lines.join('\r\n') + '\r\n' + code + '\r\n';
}

function projectStream(projectName: string, modules: VbaModuleSource[]): Buffer {
  const lines = [`ID="{${randomUUID().toUpperCase()}}"`];
  for (const m of modules) {
    if (m.type === 'document') lines.push(`Document=${m.name}/&H00000000`);
    else if (m.type === 'class') lines.push(`Class=${m.name}`);
    else lines.push(`Module=${m.name}`);
  }
  lines.push(`Name="${projectName}"`, 'HelpContextID="0"', 'VersionCompatible32="393222000"', '', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '', '[Workspace]');
  for (const m of modules) lines.push(`${m.name}=0, 0, 0, 0, C`);
  return encodeCp1252(lines.join('\r\n') + '\r\n');
}

function projectWm(modules: VbaModuleSource[]): Buffer {
  const parts: Buffer[] = [];
  for (const m of modules) parts.push(encodeCp1252(m.name), Buffer.from([0]), utf16(m.name), Buffer.from([0, 0]));
  parts.push(Buffer.from([0, 0]));
  return Buffer.concat(parts);
}

export function writeVbaProject(modules: VbaModuleSource[], projectName = 'VBAProject'): Uint8Array {
  const cfb = CFB.utils.cfb_new();
  CFB.utils.cfb_add(cfb, '/PROJECT', projectStream(projectName, modules));
  CFB.utils.cfb_add(cfb, '/PROJECTwm', projectWm(modules));
  // _VBA_PROJECT: reserved 0x61CC, version 0xFFFF (no compiled code), reserved bytes.
  CFB.utils.cfb_add(cfb, '/VBA/_VBA_PROJECT', Buffer.from([0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00]));
  CFB.utils.cfb_add(cfb, '/VBA/dir', Buffer.from(compressVba(dirStream(projectName, modules))));
  for (const m of modules) CFB.utils.cfb_add(cfb, `/VBA/${m.name}`, Buffer.from(compressVba(encodeCp1252(moduleSource(m)))));
  // cfb seeds a placeholder entry ("\u0001Sh33tJ5"); CFB.write leaves it out by itself.
  return new Uint8Array(CFB.write(cfb, { type: 'buffer' }) as Buffer);
}
