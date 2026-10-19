// Open an Office Open XML package (.xlsx/.xlsm) and hand out its parts. Only the parts a
// reader can use are decompressed: images, printer settings, themes and styles are skipped.

import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';

const SKIP = /^(xl\/media\/|xl\/printerSettings\/|xl\/theme\/|xl\/styles\.xml$|customXml\/|xl\/embeddings\/)/;

export class WorkbookPackage {
  readonly names: string[];
  private readonly parts: Record<string, Uint8Array>;

  private constructor(parts: Record<string, Uint8Array>, names: string[]) {
    this.parts = parts;
    this.names = names;
  }

  static open(file: string): WorkbookPackage {
    return WorkbookPackage.fromBuffer(readFileSync(file));
  }

  static fromBuffer(data: Uint8Array): WorkbookPackage {
    if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) {
      throw new Error('Not an Office Open XML workbook (no ZIP signature). Binary .xls and .xlsb files are not supported.');
    }
    const names: string[] = [];
    const parts = unzipSync(data, {
      filter: (f) => {
        names.push(f.name);
        return !SKIP.test(f.name);
      },
    });
    return new WorkbookPackage(parts, names);
  }

  has(name: string): boolean {
    return name in this.parts;
  }

  bytes(name: string): Uint8Array | undefined {
    return this.parts[name];
  }

  /** A part as text. Detects UTF-16LE (with or without BOM) as AfO writes it for custom properties. */
  text(name: string): string | undefined {
    const b = this.parts[name];
    if (!b) return undefined;
    return decodeText(b);
  }
}

export function decodeText(b: Uint8Array): string {
  if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) return Buffer.from(b.subarray(2)).toString('utf16le');
  if (b.length >= 3 && b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return strFromU8(b.subarray(3));
  // UTF-16LE without BOM: ASCII text has a zero in every odd byte.
  if (b.length >= 4 && b[1] === 0 && b[3] === 0 && b[0] !== 0) return Buffer.from(b).toString('utf16le');
  return strFromU8(b);
}

/** Resolve a relationship target against the folder of the part that owns the .rels file. */
export function resolveTarget(ownerPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = ownerPart.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

/** The .rels part that belongs to a part: `xl/worksheets/sheet1.xml` → `xl/worksheets/_rels/sheet1.xml.rels`. */
export function relsPartOf(part: string): string {
  const i = part.lastIndexOf('/');
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}
