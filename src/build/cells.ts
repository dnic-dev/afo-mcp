// Write single cells into an existing worksheet part (sheet XML), keeping row and cell order and
// the cell's style. Used for the formulas AfO items need in their cell (SAPSetFilterComponent …).

import { fromA1 } from './afo.js';

export interface CellWrite {
  /** A1 reference. */
  ref: string;
  /** Formula with or without the leading "="; SAP functions get the `_xll.` prefix files use. */
  formula?: string;
  value?: string | number;
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function colLetters(index: number): string {
  let s = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** A formula as stored in the file: no "=", XLL functions with `_xll.`. */
export function fileFormula(formula: string): string {
  return formula.replace(/^=/, '').replace(/(^|[^.\w])(SAP[A-Z]\w*)\s*\(/g, (_m, pre: string, fn: string) => `${pre}_xll.${fn}(`);
}

function cellXml(ref: string, w: CellWrite, style: string): string {
  if (w.formula !== undefined) return `<c r="${ref}"${style}><f>${esc(fileFormula(w.formula))}</f></c>`;
  if (typeof w.value === 'number') return `<c r="${ref}"${style}><v>${w.value}</v></c>`;
  return `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${esc(String(w.value ?? ''))}</t></is></c>`;
}

export function writeCells(sheetXml: string, writes: CellWrite[]): string {
  if (!writes.length) return sheetXml;
  let xml = sheetXml.replace(/<sheetData\s*\/>/, '<sheetData></sheetData>');
  if (!/<sheetData>/.test(xml) && !/<sheetData [^>]*>/.test(xml)) throw new Error('Worksheet without sheetData.');
  const open = /<sheetData[^>]*>/.exec(xml)!;
  const start = open.index + open[0].length;
  const end = xml.indexOf('</sheetData>', start);
  const body = xml.slice(start, end);

  // Rows as [number, xml]; self-closing rows count as empty.
  const rows = [...body.matchAll(/<row [^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)].map((m) => [Number(/\br="(\d+)"/.exec(m[0])![1]), m[0]] as [number, string]);
  for (const w of writes) {
    const { row, col } = fromA1(w.ref);
    const r = row + 1;
    const ref = `${colLetters(col)}${r}`;
    let i = rows.findIndex(([n]) => n === r);
    if (i < 0) {
      rows.push([r, `<row r="${r}"></row>`]);
      rows.sort((a, b) => a[0] - b[0]);
      i = rows.findIndex(([n]) => n === r);
    }
    let rowXml = rows[i][1].replace(/^(<row [^>]*?)\/>$/, '$1></row>');
    const cells = [...rowXml.matchAll(/<c [^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)].map((m) => m[0]);
    const existing = cells.find((c) => new RegExp(`\\br="${ref}"`).test(c));
    const style = existing ? (/\bs="\d+"/.exec(existing)?.[0] ? ` ${/\bs="\d+"/.exec(existing)![0]}` : '') : '';
    const kept = cells.filter((c) => c !== existing);
    const next = [...kept, cellXml(ref, w, style)].sort((a, b) => fromA1(/\br="([A-Z]+\d+)"/.exec(a)![1]).col - fromA1(/\br="([A-Z]+\d+)"/.exec(b)![1]).col);
    const head = /^<row [^>]*>/.exec(rowXml)![0].replace(/\s+spans="[^"]*"/, '');
    rows[i] = [r, `${head}${next.join('')}</row>`];
  }
  return xml.slice(0, start) + rows.map(([, x]) => x).join('') + xml.slice(end);
}
