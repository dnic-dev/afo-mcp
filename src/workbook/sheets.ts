// Sheets, SAP cell formulas and macro buttons from the spreadsheet parts of a workbook.

import { type WorkbookPackage, relsPartOf, resolveTarget } from './package.js';
import { type XmlNode, children, descendants, parseXml } from './xml.js';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export function readRels(pkg: WorkbookPackage, ownerPart: string): Relationship[] {
  const xml = pkg.text(relsPartOf(ownerPart));
  if (!xml) return [];
  return descendants(parseXml(xml), 'Relationship').map((r) => {
    const external = r.attrs.TargetMode === 'External';
    return {
      id: r.attrs.Id ?? '',
      type: (r.attrs.Type ?? '').split('/').pop() ?? '',
      target: external ? (r.attrs.Target ?? '') : resolveTarget(ownerPart, r.attrs.Target ?? ''),
      external,
    };
  });
}

export interface SapCall {
  function: string;
  /** Arguments as written: string literals keep their quotes, so `"DS_1"` and `B5` stay apart. */
  args: string[];
}

/** The value of a string literal argument (`"DS_1"` → `DS_1`), undefined for anything else. */
export function literal(arg: string | undefined): string | undefined {
  const m = arg ? /^"((?:[^"]|"")*)"$/.exec(arg) : null;
  return m ? m[1].replace(/""/g, '"') : undefined;
}

export interface SapFormula {
  cell: string;
  /** For shared and array formulas: the range the formula covers. */
  range?: string;
  formula: string;
  calls: SapCall[];
}

export interface Button {
  name?: string;
  text?: string;
  /** Macro name as assigned, without the `[0]!` workbook prefix. */
  macro: string;
  anchor?: string;
}

export interface Sheet {
  name: string;
  part: string;
  /** VBA code name (`Tabelle1`, `Sheet1`). */
  codeName?: string;
  /** AfO sheet identifier `<code name>|<guid>`; AfO items point to sheets by it. */
  afoId?: string;
  state?: string;
  sapFormulas: SapFormula[];
  /** Number of SAP formula cells, including the cells a shared formula fills. */
  sapFormulaCells: number;
  buttons: Button[];
  /** Name of the custom property part with the AfO definition, when this sheet carries it. */
  afoDefinitionPart?: string;
}

const SAP_CALL_RE = /\b(?:_xll\.)?(SAP[A-Z]\w*)\s*\(/g;

/** Split the argument list that starts right after an opening parenthesis. */
function splitArgs(s: string, start: number): string[] {
  const args: string[] = [];
  let depth = 0;
  let cur = '';
  let inString = false;
  for (let i = start; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') inString = !inString;
    if (inString || ch === '"') {
      cur += ch;
      continue;
    }
    if (ch === '(') depth++;
    if (ch === ')') {
      if (depth === 0) {
        args.push(cur.trim());
        return args;
      }
      depth--;
    }
    if ((ch === ',' || ch === ';') && depth === 0) {
      args.push(cur.trim());
      cur = '';
      continue;
    }
    cur += ch;
  }
  if (cur.trim()) args.push(cur.trim());
  return args;
}

export function parseSapCalls(formula: string): SapCall[] {
  const calls: SapCall[] = [];
  for (const m of formula.matchAll(SAP_CALL_RE)) {
    const args = splitArgs(formula, (m.index ?? 0) + m[0].length);
    calls.push({ function: m[1], args: args.length === 1 && args[0] === '' ? [] : args });
  }
  return calls;
}

function decodeXmlText(s: string): string {
  return s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

function cellCount(range: string): number {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
  if (!m) return 1;
  const col = (l: string) => [...l].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  return (Math.abs(col(m[3]) - col(m[1])) + 1) * (Math.abs(Number(m[4]) - Number(m[2])) + 1);
}

/**
 * SAP formulas of a sheet. Scans cell by cell with a flat pattern instead of building a tree:
 * sheet parts can be tens of megabytes.
 */
export function scanSapFormulas(sheetXml: string): { formulas: SapFormula[]; cells: number } {
  const formulas: SapFormula[] = [];
  let cells = 0;
  const sharedSap = new Set<string>();
  let cell = '';
  for (const m of sheetXml.matchAll(/<c\s[^>]*?r="([A-Z]+\d+)"|<f(\s[^>]*)?(?:\/>|>([^<]*)<\/f>)/g)) {
    if (m[1]) {
      cell = m[1];
      continue;
    }
    const attrs = m[2] ?? '';
    const body = m[3] ? decodeXmlText(m[3]) : '';
    const si = /\bsi="(\d+)"/.exec(attrs)?.[1];
    if (!body) {
      // A cell that reuses a shared formula defined elsewhere.
      if (si !== undefined && sharedSap.has(si)) cells++;
      continue;
    }
    if (!/\bSAP[A-Z]\w*\s*\(/.test(body)) continue;
    const formula = body.replace(/_xll\./g, '');
    const ref = /\bref="([^"]+)"/.exec(attrs)?.[1];
    const f: SapFormula = { cell, formula: formula.startsWith('=') ? formula : `=${formula}`, calls: parseSapCalls(formula) };
    if (ref && ref !== cell) f.range = ref;
    if (si !== undefined) sharedSap.add(si);
    formulas.push(f);
    cells += /\bt="array"/.test(attrs) && ref ? cellCount(ref) : 1;
  }
  return { formulas, cells };
}

function normaliseMacro(macro: string): string {
  const bang = macro.lastIndexOf('!');
  return (bang >= 0 ? macro.slice(bang + 1) : macro).replace(/^'+|'+$/g, '');
}

function anchorCell(col?: string, row?: string): string | undefined {
  const c = Number(col);
  const r = Number(row);
  if (!Number.isFinite(c) || !Number.isFinite(r)) return undefined;
  let n = c + 1;
  let letters = '';
  while (n > 0) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return `${letters}${r + 1}`;
}

function textOf(node: XmlNode): string {
  return [node.text, ...node.children.map(textOf)].join(' ').replace(/\s+/g, ' ').trim();
}

/** Shapes with an assigned macro in a DrawingML part. */
export function drawingButtons(xml: string): Button[] {
  const doc = parseXml(xml);
  const root = doc.children[0];
  const buttons: Button[] = [];
  for (const anchor of root?.children ?? []) {
    const from = children(anchor, 'from')[0];
    const cell = from ? anchorCell(children(from, 'col')[0]?.text, children(from, 'row')[0]?.text) : undefined;
    const visit = (node: XmlNode) => {
      if (node.attrs.macro) {
        const props = descendants(node, 'cNvPr')[0];
        const txt = descendants(node, 'txBody').map(textOf).join(' ').trim();
        buttons.push({ name: props?.attrs.name, text: txt || undefined, macro: normaliseMacro(node.attrs.macro), anchor: cell });
        return;
      }
      node.children.forEach(visit);
    };
    visit(anchor);
  }
  return buttons;
}

/** Form controls with an assigned macro in a legacy VML drawing part. */
export function vmlButtons(xml: string): Button[] {
  const doc = parseXml(xml);
  const buttons: Button[] = [];
  for (const shape of descendants(doc, 'shape')) {
    const data = children(shape, 'ClientData')[0];
    const macro = data ? children(data, 'FmlaMacro')[0]?.text.trim() : undefined;
    if (!macro) continue;
    const anchor = children(data!, 'Anchor')[0]?.text.split(',').map((s) => s.trim());
    const box = descendants(shape, 'textbox')[0];
    buttons.push({
      name: shape.attrs.alt || shape.attrs.id,
      text: box ? textOf(box) || undefined : undefined,
      macro: normaliseMacro(macro),
      anchor: anchor && anchor.length >= 4 ? anchorCell(anchor[0], anchor[2]) : undefined,
    });
  }
  return buttons;
}

export interface WorkbookStructure {
  sheets: Sheet[];
  externalLinks: string[];
  hasVba: boolean;
}

export function readStructure(pkg: WorkbookPackage, opts: { formulas: boolean; buttons: boolean }): WorkbookStructure {
  const wbXml = pkg.text('xl/workbook.xml');
  if (!wbXml) throw new Error('xl/workbook.xml not found: not an Excel workbook');
  const wb = parseXml(wbXml);
  const rels = readRels(pkg, 'xl/workbook.xml');
  const byId = new Map(rels.map((r) => [r.id, r]));
  const sheets: Sheet[] = [];
  for (const s of descendants(wb, 'sheet')) {
    const rel = byId.get(s.attrs.id ?? '');
    if (!rel || rel.type !== 'worksheet') continue;
    const sheet: Sheet = { name: s.attrs.name ?? '', part: rel.target, sapFormulas: [], sapFormulaCells: 0, buttons: [] };
    if (s.attrs.state && s.attrs.state !== 'visible') sheet.state = s.attrs.state;
    const xml = pkg.text(rel.target);
    if (xml) {
      sheet.codeName = /<sheetPr\b[^>]*\bcodeName="([^"]*)"/.exec(xml)?.[1];
      const sheetRels = readRels(pkg, rel.target);
      const relById = new Map(sheetRels.map((r) => [r.id, r]));
      for (const cp of xml.matchAll(/<customPr\s+name="([^"]*)"\s+r:id="([^"]*)"/g)) {
        const target = relById.get(cp[2])?.target;
        if (!target) continue;
        if (cp[1] === '_pios_id') sheet.afoId = pkg.text(target)?.replace(/\0+$/, '').trim();
        else if (/^serializedData/.test(cp[1])) sheet.afoDefinitionPart = target;
      }
      if (opts.formulas) {
        const scan = scanSapFormulas(xml);
        sheet.sapFormulas = scan.formulas;
        sheet.sapFormulaCells = scan.cells;
      }
      if (opts.buttons) {
        for (const r of sheetRels) {
          const part = pkg.text(r.target);
          if (!part) continue;
          if (r.type === 'drawing') sheet.buttons.push(...drawingButtons(part));
          else if (r.type === 'vmlDrawing') sheet.buttons.push(...vmlButtons(part));
        }
      }
    }
    sheets.push(sheet);
  }
  const externalLinks: string[] = [];
  for (const r of rels.filter((x) => x.type === 'externalLink')) {
    for (const lr of readRels(pkg, r.target)) if (lr.external) externalLinks.push(lr.target);
  }
  return { sheets, externalLinks, hasVba: pkg.has('xl/vbaProject.bin') };
}
