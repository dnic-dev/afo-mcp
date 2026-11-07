// Hierarchies in the stored navigation state of a data source.
//
// Switching a hierarchy on has no Analysis for Office API function, so a workbook gets it from
// its stored state: AfO opens the data source with the hierarchy active and expanded to the
// stored level. Two kinds, as AfO saves them in the BICS view:
//   - a BW hierarchy on a characteristic: <CHARACTERISTICS><CHARACTERISTIC name="X"><HIERARCHY
//     name="H"><ACTIVE>true</ACTIVE><DRILL_LEVEL>n</DRILL_LEVEL>…, a direct child of BICS_VIEW
//     (not the CHARACTERISTICS of the selection state or of an axis);
//   - the compact display of rows (several characteristics in the rows shown as one tree):
//     <AXIS name="ROWS"><HIERARCHY><ACTIVE>true</ACTIVE><DRILL_LEVEL>n</DRILL_LEVEL>.
// Nodes opened or closed by hand are stored as DRILL_OPERATIONS; they are dropped for a
// characteristic whose hierarchy is set here, so the stored level applies as given.

import { decodeCompressed } from '../workbook/bics.js';
import { compressBlock, esc } from './afo.js';

export interface HierarchySpec {
  /** Characteristic, technical name. */
  characteristic: string;
  /** BW hierarchy to show; null shows the characteristic flat (hierarchy off). */
  hierarchy: string | null;
  /** Levels shown on open: 1 = top nodes only (default 2: the top nodes and the level below). */
  drillLevel?: number;
}

/** Compact display of rows: true, or with the levels shown on open (1 = first characteristic only). */
export type CompactRowsSpec = boolean | { drillLevel?: number };

interface Span { name: string; start: number; end: number; innerStart: number; innerEnd: number }

const TAG = /<(\/?)([A-Za-z_][\w.-]*)((?:\s+[^\s=>/]+\s*=\s*"[^"]*")*)\s*(\/?)>/g;

/** The direct child elements of the range [from, to) of an XML string. */
function childSpans(xml: string, from: number, to: number): Span[] {
  const out: Span[] = [];
  const stack: Array<{ name: string; start: number; innerStart: number }> = [];
  TAG.lastIndex = from;
  for (let m = TAG.exec(xml); m && m.index < to; m = TAG.exec(xml)) {
    const [whole, closing, name, , selfClosing] = m;
    if (closing) {
      const open = stack.pop();
      if (open && stack.length === 0) out.push({ name: open.name, start: open.start, end: m.index + whole.length, innerStart: open.innerStart, innerEnd: m.index });
    } else if (selfClosing) {
      if (stack.length === 0) out.push({ name, start: m.index, end: m.index + whole.length, innerStart: m.index + whole.length, innerEnd: m.index + whole.length });
    } else {
      stack.push({ name, start: m.index, innerStart: m.index + whole.length });
    }
  }
  return out;
}

const splice = (xml: string, start: number, end: number, insert: string) => xml.slice(0, start) + insert + xml.slice(end);

/** The BICS_VIEW element of a stored state. */
function viewSpan(state: string): Span {
  const view = childSpans(state, 0, state.length).find((s) => s.name === 'BICS_VIEW');
  if (!view) throw new Error('The stored state is not a BICS view.');
  return view;
}

/** AfO counts the levels below the top nodes; the spec counts the levels shown. */
const storedLevel = (shown: number | undefined) => Math.max(0, (shown ?? 2) - 1);

function hierarchyXml(h: HierarchySpec): string {
  return h.hierarchy === null
    ? '<HIERARCHY><ACTIVE>false</ACTIVE><DRILL_LEVEL>0</DRILL_LEVEL></HIERARCHY>'
    : `<HIERARCHY name="${esc(h.hierarchy)}" dueDate="00000000"><ACTIVE>true</ACTIVE><DRILL_LEVEL>${storedLevel(h.drillLevel)}</DRILL_LEVEL></HIERARCHY>`;
}

/** Elements AfO writes after the display CHARACTERISTICS block, in its order. */
const AFTER_CHARACTERISTICS = ['CUSTOM_PARAMETERS', 'DRILL_DOWN_CHARACTERISTICS_SORTING', 'VISUAL_SELECTION_STATE_FOR_RESULT_SET', 'IS_INPUT_ENABLED', 'AXES', 'DATA_CELLS', 'RESULT_SET_SIZE_LIMIT', 'DRILL_OPERATIONS', 'EXCEPTIONS', 'CONDITIONS', 'DOCUMENT_APPLICATION_ASSIGNED'];

function setCharacteristicHierarchy(state: string, h: HierarchySpec): string {
  const view = viewSpan(state);
  const top = childSpans(state, view.innerStart, view.innerEnd);
  const block = top.find((s) => s.name === 'CHARACTERISTICS');
  const entry = `<CHARACTERISTIC name="${esc(h.characteristic)}">${hierarchyXml(h)}</CHARACTERISTIC>`;
  if (!block) {
    const before = top.find((s) => AFTER_CHARACTERISTICS.includes(s.name));
    const at = before ? before.start : view.innerEnd;
    return splice(state, at, at, `<CHARACTERISTICS>${entry}</CHARACTERISTICS>`);
  }
  if (block.innerStart === block.innerEnd) return splice(state, block.start, block.end, `<CHARACTERISTICS>${entry}</CHARACTERISTICS>`);
  const nameAttr = new RegExp(`^<CHARACTERISTIC\\s+name="${h.characteristic.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`);
  const chars = childSpans(state, block.innerStart, block.innerEnd);
  const own = chars.find((s) => s.name === 'CHARACTERISTIC' && nameAttr.test(state.slice(s.start, s.innerStart)));
  if (!own) return splice(state, block.innerEnd, block.innerEnd, entry);
  if (own.innerStart === own.innerEnd) return splice(state, own.start, own.end, entry);
  const current = childSpans(state, own.innerStart, own.innerEnd).find((s) => s.name === 'HIERARCHY');
  // AfO writes HIERARCHY after CUSTOM_PARAMETERS; without one, at the end of the characteristic.
  return current ? splice(state, current.start, current.end, hierarchyXml(h)) : splice(state, own.innerEnd, own.innerEnd, hierarchyXml(h));
}

/** Drop the stored node operations (opened or closed by hand) of these characteristics. */
function dropDrillOperations(state: string, characteristics: Set<string>): string {
  const view = viewSpan(state);
  const block = childSpans(state, view.innerStart, view.innerEnd).find((s) => s.name === 'DRILL_OPERATIONS');
  if (!block) return state;
  const ops = childSpans(state, block.innerStart, block.innerEnd).filter((s) => s.name === 'DRILL_OPERATION');
  const kept = ops.filter((s) => {
    const c = /characteristic="([^"]*)"/.exec(state.slice(s.start, s.end))?.[1];
    return !c || !characteristics.has(c);
  });
  if (kept.length === ops.length) return state;
  const inner = kept.map((s) => state.slice(s.start, s.end)).join('');
  return splice(state, block.start, block.end, inner ? `<DRILL_OPERATIONS>${inner}</DRILL_OPERATIONS>` : '<DRILL_OPERATIONS />');
}

function setCompactRows(state: string, compact: CompactRowsSpec): string {
  const view = viewSpan(state);
  const axes = childSpans(state, view.innerStart, view.innerEnd).find((s) => s.name === 'AXES');
  const rows = axes && childSpans(state, axes.innerStart, axes.innerEnd)
    .find((s) => s.name === 'AXIS' && /name="ROWS"/.test(state.slice(s.start, s.innerStart)));
  if (!rows) throw new Error('compact display of rows needs the axis layout in the stored state: give axes.rows');
  const rowChars = childSpans(state, rows.innerStart, rows.innerEnd).find((s) => s.name === 'CHARACTERISTICS');
  if (rowChars) {
    const names = [...state.slice(rowChars.innerStart, rowChars.innerEnd).matchAll(/<CHARACTERISTIC>([^<]*)<\/CHARACTERISTIC>/g)].map((m) => m[1]);
    const dropped = dropDrillOperations(state, new Set(names));
    if (dropped !== state) return setCompactRows(dropped, compact);
  }
  const on = compact !== false;
  const level = on ? storedLevel(typeof compact === 'object' ? compact.drillLevel : 1) : 0;
  // Compact rows count the levels from the first characteristic: 1 shows it alone.
  const xml = `<HIERARCHY><ACTIVE>${on}</ACTIVE><DRILL_LEVEL>${on ? level + 1 : 0}</DRILL_LEVEL></HIERARCHY>`;
  const current = childSpans(state, rows.innerStart, rows.innerEnd).find((s) => s.name === 'HIERARCHY');
  return current ? splice(state, current.start, current.end, xml) : splice(state, rows.innerStart, rows.innerStart, xml);
}

/** A stored state (BICS_VIEW XML) with the hierarchies and the compact display of rows set. */
export function setStateHierarchies(state: string, hierarchies: HierarchySpec[] | undefined, compactRows?: CompactRowsSpec): string {
  let out = state;
  for (const h of hierarchies ?? []) out = setCharacteristicHierarchy(out, h);
  if (hierarchies?.length) out = dropDrillOperations(out, new Set(hierarchies.map((h) => h.characteristic)));
  if (compactRows !== undefined) out = setCompactRows(out, compactRows);
  return out;
}

/**
 * Apply hierarchies to a data source of an AfO definition (`<DataProvider>` XML). A data source
 * without a stored state gets the minimal one first.
 */
export function withStoredHierarchies(dpXml: string, hierarchies: HierarchySpec[] | undefined, compactRows: CompactRowsSpec | undefined, minimalState: () => string): { xml: string; notes: string[] } {
  const notes: string[] = [];
  const compressed = /<BICS_VIEW_HEADER compressed="true">([^<]*)<\/BICS_VIEW_HEADER>/.exec(dpXml);
  const state = setStateHierarchies(compressed ? decodeCompressed(compressed[1]) : minimalState(), hierarchies, compactRows);
  const header = `<BICS_VIEW_HEADER compressed="true">${compressBlock(state)}</BICS_VIEW_HEADER>`;
  if (compressed) return { xml: dpXml.replace(compressed[0], () => header), notes };
  notes.push('had no stored navigation state; a minimal one now carries the hierarchy, so it opens in the query default view otherwise');
  const plain = /<BICS_VIEW_HEADER>[\s\S]*?<\/BICS_VIEW_HEADER>|<BICS_VIEW_HEADER\s*\/>/.exec(dpXml);
  if (plain) return { xml: dpXml.replace(plain[0], () => header), notes };
  if (/<DimensionSynchronization\s*\/>/.test(dpXml)) return { xml: dpXml.replace(/<DimensionSynchronization\s*\/>/, (m) => `${m}${header}`), notes };
  return { xml: dpXml.replace(/<Comments\s*\/>|<Comments>|<\/DataProvider>/, (m) => `${header}${m}`), notes };
}

/** Problems in hierarchy settings; `what` names the data source. */
export function hierarchyProblems(what: string, hierarchies: unknown, compactRows: unknown, rowsGiven: boolean | undefined): string[] {
  const errors: string[] = [];
  const level = (v: unknown) => v === undefined || (Number.isInteger(v) && (v as number) >= 1);
  if (hierarchies !== undefined) {
    if (!Array.isArray(hierarchies)) errors.push(`${what}.hierarchies must be a list.`);
    else {
      const seen = new Set<string>();
      hierarchies.forEach((h: Partial<HierarchySpec>, i) => {
        const at = `${what}.hierarchies[${i}]`;
        if (typeof h?.characteristic !== 'string' || !h.characteristic.trim()) errors.push(`${at}: characteristic is required.`);
        else if (seen.has(h.characteristic)) errors.push(`${at}: ${h.characteristic} appears twice.`);
        else seen.add(h.characteristic);
        if (h?.hierarchy !== null && (typeof h?.hierarchy !== 'string' || !h.hierarchy.trim())) errors.push(`${at}: hierarchy must be a hierarchy name, or null for the flat display.`);
        if (!level(h?.drillLevel)) errors.push(`${at}: drillLevel must be a whole number from 1 (top nodes only).`);
      });
    }
  }
  if (compactRows !== undefined) {
    if (typeof compactRows !== 'boolean' && (typeof compactRows !== 'object' || compactRows === null || !level((compactRows as { drillLevel?: unknown }).drillLevel))) {
      errors.push(`${what}.compactRows must be true, false or { drillLevel } with a whole number from 1.`);
    } else if (compactRows !== false && rowsGiven === false) {
      errors.push(`${what}.compactRows needs axes.rows (the characteristics shown as one tree).`);
    }
  }
  return errors;
}
