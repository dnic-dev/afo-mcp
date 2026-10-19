// Compressed AfO blocks and the BICS navigation state inside them.
//
// AfO stores data source state (`BICS_VIEW_HEADER`) and variable values (`Variables_Header`)
// as base64 with `compressed="true"`: a 4-byte little-endian length followed by a gzip stream
// that holds UTF-8 XML. Decoded, the state names the InfoProvider, the query, the filter per
// characteristic and the axis layout.

import { gunzipSync, inflateSync, unzlibSync, strFromU8 } from 'fflate';
import { type XmlNode, child, children, parseXml, text } from './xml.js';

export function decodeCompressed(base64: string): string {
  const buf = Buffer.from(base64.replace(/\s+/g, ''), 'base64');
  const attempts: Array<() => Uint8Array> = [
    () => gunzipSync(buf.subarray(4)),
    () => gunzipSync(buf),
    () => unzlibSync(buf.subarray(4)),
    () => inflateSync(buf.subarray(4)),
  ];
  for (const attempt of attempts) {
    try {
      return strFromU8(attempt());
    } catch {
      // try the next layout
    }
  }
  throw new Error('compressed block is neither gzip nor deflate');
}

/**
 * The content of a block that may be compressed: the parsed inner document for a
 * compressed block, the node itself otherwise.
 */
export function expandBlock(node: XmlNode | undefined): XmlNode | undefined {
  if (!node) return undefined;
  if (node.attrs.compressed !== 'true') return node;
  return parseXml(decodeCompressed(node.text));
}

const OPERATORS: Record<string, string> = {
  EQUAL: '',
  NOT_EQUAL: '<> ',
  GREATER: '> ',
  GREATER_EQUAL: '>= ',
  LESS: '< ',
  LESS_EQUAL: '<= ',
  CONTAINS_PATTERN: 'pattern ',
};

function memberText(member: XmlNode | undefined): string {
  if (!member) return '?';
  const type = member.attrs.type;
  if (type === 'LITERAL_MEMBER') return `"${text(member, 'VALUE') ?? ''}"`;
  const name = member.attrs.name ?? text(member, 'VALUE') ?? '';
  const shown = name === '' ? '(empty)' : name;
  return type === 'HIERARCHY_NODE' ? `node ${shown}` : shown;
}

/**
 * One selection as a compact string: `2024_HR3`, `not X`, `>= 202401`, `202401 .. 202412`,
 * `node <id>`, `"literal"`.
 */
export function describeSelection(sel: XmlNode): string {
  const exclude = text(sel, 'SIGN') === 'EXCLUDING' ? 'not ' : '';
  const op = sel.attrs.operator;
  const direct = child(sel, 'MEMBER');
  if (direct) return exclude + memberText(direct);
  const value = child(child(sel, 'VALUE'), 'MEMBER');
  const low = child(child(sel, 'LOW_VALUE'), 'MEMBER');
  const high = child(child(sel, 'HIGH_VALUE'), 'MEMBER');
  if (op === 'BETWEEN' || op === 'NOT_BETWEEN' || sel.attrs.type === 'INTERVAL' || (low && high)) {
    const prefix = op === 'NOT_BETWEEN' ? 'not ' : exclude;
    return `${prefix}${memberText(low ?? value)} .. ${memberText(high)}`;
  }
  const opText = op && op in OPERATORS ? OPERATORS[op] : op ? `${op.toLowerCase()} ` : '';
  return exclude + opText + memberText(value ?? low);
}

export function describeSelections(selections: XmlNode | undefined): string[] {
  return children(selections, 'SELECTION').map(describeSelection);
}

export interface Variable {
  name: string;
  type?: string;
  values: string[];
}

export function parseVariables(variables: XmlNode | undefined): Variable[] {
  return children(variables, 'VARIABLE').map((v) => ({
    name: v.attrs.name ?? '',
    type: v.attrs.type,
    values: describeSelections(child(v, 'SELECTIONS')),
  }));
}

export interface Filter {
  characteristic: string;
  selections: string[];
}

export interface BicsView {
  infoProvider?: string;
  query?: string;
  /** Planning filter name, or `!!A<aggregation level>` for an ad-hoc filter on the level. */
  selectionObject?: string;
  dueDate?: string;
  filters: Filter[];
  rows?: string[];
  columns?: string[];
  free?: string[];
  hierarchies?: Array<{ characteristic: string; hierarchy: string }>;
  exceptions?: string[];
  conditions?: string[];
  variables?: Variable[];
}

export function parseBicsView(view: XmlNode): BicsView {
  const out: BicsView = { filters: [] };
  out.infoProvider = child(view, 'INFO_PROVIDER')?.attrs.name || undefined;
  const query = child(view, 'QUERY');
  out.query = query?.attrs.name || text(view, 'SERVICE_DESCRIPTION', 'QUERY');
  const selObj = child(view, 'SELECTION_OBJECT');
  out.selectionObject = selObj?.attrs.name || undefined;
  out.dueDate = text(query, 'DUE_DATE') ?? text(selObj, 'DUE_DATE');

  const selection = child(child(view, 'SELECTION_STATE'), 'SELECTION');
  for (const ch of children(child(selection, 'CHARACTERISTICS'), 'CHARACTERISTIC')) {
    const selections = describeSelections(child(ch, 'SELECTIONS'));
    if (selections.length) out.filters.push({ characteristic: ch.attrs.name ?? '', selections });
  }

  for (const axis of children(child(view, 'AXES'), 'AXIS')) {
    const chars = children(child(axis, 'CHARACTERISTICS'), 'CHARACTERISTIC').map((c) => c.text.trim()).filter(Boolean);
    const key = axis.attrs.name === 'ROWS' ? 'rows' : axis.attrs.name === 'COLUMNS' ? 'columns' : axis.attrs.name === 'FREE' ? 'free' : undefined;
    if (key && chars.length) out[key] = chars;
  }

  const hierarchies = children(child(view, 'CHARACTERISTICS'), 'CHARACTERISTIC')
    .map((c) => ({ characteristic: c.attrs.name ?? '', hierarchy: child(c, 'HIERARCHY')?.attrs.name ?? '' }))
    .filter((h) => h.hierarchy);
  if (hierarchies.length) out.hierarchies = hierarchies;

  const exceptions = children(child(view, 'EXCEPTIONS'), 'EXCEPTION').map((e) => e.attrs.text || e.attrs.name || '?');
  if (exceptions.length) out.exceptions = exceptions;
  const conditions = children(child(view, 'CONDITIONS'), 'CONDITION').map((c) => c.attrs.text || c.attrs.name || '?');
  if (conditions.length) out.conditions = conditions;

  const variables = parseVariables(child(view, 'VARIABLES'));
  if (variables.length) out.variables = variables;
  return out;
}
