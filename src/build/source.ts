// Take parts of an existing AfO workbook over verbatim: data sources with their whole stored
// state (filters, hierarchies, display settings, design rules, exceptions, conditions),
// planning functions and sequences, crosstab settings, the system connection (SAProuter, SNC)
// and the workbook settings. Rebuilding a workbook with a new layout then keeps every AfO
// feature it uses, including ones the spec cannot describe.

import { type AfoDefinition, isAfoApplication, parseAfoDefinition } from '../workbook/afo.js';
import { WorkbookPackage } from '../workbook/package.js';
import type { AfoWorkbookSpec } from './afo.js';

export function readSourceDefinition(file: string): string {
  const pkg = WorkbookPackage.open(file);
  for (const n of pkg.names.filter((p) => /^xl\/customProperty\d+\.bin$/.test(p))) {
    const xml = pkg.text(n);
    if (xml && isAfoApplication(xml)) return xml;
  }
  throw new Error(`No Analysis for Office definition in ${file}.`);
}

/** Elements `<tag>…</tag>` (tag may be a pattern), each with its first `<Key>`. Such elements never nest. */
export function keyedElements(xml: string, tag: string): Array<{ key: string; xml: string }> {
  return [...xml.matchAll(new RegExp(`<(${tag})>[\\s\\S]*?</\\1>`, 'g'))].map((m) => ({
    key: /<Key>([^<]*)<\/Key>/.exec(m[0])?.[1] ?? '',
    xml: m[0],
  }));
}

export function keyedElement(xml: string, tag: string, key: string): string | undefined {
  return keyedElements(xml, tag).find((e) => e.key === key)?.xml;
}

/** The first element of that name, as written (`<X>…</X>` or `<X />`). */
export function firstElement(xml: string, name: string): string | undefined {
  return new RegExp(`<${name}>[\\s\\S]*?</${name}>|<${name}\\s*/>`).exec(xml)?.[0];
}

export function replaceFirstKey(xml: string, key: string, esc: (s: string) => string): string {
  return xml.replace(/<Key>[^<]*<\/Key>/, () => `<Key>${esc(key)}</Key>`);
}

/** The connection of the source as AfO wrote it, SAProuter and SNC included. */
export function sourceSystemInfo(xml: string): string | undefined {
  const design = firstElement(xml, 'PlanningSystemDesignTime');
  return (design && firstElement(design, 'SystemInfo')) ?? firstElement(xml, 'SystemInfo');
}

export interface SourceResult {
  definition: string;
  parsed: AfoDefinition;
  errors: string[];
}

/**
 * Fill what entries with `from` leave out (query, bindings, names) from the source, so the rest
 * of the builder and the agent bridge help see complete entries. Returns the problems found.
 */
export function applySource(spec: AfoWorkbookSpec, definition: string): SourceResult {
  const parsed = parseAfoDefinition(definition);
  const errors: string[] = [];
  const renames = new Map<string, string>();
  for (const [i, dp] of (spec.dataProviders ?? []).entries()) {
    if (!dp?.from) continue;
    const src = parsed.dataProviders.find((d) => d.key === dp.from);
    if (!src) { errors.push(`dataProviders[${i}]: the source has no data source ${dp.from}. Source: ${parsed.dataProviders.map((d) => d.key).join(', ')}.`); continue; }
    renames.set(dp.from, dp.key);
    dp.query ??= src.dataSource ?? src.dataSourceName ?? '';
    dp.inputReady ??= src.inputReady;
  }
  for (const [i, pf] of (spec.planningFunctions ?? []).entries()) {
    if (!pf?.from) continue;
    const src = parsed.planningObjects.find((p) => p.key === pf.from);
    if (!src) { errors.push(`planningFunctions[${i}]: the source has no planning object ${pf.from}. Source: ${parsed.planningObjects.map((p) => p.key).join(', ') || 'none'}.`); continue; }
    pf.name ??= src.name ?? '';
    pf.description ??= src.description;
    pf.aggregationLevel ??= src.aggregationLevel ?? '';
    pf.characteristics ??= [];
    if (!pf.dataProviderKey) {
      const mapped = src.dataProviderKey ? renames.get(src.dataProviderKey) : undefined;
      if (!mapped) errors.push(`planningFunctions[${i}] (${pf.key}): its data source ${src.dataProviderKey} is not taken over; take it over with from or set dataProviderKey.`);
      else pf.dataProviderKey = mapped;
    }
  }
  for (const [i, g] of (spec.grids ?? []).entries()) {
    if (!g?.from) continue;
    const src = parsed.grids.find((x) => x.key === g.from);
    if (!src) { errors.push(`grids[${i}]: the source has no crosstab ${g.from}. Source: ${parsed.grids.map((x) => x.key).join(', ')}.`); continue; }
    if (!g.dataProviderKey) {
      const mapped = src.dataProviderKey ? renames.get(src.dataProviderKey) : undefined;
      if (!mapped) errors.push(`grids[${i}] (${g.key}): its data source ${src.dataProviderKey} is not taken over; take it over with from or set dataProviderKey.`);
      else g.dataProviderKey = mapped;
    }
  }
  for (const [i, f] of (spec.formulaItems ?? []).entries()) {
    if (f?.from && !keyedElement(definition, 'Fdd', f.from)) errors.push(`formulaItems[${i}]: the source has no formula item ${f.from}. Source: ${keyedElements(definition, 'Fdd').map((x) => x.key).join(', ') || 'none'}.`);
  }
  return { definition, parsed, errors };
}
