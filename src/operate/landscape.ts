// The BW connection of a system as the SAP Logon landscape of this machine describes it
// (%APPDATA%\SAP\Common\SAPUILandscape.xml and the global landscape next to it): server,
// system number, SAProuter and SNC name. New workbooks then connect the way SAP Logon does.

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { AfoSystem } from '../build/afo.js';

const attr = (tag: string, name: string) =>
  new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1]?.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

export function landscapeFiles(appData = process.env.APPDATA ?? ''): string[] {
  const dir = join(appData, 'SAP', 'Common');
  return ['SAPUILandscape.xml', 'SAPUILandscapeGlobal.xml'].map((f) => join(dir, f)).filter((f) => existsSync(f));
}

/** The connection for a system id from the SAP Logon landscape; the first matching entry wins. */
export function landscapeSystem(systemId: string, client: string, language: string, files = landscapeFiles()): AfoSystem {
  const xml = files.map((f) => readFileSync(f, 'utf8')).join('\n');
  const services = [...xml.matchAll(/<Service\s[^>]*>/g)].map((m) => m[0])
    .filter((t) => attr(t, 'type') === 'SAPGUI' && attr(t, 'systemid')?.toUpperCase() === systemId.toUpperCase());
  if (!services.length) throw new Error(`System ${systemId} is not in the SAP Logon landscape (${files.join(', ') || 'no landscape file found'}).`);
  const s = services[0];
  const [host, port] = (attr(s, 'server') ?? '').split(':');
  const routerId = attr(s, 'routerid');
  const router = routerId ? [...xml.matchAll(/<Router\s[^>]*>/g)].map((m) => m[0]).find((t) => attr(t, 'uuid') === routerId) : undefined;
  return {
    systemId: systemId.toUpperCase(),
    systemName: attr(s, 'name') ?? systemId,
    client,
    language,
    server: host || undefined,
    systemNumber: port && /^32\d\d$/.test(port) ? port.slice(2) : undefined,
    uuid: attr(s, 'uuid'),
    sapRouter: router ? attr(router, 'router') : undefined,
    sncPartnerName: attr(s, 'sncname') || undefined,
  };
}
