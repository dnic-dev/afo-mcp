// A small XML tree parser. The AfO definition and the BICS state are regular, element-heavy
// documents; a plain tree with a few lookup helpers keeps the readers simple and predictable.

export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

const ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  if (!s.includes('&')) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, (m, e: string) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[e] ?? m;
  });
}

/** Strip a namespace prefix: `x:FmlaMacro` → `FmlaMacro`. */
export function localName(name: string): string {
  const i = name.indexOf(':');
  return i < 0 ? name : name.slice(i + 1);
}

export function parseXml(xml: string): XmlNode {
  const root: XmlNode = { name: '#document', attrs: {}, children: [], text: '' };
  const stack: XmlNode[] = [root];
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>|<\?[\s\S]*?\?>|<!DOCTYPE[^>]*>|<(\/?)([^\s/>]+)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const top = stack[stack.length - 1];
    if (m[1] !== undefined) {
      top.text += m[1];
    } else if (m[3] !== undefined) {
      if (m[2]) {
        // Close up to the matching open element; ignore a stray end tag. VML drawings carry
        // HTML-like content with unclosed elements such as <br>.
        const name = localName(m[3]);
        for (let i = stack.length - 1; i > 0; i--) {
          if (stack[i].name === name) {
            stack.length = i;
            break;
          }
        }
        continue;
      }
      const node: XmlNode = { name: localName(m[3]), attrs: parseAttrs(m[4] ?? ''), children: [], text: '' };
      top.children.push(node);
      if (!m[5]) stack.push(node);
    } else if (m[6] !== undefined) {
      top.text += decodeEntities(m[6]);
    }
  }
  return root;
}

function parseAttrs(s: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  for (const a of s.matchAll(/([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
    attrs[localName(a[1])] = decodeEntities(a[2] ?? a[3] ?? '');
  }
  return attrs;
}

export function child(node: XmlNode | undefined, name: string): XmlNode | undefined {
  return node?.children.find((c) => c.name === name);
}

export function children(node: XmlNode | undefined, name: string): XmlNode[] {
  return node ? node.children.filter((c) => c.name === name) : [];
}

/** Follow a path of child names: `path(n, 'DataSourceInfo', 'Id')`. */
export function path(node: XmlNode | undefined, ...names: string[]): XmlNode | undefined {
  let cur = node;
  for (const n of names) cur = child(cur, n);
  return cur;
}

/** Trimmed text of a descendant, or undefined when it is missing or empty. */
export function text(node: XmlNode | undefined, ...names: string[]): string | undefined {
  const t = path(node, ...names)?.text.trim();
  return t ? t : undefined;
}

export function bool(node: XmlNode | undefined, ...names: string[]): boolean | undefined {
  const t = text(node, ...names);
  if (t === undefined) return undefined;
  return t.toLowerCase() === 'true' || t === 'X';
}

export function num(node: XmlNode | undefined, ...names: string[]): number | undefined {
  const t = text(node, ...names);
  if (t === undefined) return undefined;
  const n = Number(t);
  return Number.isFinite(n) ? n : undefined;
}

/** All descendants with the given name, depth first. */
export function descendants(node: XmlNode | undefined, name: string, out: XmlNode[] = []): XmlNode[] {
  if (!node) return out;
  for (const c of node.children) {
    if (c.name === name) out.push(c);
    descendants(c, name, out);
  }
  return out;
}
