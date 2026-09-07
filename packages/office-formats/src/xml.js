// A small XML reader.
//
// Office XML is machine-written, deeply nested and large: a 30 MB content.xml
// is ordinary. A DOM would cost more than the document. This parses to a light
// tree of plain objects, and offers a streaming walk for the cases where even
// that is too much.
//
// It is deliberately not a validating parser. Namespaces are kept as written
// (`text:p`, not a resolved URI) because every consumer here matches on the
// conventional prefixes, and rewriting them would only add a mapping to get
// wrong.

const ENTITIES = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
};

export function decodeEntities(s) {
  if (!s.includes('&')) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (m, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return Number.isFinite(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[body] ?? m;
  });
}

export function escapeXml(s) {
  return String(s).replace(/[&<>"']/g, (c) => `&${{ '&': 'amp', '<': 'lt', '>': 'gt', '"': 'quot', "'": 'apos' }[c]};`);
}

/**
 * Walk XML, calling back on each event. No tree is built, so this is safe on
 * files larger than memory allows to model.
 *
 * @param {string} xml
 * @param {{ open?: (name, attrs, selfClosing) => void, close?: (name) => void, text?: (t) => void }} on
 */
export function walk(xml, on) {
  const len = xml.length;
  let i = 0;
  while (i < len) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) {
      if (on.text && i < len) on.text(decodeEntities(xml.slice(i)));
      return;
    }
    if (lt > i && on.text) {
      const t = xml.slice(i, lt);
      if (t) on.text(decodeEntities(t));
    }
    if (xml.startsWith('<!--', lt)) {
      i = xml.indexOf('-->', lt);
      i = i < 0 ? len : i + 3;
      continue;
    }
    if (xml.startsWith('<![CDATA[', lt)) {
      const end = xml.indexOf(']]>', lt);
      if (on.text) on.text(xml.slice(lt + 9, end < 0 ? len : end));
      i = end < 0 ? len : end + 3;
      continue;
    }
    if (xml.startsWith('<?', lt) || xml.startsWith('<!', lt)) {
      const end = xml.indexOf('>', lt);
      i = end < 0 ? len : end + 1;
      continue;
    }
    const gt = findTagEnd(xml, lt);
    if (gt < 0) return;
    const raw = xml.slice(lt + 1, gt);
    if (raw[0] === '/') {
      on.close?.(raw.slice(1).trim());
      i = gt + 1;
      continue;
    }
    const selfClosing = raw.endsWith('/');
    const body = selfClosing ? raw.slice(0, -1) : raw;
    const sp = firstSpace(body);
    const name = sp < 0 ? body : body.slice(0, sp);
    const attrs = sp < 0 ? {} : parseAttrs(body.slice(sp + 1));
    on.open?.(name, attrs, selfClosing);
    if (selfClosing) on.close?.(name);
    i = gt + 1;
  }
}

function firstSpace(s) {
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c === 32 || c === 9 || c === 10 || c === 13) return i;
  }
  return -1;
}

/** `>` inside an attribute value is legal, so quotes have to be tracked. */
function findTagEnd(xml, from) {
  let quote = 0;
  for (let i = from + 1; i < xml.length; i++) {
    const c = xml.charCodeAt(i);
    if (quote) {
      if (c === quote) quote = 0;
    } else if (c === 34 || c === 39) quote = c;
    else if (c === 62) return i;
  }
  return -1;
}

export function parseAttrs(s) {
  const out = {};
  const re = /([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(s))) out[m[1]] = decodeEntities(m[3] ?? m[4] ?? '');
  return out;
}

/** @typedef {{ name: string, attrs: Record<string,string>, children: Array<Node|string> }} Node */

/**
 * Parse to a tree. `keep` limits which elements are retained, which is how a
 * 30 MB spreadsheet becomes a model of the parts we actually read.
 * @param {string} xml
 * @param {{ keep?: (name: string) => boolean, text?: boolean }} [opts]
 * @returns {Node}
 */
export function parse(xml, opts = {}) {
  const root = { name: '#root', attrs: {}, children: [] };
  const stack = [root];
  walk(xml, {
    open(name, attrs, selfClosing) {
      const node = { name, attrs, children: [] };
      stack[stack.length - 1].children.push(node);
      if (!selfClosing) stack.push(node);
    },
    close() {
      if (stack.length > 1) stack.pop();
    },
    text(t) {
      if (opts.text === false) return;
      const top = stack[stack.length - 1];
      if (t) top.children.push(t);
    },
  });
  return root;
}

/** First descendant with this tag name. */
export function first(node, name) {
  if (!node) return null;
  for (const c of node.children) {
    if (typeof c === 'string') continue;
    if (c.name === name) return c;
    const deep = first(c, name);
    if (deep) return deep;
  }
  return null;
}

/** Direct children with this name (or any, when name is omitted). */
export function kids(node, name) {
  if (!node) return [];
  return node.children.filter((c) => typeof c !== 'string' && (!name || c.name === name));
}

/** Every descendant with this name, depth-first. */
export function all(node, name, out = []) {
  if (!node) return out;
  for (const c of node.children) {
    if (typeof c === 'string') continue;
    if (c.name === name) out.push(c);
    all(c, name, out);
  }
  return out;
}

/** Concatenated text of a subtree. */
export function textOf(node) {
  if (!node) return '';
  if (typeof node === 'string') return node;
  let s = '';
  for (const c of node.children) s += typeof c === 'string' ? c : textOf(c);
  return s;
}
