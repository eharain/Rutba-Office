// A small reader for the XML a camera answers with (ONVIF speaks SOAP):
// elements, attributes and text, by local name, whatever prefix the camera
// chose. No document type is read, so nothing in an answer is ever expanded
// or fetched; an answer that carries one is refused whole.

const MAX_CHARS = 1024 * 1024;
const MAX_NODES = 20000;
const MAX_DEPTH = 64;

const NAMED = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

/** Text with its five named entities and its numbered ones put back. */
export function unescapeXml(text) {
  if (!text.includes('&')) return text;
  return text.replace(/&(#x[0-9a-f]{1,6}|#[0-9]{1,7}|[a-z]{2,4});/gi, (whole, name) => {
    if (name[0] !== '#') return NAMED[name.toLowerCase()] ?? whole;
    const code = name[1] === 'x' || name[1] === 'X' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
    return code > 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : '';
  });
}

/** Text made safe to put inside an element or an attribute. */
export function escapeXml(text) {
  return String(text).replace(/[<>&"']/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' })[c]);
}

const local = (name) => {
  const at = name.indexOf(':');
  return at < 0 ? name : name.slice(at + 1);
};

function attributes(text) {
  const out = {};
  const re = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    if (m[1] === 'xmlns' || m[1].startsWith('xmlns:')) continue;
    out[local(m[1])] = unescapeXml(m[2] ?? m[3]);
  }
  return out;
}

/**
 * The document as a tree: `{ name, attrs, children, text }`, names without
 * their prefix. Null when it is not XML that can be read, is too large, or
 * carries a document type.
 */
export function parseXml(input) {
  const text = typeof input === 'string' ? input : Buffer.from(input).toString('utf8');
  if (!text || text.length > MAX_CHARS) return null;
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) return null;
  const root = { name: '', attrs: {}, children: [], text: '' };
  const stack = [root];
  let nodes = 0;
  let at = 0;
  while (at < text.length) {
    const open = text.indexOf('<', at);
    const top = stack[stack.length - 1];
    if (open < 0) {
      top.text += unescapeXml(text.slice(at));
      break;
    }
    if (open > at) top.text += unescapeXml(text.slice(at, open));
    if (text.startsWith('<!--', open)) {
      const end = text.indexOf('-->', open + 4);
      if (end < 0) return null;
      at = end + 3;
      continue;
    }
    if (text.startsWith('<![CDATA[', open)) {
      const end = text.indexOf(']]>', open + 9);
      if (end < 0) return null;
      top.text += text.slice(open + 9, end);
      at = end + 3;
      continue;
    }
    if (text.startsWith('<?', open)) {
      const end = text.indexOf('?>', open + 2);
      if (end < 0) return null;
      at = end + 2;
      continue;
    }
    const close = text.indexOf('>', open + 1);
    if (close < 0) return null;
    const inner = text.slice(open + 1, close);
    at = close + 1;
    if (inner[0] === '/') {
      if (stack.length < 2 || local(inner.slice(1).trim()) !== top.name) return null;
      stack.pop();
      continue;
    }
    if (inner[0] === '!') return null;
    const selfClosing = inner.endsWith('/');
    const body = selfClosing ? inner.slice(0, -1) : inner;
    const space = body.search(/\s/);
    const name = local(space < 0 ? body : body.slice(0, space));
    if (!name || /[<&"'=]/.test(name)) return null;
    nodes += 1;
    if (nodes > MAX_NODES) return null;
    const node = { name, attrs: space < 0 ? {} : attributes(body.slice(space)), children: [], text: '' };
    top.children.push(node);
    if (!selfClosing) {
      if (stack.length >= MAX_DEPTH) return null;
      stack.push(node);
    }
  }
  if (stack.length !== 1 || root.children.length !== 1) return null;
  return root.children[0];
}

/** The first element of that local name at or under a node, in the order of the document; null when there is none. */
export function find(node, name) {
  if (!node) return null;
  if (node.name === name) return node;
  for (const child of node.children) {
    const got = find(child, name);
    if (got) return got;
  }
  return null;
}

/** Every element of that local name under a node, not looking inside one that matched. */
export function findAll(node, name, out = []) {
  if (!node) return out;
  for (const child of node.children) {
    if (child.name === name) out.push(child);
    else findAll(child, name, out);
  }
  return out;
}

/** The text of the first element of that name, trimmed; '' when there is none. */
export function textOf(node, name) {
  const got = find(node, name);
  return got ? got.text.trim() : '';
}

/** A node along a path of local names, each a direct child of the one before; null when the path breaks. */
export function child(node, ...names) {
  let at = node;
  for (const name of names) {
    at = at?.children.find((c) => c.name === name) || null;
    if (!at) return null;
  }
  return at;
}
