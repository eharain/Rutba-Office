// A web page read into the document model odt.js reads an .odt into, so that
// it opens as a document — headings, paragraphs in their alignment, runs in
// their looks, links, lists at their levels, tables with their spans and
// pictures — where it came in as its words alone, one paragraph a block.
//
// The page is tokenised leniently, the way a browser forgives it: void
// elements need no end, a <p>, <li>, <td> or <tr> ends when the next one
// starts, an end tag closes up to the element it names, and what is inside
// <script>, <style>, <head> and comments is not text. Whitespace collapses
// as it does on screen except inside <pre>. Looks come from the tags (b,
// strong, i, em, u, s, del, sup, sub, code, a) and from style attributes
// (color, background-color, font-size, font-family, font-weight,
// font-style, text-decoration, text-align); pictures embedded as data: URIs
// come with their bytes, at the size the page gives them.

const VOID = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr']);
const RAW = new Set(['script', 'style', 'textarea', 'title']);
const BLOCK = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'table', 'tr', 'td', 'th', 'thead', 'tbody', 'tfoot', 'caption', 'blockquote', 'pre', 'hr', 'section', 'article', 'header', 'footer', 'main', 'nav', 'aside', 'figure', 'figcaption', 'dl', 'dt', 'dd', 'address', 'center', 'form', 'fieldset']);
/**
 * Implied ends, as a browser's parser has them: a start tag of a kind here
 * closes the nearest open element of its group (with whatever is open
 * inside it), but not past the list or table it belongs to.
 */
const GROUP = {
  li: [['li'], ['ul', 'ol', 'table']], dt: [['dt', 'dd'], ['dl', 'table']], dd: [['dt', 'dd'], ['dl', 'table']],
  tr: [['tr'], ['table']], td: [['td', 'th'], ['tr', 'table']], th: [['td', 'th'], ['tr', 'table']],
  thead: [['thead', 'tbody', 'tfoot'], ['table']], tbody: [['thead', 'tbody', 'tfoot'], ['table']], tfoot: [['thead', 'tbody', 'tfoot'], ['table']],
};
/** A block start tag closes an open paragraph — unless a cell, a caption, a list item or a button stands between them. */
const P_SCOPE = new Set(['table', 'td', 'th', 'caption', 'button', 'li', 'dd', 'dt']);

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', copy: '©', reg: '®', trade: '™', hellip: '…', mdash: '—', ndash: '–', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', bull: '•', middot: '·', euro: '€', pound: '£', yen: '¥', deg: '°', times: '×', divide: '÷', laquo: '«', raquo: '»', sect: '§', para: '¶', emsp: ' ', ensp: ' ', thinsp: ' ' };
const decode = (s) => String(s).replace(/&(#x[0-9a-f]+|#\d+|[a-z]+[0-9]*);?/gi, (m, e) => {
  if (e[0] === '#') { const n = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1)); return Number.isFinite(n) && n > 0 && n < 0x110000 ? String.fromCodePoint(n) : m; }
  return ENTITIES[e.toLowerCase()] ?? m;
});

/** The page as a tree of { name, attrs, children } with strings for text. */
export function parseHtml(html) {
  const root = { name: '#root', attrs: {}, children: [] };
  const stack = [root];
  const top = () => stack[stack.length - 1];
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<![^>]*>|<\?[^>]*>|<\/\s*([a-zA-Z][\w:-]*)\s*>|<([a-zA-Z][\w:-]*)((?:\s+[^\s=>/]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?)*)\s*(\/?)>/g;
  let at = 0;
  let m;
  const text = (t) => { if (t) top().children.push(decode(t)); };
  while ((m = re.exec(html))) {
    text(html.slice(at, m.index));
    at = re.lastIndex;
    if (m[1]) {
      const name = m[1].toLowerCase();
      // An end tag closes up to the element it names, if one is open.
      for (let i = stack.length - 1; i > 0; i--) if (stack[i].name === name) { stack.length = i; break; }
      continue;
    }
    if (!m[2]) continue; // comment, doctype, CDATA, processing instruction
    const name = m[2].toLowerCase();
    const attrs = {};
    for (const a of (m[3] || '').matchAll(/([^\s=>/]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g)) attrs[a[1].toLowerCase()] = decode(a[2] ?? a[3] ?? a[4] ?? '');
    if (BLOCK.has(name)) {
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i].name === 'p') { stack.length = i; break; }
        if (P_SCOPE.has(stack[i].name)) break;
      }
    }
    const group = GROUP[name];
    if (group) {
      for (let i = stack.length - 1; i > 0; i--) {
        if (group[0].includes(stack[i].name)) { stack.length = i; break; }
        if (group[1].includes(stack[i].name)) break;
      }
    }
    const el = { name, attrs, children: [] };
    top().children.push(el);
    if (RAW.has(name)) {
      const end = new RegExp(`</${name}\\s*>`, 'ig');
      end.lastIndex = at;
      const close = end.exec(html);
      // A title's or a text area's words are escaped like any text; a script's and a style's are not.
      const raw = html.slice(at, close ? close.index : html.length);
      el.children.push(name === 'title' || name === 'textarea' ? decode(raw) : raw);
      at = close ? end.lastIndex : html.length;
      re.lastIndex = at;
      continue;
    }
    if (!VOID.has(name) && !m[4]) stack.push(el);
  }
  text(html.slice(at));
  return root;
}

/** A style attribute as a map of its declarations. */
function styleOf(el) {
  const out = {};
  for (const part of String(el.attrs?.style || '').split(';')) {
    const i = part.indexOf(':');
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

const NAMED = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', gray: '#808080', grey: '#808080', silver: '#C0C0C0', maroon: '#800000', navy: '#000080', purple: '#800080', teal: '#008080', olive: '#808000', orange: '#FFA500', lime: '#00FF00', aqua: '#00FFFF', fuchsia: '#FF00FF' };
function colour(v) {
  if (!v) return null;
  const s = String(v).trim().toLowerCase();
  if (NAMED[s]) return NAMED[s];
  let m = /^#([0-9a-f]{6})$/.exec(s);
  if (m) return `#${m[1].toUpperCase()}`;
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(s);
  if (m) return `#${(m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase()}`;
  m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(s);
  if (m) return `#${[m[1], m[2], m[3]].map((x) => Math.min(255, Number(x)).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
  return null;
}
/** A CSS length in px, or null. */
function lengthPx(v, base = 16) {
  const m = /^(-?[\d.]+)\s*(px|pt|em|rem|cm|mm|in|%)?$/i.exec(String(v || '').trim());
  if (!m) return null;
  const n = Number(m[1]);
  switch ((m[2] || 'px').toLowerCase()) {
    case 'pt': return n * 4 / 3;
    case 'em': case 'rem': return n * base;
    case 'cm': return n * 96 / 2.54;
    case 'mm': return n * 96 / 25.4;
    case 'in': return n * 96;
    case '%': return null;
    default: return n;
  }
}
const FONT_SIZES = { 1: 7.5, 2: 10, 3: 12, 4: 13.5, 5: 18, 6: 24, 7: 36 };
const KEYWORD_SIZES = { 'xx-small': 7, 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24 };

/** The looks an element gives the runs inside it. */
function lookOf(el) {
  const look = {};
  switch (el.name) {
    case 'b': case 'strong': look.bold = true; break;
    case 'i': case 'em': case 'cite': case 'var': case 'dfn': look.italic = true; break;
    case 'u': case 'ins': look.underline = true; break;
    case 's': case 'strike': case 'del': look.strike = true; break;
    case 'sup': look.vertical = 'superscript'; break;
    case 'sub': look.vertical = 'subscript'; break;
    case 'code': case 'kbd': case 'samp': case 'tt': case 'pre': look.font = 'Consolas'; break;
    case 'mark': look.highlight = '#FFFF00'; break;
    case 'th': look.bold = true; break;
    case 'font': {
      if (colour(el.attrs.color)) look.color = colour(el.attrs.color);
      if (el.attrs.face) look.font = el.attrs.face.split(',')[0].replace(/['"]/g, '').trim();
      if (FONT_SIZES[el.attrs.size]) look.size = FONT_SIZES[el.attrs.size];
      break;
    }
    default: break;
  }
  const st = styleOf(el);
  if (colour(st.color)) look.color = colour(st.color);
  if (colour(st['background-color'] || st.background)) look.highlight = colour(st['background-color'] || st.background);
  if (st['font-weight']) look.bold = st['font-weight'] === 'bold' || st['font-weight'] === 'bolder' || Number(st['font-weight']) >= 600;
  if (st['font-style']) look.italic = st['font-style'] === 'italic' || st['font-style'] === 'oblique';
  if (/underline/.test(st['text-decoration'] || st['text-decoration-line'] || '')) look.underline = true;
  if (/line-through/.test(st['text-decoration'] || st['text-decoration-line'] || '')) look.strike = true;
  if (st['font-family']) look.font = st['font-family'].split(',')[0].replace(/['"]/g, '').trim();
  if (st['font-size']) {
    const px = KEYWORD_SIZES[st['font-size']] ? KEYWORD_SIZES[st['font-size']] * 4 / 3 : lengthPx(st['font-size']);
    if (px) look.size = Math.round(px * 0.75 * 2) / 2;
  }
  if (st['vertical-align'] === 'super') look.vertical = 'superscript';
  if (st['vertical-align'] === 'sub') look.vertical = 'subscript';
  return look;
}

const ALIGN = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', justify: 'both' };
const LIST_FORMAT = { decimal: '1', 'lower-alpha': 'a', 'lower-latin': 'a', 'upper-alpha': 'A', 'upper-latin': 'A', 'lower-roman': 'i', 'upper-roman': 'I', 1: '1', a: 'a', A: 'A', i: 'i', I: 'I' };
const BULLETS = ['•', '◦', '▪'];

/**
 * Read a page whole: { blocks, lists, images, title } in odt.js's shapes —
 * its pictures, where embedded, by names under Pictures/.
 */
export function readHtmlDocument(html) {
  const root = parseHtml(String(html ?? ''));
  const find = (node, name) => {
    for (const c of node.children || []) {
      if (typeof c === 'string') continue;
      if (c.name === name) return c;
      const deep = find(c, name);
      if (deep) return deep;
    }
    return null;
  };
  const titleEl = find(root, 'title');
  const title = titleEl ? String(titleEl.children[0] || '').replace(/\s+/g, ' ').trim() || null : null;
  const body = find(root, 'body') || root;
  const lists = new Map();
  const images = new Map();
  let listCount = 0;
  let imageCount = 0;

  const picture = (el) => {
    const src = String(el.attrs.src || '');
    const m = /^data:image\/(png|jpe?g|gif|bmp|webp);base64,([\s\S]*)$/i.exec(src);
    if (!m) return null;
    const ext = m[1].toLowerCase() === 'jpeg' ? 'jpg' : m[1].toLowerCase();
    const name = `Pictures/image${++imageCount}.${ext}`;
    images.set(name, new Uint8Array(Buffer.from(m[2].replace(/\s+/g, ''), 'base64')));
    const st = styleOf(el);
    const width = lengthPx(st.width) ?? (Number(el.attrs.width) || null);
    const height = lengthPx(st.height) ?? (Number(el.attrs.height) || null);
    return { href: name, width: width || 96, height: height || (width || 96), name: el.attrs.alt || el.attrs.title || null };
  };

  /** A block's paragraph looks: its alignment, a quotation's indent, its direction. */
  const paragraphLook = (el, inherited) => {
    const st = styleOf(el);
    const out = { ...inherited };
    const align = ALIGN[(st['text-align'] || el.attrs.align || '').toLowerCase()];
    if (align) out.align = align;
    if (el.name === 'center') out.align = 'center';
    if (el.name === 'blockquote') out.indentLeft = (out.indentLeft || 0) + 40;
    const ml = lengthPx(st['margin-left'] || st['padding-left']);
    if (ml && el.name !== 'ul' && el.name !== 'ol') out.indentLeft = (out.indentLeft || 0) + ml;
    const ti = lengthPx(st['text-indent']);
    if (ti) out.indentFirst = ti;
    if (/^(always|page)$/.test(st['page-break-before'] || st['break-before'] || '')) out.pageBreakBefore = true;
    if ((el.attrs.dir || '').toLowerCase() === 'rtl' || st.direction === 'rtl') out.rtl = true;
    return out;
  };

  /**
   * Blocks from a node's children: inline content gathers into paragraphs,
   * a block element starts its own. `ctx` carries the paragraph look, the
   * run look, the list a paragraph belongs to and whether whitespace keeps.
   */
  function blocksOf(node, ctx) {
    const out = [];
    let runs = [];
    let pendingBreak = false;
    const flush = () => {
      // Spaces at a paragraph's ends are not words.
      while (runs.length && runs[0].text != null && !runs[0].text.trim() && !ctx.pre) runs.shift();
      if (runs[0]?.text != null && !ctx.pre) runs[0].text = runs[0].text.replace(/^\s+/, '');
      const last = runs[runs.length - 1];
      if (last?.text != null && !ctx.pre) last.text = last.text.replace(/\s+$/, '');
      runs = runs.filter((r) => r.text !== '' || r.br || r.tab || r.image);
      if (runs.length) out.push({ type: 'paragraph', heading: ctx.heading || null, ...ctx.para, ...(ctx.list ? { list: ctx.list } : {}), ...(pendingBreak ? { pageBreakBefore: true } : {}), runs });
      if (runs.length) pendingBreak = false;
      runs = [];
    };
    const push = (text, look) => {
      let t = ctx.pre ? text : text.replace(/[\s​]+/g, ' ');
      if (!t) return;
      const last = runs[runs.length - 1];
      // Collapsed spaces do not double across tags.
      if (!ctx.pre && /^ /.test(t) && (!last || (last.text != null && / $/.test(last.text)))) t = t.slice(1);
      if (!t) return;
      if (ctx.pre && t.includes('\n')) {
        t.split('\n').forEach((piece, i) => { if (i) runs.push({ ...look, br: true }); if (piece) runs.push({ ...look, text: piece.replace(/\t/g, '    ') }); });
        return;
      }
      if (last && last.text != null && JSON.stringify({ ...last, text: '' }) === JSON.stringify({ ...look, text: '' })) last.text += t;
      else runs.push({ ...look, text: t });
    };
    const inline = (n, look) => {
      for (const c of n.children || []) {
        if (typeof c === 'string') { push(c, look); continue; }
        if (BLOCK.has(c.name)) { flush(); block(c); continue; }
        switch (c.name) {
          case 'br': runs.push({ ...look, br: true }); break;
          case 'img': { const img = picture(c); if (img) runs.push({ ...(look.link ? { link: look.link } : {}), image: img }); break; }
          case 'a': inline(c, { ...look, ...lookOf(c), ...(c.attrs.href && !/^javascript:/i.test(c.attrs.href) ? { link: c.attrs.href } : {}) }); break;
          case 'script': case 'style': case 'head': case 'title': case 'noscript': case 'template': case 'svg': case 'object': case 'iframe': case 'select': case 'button': break;
          case 'input': if (c.attrs.type === 'checkbox') push(c.attrs.checked != null ? '☒ ' : '☐ ', look); break;
          default: inline(c, { ...look, ...lookOf(c) });
        }
      }
    };
    const block = (el) => {
      const para = paragraphLook(el, ctx.para);
      const look = { ...ctx.look, ...lookOf(el) };
      switch (el.name) {
        case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6': {
          // A heading's look is its Heading style's; what the page sets on it comes with it.
          out.push(...blocksOf(el, { ...ctx, heading: Number(el.name[1]), para, look: lookOf(el) }));
          break;
        }
        case 'ul': case 'ol': out.push(...listOf(el, ctx, para)); break;
        case 'table': out.push(tableOf(el, ctx)); break;
        case 'hr': flush(); break;
        case 'pre': out.push(...blocksOf(el, { ...ctx, para, look, pre: true })); break;
        case 'li': out.push(...blocksOf(el, { ...ctx, para, look })); break;
        default: {
          if (/^(always|page)$/.test(styleOf(el)['page-break-before'] || styleOf(el)['break-before'] || '')) pendingBreak = true;
          const inner = blocksOf(el, { ...ctx, para, look });
          if (pendingBreak && inner[0]) { inner[0].pageBreakBefore = true; pendingBreak = false; }
          out.push(...inner);
        }
      }
    };
    const listOf = (el, outer, para) => {
      // A list inside an item is a level in of the same list; a list on its own starts one.
      const nested = Boolean(outer.list);
      const id = nested ? outer.list.id : `h${++listCount}`;
      const style = nested ? outer.list.style : `HTML${listCount}`;
      const level = nested ? Math.min(8, outer.list.level + 1) : 0;
      if (!lists.has(style)) lists.set(style, []);
      const levels = lists.get(style);
      if (!levels[level]) {
        const st = styleOf(el);
        const type = (st['list-style-type'] || el.attrs.type || '').trim();
        const numbered = el.name === 'ol' ? !/^(disc|circle|square|none)$/i.test(type) : /^(decimal|lower|upper)/i.test(type);
        levels[level] = numbered
          ? { kind: 'number', format: LIST_FORMAT[type] || LIST_FORMAT[type.toLowerCase()] || '1', prefix: '', suffix: '.', start: Math.max(1, Number(el.attrs.start) || 1), display: 1, indent: 48 + 24 * level, hanging: 24 }
          : { kind: 'bullet', char: /circle/i.test(type) ? '◦' : /square/i.test(type) ? '▪' : BULLETS[level % 3], indent: 48 + 24 * level, hanging: 24 };
      }
      const items = [];
      for (const c of el.children || []) {
        if (typeof c === 'string') { if (c.trim()) items.push(...blocksOf({ children: [c] }, { ...outer, para, list: { id, style, level } })); continue; }
        if (c.name === 'li') items.push(...blocksOf(c, { ...outer, para: paragraphLook(c, para), look: { ...outer.look, ...lookOf(c) }, list: { id, style, level } }));
        else if (c.name === 'ul' || c.name === 'ol') items.push(...listOf(c, { ...outer, list: { id, style, level } }, para));
        else items.push(...blocksOf({ children: [c] }, { ...outer, para, list: { id, style, level } }));
      }
      return items;
    };
    const tableOf = (el, outer) => {
      const rowsOf = (n) => (n.children || []).flatMap((c) => (typeof c === 'string' ? [] : c.name === 'tr' ? [c] : ['thead', 'tbody', 'tfoot'].includes(c.name) ? rowsOf(c) : []));
      const trs = rowsOf(el);
      const grid = [];
      const taken = new Map(); // "row:col" → covered
      let widths = [];
      const colsOf = (n) => (n.children || []).flatMap((c) => (typeof c === 'string' ? [] : c.name === 'col' ? [c] : c.name === 'colgroup' ? colsOf(c) : []));
      for (const col of colsOf(el)) {
        const span = Math.max(1, Number(col.attrs.span) || 1);
        const w = lengthPx(styleOf(col).width) ?? (Number(col.attrs.width) || null);
        for (let k = 0; k < span; k++) widths.push(w);
      }
      trs.forEach((tr, ri) => {
        const row = grid[ri] || (grid[ri] = []);
        let col = 0;
        for (const cell of tr.children || []) {
          if (typeof cell === 'string' || (cell.name !== 'td' && cell.name !== 'th')) continue;
          while (taken.has(`${ri}:${col}`)) { row[col] = { covered: true }; col += 1; }
          const colspan = Math.max(1, Math.min(64, Number(cell.attrs.colspan) || 1));
          const rowspan = Math.max(1, Math.min(trs.length - ri, Number(cell.attrs.rowspan) || 1));
          const st = styleOf(cell);
          const fill = colour(st['background-color'] || st.background || cell.attrs.bgcolor);
          const w = lengthPx(st.width) ?? (Number(cell.attrs.width) || null);
          if (w && colspan === 1 && widths[col] == null) widths[col] = w;
          const look = { ...outer.look, ...lookOf(cell) };
          const blocks = blocksOf(cell, { ...outer, para: paragraphLook(cell, {}), look, list: null, heading: null });
          row[col] = { blocks, colspan, rowspan, ...(fill ? { fill } : {}), ...(el.attrs.border && el.attrs.border !== '0' || /solid|double/.test(st.border || '') ? { border: true } : {}) };
          for (let r = 0; r < rowspan; r++) for (let k = 0; k < colspan; k++) if (r || k) taken.set(`${ri + r}:${col + k}`, true);
          for (let k = 1; k < colspan; k++) row[col + k] = { covered: true };
          col += colspan;
        }
        while (taken.has(`${ri}:${col}`)) { row[col] = { covered: true }; col += 1; }
      });
      const width = Math.max(0, ...grid.map((r) => r.length));
      for (const row of grid) for (let c = 0; c < width; c++) if (!row[c]) row[c] = taken.has(`${grid.indexOf(row)}:${c}`) ? { covered: true } : { blocks: [], colspan: 1, rowspan: 1 };
      widths = Array.from({ length: width }, (_, i) => widths[i] ?? null);
      return { type: 'table', columns: widths, rows: grid };
    };

    inline(node, ctx.look);
    flush();
    return out;
  }

  const blocks = blocksOf(body, { para: {}, look: {}, list: null, heading: null, pre: false });
  return { blocks, lists, page: null, images, title };
}
