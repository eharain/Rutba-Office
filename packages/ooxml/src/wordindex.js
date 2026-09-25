/**
 * An index, as Word builds one from XE fields.
 *
 * Mark Entry writes an XE field where the entry is marked — ` XE "Main:Sub"
 * ` with `\t "See other"` for a cross-reference, `\r bookmark` for a page
 * range, `\b` and `\i` for a bold or italic page number — and Insert Index
 * writes an INDEX field whose result is every entry, sorted, grouped under
 * its letter, with the pages it was marked on. Nothing here touches a
 * package: this reads and writes the two field codes and lays the entries
 * out, so the window and the engine agree about what an index says.
 */

/** A field code's words: quoted text as one word, `\` switches as words. */
function words(instr) {
  const out = [];
  const s = String(instr);
  let i = 0;
  while (i < s.length) {
    if (/\s/.test(s[i])) { i++; continue; }
    if (s[i] === '"') {
      let j = i + 1;
      let text = '';
      while (j < s.length && s[j] !== '"') {
        if (s[j] === '\\' && (s[j + 1] === '"' || s[j + 1] === '\\')) { text += s[j + 1]; j += 2; continue; }
        text += s[j++];
      }
      out.push({ text, quoted: true });
      i = j + 1;
      continue;
    }
    let j = i;
    while (j < s.length && !/\s/.test(s[j])) j++;
    out.push({ text: s.slice(i, j), quoted: false });
    i = j;
  }
  return out;
}

/** A colon in an entry's own words is written `\:`, since a bare one starts a subentry. */
const escEntry = (t) => String(t || '').replace(/"/g, "'").replace(/:/g, '\\:').trim();

/**
 * An XE field's code, as Mark Entry writes it. `sub` may be a string or a
 * list (Word takes subentries seven deep); `crossRef` is the whole "See …"
 * text; `bookmark` a page range.
 */
export function xeInstr({ main, sub = '', crossRef = '', bookmark = '', bold = false, italic = false } = {}) {
  const subs = (Array.isArray(sub) ? sub : [sub]).map((x) => String(x || '').trim()).filter(Boolean);
  const entry = [escEntry(main), ...subs.map(escEntry)].join(':');
  let instr = ' XE "' + entry + '" ';
  if (crossRef) instr += '\\t "' + String(crossRef).replace(/"/g, "'") + '" ';
  else if (bookmark) instr += '\\r "' + bookmark + '" ';
  if (bold) instr += '\\b ';
  if (italic) instr += '\\i ';
  return instr;
}

/** An XE field's code, read back: `{ main, subs, crossRef, bookmark, bold, italic }`, or null. */
export function parseXeInstr(instr) {
  const w = words(instr);
  if (!w.length || w[0].text.toUpperCase() !== 'XE') return null;
  const entry = w[1] && !w[1].text.startsWith('\\') ? w[1].text : (w[1]?.quoted ? w[1].text : '');
  const parts = [];
  let cur = '';
  for (let i = 0; i < entry.length; i++) {
    if (entry[i] === '\\' && entry[i + 1] === ':') { cur += ':'; i++; continue; }
    if (entry[i] === ':') { parts.push(cur.trim()); cur = ''; continue; }
    cur += entry[i];
  }
  parts.push(cur.trim());
  const out = { main: parts[0] || '', subs: parts.slice(1).filter(Boolean), crossRef: '', bookmark: '', bold: false, italic: false };
  for (let i = 2; i < w.length; i++) {
    const s = w[i].text.toLowerCase();
    if (s === '\\t') out.crossRef = w[++i]?.text || '';
    else if (s === '\\r') out.bookmark = w[++i]?.text || '';
    else if (s === '\\b') out.bold = true;
    else if (s === '\\i') out.italic = true;
    else if (s === '\\f' || s === '\\y') i++;
  }
  return out.main ? out : null;
}

/**
 * An INDEX field's code from Insert Index's choices: `\c` columns, `\e` a
 * tab (right-aligned page numbers), `\r` run-in, `\h "A"` a letter over
 * each group, `\z` the language.
 */
export function indexInstr({ columns = 2, rightAlign = false, runIn = false, headings = true, lcid = 1033 } = {}) {
  let instr = ' INDEX ';
  if (rightAlign && !runIn) instr += '\\e "\t" ';
  if (headings) instr += '\\h "A" ';
  instr += '\\c "' + Math.max(1, Math.min(4, Number(columns) || 1)) + '" ';
  if (runIn) instr += '\\r ';
  instr += '\\z "' + lcid + '" ';
  return instr;
}

/** An INDEX field's code, read back. */
export function parseIndexInstr(instr) {
  const w = words(instr);
  if (!w.length || w[0].text.toUpperCase() !== 'INDEX') return null;
  const out = { columns: 1, rightAlign: false, runIn: false, headings: false, lcid: 1033 };
  for (let i = 1; i < w.length; i++) {
    const s = w[i].text.toLowerCase();
    if (s === '\\c') out.columns = Number(w[++i]?.text) || 1;
    else if (s === '\\e') { const v = w[++i]?.text ?? ''; out.rightAlign = v === '\t'; }
    else if (s === '\\h') { i++; out.headings = true; }
    else if (s === '\\r') out.runIn = true;
    else if (s === '\\z') out.lcid = Number(w[++i]?.text) || 1033;
  }
  return out;
}

const keyOf = (t) => String(t || '').toLocaleLowerCase();
const cmp = (a, b) => {
  const c = String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
  return c || String(a).localeCompare(String(b));
};

/** The group an entry sorts under: its letter, or Symbols for anything else. */
export function letterOf(text) {
  const c = String(text || '').trim().charAt(0);
  return /\p{L}/u.test(c) ? c.toLocaleUpperCase() : 'Symbols';
}

/**
 * The index, laid out. `entries` are the marked XE fields, each with the
 * page it landed on (`page`) and, for a range, the page its bookmark ends
 * on (`pageEnd`). Answers groups `{ letter, lines }`; each line `{ level,
 * text, pages: [{ text, bold, italic }], see }` — `level` 1 for a main
 * entry, 2… for subentries — or, run-in, one line per main entry carrying
 * its subentries in `runIn`.
 */
export function buildIndex(entries, { runIn = false } = {}) {
  // A tree of main → sub → … entries, each keeping its page references.
  const root = new Map();
  for (const e of entries || []) {
    const path = [e.main, ...(e.subs || [])].filter(Boolean);
    if (!path.length) continue;
    let level = root;
    let node = null;
    for (const name of path) {
      const k = keyOf(name);
      if (!level.has(k)) level.set(k, { name, refs: [], see: [], children: new Map() });
      node = level.get(k);
      level = node.children;
    }
    if (e.crossRef) {
      if (!node.see.includes(e.crossRef)) node.see.push(e.crossRef);
    } else if (e.page !== null && e.page !== undefined && e.page !== '') {
      node.refs.push({ page: Number(e.page), pageEnd: e.pageEnd != null && e.pageEnd !== '' ? Number(e.pageEnd) : null, bold: Boolean(e.bold), italic: Boolean(e.italic) });
    }
  }

  /** A node's page references, as Word lists them: in order, once each, a range as 3–5. */
  const pagesOf = (node) => {
    const seen = new Map();
    for (const r of node.refs) {
      const text = r.pageEnd && r.pageEnd !== r.page ? `${Math.min(r.page, r.pageEnd)}–${Math.max(r.page, r.pageEnd)}` : String(r.page);
      const k = text;
      const was = seen.get(k);
      if (was) { was.bold ||= r.bold; was.italic ||= r.italic; } else seen.set(k, { text, bold: r.bold, italic: r.italic, at: Math.min(r.page, r.pageEnd ?? r.page) });
    }
    return [...seen.values()].sort((a, b) => a.at - b.at || cmp(a.text, b.text)).map(({ text, bold, italic }) => ({ text, bold, italic }));
  };

  const sorted = (map) => [...map.values()].sort((a, b) => cmp(a.name, b.name));
  const groups = new Map();
  for (const node of sorted(root)) {
    const letter = letterOf(node.name);
    if (!groups.has(letter)) groups.set(letter, []);
    const lines = groups.get(letter);
    if (runIn) {
      const subs = [];
      const walk = (n, prefix) => {
        for (const c of sorted(n.children)) {
          subs.push({ text: prefix + c.name, pages: pagesOf(c), see: c.see });
          walk(c, prefix + c.name + ': ');
        }
      };
      walk(node, '');
      lines.push({ level: 1, text: node.name, pages: pagesOf(node), see: node.see, runIn: subs });
    } else {
      const walk = (n, level) => {
        lines.push({ level, text: n.name, pages: pagesOf(n), see: n.see });
        for (const c of sorted(n.children)) walk(c, level + 1);
      };
      walk(node, 1);
    }
  }
  // Symbols first, as Word puts them, then the letters in order.
  return [...groups.entries()]
    .sort(([a], [b]) => (a === 'Symbols' ? -1 : b === 'Symbols' ? 1 : cmp(a, b)))
    .map(([letter, lines]) => ({ letter, lines }));
}

/**
 * One line's words after the entry: the page numbers (", 3, 7" or a tab
 * then "3, 7") and any "See" cross-reference (". See Fruit"), as segments
 * `{ text, bold, italic }`.
 */
export function lineTail(line, { rightAlign = false } = {}) {
  const segs = [];
  if (line.pages.length) {
    segs.push({ text: rightAlign ? '\t' : ', ' });
    line.pages.forEach((p, i) => {
      if (i) segs.push({ text: ', ' });
      segs.push({ text: p.text, ...(p.bold ? { bold: true } : {}), ...(p.italic ? { italic: true } : {}) });
    });
  }
  for (const see of line.see || []) {
    // "See" and "See also" are italic, the name they point at is not.
    const m = /^(see(?:\s+also)?)\b\s*(.*)$/i.exec(see);
    segs.push({ text: '. ' });
    if (m) {
      segs.push({ text: m[1], italic: true });
      if (m[2]) segs.push({ text: ' ' + m[2] });
    } else segs.push({ text: see });
  }
  return segs;
}
