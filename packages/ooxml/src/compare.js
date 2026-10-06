/**
 * Review → Compare: two documents read side by side and written as one —
 * the revised document, with what changed marked as Word marks a revision:
 * a word put in as an insertion, a word taken out as a deletion, a
 * paragraph added or removed whole — the compared document Word's Compare
 * makes. Paragraphs of the body are matched by their words; in a paragraph
 * that changed, words are matched one by one. A paragraph that did not
 * change is the revised one exactly as it was; one that changed keeps its
 * paragraph properties and its first run's look. Tables and the section
 * come from the revised document as they are.
 */
import { OoxmlPackage } from './package.js';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** The body's top-level elements, in order: each its XML, whether it is a paragraph, and a paragraph's words. */
function topElements(body) {
  const out = [];
  const re = /<(\/?)(w:[A-Za-z]+)\b[^>]*?(\/?)>/g;
  let depth = 0;
  let start = -1;
  let m;
  while ((m = re.exec(body))) {
    const [tag, closing, , self] = m;
    if (closing) {
      depth -= 1;
      if (depth === 0) out.push(body.slice(start, m.index + tag.length));
    } else if (self) {
      if (depth === 0) out.push(tag);
    } else {
      if (depth === 0) start = m.index;
      depth += 1;
    }
  }
  return out.map((xml) => ({ xml, para: /^<w:p[\s>]/.test(xml), text: /^<w:p[\s>]/.test(xml) ? textOf(xml) : null }));
}

/** A paragraph's words as Word shows them: its text, tabs as tabs, deleted text left out. */
function textOf(p) {
  let out = '';
  for (const m of p.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\/>|<w:br\b[^>]*\/>/g)) out += m[1] !== undefined ? unesc(m[1]) : m[0].startsWith('<w:tab') ? '\t' : '\n';
  return out;
}

/** The longest run of equal items two lists share, as pairs of their indexes. */
function lcs(a, b, same = (x, y) => x === y) {
  const n = a.length;
  const m = b.length;
  const table = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) table[i][j] = same(a[i], b[j]) ? table[i + 1][j + 1] + 1 : Math.max(table[i + 1][j], table[i][j + 1]);
  const pairs = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (same(a[i], b[j])) { pairs.push([i, j]); i++; j++; } else if (table[i + 1][j] >= table[i][j + 1]) i++; else j++;
  }
  return pairs;
}

/** How alike two paragraphs' words are, 0 to 1 — enough alike and a pair is one paragraph changed, not one gone and one new. */
function likeness(a, b) {
  const wa = a.split(/\s+/).filter(Boolean);
  const wb = b.split(/\s+/).filter(Boolean);
  if (!wa.length || !wb.length) return 0;
  return (2 * lcs(wa, wb).length) / (wa.length + wb.length);
}

/**
 * The compared body: `original` and `revised` as `word/document.xml`
 * text, the revised one's body with the changes marked by `author` at
 * `date`. Answers the new document part and how many changes it marks.
 */
export function compareBodies(original, revised, { author = 'Compare', date = new Date().toISOString().replace(/\.\d+Z$/, 'Z') } = {}) {
  const bodyOf = (xml) => /<w:body>([\s\S]*)<\/w:body>/.exec(xml);
  const ob = bodyOf(original);
  const rb = bodyOf(revised);
  if (!ob || !rb) throw new Error('a document to compare has no body');
  const A = topElements(ob[1]);
  const B = topElements(rb[1]);
  const ap = A.filter((e) => e.para);
  let id = 90000;
  let changes = 0;
  const mark = (kind) => `w:id="${id++}" w:author="${esc(author)}" w:date="${date}"`;
  const firstRPr = (p) => /<w:r\b[^>]*>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(p)?.[1] || '';
  const pPrOf = (p) => /^<w:p\b[^>]*>\s*(<w:pPr\b[^>]*\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>)?/.exec(p)?.[1] || '';
  const runText = (text, tag) => text.split(/(\t|\n)/).filter((x) => x !== '').map((x) => (x === '\t' ? '<w:tab/>' : x === '\n' ? '<w:br/>' : `<${tag} xml:space="preserve">${esc(x)}</${tag}>`)).join('');
  /** A paragraph mark marked inserted or deleted, in its pPr. */
  const withMark = (pPr, kind) => {
    const tag = `<w:${kind} ${mark(kind)}/>`;
    if (!pPr) return `<w:pPr><w:rPr>${tag}</w:rPr></w:pPr>`;
    if (/^<w:pPr\b[^>]*\/>$/.test(pPr)) return pPr.replace(/\/>$/, () => `><w:rPr>${tag}</w:rPr></w:pPr>`);
    if (/<w:rPr>/.test(pPr)) return pPr.replace('<w:rPr>', () => `<w:rPr>${tag}`);
    return pPr.replace(/<\/w:pPr>$/, () => `<w:rPr>${tag}</w:rPr></w:pPr>`);
  };
  /** A whole paragraph inserted or deleted: its runs wrapped, its mark marked. */
  const whole = (p, kind) => {
    changes += 1;
    const pPr = pPrOf(p);
    let rest = p.replace(/^<w:p\b[^>]*>/, '').replace(/<\/w:p>$/, '');
    if (pPr) rest = rest.replace(pPr, '');
    rest = rest.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (r) => `<w:${kind} ${mark(kind)}>${kind === 'del' ? r.replace(/<w:t\b/g, '<w:delText').replace(/<\/w:t>/g, '</w:delText>') : r}</w:${kind}>`);
    return `<w:p>${withMark(pPr, kind)}${rest}</w:p>`;
  };
  /** A paragraph that changed: its words matched one by one, the rest marked. */
  const changed = (was, now) => {
    const rPr = firstRPr(now.xml) || firstRPr(was.xml);
    const ta = was.text.match(/\s+|[^\s]+/g) || [];
    const tb = now.text.match(/\s+|[^\s]+/g) || [];
    const pairs = new Map(lcs(ta, tb));
    const kept = new Set(pairs.values());
    let i = 0;
    let j = 0;
    let runs = '';
    while (i < ta.length || j < tb.length) {
      if (pairs.get(i) === j) {
        // Equal words kept together in one run.
        let same = '';
        while (i < ta.length && pairs.get(i) === j) { same += ta[i]; i++; j++; }
        runs += `<w:r>${rPr}${runText(same, 'w:t')}</w:r>`;
        continue;
      }
      let del = '';
      while (i < ta.length && !pairs.has(i)) del += ta[i++];
      let ins = '';
      while (j < tb.length && !kept.has(j)) ins += tb[j++];
      if (del) { runs += `<w:del ${mark('del')}><w:r>${rPr}${runText(del, 'w:delText')}</w:r></w:del>`; changes += 1; }
      if (ins) { runs += `<w:ins ${mark('ins')}><w:r>${rPr}${runText(ins, 'w:t')}</w:r></w:ins>`; changes += 1; }
      if (!del && !ins) break;
    }
    return `<w:p>${pPrOf(now.xml)}${runs}</w:p>`;
  };

  // The paragraphs matched where their words are the same, in order.
  const bp = B.map((e, k) => ({ ...e, k })).filter((e) => e.para);
  const matched = lcs(ap.map((e) => e.text), bp.map((e) => e.text));
  const out = [];
  let ai = 0;
  let bi = 0;
  const gap = (aTo, bTo) => {
    const gone = ap.slice(ai, aTo);
    const come = bp.slice(bi, bTo);
    // Alike pairs are one paragraph changed; the rest are gone or new.
    while (gone.length && come.length) {
      if (likeness(gone[0].text, come[0].text) >= 0.3) out.push({ k: come[0].k, xml: changed(gone.shift(), come.shift()) });
      else if (come.length > gone.length) { const c = come.shift(); out.push({ k: c.k, xml: whole(c.xml, 'ins') }); }
      else out.push({ k: come[0].k - 0.5, xml: whole(gone.shift().xml, 'del') });
    }
    for (const g of gone) out.push({ k: (bp[bTo]?.k ?? B.length) - 0.5, xml: whole(g.xml, 'del') });
    for (const c of come) out.push({ k: c.k, xml: whole(c.xml, 'ins') });
  };
  for (const [x, y] of matched) {
    gap(x, y);
    out.push({ k: bp[y].k, xml: bp[y].xml });
    ai = x + 1;
    bi = y + 1;
  }
  gap(ap.length, bp.length);
  // Everything that is not a paragraph — tables, the section — where the revised document has it.
  for (const e of B.map((x, k) => ({ ...x, k })).filter((x) => !x.para)) out.push({ k: e.k, xml: e.xml });
  // In the revised document's order; a deleted paragraph just ahead of the one it stood before.
  const body = out.map((o, n) => ({ ...o, n })).sort((p, q) => p.k - q.k || p.n - q.n).map((o) => o.xml).join('');
  // The section's properties stay last, as the schema wants.
  const sect = /<w:sectPr\b[\s\S]*<\/w:sectPr>$|<w:sectPr\b[^>]*\/>$/.exec(body);
  const ordered = sect ? body.slice(0, sect.index) + body.slice(sect.index) : body;
  return { xml: revised.slice(0, rb.index) + '<w:body>' + ordered + '</w:body>' + revised.slice(rb.index + rb[0].length), changes };
}

/** Two .docx files compared: the revised document's package with its body compared against the original's. */
export function compareDocx(originalBytes, revisedBytes, opts = {}) {
  const original = OoxmlPackage.read(originalBytes);
  const revised = OoxmlPackage.read(revisedBytes);
  const { xml, changes } = compareBodies(original.text('word/document.xml'), revised.text('word/document.xml'), opts);
  revised.write_('word/document.xml', Buffer.from(xml, 'utf8'));
  return { bytes: revised.write(), changes };
}
