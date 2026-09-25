// Rutba Word: what the accessibility checker reads, the fixes it writes, and
// the text the spelling pass walks — all against a DocView (doc-view) over
// the OOXML Document (ooxml/document.js).
//
// Reading is from two places. The view's frame has what the page shows —
// styles, runs with their colours, links, the blocks the caret can reach —
// and the document's XML has what the frame leaves out: each drawing's
// `wp:docPr`, each table's first row, the parts behind headers, footers and
// notes. Writing goes through the view's `_edit`, so each fix is one undo
// step like any other edit.

import { findElements, readAltProps, writeAltProps } from './alt-text.js';
import { normaliseColour, HIGHLIGHT_HEX } from './colour.js';
import { readTitle, writeTitle, corePart } from './core-props.js';
import { paragraphsIn, paragraphText, replaceInParagraph, textBoxesIn } from './wordml.js';

const docOf = (view) => view?.doc?.doc || null;

function bodyOf(doc) {
  const open = /<w:body(\s[^>]*)?>/.exec(doc.xml);
  if (!open) return { at: 0, body: '' };
  const at = open.index + open[0].length;
  return { at, body: doc.xml.slice(at, doc.xml.lastIndexOf('</w:body>')) };
}

/** Top-level tables in the body, in order: `{ start, end, xml }` with body offsets. */
function topTables(body) {
  const out = [];
  const re = /<w:tbl\b[^>]*?(\/?)>|<\/w:tbl>/g;
  let depth = 0;
  let start = -1;
  let m;
  while ((m = re.exec(body))) {
    if (m[0] === '</w:tbl>') {
      depth -= 1;
      if (depth === 0 && start >= 0) {
        out.push({ start, end: m.index + m[0].length, xml: body.slice(start, m.index + m[0].length) });
        start = -1;
      }
    } else if (m[1] !== '/') {
      if (depth === 0) start = m.index;
      depth += 1;
    }
  }
  return out;
}

/** The kind of thing a `w:drawing` draws. */
function drawingKind(inner) {
  if (/<wps:txbx\b/.test(inner)) return 'textbox';
  if (/<a:blip\b/.test(inner) && !/<wps:wsp\b/.test(inner)) return 'picture';
  if (/<c:chart\b/.test(inner)) return 'chart';
  if (/<wpg:wgp\b/.test(inner)) return 'group';
  if (/<wps:wsp\b/.test(inner)) return 'shape';
  if (/<dgm:relIds\b/.test(inner)) return 'smartart';
  return 'object';
}

const KIND_LABEL = { picture: 'Picture', chart: 'Chart', group: 'Group', shape: 'Shape', smartart: 'SmartArt', object: 'Object' };

/**
 * Every drawing in the body with its `wp:docPr`, in document order. `ordinal`
 * counts every docPr in the main part — the address the alt-text fix writes
 * to. `block` is the edit-space paragraph that holds it; `image` its place
 * among that paragraph's pictures, which is how the page picks one.
 */
export function wordDrawings(view) {
  const doc = docOf(view);
  if (!doc) return [];
  const { at, body } = bodyOf(doc);
  const paragraphs = doc.editParagraphs();
  const all = findElements(doc.xml, 'wp:docPr');
  const out = [];
  for (const m of doc.xml.matchAll(/<w:drawing\b[^>]*>([\s\S]*?)<\/w:drawing>/g)) {
    if (m.index < at || m.index > at + body.length) continue;
    const inner = m[1];
    const docPr = findElements(inner, 'wp:docPr')[0];
    if (!docPr) continue;
    const absolute = m.index + m[0].indexOf(docPr.open);
    const ordinal = all.findIndex((e) => e.start === absolute);
    const offset = m.index - at;
    const p = paragraphs.find((q) => q.start <= offset && offset < q.end) || null;
    const kind = drawingKind(inner);
    let image = null;
    if (p && kind === 'picture') {
      const before = p.xml.slice(0, offset - p.start);
      image = [...before.matchAll(/<w:drawing\b[^>]*>([\s\S]*?)<\/w:drawing>/g)].filter((d) => /<a:blip\b/.test(d[1])).length;
    }
    out.push({ ordinal, kind, block: p ? p.index : null, image, ...readAltProps(docPr.xml) });
  }
  return out;
}

/** The document's language: the defaults in the stylesheet, else the runs' own, most common first. */
export function wordLanguage(view) {
  const doc = docOf(view);
  if (!doc) return null;
  try {
    if (doc.pkg.has('word/styles.xml')) {
      const defaults = /<w:docDefaults\b[\s\S]*?<\/w:docDefaults>/.exec(doc.pkg.text('word/styles.xml'))?.[0] || '';
      const lang = /<w:lang\b[^>]*\bw:val="([^"]+)"/.exec(defaults);
      if (lang) return lang[1];
    }
  } catch {
    /* no stylesheet to read */
  }
  const counts = new Map();
  for (const m of doc.xml.matchAll(/<w:lang\b[^>]*\bw:val="([^"]+)"/g)) counts.set(m[1], (counts.get(m[1]) || 0) + 1);
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || null;
}

const shortText = (s, n = 48) => {
  const t = String(s || '').replace(/\s+/g, ' ').trim();
  return t.length > n ? `${t.slice(0, n - 1)}…` : t;
};

/**
 * The document, described for the rules in accessibility.js. `frame` is the
 * view's own render (pages off) when the caller already has one.
 */
export function describeWord(view, { frame = null } = {}) {
  const doc = docOf(view);
  const f = frame || view.render({ pages: false });
  const blocks = f.blocks || [];
  const styles = view.docStyles || {};
  const model = { app: 'word', objects: [], headings: [], tables: [], texts: [], links: [], blankRuns: [] };

  // Pictures, charts, shapes and groups — a text box is read out as its words.
  const counters = {};
  for (const d of wordDrawings(view)) {
    if (d.kind === 'textbox') continue;
    counters[d.kind] = (counters[d.kind] || 0) + 1;
    model.objects.push({
      key: `d${d.ordinal}`,
      kind: d.kind,
      label: d.name || `${KIND_LABEL[d.kind]} ${counters[d.kind]}`,
      alt: d.descr,
      decorative: d.decorative,
      where: { block: d.block, image: d.image },
      target: { drawing: d.ordinal },
    });
  }

  // Headings, in order.
  for (const b of blocks) {
    const level = Number((/^Heading([1-9])$/.exec(b.style || '') || [])[1] || 0);
    if (!level) continue;
    model.headings.push({ key: `h${b.index}`, level, text: b.text, label: shortText(b.text) || `Heading ${level}`, styleName: `Heading ${level}`, where: { block: b.index }, target: { block: b.index } });
  }

  // Tables: the first row repeated as a header, and no merged cells.
  if (doc) {
    const { body } = bodyOf(doc);
    const paragraphs = doc.editParagraphs();
    topTables(body).forEach((t, i) => {
      const firstRow = /<w:tr\b[\s\S]*?<\/w:tr>/.exec(t.xml)?.[0] || '';
      const head = firstRow.split(/<w:tc\b/)[0];
      const hasHeader = /<w:tblHeader\b(?![^>]*w:val="(?:0|false)")/.test(head);
      const merged = /<w:gridSpan\b[^>]*w:val="(?:[2-9]|\d{2,})"/.test(t.xml) || /<w:vMerge\b/.test(t.xml) || /<w:hMerge\b/.test(t.xml);
      const first = paragraphs.find((p) => p.container && p.container.startsWith(`t${t.start}:`));
      model.tables.push({
        key: `t${i}`, label: `Table ${i + 1}`, hasHeader, merged,
        where: { block: first ? first.index : null, table: i },
        target: { table: i },
        headerFix: 'Repeat the first row as a header',
      });
    });
  }

  // Text contrast: a colour chosen by hand against what is behind it. Word
  // draws "Automatic" text black or white by what it sits on, so only a
  // stated colour — the run's, its character style's or the paragraph
  // style's — can fall short.
  const page = normaliseColour(safely(() => doc?.pageColour?.())) || '#FFFFFF';
  for (const b of blocks) {
    const style = styles[b.style || ''] || styles['*default*'] || {};
    for (const r of b.runs || []) {
      if (!String(r.text || '').trim() || r.del) continue;
      const fg = normaliseColour(r.fontColour) || normaliseColour(style.colour);
      if (!fg) continue;
      const bg = normaliseColour(r.highlight ? HIGHLIGHT_HEX[r.highlight] || r.highlight : null) || normaliseColour(b.shading) || page;
      const sizePt = Number(r.fontSize) || (style.sizePx ? (style.sizePx * 72) / 96 : 11);
      model.texts.push({ key: `b${b.index}`, label: shortText(b.text) || `Paragraph ${b.index + 1}`, where: { block: b.index }, fg, bg, sizePt, bold: Boolean(r.bold ?? style.bold), target: { block: b.index, length: String(b.text || '').length } });
    }
  }

  // Links: consecutive runs to the same address are one link.
  for (const b of blocks) {
    let current = null;
    const flush = () => {
      if (current) model.links.push({ key: `l${b.index}:${model.links.length}`, text: current.text, url: current.url, label: `"${shortText(current.text, 40)}"`, where: { block: b.index } });
      current = null;
    };
    for (const r of b.runs || []) {
      if (r.link && current && current.url === r.link) current.text += r.text || '';
      else {
        flush();
        if (r.link) current = { url: r.link, text: r.text || '' };
      }
    }
    flush();
  }

  // Empty paragraphs in a row, used as spacing. A paragraph that carries a
  // section break, a page break, a picture or a field is not empty.
  if (doc) {
    const raw = doc.editParagraphs();
    let run = [];
    const flush = () => {
      if (run.length >= 2) {
        model.blankRuns.push({ key: `e${run[0]}`, label: `${run.length} empty paragraphs after paragraph ${run[0]}`, where: { block: run[0] }, target: { blocks: run.slice(1) } });
      }
      run = [];
    };
    for (const b of blocks) {
      const p = raw[b.index];
      const empty = !b.container && !String(b.text || '').trim() && !b.images?.length && !b.textBoxes?.length && !b.structural
        && p && !/<w:(sectPr|br|drawing|pict|object|fldChar|fldSimple|instrText|footnoteReference|endnoteReference|bookmarkStart)\b|<m:oMath/.test(p.xml);
      if (empty) run.push(b.index);
      else flush();
    }
    flush();
  }

  // The title in the document's properties.
  if (doc) {
    model.title = readTitle(doc.pkg);
    const first = blocks.find((b) => (b.style === 'Title' || /^Heading1$/.test(b.style || '')) && String(b.text || '').trim());
    model.titleSuggestion = shortText(first?.text || '', 120);
  }
  return model;
}

function safely(fn) {
  try {
    return fn();
  } catch {
    return null;
  }
}

/* ── the fixes ─────────────────────────────────────────────────────────── */

/** Write the document's XML through the view as one undo step. */
function editXml(view, label, change) {
  const doc = docOf(view);
  if (!doc) throw new Error('this document has no XML to change');
  return view._edit(label, null, () => {
    const next = change(doc.xml);
    if (next !== doc.xml) {
      doc.xml = next;
      doc.dirty = true;
    }
    view._invalidate();
    return true;
  });
}

/** Alt text on the `ordinal`-th `wp:docPr` of the main part — a description, or decorative. */
export function setWordAltText(view, { drawing, descr = '', decorative = false }) {
  return editXml(view, 'alt text', (xml) => {
    const hit = findElements(xml, 'wp:docPr')[Number(drawing)];
    if (!hit) throw new Error('that picture is no longer in the document');
    // The picture's own cNvPr (pic:cNvPr) carries a description too; Word
    // writes the same words to both, and a reader that looks at either finds them.
    let next = xml.slice(0, hit.start) + writeAltProps(hit.xml, { descr, decorative }) + xml.slice(hit.end);
    const close = next.indexOf('</w:drawing>', hit.start);
    if (close > 0) {
      const inner = next.slice(hit.start, close);
      const pic = findElements(inner, 'pic:cNvPr')[0];
      if (pic) {
        const rewritten = writeAltProps(pic.xml, { descr: decorative ? '' : descr, decorative: false });
        next = next.slice(0, hit.start + pic.start) + rewritten + next.slice(hit.start + pic.end);
      }
    }
    return next;
  });
}

/** Repeat the first row of the `table`-th top-level table as a header row (`w:tblHeader`). */
export function setWordTableHeader(view, { table }) {
  return editXml(view, 'table header', (xml) => {
    const open = /<w:body(\s[^>]*)?>/.exec(xml);
    const at = open.index + open[0].length;
    const t = topTables(xml.slice(at))[Number(table)];
    if (!t) throw new Error('that table is no longer in the document');
    const rowAt = t.xml.search(/<w:tr\b/);
    if (rowAt < 0) return xml;
    const rowOpen = /^<w:tr\b[^>]*>/.exec(t.xml.slice(rowAt))[0];
    let i = rowAt + rowOpen.length;
    let rest = t.xml.slice(i);
    const prEx = /^\s*<w:tblPrEx\b[^>]*(?:\/>|>[\s\S]*?<\/w:tblPrEx>)/.exec(rest);
    if (prEx) { i += prEx[0].length; rest = t.xml.slice(i); }
    let row;
    const trPr = /^\s*<w:trPr\b[^>]*(?:\/>|>([\s\S]*?)<\/w:trPr>)/.exec(rest);
    if (trPr) {
      if (/<w:tblHeader\b/.test(trPr[0])) row = t.xml.slice(0, i) + rest.replace(/<w:tblHeader\b[^>]*\/>/, '<w:tblHeader/>');
      else if (trPr[0].endsWith('/>')) row = t.xml.slice(0, i) + rest.replace(/^(\s*)<w:trPr\b([^>]*)\/>/, '$1<w:trPr$2><w:tblHeader/></w:trPr>');
      else row = t.xml.slice(0, i) + rest.replace('</w:trPr>', '<w:tblHeader/></w:trPr>');
    } else {
      row = t.xml.slice(0, i) + '<w:trPr><w:tblHeader/></w:trPr>' + rest;
    }
    return xml.slice(0, at + t.start) + row + xml.slice(at + t.end);
  });
}

/** Take out empty paragraphs (edit-space indices), last first so the others keep their numbers. */
export function removeWordParagraphs(view, { blocks }) {
  const doc = docOf(view);
  const list = [...new Set((blocks || []).map(Number))].sort((a, b) => b - a);
  return view._edit('remove empty paragraphs', null, () => {
    for (const i of list) {
      const p = doc.editParagraphs()[i];
      if (!p || String(p.text || '').trim()) continue;
      doc.removeEditParagraph(i);
    }
    doc.dirty = true;
    view._invalidate();
    const focus = Math.min(view.focus.block, Math.max(0, view.blocks.length - 1));
    view.anchor = { block: focus, offset: 0 };
    view.focus = { block: focus, offset: 0 };
    return true;
  });
}

/** The title in the document's properties. Undo puts back the part as it was, when there was one. */
export function setWordTitle(view, { title }) {
  const doc = docOf(view);
  const part = corePart(doc.pkg);
  if (part) doc._undoParts?.add(part);
  return view._edit('document title', null, () => {
    writeTitle(doc.pkg, title);
    doc.dirty = true;
    view.touched = true;
    return true;
  });
}

/* ── the words, for the spelling pass ──────────────────────────────────── */

/**
 * Every piece of text in the document in the order the pass walks it: the
 * body's paragraphs (a table's cells among them) with each text box after
 * the paragraph that anchors it, then the footnotes, the endnotes, and the
 * headers and footers. `key` addresses a segment; `where` is what the window
 * needs to show it.
 */
export function wordSegments(view) {
  const doc = docOf(view);
  const out = [];
  const raw = doc ? doc.editParagraphs() : [];
  for (const b of view.blocks) {
    out.push({ key: `b:${b.index}`, text: b.text || '', where: { story: 'body', block: b.index } });
    const p = raw[b.index];
    if (p && p.xml.includes('<w:txbxContent')) {
      textBoxesIn(p.xml).forEach((box, bi) => box.paragraphs.forEach((q, qi) => {
        out.push({ key: `x:${b.index}:${bi}:${qi}`, text: paragraphText(q.xml).text, where: { story: 'textbox', block: b.index } });
      }));
    }
  }
  if (doc) {
    for (const [part, tag, story] of [['word/footnotes.xml', 'footnote', 'footnote'], ['word/endnotes.xml', 'endnote', 'endnote']]) {
      if (!doc.pkg.has(part)) continue;
      const xml = doc.pkg.text(part);
      for (const m of xml.matchAll(new RegExp(`<w:${tag}\\b([^>]*)>([\\s\\S]*?)<\\/w:${tag}>`, 'g'))) {
        const type = /\bw:type="([^"]+)"/.exec(m[1])?.[1];
        if (type && type !== 'normal') continue;
        const id = /\bw:id="([^"]+)"/.exec(m[1])?.[1];
        paragraphsIn(m[2]).forEach((q, qi) => out.push({ key: `n:${part}:${id}:${qi}`, text: paragraphText(q.xml).text, where: { story, id } }));
      }
    }
    const bands = safely(() => doc.headerFooters()) || {};
    const seen = new Set();
    for (const [group, story] of [['headers', 'header'], ['footers', 'footer']]) {
      for (const band of Object.values(bands[group] || {})) {
        if (!band?.part || seen.has(band.part) || !doc.pkg.has(band.part)) continue;
        seen.add(band.part);
        paragraphsIn(doc.pkg.text(band.part)).forEach((q, qi) => out.push({ key: `h:${band.part}:${qi}`, text: paragraphText(q.xml).text, where: { story } }));
      }
    }
  }
  return out;
}

/** Where a pass starts: the caret's paragraph and offset. */
export function wordStart(view) {
  return { key: `b:${view.focus?.block ?? 0}`, offset: view.focus?.offset ?? 0 };
}

/**
 * Replace `[from, to)` of a segment with `text`. The body goes through the
 * view's own selection and typing, so a tracked-changes document records
 * the change; the other stories are rewritten in their XML. One undo step
 * for however many `edits` (each `{ key, from, to, text }`), which is what
 * Change All needs. Edits in one segment are applied right to left.
 */
export function replaceWordText(view, edits) {
  const doc = docOf(view);
  const list = [...edits].sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : b.from - a.from));
  // A part outside the main one goes on the undo list BEFORE the step's
  // snapshot is taken, or undo would not put it back.
  for (const e of list) {
    const [kind, ...rest] = e.key.split(':');
    if (kind === 'n') doc._undoParts?.add(rest[0]);
    if (kind === 'h') doc._undoParts?.add(rest.slice(0, -1).join(':'));
  }
  return view._edit('spelling', null, () => {
    for (const e of list) {
      const [kind, ...rest] = e.key.split(':');
      if (kind === 'b') {
        const block = Number(rest[0]);
        const b = view.block(block);
        if (b && !b.structural) {
          view.setSelection({ block, offset: e.from }, { block, offset: e.to });
          view.insertText(e.text);
          continue;
        }
        // A paragraph the caret cannot type into — it anchors a text box,
        // say — has its words rewritten in its XML, when its text there is
        // the text the pass read.
        const p = doc.editParagraphs()[block];
        if (!p || !b || paragraphText(p.xml).text !== b.text) continue;
        const { at } = bodyOf(doc);
        doc.xml = doc.xml.slice(0, at + p.start) + replaceInParagraph(p.xml, e.from, e.to, e.text) + doc.xml.slice(at + p.end);
        doc.dirty = true;
        continue;
      }
      if (kind === 'x') {
        const [block, bi, qi] = rest.map(Number);
        const p = doc.editParagraphs()[block];
        if (!p) continue;
        const box = textBoxesIn(p.xml)[bi];
        const q = box?.paragraphs[qi];
        if (!q) continue;
        const nextP = p.xml.slice(0, box.start + q.start) + replaceInParagraph(q.xml, e.from, e.to, e.text) + p.xml.slice(box.start + q.end);
        const { at } = bodyOf(doc);
        doc.xml = doc.xml.slice(0, at + p.start) + nextP + doc.xml.slice(at + p.end);
        doc.dirty = true;
        continue;
      }
      // A footnote, an endnote, a header or a footer: its own part.
      const part = kind === 'n' ? rest[0] : rest.slice(0, -1).join(':');
      if (!doc.pkg.has(part)) continue;
      let xml = doc.pkg.text(part);
      let scope = { start: 0, xml };
      if (kind === 'n') {
        const [, id] = rest;
        const tag = part.includes('endnotes') ? 'endnote' : 'footnote';
        const m = [...xml.matchAll(new RegExp(`<w:${tag}\\b([^>]*)>([\\s\\S]*?)<\\/w:${tag}>`, 'g'))].find((x) => /\bw:id="([^"]+)"/.exec(x[1])?.[1] === id);
        if (!m) continue;
        scope = { start: m.index + m[0].indexOf('>') + 1, xml: m[2] };
      }
      const qi = Number(rest[rest.length - 1]);
      const q = paragraphsIn(scope.xml)[qi];
      if (!q) continue;
      xml = xml.slice(0, scope.start + q.start) + replaceInParagraph(q.xml, e.from, e.to, e.text) + xml.slice(scope.start + q.end);
      doc.pkg.write_(part, xml);
      doc.dirty = true;
    }
    view._invalidate();
    return true;
  });
}
