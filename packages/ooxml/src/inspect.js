/**
 * The Document Inspector, File → Info → Check for Issues → Inspect, for a
 * document, a workbook and a presentation: what a file carries that its
 * author may not want to hand on, found part by part, and removed on
 * request, as Word's, Excel's and PowerPoint's inspectors remove it.
 *
 * Each kind is `{ id, found, items }`: how many were found and a few of
 * them by name. `removeFromPackage` takes away the kinds named; a document's
 * tracked changes are accepted (the caller runs the document's own Accept
 * All first, and what is left is accepted here), its comments, hidden words,
 * invisible objects, headers and footers taken out, its personal details
 * blanked and Word told to keep them out (w:removePersonalInformation).
 */
import { OoxmlPackage, esc } from './package.js';
import { unesc } from './workbook.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The parts of a package a test names. */
const partsWhere = (pkg, test) => pkg.partNames().filter((n) => test(n));

/** Take a part away with its relationships from every other part. */
function dropPart(pkg, name) {
  if (!pkg.has(name)) return;
  const target = name.split('/').pop();
  for (const rels of pkg.partNames().filter((n) => n.endsWith('.rels'))) {
    const xml = pkg.text(rels);
    const owner = rels === '_rels/.rels' ? '' : rels.replace(/_rels\/([^/]+)\.rels$/, '$1');
    const next = xml.replace(/<Relationship\b[^>]*\/>/g, (tag) => {
      const t = /Target="([^"]*)"/.exec(tag)?.[1];
      if (!t || /TargetMode="External"/.test(tag)) return tag;
      try {
        return OoxmlPackage.resolveTarget(owner, unesc(t)) === name ? '' : tag;
      } catch {
        return t.endsWith(target) ? '' : tag;
      }
    });
    if (next !== xml) pkg.write_(rels, next);
  }
  pkg.removePart(name);
}

/** The relationship ids a part has to parts matching a test. */
function relIdsTo(pkg, from, test) {
  return pkg.rels(from).filter((r) => r.TargetMode !== 'External' && test(OoxmlPackage.resolveTarget(from, r.Target), r)).map((r) => r.Id);
}

const xmlParts = (pkg) => pkg.partNames().filter((n) => /\.xml$/i.test(n));

/* ── properties and people, the same in every kind ───────────────────────── */

const CORE_PERSONAL = ['dc:creator', 'cp:lastModifiedBy'];
const CORE_PROPERTIES = ['dc:title', 'dc:subject', 'cp:keywords', 'dc:description', 'cp:category', 'cp:contentStatus', 'cp:version', 'cp:revision'];
const APP_PERSONAL = ['Company', 'Manager'];

function elementWords(xml, tag) {
  const m = new RegExp('<' + tag + '\\b[^>]*>([\\s\\S]*?)</' + tag + '>').exec(xml);
  return m ? unesc(m[1]).trim() : '';
}

function blankElements(xml, tags) {
  let out = xml;
  for (const tag of tags) out = out.replace(new RegExp('<' + tag + '\\b[^>]*>[\\s\\S]*?</' + tag + '>|<' + tag + '\\b[^>]*/>', 'g'), '');
  return out;
}

/** The authors' names a package's comments and changes carry, every kind of comment and person part. */
function authorNames(pkg) {
  const names = new Set();
  for (const name of xmlParts(pkg)) {
    const xml = pkg.text(name);
    for (const m of xml.matchAll(/\s(?:w:author|w15:author|author)="([^"]*)"/g)) if (m[1]) names.add(unesc(m[1]));
    if (/comment|person|people|authors/i.test(name)) {
      for (const m of xml.matchAll(/<(?:p:cmAuthor|p188:author|person)\b[^>]*\bname="([^"]*)"|<(?:person)\b[^>]*displayName="([^"]*)"/g)) names.add(unesc(m[1] ?? m[2]));
      for (const m of xml.matchAll(/<author>([^<]*)<\/author>/g)) names.add(unesc(m[1]));
    }
  }
  names.delete('');
  return [...names];
}

function inspectPersonal(pkg) {
  const items = [];
  if (pkg.has('docProps/core.xml')) {
    const core = pkg.text('docProps/core.xml');
    for (const tag of CORE_PERSONAL) { const v = elementWords(core, tag); if (v) items.push(v); }
  }
  if (pkg.has('docProps/app.xml')) {
    const app = pkg.text('docProps/app.xml');
    for (const tag of APP_PERSONAL) { const v = elementWords(app, tag); if (v) items.push(v); }
  }
  for (const n of authorNames(pkg)) items.push(n);
  const unique = [...new Set(items)];
  return { id: 'personal', found: unique.length, items: unique.slice(0, 6) };
}

function inspectProperties(pkg) {
  const items = [];
  if (pkg.has('docProps/core.xml')) {
    const core = pkg.text('docProps/core.xml');
    for (const tag of CORE_PROPERTIES) { const v = elementWords(core, tag); if (v) items.push(v); }
  }
  if (pkg.has('docProps/custom.xml')) for (const m of pkg.text('docProps/custom.xml').matchAll(/<property\b[^>]*\bname="([^"]*)"/g)) items.push(unesc(m[1]));
  return { id: 'properties', found: items.length, items: items.slice(0, 6) };
}

function inspectCustomXml(pkg) {
  const items = partsWhere(pkg, (n) => /^customXml\/item\d+\.xml$/i.test(n));
  return { id: 'customXml', found: items.length, items };
}

/** Blank the authors in every part, as the inspectors do: "Author", initials "A". */
function anonymise(pkg) {
  for (const name of xmlParts(pkg)) {
    const xml = pkg.text(name);
    let next = xml.replace(/(\s(?:w:author|w15:author)=")[^"]*"/g, '$1Author"').replace(/(\sw:initials=")[^"]*"/g, '$1A"');
    if (/comment|authors|person/i.test(name)) {
      next = next.replace(/(<(?:p:cmAuthor|p188:author)\b[^>]*\bname=")[^"]*"/g, '$1Author"').replace(/(<(?:p:cmAuthor|p188:author)\b[^>]*\binitials=")[^"]*"/g, '$1A"')
        .replace(/(<(?:p:cmAuthor|p188:author)\b[^>]*\b(?:userId|providerId)=")[^"]*"/g, '$1"')
        .replace(/(<person\b[^>]*\bdisplayName=")[^"]*"/g, '$1Author"').replace(/(<person\b[^>]*\b(?:userId|providerId)=")[^"]*"/g, '$1"')
        .replace(/<author>[^<]*<\/author>/g, '<author>Author</author>');
    }
    if (next !== xml) pkg.write_(name, next);
  }
}

function removePersonal(pkg, kind) {
  if (pkg.has('docProps/core.xml')) pkg.write_('docProps/core.xml', blankElements(pkg.text('docProps/core.xml'), CORE_PERSONAL));
  if (pkg.has('docProps/app.xml')) pkg.write_('docProps/app.xml', blankElements(pkg.text('docProps/app.xml'), APP_PERSONAL));
  anonymise(pkg);
  for (const n of partsWhere(pkg, (x) => /^word\/people\.xml$/.test(x))) dropPart(pkg, n);
  // Word keeps them out from now on: Remove personal information from file properties on save.
  if (kind === 'document' && pkg.has('word/settings.xml')) {
    const xml = pkg.text('word/settings.xml');
    if (!/<w:removePersonalInformation\b/.test(xml)) {
      const flags = '<w:removePersonalInformation/><w:removeDateAndTime/>';
      const after = /<w:zoom\b[^>]*\/>|<w:view\b[^>]*\/>|<w:writeProtection\b[^>]*\/>/.exec(xml);
      const next = after ? xml.slice(0, after.index + after[0].length) + flags + xml.slice(after.index + after[0].length) : xml.replace(/<w:settings\b[^>]*>/, (open) => open + flags);
      pkg.write_('word/settings.xml', next);
    }
  }
}

function removeProperties(pkg) {
  if (pkg.has('docProps/core.xml')) pkg.write_('docProps/core.xml', blankElements(pkg.text('docProps/core.xml'), CORE_PROPERTIES));
  if (pkg.has('docProps/custom.xml')) dropPart(pkg, 'docProps/custom.xml');
}

function removeCustomXml(pkg) {
  for (const n of partsWhere(pkg, (x) => /^customXml\//.test(x) && !x.includes('_rels/'))) dropPart(pkg, n);
}

/* ── a document ──────────────────────────────────────────────────────────── */

/** The parts a document's words live in: the body, headers, footers, notes, comments. */
const docStories = (pkg) => partsWhere(pkg, (n) => /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/.test(n));

const REVISION_RE = /<w:(?:ins|del|moveFrom|moveTo)\b(?![^>]*\/>)[^>]*>|<w:(?:rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange|tblGridChange|numberingChange)\b|<w:(?:ins|del)\b[^>]*\/>|<w:cell(?:Ins|Del|Merge)\b/g;
const HIDDEN_RUN_RE = /<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:rPr\b[^>]*>(?:(?!<\/w:rPr>)[\s\S])*?<w:vanish\b(?![^>]*w:val="(?:0|false|off)")[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g;
const INVISIBLE_DRAWING_RE = /<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<wp:docPr\b[^>]*\bhidden="(?:1|true)"[\s\S]*?<\/w:r>/g;

function inspectDocument(pkg) {
  const stories = docStories(pkg);
  const text = (n) => pkg.text(n);
  const commentsPart = partsWhere(pkg, (n) => /^word\/comments\.xml$/.test(n))[0];
  const comments = commentsPart ? [...text(commentsPart).matchAll(/<w:comment\b[^>]*>([\s\S]*?)<\/w:comment>/g)].map((m) => (unesc(m[1].replace(/<[^>]*>/g, '')) || '').trim().slice(0, 60)) : [];
  let revisions = 0;
  let hidden = [];
  let invisible = 0;
  for (const n of stories.filter((s) => !/comments/.test(s))) {
    const xml = text(n);
    revisions += (xml.match(REVISION_RE) || []).length;
    hidden = hidden.concat([...xml.matchAll(HIDDEN_RUN_RE)].map((m) => unesc((m[0].match(/<w:t\b[^>]*>([^<]*)<\/w:t>/g) || []).map((x) => x.replace(/<[^>]*>/g, '')).join(''))).filter(Boolean));
    invisible += (xml.match(INVISIBLE_DRAWING_RE) || []).length;
  }
  const headers = partsWhere(pkg, (n) => /^word\/(?:header|footer)\d*\.xml$/.test(n));
  const watermark = headers.some((n) => /PowerPlusWaterMarkObject|<v:textpath\b/.test(text(n)));
  const embedded = partsWhere(pkg, (n) => /^word\/embeddings\//.test(n));
  const ink = stories.reduce((k, n) => k + (text(n).match(/<w14:contentPart\b/g) || []).length, 0);
  return [
    { id: 'comments', found: comments.length, items: comments.slice(0, 6) },
    { id: 'revisions', found: revisions, items: [] },
    inspectPersonal(pkg),
    inspectProperties(pkg),
    inspectCustomXml(pkg),
    { id: 'headers', found: headers.length, items: headers.map((h) => h.split('/').pop()), ...(watermark ? { watermark: true } : {}) },
    { id: 'hidden', found: hidden.length, items: hidden.slice(0, 6) },
    { id: 'invisible', found: invisible, items: [] },
    { id: 'ink', found: ink, items: [] },
    { id: 'embedded', found: embedded.length, items: embedded.map((e) => e.split('/').pop()), keep: true },
  ];
}

/** Every tracked change in a stretch of XML accepted: insertions kept, deletions gone, records of before dropped. */
export function acceptAllXml(xml) {
  let out = String(xml);
  // Deleted and moved-away words go, whole.
  out = out.replace(/<w:(del|moveFrom)\b(?![^>]*\/>)[^>]*>[\s\S]*?<\/w:\1>/g, '');
  // Inserted and moved-to words stay, their wrapper taken off.
  out = out.replace(/<w:(ins|moveTo)\b(?![^>]*\/>)[^>]*>([\s\S]*?)<\/w:\1>/g, '$2');
  // A paragraph mark put in, a range marker of a move, a record of the formatting before.
  out = out.replace(/<w:(?:ins|del)\b[^>]*\/>/g, '').replace(/<w:move(?:From|To)Range(?:Start|End)\b[^>]*\/>/g, '');
  out = out.replace(/<w:(rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange|tblGridChange|numberingChange)\b[^>]*(?:\/>|>[\s\S]*?<\/w:\1>)/g, '');
  out = out.replace(/<w:cell(?:Ins|Del|Merge)\b[^>]*\/>/g, '');
  return out;
}

function removeFromDocument(pkg, ids) {
  const stories = () => docStories(pkg);
  const edit = (n, fn) => { const xml = pkg.text(n); const next = fn(xml); if (next !== xml) pkg.write_(n, next); };
  if (ids.includes('comments')) {
    for (const n of stories()) {
      if (/comments\.xml$/.test(n)) continue;
      edit(n, (xml) => xml.replace(/<w:comment(?:RangeStart|RangeEnd)\b[^>]*\/>/g, '').replace(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:commentReference\b[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g, ''));
    }
    for (const n of partsWhere(pkg, (x) => /^word\/comments(?:Extended|Ids|Extensible)?\.xml$/.test(x))) dropPart(pkg, n);
  }
  if (ids.includes('revisions')) for (const n of stories()) edit(n, acceptAllXml);
  if (ids.includes('hidden')) for (const n of stories()) edit(n, (xml) => xml.replace(HIDDEN_RUN_RE, ''));
  if (ids.includes('invisible')) for (const n of stories()) edit(n, (xml) => xml.replace(INVISIBLE_DRAWING_RE, ''));
  if (ids.includes('ink')) for (const n of stories()) edit(n, (xml) => xml.replace(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<mc:AlternateContent\b(?:(?!<\/mc:AlternateContent>)[\s\S])*?<w14:contentPart\b[\s\S]*?<\/mc:AlternateContent>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g, ''));
  if (ids.includes('headers')) {
    edit('word/document.xml', (xml) => xml.replace(/<w:(?:header|footer)Reference\b[^>]*\/>/g, '').replace(/<w:titlePg\b[^>]*\/>/g, ''));
    for (const n of partsWhere(pkg, (x) => /^word\/(?:header|footer)\d*\.xml$/.test(x))) dropPart(pkg, n);
    if (pkg.has('word/settings.xml')) edit('word/settings.xml', (xml) => xml.replace(/<w:evenAndOddHeaders\b[^>]*\/>/g, ''));
  }
  if (ids.includes('personal')) removePersonal(pkg, 'document');
  if (ids.includes('properties')) removeProperties(pkg);
  if (ids.includes('customXml')) removeCustomXml(pkg);
}

/* ── a presentation ──────────────────────────────────────────────────────── */

const slideParts = (pkg) => partsWhere(pkg, (n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
const SHAPE_RE = /<p:(sp|pic|graphicFrame|grpSp|cxnSp|contentPart)\b[^>]*>[\s\S]*?<\/p:\1>/g;

/** The slide's size, in EMU, from the presentation. */
function slideSize(pkg) {
  const pres = pkg.has('ppt/presentation.xml') ? pkg.text('ppt/presentation.xml') : '';
  const m = /<p:sldSz\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(pres);
  return m ? { cx: Number(m[1]), cy: Number(m[2]) } : { cx: 12192000, cy: 6858000 };
}

/** Top-level shapes of a slide's tree, with where they stand. */
function topShapes(xml) {
  const tree = /<p:spTree\b[^>]*>([\s\S]*)<\/p:spTree>/.exec(xml);
  if (!tree) return [];
  const inner = tree[1];
  const base = tree.index + tree[0].indexOf(inner);
  const out = [];
  const re = /<(p:(?:sp|pic|graphicFrame|grpSp|cxnSp|contentPart)|mc:AlternateContent)\b[^>]*>/g;
  let m;
  while ((m = re.exec(inner))) {
    const name = m[1];
    let depth = 0;
    const tags = new RegExp('<' + name + '\\b[^>]*?(/?)>|</' + name + '>', 'g');
    tags.lastIndex = m.index;
    let t;
    let end = -1;
    while ((t = tags.exec(inner))) {
      if (t[0].startsWith('</')) { depth -= 1; if (depth === 0) { end = t.index + t[0].length; break; } } else if (!t[1]) depth += 1;
    }
    if (end < 0) break;
    const shape = inner.slice(m.index, end);
    const off = /<a:off\b[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/.exec(shape);
    const ext = /<a:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(shape);
    out.push({
      start: base + m.index, end: base + end, xml: shape,
      name: unesc(/<p:cNvPr\b[^>]*\bname="([^"]*)"/.exec(shape)?.[1] ?? ''),
      hidden: /<p:cNvPr\b[^>]*\bhidden="(?:1|true)"/.test(shape),
      ink: name === 'p:contentPart' || /<p14:contentPart\b|<p:contentPart\b/.test(shape),
      box: off && ext ? { x: Number(off[1]), y: Number(off[2]), w: Number(ext[1]), h: Number(ext[2]) } : null,
    });
    re.lastIndex = end;
  }
  return out;
}

function inspectPresentation(pkg) {
  const size = slideSize(pkg);
  const commentParts = partsWhere(pkg, (n) => /^ppt\/comments\/(?:comment|modernComment)[^/]*\.xml$/.test(n));
  // A legacy comment's p:text, or a PowerPoint 365 comment (p188:cm) or reply (p188:reply) with its words in a txBody.
  const comments = commentParts.flatMap((n) => [...pkg.text(n).matchAll(/<p:cm\b[\s\S]*?<p:text>([\s\S]*?)<\/p:text>|<p188:(?:cm|reply)\b[^>]*>[\s\S]*?<p188:txBody>([\s\S]*?)<\/p188:txBody>/g)].map((m) => unesc((m[1] ?? m[2] ?? '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ')).trim().slice(0, 60)));
  const notes = partsWhere(pkg, (n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n)).filter((n) => notesWords(pkg.text(n)).trim());
  const invisible = [];
  const offSlide = [];
  let ink = 0;
  for (const s of slideParts(pkg)) {
    for (const sh of topShapes(pkg.text(s))) {
      if (sh.ink) ink += 1;
      if (sh.hidden) invisible.push(sh.name);
      else if (sh.box && (sh.box.x >= size.cx || sh.box.y >= size.cy || sh.box.x + sh.box.w <= 0 || sh.box.y + sh.box.h <= 0)) offSlide.push(sh.name);
    }
  }
  const embedded = partsWhere(pkg, (n) => /^ppt\/embeddings\//.test(n));
  return [
    { id: 'comments', found: comments.length, items: comments.slice(0, 6) },
    { id: 'notes', found: notes.length, items: [] },
    inspectPersonal(pkg),
    inspectProperties(pkg),
    inspectCustomXml(pkg),
    { id: 'invisible', found: invisible.length, items: invisible.slice(0, 6) },
    { id: 'offSlide', found: offSlide.length, items: offSlide.slice(0, 6) },
    { id: 'ink', found: ink, items: [] },
    { id: 'embedded', found: embedded.length, items: embedded.map((e) => e.split('/').pop()), keep: true },
  ];
}

/** A notes slide's own words: its body placeholder, not the slide image or the number. */
function notesWords(xml) {
  return topShapes(xml).filter((s) => /<p:ph\b[^>]*type="body"/.test(s.xml)).map((s) => (s.xml.match(/<a:t>([^<]*)<\/a:t>/g) || []).map((t) => unesc(t.replace(/<[^>]*>/g, ''))).join(' ')).join(' ');
}

function removeFromPresentation(pkg, ids) {
  const size = slideSize(pkg);
  if (ids.includes('comments')) {
    for (const n of partsWhere(pkg, (x) => /^ppt\/comments\//.test(x) && !x.includes('_rels/'))) dropPart(pkg, n);
    for (const n of partsWhere(pkg, (x) => /^ppt\/(?:commentAuthors|authors)\.xml$/.test(x))) dropPart(pkg, n);
    // A slide's own record of its modern comments goes with them.
    for (const s of slideParts(pkg)) {
      const xml = pkg.text(s);
      const next = xml.replace(/<p:ext\b[^>]*uri="\{6950BFC3-D8DA-4A85-94F7-54DA5524770B\}"[^>]*>[\s\S]*?<\/p:ext>/g, '');
      if (next !== xml) pkg.write_(s, next);
    }
  }
  if (ids.includes('notes')) {
    for (const n of partsWhere(pkg, (x) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(x))) {
      const xml = pkg.text(n);
      let next = xml;
      for (const sh of topShapes(xml).filter((s) => /<p:ph\b[^>]*type="body"/.test(s.xml)).reverse()) {
        const emptied = sh.xml.replace(/<p:txBody>([\s\S]*?)<\/p:txBody>/, (m, inner) => '<p:txBody>' + (/<a:bodyPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:bodyPr>)/.exec(inner)?.[0] ?? '<a:bodyPr/>') + (/<a:lstStyle\b[^>]*(?:\/>|>[\s\S]*?<\/a:lstStyle>)/.exec(inner)?.[0] ?? '<a:lstStyle/>') + '<a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody>');
        next = next.slice(0, sh.start) + emptied + next.slice(sh.end);
      }
      if (next !== xml) pkg.write_(n, next);
    }
  }
  if (ids.includes('invisible') || ids.includes('offSlide') || ids.includes('ink')) {
    for (const s of slideParts(pkg)) {
      const xml = pkg.text(s);
      let next = xml;
      const gone = topShapes(xml).filter((sh) => (ids.includes('invisible') && sh.hidden)
        || (ids.includes('ink') && sh.ink)
        || (ids.includes('offSlide') && !sh.hidden && sh.box && (sh.box.x >= size.cx || sh.box.y >= size.cy || sh.box.x + sh.box.w <= 0 || sh.box.y + sh.box.h <= 0)));
      for (const sh of gone.reverse()) next = next.slice(0, sh.start) + next.slice(sh.end);
      if (next !== xml) pkg.write_(s, next);
    }
  }
  if (ids.includes('personal')) removePersonal(pkg, 'presentation');
  if (ids.includes('properties')) removeProperties(pkg);
  if (ids.includes('customXml')) removeCustomXml(pkg);
}

/* ── a workbook ──────────────────────────────────────────────────────────── */

const sheetPartsOf = (pkg) => {
  const wb = pkg.has('xl/workbook.xml') ? pkg.text('xl/workbook.xml') : '';
  const rels = new Map(pkg.rels('xl/workbook.xml').map((r) => [r.Id, OoxmlPackage.resolveTarget('xl/workbook.xml', r.Target)]));
  return [...wb.matchAll(/<sheet\b([^>]*)\/>/g)].map((m, index) => {
    const a = m[1];
    return { index, tag: m[0], name: unesc(/\bname="([^"]*)"/.exec(a)?.[1] ?? ''), state: /\bstate="([^"]*)"/.exec(a)?.[1] || 'visible', part: rels.get(/\br:id="([^"]*)"/.exec(a)?.[1]) || null };
  });
};

function inspectWorkbook(pkg) {
  const sheets = sheetPartsOf(pkg);
  const comments = partsWhere(pkg, (n) => /^xl\/(?:comments\d*|threadedComments\/[^/]+)\.xml$/.test(n))
    .flatMap((n) => [...pkg.text(n).matchAll(/<(?:comment|threadedComment)\b[\s\S]*?<\/(?:comment|threadedComment)>/g)]).length;
  const hiddenSheets = sheets.filter((s) => s.state !== 'visible').map((s) => s.name);
  let rows = 0;
  let cols = 0;
  let headers = 0;
  const hiddenObjects = [];
  for (const s of sheets) {
    if (!s.part || !pkg.has(s.part)) continue;
    const xml = pkg.text(s.part);
    rows += (xml.match(/<row\b[^>]*\bhidden="(?:1|true)"/g) || []).length;
    cols += [...xml.matchAll(/<col\b([^>]*)\/>/g)].filter((m) => /\bhidden="(?:1|true)"/.test(m[1])).reduce((k, m) => k + (Number(/\bmax="(\d+)"/.exec(m[1])?.[1]) - Number(/\bmin="(\d+)"/.exec(m[1])?.[1]) + 1 || 1), 0);
    if (/<headerFooter\b[^>]*>[\s\S]*?<(?:odd|even|first)(?:Header|Footer)>[^<]+/.test(xml)) headers += 1;
  }
  for (const d of partsWhere(pkg, (n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))) {
    for (const m of pkg.text(d).matchAll(/<xdr:cNvPr\b[^>]*\bname="([^"]*)"[^>]*\bhidden="(?:1|true)"/g)) hiddenObjects.push(unesc(m[1]));
  }
  const wb = pkg.has('xl/workbook.xml') ? pkg.text('xl/workbook.xml') : '';
  const hiddenNames = [...wb.matchAll(/<definedName\b([^>]*)>/g)].filter((m) => /\bhidden="(?:1|true)"/.test(m[1]) && !/\bname="_xlnm\./.test(m[1])).map((m) => unesc(/\bname="([^"]*)"/.exec(m[1])?.[1] ?? ''));
  return [
    { id: 'comments', found: comments, items: [] },
    inspectPersonal(pkg),
    inspectProperties(pkg),
    inspectCustomXml(pkg),
    { id: 'headers', found: headers, items: [] },
    { id: 'hiddenSheets', found: hiddenSheets.length, items: hiddenSheets },
    { id: 'hiddenRowsColumns', found: rows + cols, items: [], rows, columns: cols },
    { id: 'hiddenNames', found: hiddenNames.length, items: hiddenNames },
    { id: 'invisible', found: hiddenObjects.length, items: hiddenObjects },
  ];
}

function removeFromWorkbook(pkg, ids) {
  const sheets = sheetPartsOf(pkg);
  if (ids.includes('comments')) {
    for (const s of sheets) {
      if (!s.part || !pkg.has(s.part)) continue;
      const vml = relIdsTo(pkg, s.part, (target) => /vmlDrawing/.test(target) && pkg.has(target) && /ObjectType="Note"/.test(pkg.text(target)));
      let xml = pkg.text(s.part);
      for (const id of vml) xml = xml.replace(new RegExp('<legacyDrawing\\b[^>]*r:id="' + id + '"[^>]*/>'), '');
      if (xml !== pkg.text(s.part)) pkg.write_(s.part, xml);
      for (const r of pkg.rels(s.part)) {
        const target = OoxmlPackage.resolveTarget(s.part, r.Target);
        if (vml.includes(r.Id) || /\/comments\d*\.xml$|threadedComments/.test(target)) dropPart(pkg, target);
      }
    }
    for (const n of partsWhere(pkg, (x) => /^xl\/persons\//.test(x) && !x.includes('_rels/'))) dropPart(pkg, n);
  }
  if (ids.includes('headers')) {
    for (const s of sheets) if (s.part && pkg.has(s.part)) pkg.write_(s.part, pkg.text(s.part).replace(/<headerFooter\b[^>]*(?:\/>|>[\s\S]*?<\/headerFooter>)/g, ''));
  }
  if (ids.includes('hiddenNames') && pkg.has('xl/workbook.xml')) {
    pkg.write_('xl/workbook.xml', pkg.text('xl/workbook.xml').replace(/<definedName\b([^>]*)>[\s\S]*?<\/definedName>/g, (whole, a) => (/\bhidden="(?:1|true)"/.test(a) && !/\bname="_xlnm\./.test(a) ? '' : whole)));
  }
  if (ids.includes('invisible')) {
    for (const d of partsWhere(pkg, (n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))) {
      const xml = pkg.text(d);
      const next = xml.replace(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/xdr:\1>/g, (a) => (/<xdr:cNvPr\b[^>]*\bhidden="(?:1|true)"/.test(a) ? '' : a));
      if (next !== xml) pkg.write_(d, next);
    }
  }
  if (ids.includes('hiddenSheets')) removeHiddenSheets(pkg);
  if (ids.includes('personal')) removePersonal(pkg, 'workbook');
  if (ids.includes('properties')) removeProperties(pkg);
  if (ids.includes('customXml')) removeCustomXml(pkg);
}

/** Hidden sheets taken out: their parts, their entries, names that belong to them; the calculation chain is rebuilt by Excel. */
function removeHiddenSheets(pkg) {
  const sheets = sheetPartsOf(pkg);
  const gone = sheets.filter((s) => s.state !== 'visible');
  if (!gone.length || gone.length === sheets.length) return;
  let wb = pkg.text('xl/workbook.xml');
  const kept = sheets.filter((s) => s.state === 'visible').map((s) => s.index);
  for (const s of gone) wb = wb.replace(s.tag, '');
  // Names local to a sheet follow it: removed with it, renumbered for the rest.
  wb = wb.replace(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g, (whole, a, body) => {
    const local = /\blocalSheetId="(\d+)"/.exec(a);
    if (!local) return whole;
    const at = kept.indexOf(Number(local[1]));
    return at < 0 ? '' : '<definedName' + a.replace(/\blocalSheetId="\d+"/, () => 'localSheetId="' + at + '"') + '>' + body + '</definedName>';
  });
  // The active tab must be one that is still there.
  wb = wb.replace(/\bactiveTab="\d+"/, 'activeTab="0"').replace(/\bfirstSheet="\d+"/, 'firstSheet="0"');
  pkg.write_('xl/workbook.xml', wb);
  for (const s of gone) if (s.part) dropPart(pkg, s.part);
  if (pkg.has('xl/calcChain.xml')) dropPart(pkg, 'xl/calcChain.xml');
  if (pkg.has('docProps/app.xml')) {
    // The list of titles Excel keeps names the sheets; the hidden ones go from it.
    let app = pkg.text('docProps/app.xml');
    for (const s of gone) app = app.replace('<vt:lpstr>' + esc(s.name) + '</vt:lpstr>', '');
    pkg.write_('docProps/app.xml', app);
  }
}

/* ── the doors ───────────────────────────────────────────────────────────── */

/** What a package holds, kind by kind: `kind` is 'document', 'workbook' or 'presentation'. */
export function inspectPackage(pkg, kind) {
  if (kind === 'document') return inspectDocument(pkg);
  if (kind === 'presentation') return inspectPresentation(pkg);
  if (kind === 'workbook') return inspectWorkbook(pkg);
  throw new Error('Inspect works on a document, a workbook or a presentation.');
}

/** The kinds named in `ids` taken out of the package. Answers the ids it took out. */
export function removeFromPackage(pkg, kind, ids = []) {
  const known = new Set(inspectPackage(pkg, kind).filter((k) => !k.keep).map((k) => k.id));
  const asked = ids.filter((id) => known.has(id));
  if (kind === 'document') removeFromDocument(pkg, asked);
  else if (kind === 'presentation') removeFromPresentation(pkg, asked);
  else removeFromWorkbook(pkg, asked);
  return asked;
}

export { REL as RELATIONSHIPS_NS };
