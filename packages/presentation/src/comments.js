// Comments on slides — Review → New Comment, replies, Resolve, Delete.
//
// PowerPoint 365 writes "modern" comments: one part per slide,
// `ppt/comments/modernComment_<slide id>_<creation id>.xml`, a `p188:cmLst`
// of threads, each anchored to the slide or to one of its shapes by the
// "moniker" lists PowerPoint's own commands use (`pc:sldMk`, `ac:spMk`),
// with its replies inside it; the people are in `ppt/authors.xml`. Older
// PowerPoint wrote `p:cmLst` parts and `ppt/commentAuthors.xml`; those are
// read and can be deleted, and — as in PowerPoint 365 — are not replied to.
//
// This module is the XML: reading both formats into plain threads and
// writing the modern one. The deck owns the parts and relationships.

import { parse, kids, first, all, textOf, escapeXml } from '@rutba/office-formats/xml';

export const NS = {
  p188: 'http://schemas.microsoft.com/office/powerpoint/2018/8/main',
  pc: 'http://schemas.microsoft.com/office/powerpoint/2013/main/command',
  ac: 'http://schemas.microsoft.com/office/drawing/2013/main/command',
  a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

export const COMMENT_REL = {
  modern: 'http://schemas.microsoft.com/office/2018/10/relationships/comments',
  authors: 'http://schemas.microsoft.com/office/2018/10/relationships/authors',
  legacy: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  legacyAuthors: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors',
};

export const COMMENT_CT = {
  modern: 'application/vnd.ms-powerpoint.comments+xml',
  authors: 'application/vnd.ms-powerpoint.authors+xml',
};

/** The slide extension that names its modern comments part. */
export const COMMENT_REL_EXT = '{6950BFC3-D8DA-4A85-94F7-54DA5524770B}';
/** The slide extension that carries its creation id — the `cId` a comment's moniker names. */
export const CREATION_ID_EXT = '{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export function guid() {
  const hex = () => Math.floor(Math.random() * 16).toString(16).toUpperCase();
  return '{' + [8, 4, 4, 4, 12].map((n) => Array.from({ length: n }, hex).join('')).join('-') + '}';
}

/** A time as PowerPoint writes one in a comment: local-free ISO with milliseconds, no zone letter. */
export function stamp(date = new Date()) {
  return date.toISOString().replace(/Z$/, '');
}

/** Initials from a name: the first letters of the first and last words. */
export function initialsOf(name) {
  const words = String(name || '').trim().split(/\s+/).filter(Boolean);
  if (!words.length) return '?';
  return (words.length === 1 ? words[0].slice(0, 2) : words[0][0] + words[words.length - 1][0]).toUpperCase();
}

/** A text body's words, one line a paragraph. */
function bodyText(node) {
  if (!node) return '';
  return kids(node, 'a:p').map((p) => kids(p).filter((k) => k.name === 'a:r' || k.name === 'a:br' || k.name === 'a:fld').map((k) => (k.name === 'a:br' ? '\n' : textOf(first(k, 'a:t')))).join('')).join('\n');
}

/** Words as a comment's text body, a paragraph a line. */
export function textBodyXml(text) {
  const lines = String(text ?? '').split('\n');
  return '<p188:txBody><a:bodyPr/><a:lstStyle/>' +
    lines.map((l) => (l ? `<a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(l)}</a:t></a:r></a:p>` : '<a:p><a:endParaRPr lang="en-US"/></a:p>')).join('') +
    '</p188:txBody>';
}

/** `ppt/authors.xml` → Map id → { id, name, initials }. */
export function readAuthors(xml) {
  const out = new Map();
  if (!xml) return out;
  for (const a of all(parse(xml), 'p188:author')) out.set(a.attrs.id, { id: a.attrs.id, name: a.attrs.name || 'Unknown', initials: a.attrs.initials || initialsOf(a.attrs.name) });
  return out;
}

export function authorsXml(authors) {
  return `${DECL}<p188:authorLst xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p188="${NS.p188}">` +
    authors.map((a) => `<p188:author id="${a.id}" name="${escapeXml(a.name)}" initials="${escapeXml(a.initials)}" userId="${escapeXml(a.name)}" providerId="None"/>`).join('') +
    '</p188:authorLst>';
}

/**
 * A modern comments part → its threads: id, author, when, words, status,
 * what it is anchored to (a shape's id, or the slide), a position when it
 * has one (EMU), and its replies.
 */
export function readModernComments(xml, authors) {
  if (!xml) return [];
  const root = parse(xml);
  const who = (id) => authors.get(id) || { id, name: 'Unknown', initials: '?' };
  return all(root, 'p188:cm').map((cm) => {
    const sp = first(cm, 'ac:spMk');
    const pos = kids(cm, 'p188:pos')[0];
    const author = who(cm.attrs.authorId);
    return {
      id: cm.attrs.id,
      authorId: cm.attrs.authorId,
      author: author.name,
      initials: author.initials,
      created: cm.attrs.created || null,
      status: cm.attrs.status || 'active',
      text: bodyText(kids(cm, 'p188:txBody')[0]),
      shapeId: sp?.attrs.id ?? null,
      pos: pos ? { x: Number(pos.attrs.x) || 0, y: Number(pos.attrs.y) || 0 } : null,
      replies: kids(first(cm, 'p188:replyLst') || { children: [] }, 'p188:reply').map((r) => {
        const ra = who(r.attrs.authorId);
        return { id: r.attrs.id, authorId: r.attrs.authorId, author: ra.name, initials: ra.initials, created: r.attrs.created || null, text: bodyText(kids(r, 'p188:txBody')[0]) };
      }),
      legacy: false,
    };
  });
}

/** `ppt/commentAuthors.xml` → Map id → { name, initials }. */
export function readLegacyAuthors(xml) {
  const out = new Map();
  if (!xml) return out;
  for (const a of all(parse(xml), 'p:cmAuthor')) out.set(String(a.attrs.id), { name: a.attrs.name || 'Unknown', initials: a.attrs.initials || initialsOf(a.attrs.name) });
  return out;
}

/**
 * An older PowerPoint's `p:cmLst` → threads the same shape as the modern
 * ones: a comment whose threading extension names a parent is that
 * parent's reply. Positions are in eighths of a point.
 */
export function readLegacyComments(xml, authors) {
  if (!xml) return [];
  const list = all(parse(xml), 'p:cm').map((cm) => {
    const a = authors.get(String(cm.attrs.authorId)) || { name: 'Unknown', initials: '?' };
    const pos = kids(cm, 'p:pos')[0];
    const parent = first(cm, 'p15:parentCm');
    return {
      id: `legacy:${cm.attrs.authorId}:${cm.attrs.idx}`,
      authorId: String(cm.attrs.authorId),
      idx: String(cm.attrs.idx),
      author: a.name,
      initials: a.initials,
      created: cm.attrs.dt || null,
      status: 'active',
      text: textOf(kids(cm, 'p:text')[0]),
      shapeId: null,
      // Eighths of a point to EMU (12700 a point).
      pos: pos ? { x: Math.round((Number(pos.attrs.x) || 0) * 1587.5), y: Math.round((Number(pos.attrs.y) || 0) * 1587.5) } : null,
      parent: parent ? `legacy:${parent.attrs.authorId}:${parent.attrs.idx}` : null,
      replies: [],
      legacy: true,
    };
  });
  const byId = new Map(list.map((c) => [c.id, c]));
  const threads = [];
  for (const c of list) {
    const p = c.parent && byId.get(c.parent);
    if (p) p.replies.push({ id: c.id, authorId: c.authorId, author: c.author, initials: c.initials, created: c.created, text: c.text });
    else threads.push(c);
  }
  return threads;
}

/** One modern thread as PowerPoint writes it: anchor, position, replies, words — in the schema's order. */
export function commentXml({ id, authorId, created, status = 'active', sldId, cId, shape = null, pos = null, replies = [], text }) {
  const sldMk = `<pc:sldMk xmlns:pc="${NS.pc}" cId="${cId}" sldId="${sldId}"/>`;
  const docMk = `<pc:docMk xmlns:pc="${NS.pc}"/>`;
  const anchor = shape
    ? `<ac:deMkLst xmlns:ac="${NS.ac}">${docMk}${sldMk}<ac:spMk id="${shape.id}" creationId="${shape.creationId}"/></ac:deMkLst>`
    : `<pc:sldMkLst xmlns:pc="${NS.pc}"><pc:docMk/><pc:sldMk cId="${cId}" sldId="${sldId}"/></pc:sldMkLst>`;
  const posXml = pos ? `<p188:pos x="${Math.round(pos.x)}" y="${Math.round(pos.y)}"/>` : '';
  const repliesXml = replies.length
    ? `<p188:replyLst>${replies.map((r) => `<p188:reply id="${r.id}" authorId="${r.authorId}" created="${r.created}">${textBodyXml(r.text)}</p188:reply>`).join('')}</p188:replyLst>`
    : '';
  return `<p188:cm id="${id}" authorId="${authorId}"${status && status !== 'active' ? ` status="${status}"` : ''} created="${created}">${anchor}${posXml}${repliesXml}${textBodyXml(text)}</p188:cm>`;
}

export function commentListXml(inner) {
  return `${DECL}<p188:cmLst xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p188="${NS.p188}">${inner}</p188:cmLst>`;
}

/** The range of one `p188:cm` (by id) in a part, or null. */
export function threadRange(xml, id) {
  const re = /<p188:cm\b[^>]*>[\s\S]*?<\/p188:cm>/g;
  let m;
  while ((m = re.exec(xml))) {
    if (new RegExp(`^<p188:cm\\b[^>]*\\bid="${id.replace(/[{}]/g, (c) => `\\${c}`)}"`).test(m[0])) return { start: m.index, end: m.index + m[0].length, xml: m[0] };
  }
  return null;
}
