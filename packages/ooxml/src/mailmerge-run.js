/**
 * Finish & Merge — the main document run once per record.
 *
 * The rules for what each field says live in `mailmerge.js`; this is the
 * file half. A merge walks the body's fields in document order, replacing
 * each mail merge field with its result for the record in hand — in the
 * formatting its own result run carried, so a bold «LastName» merges bold —
 * and leaves every other field (a page number, a cross-reference) as it
 * stands. A «Next Record» moves the same copy on to the next record, which
 * is how a sheet of labels holds twenty-one of them; a true «Skip Record
 * If» throws the record's copy away. The copies become one document — a
 * section each, as Word's Edit Individual Documents makes them, or one after
 * another with no break for a directory — or one e-mail message each.
 */
import { OoxmlPackage } from './package.js';
import { unesc } from './workbook.js';
import { Document } from './document.js';
import { topComplexFields, textOf, parseRuns, renderRun, RPR_RE } from './runs.js';
import { MERGE_KINDS, readInstr, evaluateField, fieldValue } from './mailmerge.js';

export { rangeOf } from './mailmerge.js';

/**
 * Every field in a fragment of the body, in document order: the complex
 * fields at the top (nesting honoured — an IF carries its MERGEFIELD) and
 * the simple ones outside them. Each knows its span and where its cached
 * result lies.
 */
export function fieldsIn(xml) {
  const source = String(xml);
  const complex = topComplexFields(source);
  const inside = (at) => complex.some((f) => at >= f.start && at < f.end);
  const simple = [];
  for (const m of source.matchAll(/<w:fldSimple\b([^>]*?)(?:\/>|>([\s\S]*?)<\/w:fldSimple>)/g)) {
    if (inside(m.index)) continue;
    const instr = /\bw:instr="([^"]*)"/.exec(m[1]);
    const openLen = m[0].indexOf('>') + 1;
    const selfClosing = m[2] === undefined;
    simple.push({
      start: m.index,
      end: m.index + m[0].length,
      instr: instr ? unesc(instr[1]) : '',
      resultStart: selfClosing ? m.index + m[0].length : m.index + openLen,
      resultEnd: selfClosing ? m.index + m[0].length : m.index + m[0].length - '</w:fldSimple>'.length,
    });
  }
  return [...complex, ...simple].sort((a, b) => a.start - b.start);
}

/** The words a field's place takes: its own result run's formatting, the value as runs, a line break for each "\n". */
function resultXml(xml, field, text) {
  if (!text) return '';
  const result = xml.slice(field.resultStart, field.resultEnd);
  const first = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/.exec(result) || /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/.exec(xml.slice(field.start, field.end));
  let rPr = first ? RPR_RE.exec(first[1])?.[0] ?? null : null;
  // Merged words are words: the proofing Word switched off on the field's
  // chevrons comes back on.
  if (rPr) {
    rPr = rPr.replace(/<w:noProof\b[^>]*\/>/g, '');
    if (/^<w:rPr\b[^>]*>\s*<\/w:rPr>$/.test(rPr)) rPr = null;
  }
  return renderRun(rPr, text);
}

/**
 * The copies a merge makes: `order` is the records to walk (indices into
 * `source.records`, in merge order), `first` the 1-based sequence number the
 * first copy takes. Each copy is `{ records, xml }` — the records it used
 * (more than one on a sheet of labels) and its body with every mail merge
 * field replaced.
 */
export function mergeCopies(bodyXml, source, order, { mapping = {} } = {}) {
  const xml = String(bodyXml);
  const fields = fieldsIn(xml).filter((f) => MERGE_KINDS.has(readInstr(f.instr).kind));
  const copies = [];
  let i = 0;
  let sequence = 0;
  while (i < order.length) {
    let pointer = i;
    let skip = false;
    const used = [order[i]];
    let out = '';
    let at = 0;
    for (const f of fields) {
      out += xml.slice(at, f.start);
      at = f.end;
      const index = pointer < order.length ? order[pointer] : null;
      const record = index === null ? null : source.records[index];
      const got = record
        ? evaluateField(f.instr, { source, record, recordNumber: index + 1, sequence: sequence + 1, mapping })
        : { text: '', next: readInstr(f.instr).kind === 'next' };
      if (got.skip) skip = true;
      if (got.next) {
        pointer += 1;
        if (pointer < order.length) used.push(order[pointer]);
      }
      out += resultXml(xml, f, got.text);
    }
    out += xml.slice(at);
    i = pointer + 1;
    if (skip) continue;
    sequence += 1;
    copies.push({ records: used, xml: out });
  }
  return copies;
}

/**
 * The last paragraph at the top of a fragment — the one that ends a record's
 * section — as `{ start, end }`, or null when the fragment ends in a table
 * with no paragraph after it.
 */
function lastTopParagraph(xml) {
  const re = /<w:(tbl|sdt|txbxContent)\b[^>]*?(\/?)>|<\/w:(tbl|sdt|txbxContent)>|<w:p\b[^>]*?(\/?)>|<\/w:p>/g;
  let depth = 0;
  let start = -1;
  let last = null;
  let tailIsTable = false;
  let m;
  while ((m = re.exec(xml))) {
    const tag = m[0];
    if (m[1]) { if (m[2] !== '/') depth += 1; continue; }
    if (m[3]) {
      depth -= 1;
      if (depth === 0 && m[3] === 'tbl') tailIsTable = true;
      continue;
    }
    if (depth > 0) continue;
    if (tag.startsWith('<w:p')) {
      if (tag.endsWith('/>')) { last = { start: m.index, end: m.index + tag.length }; tailIsTable = false; } else start = m.index;
      continue;
    }
    if (tag === '</w:p>' && start >= 0) {
      last = { start, end: m.index + tag.length };
      tailIsTable = false;
      start = -1;
    }
  }
  return tailIsTable ? null : last;
}

/** A paragraph given a section break: `w:sectPr` in its `w:pPr`, where the schema puts it. */
function withSectionBreak(paragraphXml, sectPr) {
  let p = paragraphXml;
  if (/^<w:p\b[^>]*\/>$/.test(p)) p = p.replace(/\/>$/, '>') + '</w:p>';
  const open = /^<w:p\b[^>]*>/.exec(p)[0];
  const pPr = /<w:pPr\b[^>]*\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.exec(p.slice(open.length, open.length + 20000));
  if (!pPr || pPr.index !== 0) return open + '<w:pPr>' + sectPr + '</w:pPr>' + p.slice(open.length);
  const at = open.length;
  let inner = pPr[0];
  if (/^<w:pPr\b[^>]*\/>$/.test(inner)) inner = '<w:pPr>' + sectPr + '</w:pPr>';
  else {
    inner = inner.replace(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/, '');
    const change = inner.indexOf('<w:pPrChange');
    inner = change >= 0 ? inner.slice(0, change) + sectPr + inner.slice(change) : inner.replace(/<\/w:pPr>$/, sectPr + '</w:pPr>');
  }
  return p.slice(0, at) + inner + p.slice(at + pPr[0].length);
}

/** The body split into its content and the document's own `w:sectPr` at its end. */
function splitBody(doc) {
  const { prefix, body, suffix } = doc._body();
  const m = /<w:sectPr\b[^>]*\/>\s*$|<w:sectPr\b[^>]*>(?:(?!<w:sectPr\b)[\s\S])*?<\/w:sectPr>\s*$/.exec(body);
  return m
    ? { prefix, content: body.slice(0, m.index), sectPr: m[0].trim(), suffix }
    : { prefix, content: body, sectPr: '<w:sectPr/>', suffix };
}


/**
 * Edit Individual Documents: a new document of every copy — a section each,
 * the document's own page set up again at every break, as Word's merge
 * writes it; for a directory, one after another in the one section. The
 * merged document is an ordinary one: no mail merge settings, no fields
 * left to merge. Answers `{ bytes, copies, records }`.
 */
export function mergeToDocument(doc, source, order, { type = 'formLetters', mapping = {} } = {}) {
  const { prefix, content, sectPr, suffix } = splitBody(doc);
  const copies = mergeCopies(content, source, order, { mapping });
  let body = '';
  copies.forEach((copy, k) => {
    let xml = copy.xml;
    if (type !== 'catalog' && k < copies.length - 1) {
      // The record's section ends at its last paragraph — or, when it ends in
      // a table (a sheet of labels), at a paragraph of its own after it.
      const last = lastTopParagraph(xml);
      xml = last
        ? xml.slice(0, last.start) + withSectionBreak(xml.slice(last.start, last.end), sectPr) + xml.slice(last.end)
        : xml + '<w:p><w:pPr>' + sectPr + '</w:pPr></w:p>';
    }
    body += xml;
  });
  const pkg = OoxmlPackage.read(doc.save());
  const out = new Document(pkg);
  out.xml = prefix + body + sectPr + suffix;
  out.dirty = true;
  out.setMailMerge(null);
  return { bytes: out.save(), copies: copies.length, records: copies.reduce((n, c) => n + c.records.length, 0) };
}

const escHtml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** The paragraphs of a fragment in reading order, text boxes aside. */
function paragraphsOf(xml) {
  const out = [];
  const clean = String(xml).replace(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/g, '');
  for (const m of clean.matchAll(/<w:p\b[^>]*\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)) out.push(m[0]);
  return out;
}

/** A merged copy as a message body: plain text, a line per paragraph. */
export function bodyText(xml) {
  return paragraphsOf(xml).map((p) => textOf(p).replace(/\uFFFC/g, '')).join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

/** A merged copy as HTML: a `<p>` per paragraph, bold, italic and underline kept. */
export function bodyHtml(xml) {
  const paras = paragraphsOf(xml).map((p) => {
    const runs = parseRuns(p).filter((r) => !r.del);
    const inner = runs.map((r) => {
      let t = escHtml(String(r.text || '').replace(/\uFFFC/g, '')).replace(/\n/g, '<br>').replace(/\t/g, '&emsp;');
      if (!t) return '';
      if (r.bold) t = '<b>' + t + '</b>';
      if (r.italic) t = '<i>' + t + '</i>';
      if (r.underline) t = '<u>' + t + '</u>';
      return t;
    }).join('');
    return '<p style="margin:0 0 10px">' + (inner || '&nbsp;') + '</p>';
  });
  return '<!doctype html><html><body style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' + paras.join('') + '</body></html>';
}

/**
 * Send E-mail Messages: a message per copy — to the address in the column
 * the To box names, with the Subject line as typed, the merged document as
 * its body (plain, or HTML with a plain copy beside it). Each carries the
 * record it came from, so a failure can say whose it was.
 */
export function mergeMessages(doc, source, order, { toField, subject = '', format = 'html', mapping = {} } = {}) {
  const { content } = splitBody(doc);
  return mergeCopies(content, source, order, { mapping }).map((copy) => {
    const record = source.records[copy.records[0]];
    const text = bodyText(copy.xml);
    return {
      record: copy.records[0] + 1,
      to: fieldValue(source, record, toField).trim(),
      subject: String(subject || ''),
      text,
      ...(format === 'html' ? { html: bodyHtml(copy.xml) } : {}),
    };
  });
}
