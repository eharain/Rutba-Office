'use strict';

/**
 * A business document, laid out.
 *
 * Everything below the letterhead is the same on an invoice, a quote, a
 * credit note and a statement: who it is from, who it is to, what it refers
 * to, a table of things, what they come to, and the small print. Only the
 * words change. So this takes a plain description - no database rows, no
 * Strapi, nothing that knows what an invoice is - and draws it.
 *
 * That split is the point. The service that owns invoices builds a
 * description; this owns the paper. Neither has to learn the other's job,
 * and a document type nobody has thought of yet is a new description rather
 * than a new layout.
 *
 * == One decision that shows on every page ==============================
 *
 * **The currency is named once, in the column heading, and never repeated
 * in a cell.** "Amount (GBP)" over a column of bare numbers is what a
 * finance department expects, it survives a customer whose reader lacks the
 * symbol, and it removes any chance of a document showing two currencies'
 * symbols in one table because one row was priced before a rate changed.
 */

const { Sheet, INK } = require('./sheet');

const isFilled = (v) => v !== null && v !== undefined && v !== '';

/**
 * Draw a described document and return the Sheet, still open, so a caller
 * can append something this does not know about before finishing.
 */
function drawPaper(paper = {}, options = {}) {
  const sheet = new Sheet({
    size: paper.size || 'A4',
    landscape: Boolean(paper.landscape),
    title: [paper.title, paper.reference].filter(Boolean).join(' '),
    author: paper.letterhead?.name || '',
    subject: paper.subject || '',
    created: paper.created || options.created || null,
    footer: paper.footer || '',
  });

  letterhead(sheet, paper);
  parties(sheet, paper);

  if (isFilled(paper.intro)) {
    sheet.paragraph(paper.intro, { size: 9.5, gap: 8 });
  }

  if (paper.table && (paper.table.columns || []).length) {
    sheet.table({
      columns: paper.table.columns,
      rows: paper.table.rows || [],
      zebra: paper.table.zebra !== false,
      size: paper.table.size || 9,
    });
    if (!(paper.table.rows || []).length) {
      sheet.paragraph(paper.table.empty || 'Nothing to show.',
        { size: 9, colour: INK.quiet, gap: 8 });
    }
  }

  if ((paper.totals || []).length) {
    sheet.totals(paper.totals, { size: 9.5 });
    sheet.spacer(10);
  }

  for (const section of paper.sections || []) {
    if (!isFilled(section.body) && !(section.rows || []).length) continue;
    sheet.spacer(4);
    if (section.heading) {
      sheet.paragraph(section.heading, { size: 9, font: 'Helvetica-Bold', gap: 2 });
    }
    if (isFilled(section.body)) {
      sheet.paragraph(section.body, { size: 8.8, colour: INK.quiet, gap: 6 });
    }
    if ((section.rows || []).length) {
      sheet.keyValues(section.rows, {
        width: Math.min(320, sheet.contentWidth), labelWidth: 130, align: 'left', size: 8.8,
      });
      sheet.spacer(4);
    }
  }

  signatures(sheet, paper);

  return sheet;
}

/**
 * The adopted marks (sign S5). Each entry is one party's execution: a
 * decoded image (images.js shape) drawn at up to 170x46pt with its aspect
 * kept, or a typed adoption rendered in an italic face - which is exactly
 * what the ceremony showed the signer when they adopted it. The name and
 * detail lines under each mark are the certificate's own words; nothing
 * here invents a claim.
 */
function signatures(sheet, paper) {
  const entries = (paper.signatures || []).filter((s) => s && (s.image || isFilled(s.typedName)));
  if (!entries.length) return;

  sheet.spacer(6);
  sheet.paragraph('Signatures as adopted', { size: 9, font: 'Helvetica-Bold', gap: 4 });

  for (const entry of entries) {
    // Keep a mark and its caption together across a page break.
    sheet.space(46 + 26);
    if (entry.image) {
      const maxW = Math.min(170, sheet.contentWidth);
      const maxH = 46;
      const scale = Math.min(maxW / entry.image.width, maxH / entry.image.height, 1);
      const w = Math.max(24, entry.image.width * scale);
      const h = Math.max(12, entry.image.height * scale);
      const ref = sheet.doc.addImage(entry.image);
      sheet.image(ref, w, h, { gap: 2 });
    } else {
      sheet.paragraph(entry.typedName, { size: 19, font: 'Helvetica-Oblique', gap: 2 });
    }
    if (isFilled(entry.name)) {
      sheet.paragraph(entry.name, { size: 9, font: 'Helvetica-Bold', gap: 0 });
    }
    if (isFilled(entry.detail)) {
      sheet.paragraph(entry.detail, { size: 8.4, colour: INK.quiet, gap: 4 });
    }
    sheet.rule({ gap: 6, colour: INK.faint });
  }
}

/** The description in, the bytes out. */
function renderPaper(paper, options = {}) {
  return drawPaper(paper, options).toBuffer();
}

/**
 * Company block on the left, document name on the right.
 *
 * The document's own name is the largest thing on the page, because the
 * first question a person holding it asks is "what is this?" and the second
 * is "which one?" - so the reference sits directly under it.
 */
function letterhead(sheet, paper) {
  const head = paper.letterhead || {};
  const top = sheet.y;

  const left = [];
  if (isFilled(head.name)) left.push({ text: head.name, font: 'Helvetica-Bold', size: 13 });
  for (const line of head.lines || []) if (isFilled(line)) left.push({ text: line, size: 8.6, colour: INK.quiet });
  if ((head.contact || []).length) {
    left.push({ gap: 3 });
    for (const line of head.contact) if (isFilled(line)) left.push({ text: line, size: 8.6, colour: INK.quiet });
  }

  const right = [{ text: String(paper.title || '').toUpperCase(), font: 'Helvetica-Bold', size: 20, align: 'right' }];
  if (isFilled(paper.reference)) {
    right.push({ text: String(paper.reference), size: 11, align: 'right' });
  }
  if (isFilled(paper.status)) {
    right.push({ gap: 2 });
    right.push({ text: String(paper.status).toUpperCase(), size: 9, align: 'right', font: 'Helvetica-Bold', colour: paper.status_colour || INK.quiet });
  }

  sheet.y = top;
  sheet.columns([
    { lines: left, flex: 1.4 },
    { lines: right, flex: 1 },
  ]);
  sheet.rule({ gap: 8 });
}

/** Who it is addressed to, and the handful of dates that qualify it. */
function parties(sheet, paper) {
  const party = paper.party;
  const meta = (paper.meta || []).filter(([, value]) => isFilled(value));
  if (!party && !meta.length) return;

  const top = sheet.y;
  let afterMeta = top;
  if (meta.length) {
    const width = Math.min(250, sheet.contentWidth * 0.45);
    sheet.keyValues(meta, {
      x: sheet.right - width, width, labelWidth: width * 0.5, align: 'right', size: 9,
    });
    afterMeta = sheet.y;
  }

  sheet.y = top;
  if (party) {
    const lines = [];
    if (isFilled(party.label)) lines.push({ text: party.label, size: 8, colour: INK.quiet });
    if (isFilled(party.name)) lines.push({ text: party.name, font: 'Helvetica-Bold', size: 10.5 });
    for (const line of party.lines || []) if (isFilled(line)) lines.push({ text: line, size: 9, colour: INK.quiet });
    sheet.columns([{ lines, width: sheet.contentWidth * 0.5 }]);
  }

  sheet.y = Math.max(afterMeta, sheet.y) + 12;
}

module.exports = { drawPaper, renderPaper };
