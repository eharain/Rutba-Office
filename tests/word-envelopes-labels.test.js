// Mailings → Envelopes and Labels — the engine's half.
//
// A sheet of labels is a table the size of the sheet, the labels its cells at
// the sheet's own pitch; an envelope is a section of its own at the front of
// the letter, on its own paper. These pin the XML Word would recognise as its
// own, and — because a label that prints across two labels is no label — the
// PDF: its paper, and where every address lands on it.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { buildDocx } from '@rutba/ooxml/build';
import { LABEL_PRODUCTS, labelProduct, ENVELOPE_SIZES } from '@rutba/ooxml/labels';
import { sourceFromRows, parseDelimited } from '@rutba/ooxml/mailmerge';

const ADDRESS = ['Mr. Joshua Randall', '12 High Street', 'London SW1A 1AA'];
const PT = 72 / 25.4; // points per millimetre

/** Every text a PDF draws, with where: `{ text, x, y }` in points from the page's bottom-left. */
function pdfTexts(buffer) {
  const s = buffer.toString('latin1');
  const pages = [];
  for (const stream of s.matchAll(/stream\r?\n([\s\S]*?)endstream/g)) {
    const out = [];
    for (const m of stream[1].matchAll(/1 0 0 1 ([\d.-]+) ([\d.-]+) Tm\s*\(((?:\\.|[^\\)])*)\) Tj/g)) out.push({ text: m[3], x: Number(m[1]), y: Number(m[2]) });
    if (out.length) pages.push(out);
  }
  return pages;
}
const mediaBoxes = (buffer) => [...buffer.toString('latin1').matchAll(/\/MediaBox \[([^\]]*)\]/g)].map((m) => m[1].split(' ').map(Number));

const blank = () => openDocx(buildDocx({ paragraphs: [{ text: '' }] }));
const letter = () => openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Dear Joshua,' }, { text: 'Thank you for your order.' }] }));

test('an Avery L7160 sheet is Word\'s table: three labels and two gaps across, seven exact rows, no borders, the sheet\'s margins', () => {
  const view = blank();
  view.setLabelSheet({ product: 'L7160', lines: ADDRESS });
  const xml = view.doc.doc.xml;
  assert.match(xml, /<w:tblGrid><w:gridCol w:w="3600"\/><w:gridCol w:w="144"\/><w:gridCol w:w="3600"\/><w:gridCol w:w="144"\/><w:gridCol w:w="3600"\/><\/w:tblGrid>/, 'label, gap, label, gap, label — 63.5 mm and 2.54 mm');
  assert.equal((xml.match(/<w:trHeight w:val="2160" w:hRule="exact"\/>/g) || []).length, 7, 'seven rows of exactly 38.1 mm');
  assert.match(xml, /<w:tblLayout w:type="fixed"\/><w:tblBorders><w:top w:val="nil"\/>/);
  assert.match(xml, /<w:pgSz w:w="11906" w:h="16838"\/><w:pgMar w:top="859" w:right="\d+" w:bottom="0" w:left="411"/, 'A4, 15.15 mm at the top, 7.25 mm at the side');
  const cells = view.blocks.filter((b) => b.container);
  const labels = cells.filter((b) => b.text);
  assert.equal(labels.length, 21, 'twenty-one labels, all the same address');
  assert.ok(labels.every((b) => b.text === ADDRESS.join('\n')));
  assert.equal(cells.length, 35, '21 labels and the 14 narrow cells between them');
  assert.deepEqual(labels[0].tableLook, { bare: true, fixed: true, cellMarginPx: { left: 1, right: 1, top: 0, bottom: 0 } });
  assert.equal(labels[0].rowRule, 'exact');
  assert.equal(labels[0].cellVAlign, 'center');
  const s = view.section;
  assert.equal(Math.round(s.widthPx), 794);
  assert.equal(Math.round(s.heightPx), 1123);
  assert.equal(Math.round(s.margins.top), 57);
});

test('the label sheet prints on A4, one page, every address inside its own label', () => {
  const view = blank();
  const p = labelProduct('L7160');
  view.setLabelSheet({ product: 'L7160', lines: ADDRESS });
  const pdf = renderPdf(view, {});
  assert.equal(pdf.pages, 1, 'the sheet fits one page — the last small paragraph with it');
  const [box] = mediaBoxes(pdf.buffer);
  assert.deepEqual(box.map((n) => Math.round(n)), [0, 0, 595, 842]);
  const texts = pdfTexts(pdf.buffer)[0];
  const firsts = texts.filter((t) => t.text === ADDRESS[0]);
  assert.equal(firsts.length, 21);
  const pageH = 842 * (841.89 / 842);
  for (const t of firsts) {
    const col = Math.floor((t.x / PT - p.side) / p.pitchX);
    const row = Math.floor(((pageH - t.y) / PT - p.top) / p.pitchY);
    assert.ok(col >= 0 && col < p.cols && row >= 0 && row < p.rows, `in the sheet: ${t.x}, ${t.y}`);
    const left = (p.side + col * p.pitchX) * PT;
    const top = (p.top + row * p.pitchY) * PT;
    assert.ok(t.x >= left && t.x + 110 <= left + p.w * PT + 1, `across, inside label ${row + 1}:${col + 1} (${t.x} in ${left}…${left + p.w * PT})`);
    assert.ok(pageH - t.y >= top && pageH - t.y <= top + p.h * PT, `down, inside label ${row + 1}:${col + 1}`);
  }
  // The three lines of each address, one under another in the label.
  const lines = texts.filter((t) => Math.abs(t.x - firsts[0].x) < 0.01 && t.y <= firsts[0].y && t.y > firsts[0].y - 50).map((t) => t.text);
  assert.deepEqual(lines, ADDRESS);
});

test('a single label goes at the row and column asked; a US Letter sheet is Letter paper', () => {
  const view = blank();
  view.setLabelSheet({ product: 'L7160', lines: ADDRESS, mode: 'single', row: 2, col: 3 });
  const labels = view.blocks.filter((b) => b.container && b.text);
  assert.equal(labels.length, 1);
  assert.equal(labels[0].container.replace(/^t\d+:/, ''), 'r1:c4', 'row 2, the third label — the fifth cell, past two gaps');

  const us = blank();
  us.setLabelSheet({ product: '5160', lines: ADDRESS });
  assert.equal(us.blocks.filter((b) => b.container && b.text).length, 30);
  const [box] = mediaBoxes(renderPdf(us, {}).buffer);
  assert.deepEqual(box.map((n) => Math.round(n)), [0, 0, 612, 792]);
  assert.ok(LABEL_PRODUCTS.length >= 12 && ['L7160', 'L7163', 'L7173', '5160', '5161', '5163', '5164', '3422'].every((id) => LABEL_PRODUCTS.some((x) => x.id === id)));
});

test('Add to Document puts a DL envelope in front as a section of its own, landscape, the delivery address in Word\'s frame', () => {
  const view = letter();
  view.addEnvelope({ size: 'DL', delivery: ADDRESS, returnAddress: ['Rutba Office', '1 Main Road', 'Leeds LS1 1AA'] });
  const engine = view.doc.doc;
  const sections = engine.sections();
  assert.equal(sections.length, 2);
  assert.equal(sections[0].envelope, true);
  assert.equal(sections[0].code, 27);
  assert.equal(sections[0].orientation, 'landscape');
  assert.equal(Math.round(sections[0].widthPx), 831, '220 mm across');
  assert.equal(Math.round(sections[0].heightPx), 416, '110 mm down');
  assert.match(engine.xml, /<w:pgSz w:w="12472" w:h="6236" w:orient="landscape" w:code="27"\/><w:pgMar w:top="360" w:right="576" w:bottom="720" w:left="576"/);
  assert.match(engine.xml, /<w:pgNumType w:start="0"\/>/, 'the letter\'s pages count from 1 again');
  const dl = ENVELOPE_SIZES.find((s) => s.id === 'DL');
  assert.match(engine.xml, new RegExp(`<w:pStyle w:val="EnvelopeAddress"/><w:framePr w:w="7920" w:h="1980" w:hRule="exact" w:hSpace="180" w:wrap="auto" w:hAnchor="page" w:vAnchor="page" w:x="${Math.round((dl.w - 7920) / 2)}" w:y="${dl.h - 1980}"/></w:pPr>`));
  assert.equal((engine.pkg.text('word/styles.xml').match(/w:styleId="EnvelopeAddress"/g) || []).length, 1);
  // The letter keeps its own page — the editor still draws A4.
  assert.equal(Math.round(view.section.heightPx), 1123);
  const frame = view.blocks.find((b) => b.frame);
  assert.equal(frame.text, ADDRESS.join('\n'));
  assert.equal(view.blocks.find((b) => b.text === 'Dear Joshua,').pageBreakBefore, true, 'the letter starts on a page of its own');
  const summary = view.render({ pages: false }).envelope;
  assert.equal(summary.size, 'DL');
  assert.deepEqual(summary.delivery, ADDRESS);

  // Change Document: the same envelope section, rewritten — not a second one.
  view.addEnvelope({ size: 'C5', delivery: ['Ms. Cynthia Gartner', 'Contoso Ltd'], returnAddress: null });
  const again = engine.sections();
  assert.equal(again.length, 2);
  assert.equal(again[0].code, 28);
  assert.deepEqual(engine.envelope().delivery, ['Ms. Cynthia Gartner', 'Contoso Ltd']);
  assert.deepEqual(engine.envelope().returnAddress, []);
  assert.ok(view.undo(), 'one undo takes the change back');
  assert.equal(engine.envelope().size, 'DL');
});

test('the envelope prints on its own paper and the letter on A4, each address where the envelope needs it', () => {
  const view = letter();
  view.addEnvelope({ size: 'DL', delivery: ADDRESS, returnAddress: ['Rutba Office', 'Leeds'] });
  const pdf = renderPdf(view, {});
  assert.equal(pdf.pages, 2);
  const boxes = mediaBoxes(pdf.buffer).map((b) => b.map((n) => Math.round(n)));
  assert.deepEqual(boxes, [[0, 0, 624, 312], [0, 0, 595, 842]], 'DL landscape, then A4');
  const [env, page] = pdfTexts(pdf.buffer);
  const del = env.find((t) => t.text === ADDRESS[0]);
  const dl = ENVELOPE_SIZES.find((s) => s.id === 'DL');
  const frameLeft = ((dl.w - 7920) / 2 + 2880) / 20;
  const frameTop = (dl.h - 1980) / 20;
  assert.ok(Math.abs(del.x - frameLeft) < 1, `the delivery address starts where the frame's indent puts it (${del.x} vs ${frameLeft})`);
  assert.ok(312 - del.y > frameTop && 312 - del.y < frameTop + 99, 'inside the frame, down');
  const ret = env.find((t) => t.text === 'Rutba Office');
  assert.ok(Math.abs(ret.x - 576 / 20) < 1 && 312 - ret.y < 40, 'the return address in the top-left corner, at the envelope\'s margins');
  assert.ok(page.some((t) => t.text === 'Dear Joshua,'));
});

test('Start Mail Merge → Labels: every label after the first begins «Next Record»; Update Labels copies the first; the merge fills a label per recipient', () => {
  const view = blank();
  view.setLabelSheet({ product: 'L7163', mode: 'merge' });
  assert.equal(view.merge.type, 'mailingLabels');
  const cells = () => view.blocks.filter((b) => /:c[02]$/.test(b.container || ''));
  assert.equal(cells().filter((b) => b.text === '«Next Record»').length, 13, 'fourteen labels, the first empty');
  view.attachMergeSource(sourceFromRows(parseDelimited('Title,First Name,Last Name,Address Line 1,City,ZIP Code\nMr.,Joshua,Randall,12 High Street,London,SW1A 1AA\nMs.,Cynthia,Gartner,1 Main Road,Leeds,LS1 1AA\nDr.,Amira,Haddad,4 Rue Neuve,Paris,75001\n'), { kind: 'csv', name: 'l.csv' }));
  const first = view.blocks.find((b) => b.container?.endsWith('r0:c0'));
  view.setSelection({ block: first.index, offset: 0 });
  view.insertAddressBlock({ country: 'never' });
  const n = view.updateLabels();
  assert.equal(n, 13);
  assert.ok(cells().filter((b) => b.container !== first.container).every((b) => b.text === '«Next Record»«AddressBlock»'));
  view.setMergePreview({ on: true, record: 1 });
  const shown = view.render({ pages: false }).blocks.filter((b) => b.container && b.text).map((b) => b.text.split('\n')[0]);
  assert.deepEqual(shown.slice(0, 3), ['Mr. Joshua Randall', 'Ms. Cynthia Gartner', 'Dr. Amira Haddad']);
  const merged = openDocx(view.mergeToDocument().bytes);
  const filled = merged.blocks.filter((b) => b.container && b.text).map((b) => b.text.split('\n')[0]);
  assert.deepEqual(filled, ['Mr. Joshua Randall', 'Ms. Cynthia Gartner', 'Dr. Amira Haddad'], 'three labels on the one sheet, the rest left blank');
  assert.equal(renderPdf(merged, {}).pages, 1);
});

test('Start Mail Merge → Envelopes: the document is one envelope, an empty delivery frame for the Address Block, and each recipient an envelope of that size', () => {
  const view = letter();
  view.setEnvelopeDocument({ size: 'C6', returnAddress: ['Rutba Office', 'Leeds'] });
  assert.equal(view.merge.type, 'envelopes');
  assert.equal(Math.round(view.section.widthPx), 612, 'C6 across');
  const frame = view.blocks.find((b) => b.frame);
  assert.equal(frame.text, '');
  assert.equal(view.focus.block, frame.index, 'the caret waits in the frame');
  view.attachMergeSource(sourceFromRows([['First Name', 'Last Name', 'City'], ['Joshua', 'Randall', 'London'], ['Cynthia', 'Gartner', 'Leeds']]));
  view.insertAddressBlock({ company: false, country: 'never' });
  const merged = openDocx(view.mergeToDocument().bytes);
  assert.equal(merged.doc.doc.sections().length, 2);
  const pdf = renderPdf(merged, {});
  assert.equal(pdf.pages, 2);
  assert.ok(mediaBoxes(pdf.buffer).every((b) => Math.round(b[2]) === 459 && Math.round(b[3]) === 323), 'every page a C6 envelope');
  assert.deepEqual(merged.blocks.filter((b) => b.frame).map((b) => b.text), ['Joshua Randall\nLondon', 'Cynthia Gartner\nLeeds']);
});

test('a letter with an envelope saved and opened again still has both, and the envelope is read back', () => {
  const view = letter();
  view.addEnvelope({ size: 'No10', delivery: ADDRESS, returnAddress: ['Rutba Office'] });
  const reopened = openDocx(view.save());
  const e = reopened.doc.doc.envelope();
  assert.equal(e.size, 'No10');
  assert.deepEqual(e.delivery, ADDRESS);
  assert.deepEqual(e.returnAddress, ['Rutba Office']);
  assert.equal(reopened.pages.pages.length, 2);
  assert.equal(reopened.pages.pages[0].header, null, 'no header on the envelope');
  assert.equal(Math.round(reopened.pages.pages[0].section.widthPx), 912, 'No. 10 is 9½ inches across');
});
