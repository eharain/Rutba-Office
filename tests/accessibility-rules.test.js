// Review → Check Accessibility: the rules themselves, apart from any file.
//
// Contrast is WCAG 2.1's, measured the way Office's checker measures it —
// 4.5:1 for text, 3:1 for large text (18 pt, or 14 pt bold). Alternative text
// is written where every format keeps it, on the object's non-visual
// properties, with "decorative" as the extension Office 2019 writes.

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  contrastRatio, requiredRatio, readableOn, normaliseColour,
  checkAccessibility, groupIssues, unclearLinkText, isDefaultSheetName, visualOrder,
  readAltProps, writeAltProps, DECORATIVE_URI,
} from '@rutba/proofing';

test('contrast: WCAG ratios, the large-text bar, and the colour a fix picks', () => {
  assert.equal(Math.round(contrastRatio('#000000', '#FFFFFF')), 21);
  assert.equal(contrastRatio('777777', '#FFFFFF').toFixed(2), '4.48', 'the grey just under the bar');
  assert.ok(contrastRatio('#767676', 'white') >= 4.5, 'and the one just over it');
  assert.equal(requiredRatio(11, false), 4.5);
  assert.equal(requiredRatio(18, false), 3, '18 pt is large');
  assert.equal(requiredRatio(14, true), 3, '14 pt bold is large');
  assert.equal(requiredRatio(14, false), 4.5);
  assert.equal(readableOn('#1F3864'), '#FFFFFF', 'white on dark blue');
  assert.equal(readableOn('#FFF2CC'), '#000000', 'black on pale yellow');
  assert.equal(normaliseColour('darkBlue'), '#00008B', 'a Word highlight name');
  assert.equal(normaliseColour('auto'), null, 'automatic is no colour at all');
});

test('the rules: tiers, grouping, the verdict, and a heading level fix that names the right level', () => {
  const model = {
    objects: [
      { key: 'a', label: 'Picture 1', alt: '', where: { block: 1 }, target: { drawing: 0 } },
      { key: 'b', label: 'Picture 2', alt: 'A bar chart of sales', where: { block: 2 }, target: { drawing: 1 } },
      { key: 'c', label: 'Logo', alt: '', decorative: true, where: { block: 3 }, target: { drawing: 2 } },
    ],
    headings: [
      { key: '1', level: 1, text: 'Report', where: { block: 0 } },
      { key: '2', level: 3, text: 'Detail', where: { block: 4 } },
      { key: '3', level: 2, text: '   ', where: { block: 5 } },
    ],
    texts: [{ key: 't', label: 'Pale', fg: '#BBBBBB', bg: '#FFFFFF', sizePt: 11, where: { block: 6 }, target: { block: 6 } }],
    links: [{ key: 'l', text: 'click here', url: 'https://example.com', where: { block: 7 } }, { key: 'm', text: 'the 2026 price list', url: 'https://example.com/p' }],
    title: '',
  };
  const { issues, counts, verdict } = checkAccessibility(model);
  assert.deepEqual(issues.map((i) => i.rule), ['altText', 'contrast', 'headingOrder', 'emptyHeading', 'linkText', 'titleMissing']);
  assert.deepEqual(counts, { error: 1, warning: 5, tip: 0 });
  assert.equal(verdict, 'investigate');
  const skip = issues.find((i) => i.rule === 'headingOrder');
  assert.deepEqual(skip.fixes[0], { kind: 'headingLevel', label: 'Change to Heading 2', level: 2, target: undefined });
  assert.equal(issues.find((i) => i.rule === 'contrast').fixes[0].colour, '#000000');
  assert.deepEqual(issues.find((i) => i.rule === 'altText').fixes.map((f) => f.kind), ['altText', 'decorative']);
  const grouped = groupIssues(issues);
  assert.deepEqual(grouped.map((g) => [g.label, g.count]), [['Errors', 1], ['Warnings', 5]]);
  assert.equal(grouped[1].rules[0].title, 'Hard-to-read text contrast');
  assert.equal(checkAccessibility({ objects: [{ key: 'x', alt: 'Described', label: 'x' }] }).verdict, 'good');
});

test('unclear links, default sheet names and a slide\'s reading order', () => {
  for (const t of ['click here', 'Here', 'Read more.', 'https://rutba.io/office', 'www.example.com']) assert.ok(unclearLinkText(t), t);
  for (const t of ['the 2026 price list', 'Rutba Office downloads']) assert.ok(!unclearLinkText(t), t);
  assert.ok(isDefaultSheetName('Sheet1') && isDefaultSheetName('Sheet 12'));
  assert.ok(!isDefaultSheetName('Sales 2026'));
  const shapes = [{ id: 'body', x: 40, y: 140, w: 600, h: 300 }, { id: 'title', x: 40, y: 30, w: 600, h: 80 }, { id: 'note', x: 700, y: 150, w: 200, h: 60 }];
  assert.deepEqual(visualOrder(shapes).map((s) => s.id), ['title', 'body', 'note'], 'the title first, then left to right on the same band');
  const flagged = checkAccessibility({ slides: [{ index: 0, title: 'One', where: { slide: 0 }, shapes }] });
  assert.equal(flagged.issues[0].rule, 'readingOrder');
  assert.equal(flagged.issues[0].fixes[0].kind, 'layers');
});

test('alt text is written on the element as a description, or as Office\'s decorative extension', () => {
  const bare = '<p:cNvPr id="4" name="Picture 3"/>';
  const described = writeAltProps(bare, { descr: 'A red "sale" sign' });
  assert.equal(described, '<p:cNvPr id="4" name="Picture 3" descr="A red &quot;sale&quot; sign"/>');
  assert.deepEqual([readAltProps(described).descr, readAltProps(described).decorative], ['A red "sale" sign', false]);

  const withLink = '<wp:docPr id="1" name="Logo" descr="old"><a:hlinkClick r:id="rId9"/></wp:docPr>';
  const decorative = writeAltProps(withLink, { decorative: true });
  assert.ok(!/descr=/.test(decorative), 'a decorative object carries no description');
  assert.ok(decorative.includes('<a:hlinkClick r:id="rId9"/>'), 'the picture\'s link is kept');
  assert.ok(decorative.includes(`uri="${DECORATIVE_URI}"`) && /<adec:decorative [^>]*val="1"\/>/.test(decorative));
  assert.equal(readAltProps(decorative).decorative, true);

  const again = writeAltProps(decorative, { descr: 'Company logo' });
  assert.equal(readAltProps(again).decorative, false, 'a description takes the decorative mark off');
  assert.ok(!/extLst/.test(again), 'and the empty extension list with it');
  assert.equal(again, '<wp:docPr id="1" name="Logo" descr="Company logo"><a:hlinkClick r:id="rId9"/></wp:docPr>');
});

test('the document service answers doc.proof accessibility, and a fix is an op: one undo step, the file marked changed', async () => {
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const { createProofing } = await import('../apps/desktop/main/proofing.js');
  const { gradientPng } = await import('../apps/desktop/main/sample-picture.js');
  const proofing = createProofing({ worker: false });
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }), proofing });
  const { id } = docs.new({ kind: 'doc' });
  docs.apply({ id, ops: [{ op: 'insertImage', name: 'scan.png', contentType: 'image/png', data: gradientPng(10, 10, [0, 0, 0], [255, 255, 255]), widthPx: 40, heightPx: 40 }] });
  const first = await docs.proof({ id, action: 'accessibility' });
  const alt = first.issues.find((i) => i.rule === 'altText');
  assert.ok(alt, 'the picture is found');
  assert.equal(first.rules.altText.title, 'Missing alternative text', 'the pane\'s words ride along');

  // Right-click → Edit Alt Text addresses the picture the page picked: its paragraph and its place there.
  const picked = { block: alt.where.block, image: alt.where.image };
  assert.equal((await docs.proof({ id, action: 'altTextOf', ...picked })).target.drawing, alt.fixes[0].target.drawing);
  const before = docs.meta({ id }).version;
  docs.apply({ id, ops: [{ op: 'setAltText', ...picked, descr: 'A scanned signature' }] });
  assert.ok(docs.meta({ id }).dirty && docs.meta({ id }).version > before);
  assert.equal((await docs.proof({ id, action: 'altTextOf', ...picked })).descr, 'A scanned signature');
  assert.ok(!(await docs.proof({ id, action: 'accessibility' })).issues.some((i) => i.rule === 'altText'));
  docs.undo({ id });
  assert.ok((await docs.proof({ id, action: 'accessibility' })).issues.some((i) => i.rule === 'altText'), 'undo takes it back');
});
