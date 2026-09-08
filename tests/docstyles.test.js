/**
 * Styles, numbering and pictures — the document reading its OWN definitions.
 *
 * The bug class these guard against is quiet mislabelling: a style that
 * resolves to the wrong size is not just cosmetic, it moves every page break
 * after it; a list that miscounts reads as a drafting error in a contract; a
 * picture that silently vanishes is a letterhead without the company on it.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readParagraphStyles, readNumberingDefs } from '@rutba/ooxml';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { computeListLabels, formatCounter } from '@rutba/doc-view/lists';
import { paginate } from '@rutba/doc-view/paginate';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildComplexDocument } from './fixtures/complex-document.js';

const LETTER = buildComplexDocument();

// ------------------------------------------------------------------ styles --

test('a style chain flattens through basedOn down to docDefaults', () => {
  const styles = readParagraphStyles(`
    <w:styles>
      <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/>
        <w:rPr><w:color w:val="333333"/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>
        <w:basedOn w:val="Base"/><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
    </w:styles>`);

  const h2 = styles.Heading2;
  assert.equal(h2.bold, true, 'its own bold');
  assert.equal(h2.sizePx, 13 * 96 / 72, 'its own size: 26 half-points is 13pt');
  assert.equal(h2.colour, '#333333', 'inherited from Base');
  // Base itself has no size, so it inherits the document default: 11pt.
  assert.equal(styles.Base.sizePx, 11 * 96 / 72);
});

test('an unstyled paragraph still obeys docDefaults', () => {
  // An 11pt house document must not render at our hardcoded fallback size.
  const styles = readParagraphStyles(
    '<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>');
  assert.equal(styles['*default*'].sizePx, 12 * 96 / 72);
});

test('a basedOn cycle resolves to defaults instead of recursing forever', () => {
  const styles = readParagraphStyles(`
    <w:styles>
      <w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:basedOn w:val="B"/><w:rPr><w:b/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/><w:basedOn w:val="A"/><w:rPr><w:i/></w:rPr></w:style>
    </w:styles>`);
  assert.equal(styles.A.bold, true);
  assert.equal(styles.B.italic, true);
});

test('a resolved style moves the page breaks, not just the paint', () => {
  // The same forty paragraphs, once as body text and once as a style that
  // resolves twice the size, must not break in the same places.
  const blocks = Array.from({ length: 40 }, (_, i) => ({
    index: i, text: 'Paragraph ' + i + '. ' + 'word '.repeat(24), style: 'Big',
    runs: [], structural: false, structuralTags: [],
  }));
  const flow = blocks.map((b) => ({ kind: 'paragraph', paragraphIndex: b.index }));
  const section = {
    widthPx: 794, heightPx: 1123, contentWidthPx: 602,
    margins: { top: 96, right: 96, bottom: 96, left: 96, header: 47, footer: 47, gutter: 0 },
    titlePage: false, evenAndOdd: false, declared: true,
  };
  const small = paginate({ flow, blocks, section, styles: { Big: { sizePx: 14 } } });
  const large = paginate({ flow, blocks, section, styles: { Big: { sizePx: 28, bold: true } } });
  assert.ok(large.count > small.count,
    'twice the font size must need more sheets: ' + small.count + ' vs ' + large.count);
  const fragment = large.pages[0].fragments[0];
  assert.equal(fragment.sizePx, 28);
  assert.equal(fragment.weight, 'bold');
});

// --------------------------------------------------------------- numbering --

test('counters format the way Word formats them', () => {
  assert.equal(formatCounter('decimal', 3), '3');
  assert.equal(formatCounter('lowerLetter', 1), 'a');
  assert.equal(formatCounter('lowerLetter', 27), 'aa', 'spreadsheet-column style past z');
  assert.equal(formatCounter('upperRoman', 4), 'IV');
  assert.equal(formatCounter('lowerRoman', 1998), 'mcmxcviii');
  assert.equal(formatCounter('bullet', 9), '•');
});

test('a deeper level restarts when its parent increments', () => {
  const defs = {
    5: [
      { format: 'decimal', lvlText: '%1.', start: 1, indentPx: 24 },
      { format: 'decimal', lvlText: '%1.%2.', start: 1, indentPx: 48 },
    ],
  };
  const items = [
    [0, '1.'], [1, '1.1.'], [1, '1.2.'], [0, '2.'], [1, '2.1.'],
  ];
  const blocks = items.map(([level], i) => ({ index: i, numbering: { numId: '5', level } }));
  const flow = blocks.map((b) => ({ kind: 'paragraph', paragraphIndex: b.index }));

  const labels = computeListLabels(flow, blocks, defs);
  for (const [i, [, expected]] of items.entries()) {
    assert.equal(labels.get(i).label, expected, 'item ' + i);
  }
});

test('two lists never share a counter, and bullets never count', () => {
  const defs = {
    1: [{ format: 'decimal', lvlText: '%1.', start: 1, indentPx: 24 }],
    2: [{ format: 'decimal', lvlText: '%1)', start: 4, indentPx: 24 }],
    3: [{ format: 'bullet', lvlText: '', start: 1, indentPx: 24 }],
  };
  const blocks = [
    { index: 0, numbering: { numId: '1', level: 0 } },
    { index: 1, numbering: { numId: '2', level: 0 } },   // its own list, its own start
    { index: 2, numbering: { numId: '3', level: 0 } },   // a bullet in between
    { index: 3, numbering: { numId: '1', level: 0 } },   // list 1 continues
    { index: 4, numbering: { numId: '2', level: 0 } },
  ];
  const flow = blocks.map((b) => ({ kind: 'paragraph', paragraphIndex: b.index }));
  const labels = computeListLabels(flow, blocks, defs);
  assert.equal(labels.get(0).label, '1.');
  assert.equal(labels.get(1).label, '4)', 'w:start is honoured');
  assert.equal(labels.get(2).label, '•');
  assert.equal(labels.get(3).label, '2.', 'the bullet did not consume a number');
  assert.equal(labels.get(4).label, '5)');
});

test('deleting an item renumbers the ones after it', () => {
  const view = openDocx(LETTER);
  const before = view.render();
  const listFragments = (frame) => frame.pages.pages.flatMap((p) => p.fragments)
    .filter((f) => f.listLabel).map((f) => [f.listLabel, f.lines[0].text]);

  assert.deepEqual(listFragments(before).map(([l]) => l), ['1.', '2.']);

  // Select the whole first list item and its trailing break, then delete.
  const first = before.blocks.findIndex((b) => b.text === 'Invoices raised in the period');
  view.setSelection({ block: first, offset: 0 }, { block: first + 1, offset: 0 });
  view.deleteSelection();

  const after = listFragments(view.render());
  assert.equal(after.length, 1);
  assert.equal(after[0][0], '1.', 'the survivor is renumbered, not left as 2.');
  assert.match(after[0][1], /Payments received/);
});

// ---------------------------------------------------------------- pictures --

test('an inline picture reaches the page as a data URI at its stated size', () => {
  const frame = openDocx(LETTER).render();
  const images = frame.pages.pages.flatMap((p) => p.fragments).filter((f) => f.kind === 'images');
  assert.equal(images.length, 1);
  const [logo] = images[0].images;
  assert.equal(logo.name, 'Company logo');
  assert.match(logo.href, /^data:image\/png;base64,/);
  // 914400x457200 EMU is 96x48 px.
  assert.equal(Math.round(logo.widthPx), 96);
  assert.equal(Math.round(logo.heightPx), 48);
});

test('a picture wider than the column scales down, keeping its shape', () => {
  const blocks = [{
    index: 0, text: '', style: null, runs: [], structural: false, structuralTags: [],
    images: [{ href: 'data:image/png;base64,AA==', widthPx: 1200, heightPx: 600, name: 'wide' }],
  }];
  const laid = paginate({
    flow: [{ kind: 'paragraph', paragraphIndex: 0 }], blocks,
    section: {
      widthPx: 794, heightPx: 1123, contentWidthPx: 602,
      margins: { top: 96, right: 96, bottom: 96, left: 96, header: 47, footer: 47, gutter: 0 },
      titlePage: false, evenAndOdd: false, declared: true,
    },
  });
  const [img] = laid.pages[0].fragments.find((f) => f.kind === 'images').images;
  assert.equal(Math.round(img.widthPx), 602);
  assert.equal(Math.round(img.heightPx), 301, 'the aspect ratio survived');
});

test('editing the text around a picture leaves the picture untouched', () => {
  const view = openDocx(LETTER);
  const target = view.render().blocks.findIndex((b) => !b.structural && b.text.length > 5);
  view.setSelection({ block: target, offset: 0 });
  view.insertText('Amended: ');
  const images = view.render().pages.pages.flatMap((p) => p.fragments).filter((f) => f.kind === 'images');
  assert.equal(images.length, 1, 'still there, still drawn');
  // and the file's media part is untouched on save — the round-trip suite
  // covers the bytes; here we only care that layout did not eat it.
});

test('a run that says colour auto is black, whatever colour its style says', () => {
  const styles =
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:color w:val="2E74B5" w:themeColor="accent1"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="plain"/><w:rPr><w:color w:val="auto"/></w:rPr></w:style>' +
    '</w:styles>';
  const read = readParagraphStyles(styles);
  assert.equal(read.Heading1.colour, '#2e74b5');
  assert.equal(read.Plain.colour, '#000000', 'auto is a colour — the default one — not the absence of one');
});

test('a numbering level remembers the font its bullet character was stored for', () => {
  const numbering =
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>' +
    '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
  const defs = readNumberingDefs(numbering);
  assert.equal(defs['1'][0].font, 'Wingdings');
  assert.equal(defs['1'][0].lvlText, '');
});

test('the page honours colour auto on a run, contextual spacing between list items, and the hanging marker', () => {
  const pkg = OoxmlPackage.read(LETTER);
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  pkg.write_('word/styles.xml',
    `<w:styles ${W}>` +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:color w:val="2E74B5"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
    '</w:styles>');
  pkg.write_('word/numbering.xml',
    `<w:numbering ${W}>` +
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>');
  const item = (t) => `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
  pkg.write_('word/document.xml',
    `<w:document ${W}><w:body>` +
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>Products</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Intro</w:t></w:r></w:p>' +
    item('One') + item('Two') + item('Three') +
    '<w:p><w:r><w:t>After</w:t></w:r></w:p>' +
    '<w:sectPr/></w:body></w:document>');
  const view = openDocx(pkg.write());
  const { blocks, listLabels } = view.render({ pages: false });
  assert.equal(blocks[0].runs[0].fontColour, '000000', 'auto beats the heading style\'s blue');
  assert.equal(blocks[2].spaceAfterPx, 0, 'no gap after the first item');
  assert.equal(blocks[3].spaceBeforePx, 0);
  assert.equal(blocks[3].spaceAfterPx, 0, 'nor around the middle one');
  assert.equal(blocks[4].spaceBeforePx, 0);
  assert.equal(blocks[4].spaceAfterPx, null, 'the last item keeps its gap from the prose after the list');
  assert.equal(blocks[1].spaceAfterPx, null, 'and the prose before it keeps its own');
  const label = listLabels.get ? listLabels.get(2) : listLabels[2];
  assert.equal(label.label, '•', 'Symbol U+F0B7 is a round bullet');
  assert.equal(label.indentPx, 48);
  assert.equal(label.hangingPx, 24, 'the marker hangs a quarter inch to the left of the text');
});
