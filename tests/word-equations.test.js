// Equations — reading and drawing Office Math.
//
// An equation in a .docx is an `m:oMathPara` (display: its own line,
// centred) or an `m:oMath` (inline, in the run of the words) among a
// paragraph's runs. The engine reads each as ONE character of the text — a
// run carrying the OMML untouched, so a save writes it back byte for byte —
// and the page draws its MathML. These pin the reading, the MathML each
// construct becomes, the editing around it, the save, and the print path.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';
import { ommlToMathml, ommlToLinear, ommlText, ommlInfo } from '@rutba/ooxml/math';
import { QUADRATIC, QUADRATIC_BODY, display, inline, mr, CTRL, withEquations } from './fixtures/equations.js';

const RICH = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rich');
const MARK = '\uFFFC';

const docWith = (paragraphs, equations) => openDocx(withEquations(buildDocx({ styles: true, paragraphs }), equations));
const documentXml = (bytes) => OoxmlPackage.read(bytes).text('word/document.xml');
const mathOf = (view, block) => view.render({ pages: false }).blocks[block].runs.find((r) => r.math)?.math;

test('the quadratic formula Word writes reads as one display equation, drawn as a fraction over a radical', () => {
  const view = docWith([{ text: 'Before' }, { text: '@Q@' }, { text: 'After' }], { Q: QUADRATIC });
  const block = view.blocks[1];
  assert.equal(block.text, MARK, 'the equation is ONE character of its paragraph');
  assert.equal(block.runs.length, 1);
  assert.equal(block.runs[0].math.display, true);
  assert.equal(block.runs[0].math.xml, QUADRATIC, 'the OMML is carried untouched');
  assert.equal(block.structural, false, 'the paragraph stays editable');

  const math = mathOf(view, 1);
  assert.match(math.mathml, /^<math display="block">/);
  assert.match(math.mathml, /<mfrac><mrow><mo>−<\/mo><mi>b<\/mi><mo>±<\/mo><msqrt>/, 'minus, b, plus-or-minus, then the radical, over the bar');
  assert.match(math.mathml, /<msqrt><mrow><msup><mrow><mi>b<\/mi><\/mrow><mrow><mn>2<\/mn><\/mrow><\/msup><mo>−<\/mo><mn>4<\/mn><mi>a<\/mi><mi>c<\/mi><\/mrow><\/msqrt>/);
  assert.match(math.mathml, /<mrow><mn>2<\/mn><mi>a<\/mi><\/mrow><\/mfrac>/, 'the denominator 2a');
  assert.equal(math.linear, 'x=(-b±√(b^2-4ac))/2a', 'the linear form Word itself shows');
  assert.equal(math.jc, 'centerGroup');
});

test('the showcase document written by Word: its quadratic formula is in the model now, not absent', { skip: !existsSync(join(RICH, 'showcase.docx')) && 'fixture not generated' }, () => {
  const view = openDocx(readFileSync(join(RICH, 'showcase.docx')));
  const frame = view.render({ pages: false });
  const found = frame.blocks.flatMap((b) => b.runs).filter((r) => r.math);
  assert.equal(found.length, 1, 'one equation');
  assert.equal(found[0].math.linear, 'x=(-b±√(b^2-4ac))/2a');
  assert.match(found[0].math.mathml, /<mfrac>[\s\S]*<msqrt>/);
  // Word's own highlight on the equation's runs is drawn too.
  assert.match(found[0].math.mathml, /background:#FFFF00/);
});

test('inline and display: an m:oMath in the words is inline, a character among them', () => {
  const area = inline(mr('π') + `<m:sSup><m:sSupPr>${CTRL}</m:sSupPr><m:e>${mr('r')}</m:e><m:sup>${mr('2')}</m:sup></m:sSup>`);
  const view = docWith([{ text: 'The area is @A@ square units.' }, { text: '@Q@' }], { A: area, Q: QUADRATIC });
  const b = view.blocks[0];
  assert.equal(b.text, `The area is ${MARK} square units.`);
  assert.equal(b.runs.length, 3);
  assert.equal(b.runs[1].math.display, false);
  assert.equal(b.runs[1].math.xml, area);
  const math = mathOf(view, 0);
  assert.match(math.mathml, /^<math display="inline"><mi>π<\/mi><msup><mrow><mi>r<\/mi><\/mrow><mrow><mn>2<\/mn><\/mrow><\/msup><\/math>$/);
  assert.equal(math.display, false);
  assert.equal(mathOf(view, 1).display, true);
  assert.deepEqual(ommlInfo(area), { display: false, jc: null, count: 1 });
});

// Each construct Word writes, and the MathML Chromium draws it with.
const e = (body) => `<m:e>${body}</m:e>`;
const CONSTRUCTS = [
  ['a fraction', `<m:f><m:fPr>${CTRL}</m:fPr><m:num>${mr('a')}</m:num><m:den>${mr('b')}</m:den></m:f>`, /<mfrac><mrow><mi>a<\/mi><\/mrow><mrow><mi>b<\/mi><\/mrow><\/mfrac>/],
  ['a binomial (no bar)', `<m:f><m:fPr><m:type m:val="noBar"/>${CTRL}</m:fPr><m:num>${mr('n')}</m:num><m:den>${mr('k')}</m:den></m:f>`, /<mfrac linethickness="0">/],
  ['a linear fraction', `<m:f><m:fPr><m:type m:val="lin"/>${CTRL}</m:fPr><m:num>${mr('a')}</m:num><m:den>${mr('b')}</m:den></m:f>`, /<mi>a<\/mi><\/mrow><mo>\/<\/mo><mrow><mi>b<\/mi>/],
  ['a cube root (m:deg)', `<m:rad><m:radPr>${CTRL}</m:radPr><m:deg>${mr('3')}</m:deg>${e(mr('x'))}</m:rad>`, /<mroot><mrow><mi>x<\/mi><\/mrow><mrow><mn>3<\/mn><\/mrow><\/mroot>/],
  ['a square root (degHide)', `<m:rad><m:radPr><m:degHide m:val="1"/>${CTRL}</m:radPr><m:deg/>${e(mr('x'))}</m:rad>`, /<msqrt><mrow><mi>x<\/mi><\/mrow><\/msqrt>/],
  ['a subscript', `<m:sSub><m:sSubPr>${CTRL}</m:sSubPr>${e(mr('a'))}<m:sub>${mr('i')}</m:sub></m:sSub>`, /<msub><mrow><mi>a<\/mi><\/mrow><mrow><mi>i<\/mi><\/mrow><\/msub>/],
  ['a sub-superscript', `<m:sSubSup><m:sSubSupPr>${CTRL}</m:sSubSupPr>${e(mr('x'))}<m:sub>${mr('1')}</m:sub><m:sup>${mr('2')}</m:sup></m:sSubSup>`, /<msubsup><mrow><mi>x<\/mi><\/mrow><mrow><mn>1<\/mn><\/mrow><mrow><mn>2<\/mn><\/mrow><\/msubsup>/],
  ['a sum with limits under and over', `<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/>${CTRL}</m:naryPr><m:sub>${mr('i=1')}</m:sub><m:sup>${mr('n')}</m:sup>${e(mr('i'))}</m:nary>`, /<munderover><mo largeop="true" movablelimits="true">∑<\/mo><mrow><mi>i<\/mi><mo lspace="0.05em" rspace="0.05em">=<\/mo><mn>1<\/mn><\/mrow><mrow><mi>n<\/mi><\/mrow><\/munderover><mrow><mi>i<\/mi><\/mrow>/],
  ['an integral (no chr: ∫, limits beside)', `<m:nary><m:naryPr>${CTRL}</m:naryPr><m:sub>${mr('0')}</m:sub><m:sup>${mr('1')}</m:sup>${e(mr('f(x)dx'))}</m:nary>`, /<msubsup><mo largeop="true" movablelimits="false">∫<\/mo><mrow><mn>0<\/mn><\/mrow><mrow><mn>1<\/mn><\/mrow><\/msubsup>/],
  ['a product with the upper limit hidden', `<m:nary><m:naryPr><m:chr m:val="∏"/><m:supHide m:val="1"/>${CTRL}</m:naryPr><m:sub>${mr('k')}</m:sub><m:sup/>${e(mr('a'))}</m:nary>`, /<munder><mo largeop="true" movablelimits="true">∏<\/mo><mrow><mi>k<\/mi><\/mrow><\/munder>/],
  ['brackets with separators', `<m:d><m:dPr><m:begChr m:val="["/><m:sepChr m:val=","/><m:endChr m:val="]"/>${CTRL}</m:dPr>${e(mr('a'))}${e(mr('b'))}</m:d>`, /<mo fence="true" stretchy="true" symmetric="true" form="prefix">\[<\/mo><mrow><mi>a<\/mi><\/mrow><mo separator="true" stretchy="true">,<\/mo><mrow><mi>b<\/mi><\/mrow><mo fence="true" stretchy="true" symmetric="true" form="postfix">]<\/mo>/],
  ['parentheses by default, a missing end left open', `<m:d><m:dPr><m:endChr m:val=""/>${CTRL}</m:dPr>${e(mr('x'))}</m:d>`, /^<math display="inline"><mrow><mo [^>]*form="prefix">\(<\/mo><mrow><mi>x<\/mi><\/mrow><\/mrow><\/math>$/],
  ['a function name, upright', `<m:func><m:funcPr>${CTRL}</m:funcPr><m:fName>${mr('sin', '<m:sty m:val="p"/>')}</m:fName>${e(mr('θ'))}</m:func>`, /<mrow><mrow><mi>sin<\/mi><\/mrow><mspace width="0.1667em"\/><mo>\u2061<\/mo><mrow><mi>θ<\/mi><\/mrow><\/mrow>/],
  ['an accent (hat by default)', `<m:acc><m:accPr>${CTRL}</m:accPr>${e(mr('x'))}</m:acc>`, /<mover accent="true"><mrow><mi>x<\/mi><\/mrow><mo stretchy="true">ˆ<\/mo><\/mover>/],
  ['a vector arrow', `<m:acc><m:accPr><m:chr m:val="\u20D7"/>${CTRL}</m:accPr>${e(mr('v'))}</m:acc>`, /<mo stretchy="true">→<\/mo><\/mover>/],
  ['an overbar', `<m:bar><m:barPr><m:pos m:val="top"/>${CTRL}</m:barPr>${e(mr('z'))}</m:bar>`, /^<math display="inline"><mrow style="border-top:0.06em solid currentColor;padding-top:0.08em"><mi>z<\/mi><\/mrow><\/math>$/],
  ['an underbar (the default position)', `<m:bar><m:barPr>${CTRL}</m:barPr>${e(mr('z'))}</m:bar>`, /<mrow style="border-bottom:0.06em solid currentColor;padding-bottom:0.08em"><mi>z<\/mi><\/mrow>/],
  ['a grouping brace under', `<m:groupChr><m:groupChrPr>${CTRL}</m:groupChrPr>${e(mr('a+b'))}</m:groupChr>`, /<munder><mrow>[\s\S]*<\/mrow><mo stretchy="true">⏟<\/mo><\/munder>/],
  ['a limit under lim', `<m:func><m:funcPr>${CTRL}</m:funcPr><m:fName><m:limLow><m:limLowPr>${CTRL}</m:limLowPr>${e(mr('lim', '<m:sty m:val="p"/>'))}<m:lim>${mr('n→∞')}</m:lim></m:limLow></m:fName>${e(mr('a'))}</m:func>`, /<munder><mrow><mi>lim<\/mi><\/mrow><mrow><mi>n<\/mi><mo lspace="0.05em" rspace="0.05em">→<\/mo><mi mathvariant="normal">∞<\/mi><\/mrow><\/munder>/],
  ['a limit over', `<m:limUpp><m:limUppPr>${CTRL}</m:limUppPr>${e(mr('x'))}<m:lim>${mr('def')}</m:lim></m:limUpp>`, /<mover><mrow><mi>x<\/mi><\/mrow><mrow><mi>d<\/mi><mi>e<\/mi><mi>f<\/mi><\/mrow><\/mover>/],
  ['a matrix', `<m:m><m:mPr>${CTRL}</m:mPr><m:mr>${e(mr('a'))}${e(mr('b'))}</m:mr><m:mr>${e(mr('c'))}${e(mr('d'))}</m:mr></m:m>`, /<mtable><mtr><mtd><mrow><mi>a<\/mi><\/mrow><\/mtd><mtd><mrow><mi>b<\/mi><\/mrow><\/mtd><\/mtr><mtr><mtd><mrow><mi>c<\/mi><\/mrow><\/mtd><mtd><mrow><mi>d<\/mi><\/mrow><\/mtd><\/mtr><\/mtable>/],
  ['an equation array aligned at &', `<m:eqArr><m:eqArrPr>${CTRL}</m:eqArrPr>${e(mr('x&amp;=1'))}${e(mr('y&amp;=2'))}</m:eqArr>`, /<mtable><mtr><mtd style="text-align:right;padding:0.1em 0"><mrow><mi>x<\/mi><\/mrow><\/mtd><mtd style="text-align:left;padding:0.1em 0"><mrow><mo>=<\/mo><mn>1<\/mn><\/mrow><\/mtd><\/mtr>/],
  ['a box', `<m:box><m:boxPr>${CTRL}</m:boxPr>${e(mr('a'))}</m:box>`, /^<math display="inline"><mrow><mi>a<\/mi><\/mrow><\/math>$/],
  ['a border box', `<m:borderBox><m:borderBoxPr>${CTRL}</m:borderBoxPr>${e(mr('E=mc'))}</m:borderBox>`, /<mrow style="border-top:0.06em solid currentColor;border-bottom:0.06em solid currentColor;border-left:0.06em solid currentColor;border-right:0.06em solid currentColor;/],
  ['scripts before (m:sPre)', `<m:sPre><m:sPrePr>${CTRL}</m:sPrePr><m:sub>${mr('6')}</m:sub><m:sup>${mr('14')}</m:sup>${e(mr('C'))}</m:sPre>`, /<mmultiscripts><mrow><mi>C<\/mi><\/mrow><mprescripts\/><mrow><mn>6<\/mn><\/mrow><mrow><mn>14<\/mn><\/mrow><\/mmultiscripts>/],
  ['plain style (m:sty p) keeps the letters together, upright', mr('max', '<m:sty m:val="p"/>'), /<mi>max<\/mi>/],
  ['bold (m:sty b)', mr('v', '<m:sty m:val="b"/>'), /<mi mathvariant="normal" style="font-weight:bold">v<\/mi>/],
  ['double-struck (m:scr)', mr('R', '<m:scr m:val="double-struck"/>'), /<mi mathvariant="normal">ℝ<\/mi>/],
  ['script (m:scr)', mr('L', '<m:scr m:val="script"/>'), /<mi mathvariant="normal">ℒ<\/mi>/],
  ['normal text (m:nor)', mr('if x > 0', '<m:nor/>'), /<mtext>if x &gt; 0<\/mtext>/],
  ['a capital Greek letter is upright, a small one italic', mr('Δθ'), /<mi mathvariant="normal">Δ<\/mi><mi>θ<\/mi>/],
  ['a decimal number is one mn; a hyphen is a minus', mr('3.14-x'), /<mn>3\.14<\/mn><mo>−<\/mo><mi>x<\/mi>/],
];
for (const [name, body, expected] of CONSTRUCTS) {
  test(`OMML → MathML: ${name}`, () => {
    assert.match(ommlToMathml(inline(body)), expected);
  });
}

test('an element this reader does not know is drawn as its words, never dropped', () => {
  const xml = inline(mr('a+') + `<m:weird><m:weirdPr/><m:e>${mr('xyz')}</m:e></m:weird>`);
  const mathml = ommlToMathml(xml);
  assert.match(mathml, /<mi>x<\/mi><mi>y<\/mi><mi>z<\/mi>/);
  assert.equal(ommlText(xml), 'a+xyz');
  assert.equal(ommlToLinear(xml), 'a+xyz');
  // An empty equation still draws something to click.
  assert.match(ommlToMathml('<m:oMath></m:oMath>'), /⬚/);
});

test('saving writes every equation back byte for byte — untouched, and when its paragraph was edited', () => {
  const area = inline(mr('π') + `<m:sSup><m:sSupPr>${CTRL}</m:sSupPr><m:e>${mr('r')}</m:e><m:sup>${mr('2')}</m:sup></m:sSup>`);
  const bytes = withEquations(buildDocx({ styles: true, paragraphs: [{ text: 'The area is @A@ square units.' }, { text: '@Q@' }, { text: 'Last' }] }), { A: area, Q: QUADRATIC });
  const before = documentXml(bytes);

  // Untouched: the same bytes.
  assert.equal(documentXml(openDocx(bytes).save()), before);

  // Edited around: typing after the inline equation and after the display
  // one rebuilds both paragraphs, and both equations come back exactly —
  // each in its place among the words, not moved to the paragraph's end.
  const view = openDocx(bytes);
  view.collapseTo({ block: 0, offset: view.blocks[0].text.length });
  view.insertText('!');
  view.collapseTo({ block: 1, offset: 1 });
  view.insertText(' (1)');
  const after = documentXml(view.save());
  assert.ok(after.includes(area), 'the inline equation, byte for byte');
  assert.ok(after.includes(QUADRATIC), 'the display equation, byte for byte');
  assert.ok(after.indexOf('The area is') < after.indexOf(area) && after.indexOf(area) < after.indexOf(' square units.!'), 'the inline equation stays between its words');
  assert.equal(after.split('<m:oMath>').length - 1, 2, 'each written once — not also kept as a leftover fragment');
  const reread = openDocx(view.save());
  assert.equal(reread.blocks[0].text, `The area is ${MARK} square units.!`);
  assert.equal(reread.blocks[1].text, `${MARK} (1)`);
});

test('the caret steps over an equation as one character; typing beside it never goes inside', () => {
  const view = docWith([{ text: 'a @A@ b' }], { A: inline(mr('x')) });
  view.collapseTo({ block: 0, offset: 2 });
  view.insertText('<');
  view.moveCaret('right');
  view.insertText('>');
  assert.equal(view.blocks[0].text, `a <${MARK}> b`);
  const runs = view.blocks[0].runs;
  assert.ok(runs.find((r) => r.math).math.xml === inline(mr('x')), 'the equation untouched');
});

test('Backspace after an equation removes it whole, and Undo puts back the same XML', () => {
  const bytes = withEquations(buildDocx({ styles: true, paragraphs: [{ text: 'Before' }, { text: '@Q@' }, { text: 'x @A@ y' }] }), { Q: QUADRATIC, A: inline(mr('a')) });
  const view = openDocx(bytes);
  view.collapseTo({ block: 1, offset: 1 });
  view.deleteBackward();
  assert.equal(view.blocks[1].text, '');
  assert.ok(!documentXml(view.serialize()).includes('<m:oMathPara>'), 'gone from the file');
  view.undo();
  assert.equal(view.blocks[1].runs[0].math.xml, QUADRATIC, 'back, exactly');

  // A selection across an inline equation deletes it with the words.
  view.setSelection({ block: 2, offset: 1 }, { block: 2, offset: 4 });
  view.deleteSelection();
  assert.equal(view.blocks[2].text, 'xy');
  view.undo();
  assert.equal(view.blocks[2].text, `x ${MARK} y`);
  // Delete in front of it takes it too.
  view.collapseTo({ block: 2, offset: 2 });
  view.deleteForward();
  assert.equal(view.blocks[2].text, 'x  y');
  view.undo();
  assert.equal(documentXml(view.save()), documentXml(bytes), 'after the undos, the file is the file');
});

test('pagination: a display equation is a block of its own height — measured when the desktop measured it', () => {
  const lines = Array.from({ length: 25 }, (_, i) => ({ text: `Line ${i + 1} of the page.` }));
  const view = docWith([...lines, { text: '@Q@' }, { text: 'After the equation.' }], { Q: QUADRATIC });
  const eqIndex = lines.length;
  const fragments = () => view.pages.pages.map((p) => p.fragments);
  const pageOf = (index, kind) => fragments().findIndex((f) => f.some((x) => x.paragraphIndex === index && (!kind || x.kind === kind)));

  // With no MathML layout to hand the equation is its linear form, one line tall.
  const plain = fragments().flat().find((f) => f.kind === 'equation');
  assert.ok(plain, 'placed as an equation fragment');
  assert.equal(plain.box.text, 'x=(-b±sqrt(b^2-4ac))/2a');
  assert.equal(plain.align, 'center');
  const plainPage = pageOf(eqIndex, 'equation');
  assert.equal(plainPage, 0, 'one line of it fits on the first page');

  // Measured: three and a half ems tall, as Chromium lays the fraction out —
  // the paginator leaves that much room, and the words after it move down.
  const xml = view.blocks[eqIndex].runs[0].math.xml;
  view.setMathMeasures(new Map([[xml, { widthEm: 9, heightEm: 3.5, baselineEm: 2.2 }]]));
  const measured = fragments().flat().find((f) => f.kind === 'equation');
  const sizePx = measured.sizePx;
  assert.ok(Math.abs(measured.box.heightPx - 3.5 * sizePx) < 0.01, 'the measured height');
  assert.ok(Math.abs(measured.box.widthPx - 9 * sizePx) < 0.01, 'the measured width');
  assert.ok(measured.box.heightPx > plain.box.heightPx * 2);

  // Tall enough, it goes whole onto the next page rather than being cut.
  view.setMathMeasures(new Map([[xml, { widthEm: 9, heightEm: 50, baselineEm: 25 }]]));
  assert.equal(pageOf(eqIndex, 'equation'), plainPage + 1, 'pushed to the next page, whole');
  assert.equal(pageOf(eqIndex + 1), plainPage + 1, 'and the words after it follow');
});

test('pagination: an inline equation is a word as wide as it draws, and its lines are as tall as it', () => {
  const view = docWith([{ text: 'The area is @A@ square units, which is what the formula says.' }], { A: inline(mr('πr')) });
  const xml = view.blocks[0].runs[1].math.xml;
  const plain = view.pages.pages[0].fragments.find((f) => f.paragraphIndex === 0);
  view.setMathMeasures(new Map([[xml, { widthEm: 30, heightEm: 2.4, baselineEm: 1.6 }]]));
  const measured = view.pages.pages[0].fragments.filter((f) => f.paragraphIndex === 0);
  assert.equal(measured[0].kind, 'paragraph');
  assert.ok(measured[0].lineHeightPx > plain.lineHeightPx, 'the line grows to hold the equation');
  assert.ok(measured[0].lines.length > plain.lines.length, 'thirty ems of equation push words onto the next line');
  assert.equal(measured[0].lines[0].start, 0);
  assert.equal(measured[measured.length - 1].lines.at(-1).end, view.blocks[0].text.length, 'the lines still cover the whole paragraph');
});

test('print: the PDF places the desktop\'s picture of the equation, or its linear form where there is none', () => {
  const view = docWith([{ text: '@Q@' }, { text: 'Area @A@ here.' }], { Q: QUADRATIC, A: inline(mr('πr')) });
  const plain = renderPdf(view, { created: '2026-09-25T00:00:00Z' }).buffer.toString('latin1');
  // (± is WinAnsi 0xB1, which the writer spells as the octal escape \261.)
  assert.match(plain, /\(x=\\\(-b\\261sqrt\\\(b\^2-4ac\\\)\\\)\/2a\) Tj/, 'the linear form, spelt in the PDF font\'s own characters');
  assert.ok(!plain.includes('/Subtype /Image'));

  const white = (w, h) => ({ width: w, height: h, data: Buffer.alloc(w * h * 3, 255) });
  view.setMathMeasures(new Map([
    [view.blocks[0].runs[0].math.xml, { widthEm: 9, heightEm: 3.5, baselineEm: 2.2, raster: white(30, 12) }],
    [view.blocks[1].runs[1].math.xml, { widthEm: 1.4, heightEm: 1.1, baselineEm: 0.8, raster: white(6, 5) }],
  ]));
  const pictured = renderPdf(view, { created: '2026-09-25T00:00:00Z' }).buffer.toString('latin1');
  assert.equal(pictured.split('/Subtype /Image').length - 1, 2, 'both equations drawn from their pictures');
  assert.ok(!pictured.includes('sqrt'), 'no linear fallback once the pictures are there');
});
