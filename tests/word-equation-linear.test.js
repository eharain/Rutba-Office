// Equations — typing them in Word's linear format.
//
// Insert → Equation (Alt+=) takes a line of UnicodeMath, as Word's editor
// does — `x=(-b±√(b^2-4ac))/2a`, `\sum_(i=1)^n i`, `\int_0^1 f(x)dx` — and
// builds it up into OMML written the way Word writes it. An equation opened
// again comes back as its linear form. These pin the parser, the OMML's
// shape, the round trip, the symbols, the errors, and the engine's insert,
// replace and paste.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';
import { ommlToLinear, ommlToMathml } from '@rutba/ooxml/math';
import { linearToOmml, parseLinear, EQUATION_GALLERY, KEYWORDS } from '@rutba/ooxml/math-linear';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { QUADRATIC, CTRL, mr } from './fixtures/equations.js';

const MARK = '\uFFFC';
const build = (linear, display = false) => {
  const r = linearToOmml(linear, { display });
  assert.ok(r.ok, `${linear}: ${r.error}`);
  return r.xml;
};
const documentXml = (bytes) => OoxmlPackage.read(bytes).text('word/document.xml');

test('the quadratic formula typed in the linear form builds up to exactly the OMML Word writes', () => {
  assert.equal(build('x=(-b±√(b^2-4ac))/2a', true), QUADRATIC);
  // And Word's OMML reads back as the line Word shows.
  assert.equal(ommlToLinear(QUADRATIC), 'x=(-b±√(b^2-4ac))/2a');
});

test('a display equation is m:oMathPara; one in the words is a bare m:oMath; every run is Cambria Math', () => {
  const display = build('x^2', true);
  const inline = build('x^2', false);
  assert.match(display, /^<m:oMathPara><m:oMath>[\s\S]*<\/m:oMath><\/m:oMathPara>$/);
  assert.match(inline, /^<m:oMath><m:sSup>/);
  assert.equal(inline, `<m:oMath><m:sSup><m:sSupPr>${CTRL}</m:sSupPr><m:e>${mr('x')}</m:e><m:sup>${mr('2')}</m:sup></m:sSup></m:oMath>`);
});

test('large operators: a sum takes its limits over and under, its body the next term; an integral takes them beside', () => {
  const sum = build('\\sum_(i=1)^n i=n(n+1)/2');
  assert.match(sum, /<m:nary><m:naryPr><m:chr m:val="∑"\/><m:limLoc m:val="undOvr"\/>/);
  assert.match(sum, /<m:sub><m:r>[\s\S]*<m:t>i=1<\/m:t><\/m:r><\/m:sub><m:sup><m:r>[\s\S]*<m:t>n<\/m:t><\/m:r><\/m:sup><m:e><m:r>[\s\S]*<m:t>i<\/m:t><\/m:r><\/m:e><\/m:nary>/);
  assert.match(sum, /<\/m:nary><m:r>[\s\S]*<m:t>=<\/m:t><\/m:r><m:f>/, 'the body ends at =, and the fraction follows');
  const integral = build('\\int_0^1 f(x)dx');
  assert.match(integral, /<m:naryPr><m:limLoc m:val="subSup"\/>/, 'no chr: ∫ is the default');
  assert.match(integral, /<m:e><m:r>[\s\S]*<m:t>f<\/m:t><\/m:r><m:d>[\s\S]*<m:t>x<\/m:t>[\s\S]*<\/m:d><m:r>[\s\S]*<m:t>dx<\/m:t><\/m:r><\/m:e><\/m:nary>/);
});

test('scripts, radicals and fractions: a_i, x^2, \\sqrt(x), √(3&x), (a+b)/2 and (a/b)^2', () => {
  assert.match(build('a_i'), /<m:sSub><m:sSubPr>[\s\S]*<m:e>[\s\S]*<m:t>a<\/m:t>[\s\S]*<m:sub>[\s\S]*<m:t>i<\/m:t>/);
  assert.match(build('\\sqrt(x)'), /<m:rad><m:radPr><m:degHide m:val="1"\/>[\s\S]*<m:deg\/><m:e>[\s\S]*<m:t>x<\/m:t>/);
  assert.match(build('√(3&x)'), /<m:rad><m:radPr><m:ctrlPr>[\s\S]*<m:deg><m:r>[\s\S]*<m:t>3<\/m:t>[\s\S]*<\/m:deg><m:e>/);
  assert.match(build('(a+b)/2'), /<m:f><m:fPr>[\s\S]*<m:num><m:r>[\s\S]*<m:t>a\+b<\/m:t><\/m:r><\/m:num>/, 'the brackets round a numerator are dropped');
  assert.match(build('(a/b)^2'), /<m:sSup>[\s\S]*<m:e><m:d>[\s\S]*<m:f>/, 'a script base keeps its brackets');
  assert.match(build('n¦k'), /<m:type m:val="noBar"\/>/);
  assert.match(build('1/2(α±β)'), /<m:den><m:r>[\s\S]*<m:t>2<\/m:t><\/m:r><\/m:den><\/m:f><m:d>/, 'a denominator ends at a bracket');
});

test('functions are upright, a limit sits under lim, and matrices take & between cells and @ between rows', () => {
  assert.match(build('sin θ'), /<m:func><m:funcPr>[\s\S]*<m:fName><m:r><m:rPr><m:sty m:val="p"\/><\/m:rPr>[\s\S]*<m:t>sin<\/m:t><\/m:r><\/m:fName><m:e>[\s\S]*<m:t>θ<\/m:t>/);
  assert.match(build('lim_(n→∞) a_n'), /<m:fName><m:limLow>[\s\S]*<m:t>lim<\/m:t>[\s\S]*<m:lim>[\s\S]*<m:t>n→∞<\/m:t>[\s\S]*<\/m:limLow><\/m:fName><m:e><m:sSub>/);
  const m = build('(a&b@c&d)');
  assert.match(m, /<m:d><m:dPr>[\s\S]*<m:e><m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"\/><m:mcJc m:val="center"\/><\/m:mcPr><\/m:mc><\/m:mcs>/);
  assert.equal((m.match(/<m:mr>/g) || []).length, 2);
  assert.match(build('■(1&0@0&1)'), /^<m:oMath><m:m>/);
  assert.match(build('█(x&=1@y&=2)'), /<m:eqArr>[\s\S]*<m:t>x&amp;=1<\/m:t>/, 'the alignment mark stays in the text, as Word keeps it');
});

test('keywords are symbols: \\alpha, \\pm, \\infty, \\rightarrow, \\le — and accents follow what they sit on', () => {
  assert.match(build('\\alpha+\\beta'), /<m:t>α\+β<\/m:t>/);
  assert.match(build('a\\pm b\\le\\infty'), /<m:t>a±b≤∞<\/m:t>/);
  assert.match(build('x\\rightarrow 0'), /<m:t>x→0<\/m:t>/);
  assert.match(build('x\\hat'), /<m:acc><m:accPr><m:chr m:val="\u0302"\/>/);
  assert.match(build('v\\vec'), /<m:chr m:val="\u20D7"\/>/);
  assert.match(build('¯(AB)'), /<m:bar><m:barPr><m:pos m:val="top"\/>/);
  assert.ok(Object.keys(KEYWORDS).length > 150, 'the keyword table');
  assert.match(build('"if " x>0'), /<m:rPr><m:nor\/><\/m:rPr>[\s\S]*<m:t xml:space="preserve">if <\/m:t>/, 'quoted words are normal text');
});

test('a mistake in the linear form is reported, with where, never thrown', () => {
  for (const [input, expected, at] of [
    ['\\foo+1', /\\foo is not a keyword/, 0],
    ['x+(a+b', /not closed/, 2],
    ['■(a&b', /matrix opened here is not closed/, 1],
    ['"half', /quotes is not closed/, 0],
    ['   ', /Type an equation/, 0],
  ]) {
    let result;
    assert.doesNotThrow(() => { result = linearToOmml(input); });
    assert.equal(result.ok, false, input);
    assert.match(result.error, expected);
    assert.equal(result.at, at, `${input}: where`);
  }
  assert.equal(parseLinear(null).ok, false);
});

test('linear → OMML → linear → OMML is stable, for Word\'s gallery and the editor\'s own templates', () => {
  const cases = [
    ...EQUATION_GALLERY.map((g) => g.linear),
    '\\sum_(i=1)^n i', '\\int_0^1 f(x)dx', 'a_i', 'x^2', 'sin(x+1)', 'log_2 x', 'e^-x', '_6^14 C', '∑_k a_k b_k', 'x\\hat', '√(3&x)',
  ];
  for (const linear of cases) {
    const once = build(linear, true);
    const back = ommlToLinear(once);
    assert.equal(build(back, true), once, `${linear} → ${back}`);
  }
  // Word's own quadratic reads back to the line it was typed as.
  assert.equal(ommlToLinear(build('x=(-b±√(b^2-4ac))/2a', true)), 'x=(-b±√(b^2-4ac))/2a');
});

test('the built-in gallery holds Word\'s equations, each one building and drawing', () => {
  const names = EQUATION_GALLERY.map((g) => g.name);
  for (const name of ['Quadratic Formula', 'Area of Circle', 'Binomial Theorem', 'Pythagorean Theorem', 'Fourier Series']) assert.ok(names.includes(name), name);
  for (const g of EQUATION_GALLERY) {
    const xml = build(g.linear, true);
    assert.match(ommlToMathml(xml), /^<math display="block">/, g.name);
  }
  assert.match(ommlToMathml(build(EQUATION_GALLERY.find((g) => g.name === 'Binomial Theorem').linear, true)), /<mfrac linethickness="0">/, 'the binomial coefficient');
});

test('the engine inserts an equation at the caret — display in an empty paragraph, inline in words — declares xmlns:m, and Undo takes it out', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'The area is  here.' }, { text: '' }] }));
  assert.ok(!documentXml(view.serialize()).includes('xmlns:m='), 'no math namespace to begin with');
  view.collapseTo({ block: 0, offset: 12 });
  view.insertEquation({ xml: build('πr^2', false) });
  assert.equal(view.blocks[0].text, `The area is ${MARK} here.`);
  assert.deepEqual(view.focus, { block: 0, offset: 13 }, 'the caret after it');
  view.collapseTo({ block: 1, offset: 0 });
  view.insertEquation({ xml: build('x=(-b±√(b^2-4ac))/2a', true) });
  const saved = documentXml(view.save());
  assert.match(saved, /<w:document xmlns:m="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math"/);
  assert.ok(saved.includes(QUADRATIC), 'the display equation, as Word writes it');
  assert.match(saved, /<w:t xml:space="preserve">The area is <\/w:t><\/w:r><m:oMath>/);
  view.undo();
  assert.equal(view.blocks[1].text, '');
  view.undo();
  assert.equal(view.blocks[0].text, 'The area is  here.');
});

test('replacing an equation opened in the editor, and a paste from within the suite that carries one', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: '' }, { text: 'After' }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertEquation({ xml: build('x=(-b±√(b^2-4ac))/2a', true) });
  view.replaceEquation({ block: 0, offset: 0, xml: build('x=(-b±√(b^2-4ac))/2 a', true) });
  const run = view.blocks[0].runs[0];
  assert.equal(ommlToLinear(run.math.xml), 'x=(-b±√(b^2-4ac))/2 a', 'the denominator is 2 now, and a follows it');
  assert.deepEqual([view.selection.from.offset, view.selection.to.offset], [0, 1], 'left selected, as Word leaves it');
  view.undo();
  assert.equal(ommlToLinear(view.blocks[0].runs[0].math.xml), 'x=(-b±√(b^2-4ac))/2a');
  assert.throws(() => view.replaceEquation({ block: 1, offset: 0, xml: build('x', true) }), /no equation there/);

  // A paste of "Area: \uFFFC" and a second line puts the equation back as itself.
  view.collapseTo({ block: 1, offset: 5 });
  view.pasteRuns([[{ text: ' is ' }, { math: { xml: build('πr^2', false) } }], [{ text: 'next line' }]]);
  assert.equal(view.blocks[1].text, `After is ${MARK}`);
  assert.equal(view.blocks[2].text, 'next line');
  assert.equal(ommlToLinear(view.blocks[1].runs.find((r) => r.math).math.xml), 'πr^2');
  view.undo();
  assert.equal(view.blocks[1].text, 'After', 'one undo for the whole paste');
});

test('the Word window\'s ops: insertEquation from the linear form, replaceEquation, and a sentence for a mistake', () => {
  const service = createDocumentService({ holdBlob: () => ({}) });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-eq-'));
  const file = path.join(dir, 'eq.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: '' }] }));
  const session = service.open({ path: file, kind: 'doc' });
  service.apply({ id: session.id, ops: [{ op: 'insertEquation', linear: 'a^2+b^2=c^2', display: true }] });
  let model = service.model({ id: session.id });
  const math = model.blocks[0].runs[0].math;
  assert.equal(math.display, true);
  assert.equal(math.linear, 'a^2+b^2=c^2');
  assert.match(math.mathml, /<msup>/);
  service.apply({ id: session.id, ops: [{ op: 'replaceEquation', block: 0, offset: 0, linear: 'a^2+b^2=c^3', display: true }] });
  model = service.model({ id: session.id });
  assert.equal(model.blocks[0].runs[0].math.linear, 'a^2+b^2=c^3');
  assert.throws(() => service.apply({ id: session.id, ops: [{ op: 'insertEquation', linear: '(a+b', display: true }] }), /That equation cannot be built: the \( opened here is not closed/);
  service.close({ id: session.id });
  fs.rmSync(dir, { recursive: true, force: true });
});
