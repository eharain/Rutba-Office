// Equations in a deck — the engine's half.
//
// PowerPoint writes an equation as a text box holding Office Math in
// `a14:m`, wrapped in an mc:AlternateContent whose fallback is the same box
// filled with a picture of it. Read, it is one shape drawn as MathML; kept,
// it is written back byte for byte; inserted, it is written the way
// PowerPoint writes it.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide, sceneText } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { linearToOmml } from '@rutba/ooxml/math-linear';

const PNG = Buffer.from('89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C4890000000D49444154789C6360000002000154A24F5D0000000049454E44AE426082', 'hex');
const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const RPR = '<a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>';
/** A fraction x over 2, as PowerPoint 365 writes an equation, with its picture fallback. */
const POWERPOINT_EQUATION = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">' +
  '<p:sp><p:nvSpPr><p:cNvPr id="7" name="TextBox 6"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="4000000" y="3000000"/><a:ext cx="1200000" cy="800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a14:m><m:oMathPara ${M}><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath><m:f><m:fPr><m:ctrlPr>${RPR}</m:ctrlPr></m:fPr><m:num><m:r>${RPR}<m:t>x</m:t></m:r></m:num><m:den><m:r>${RPR}<m:t>2</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>` +
  '</mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="7" name="TextBox 6"/><p:cNvSpPr txBox="1"><a:spLocks noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="4000000" y="3000000"/><a:ext cx="1200000" cy="800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:noFill/></a:rPr><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>';
/** Words with an inline equation among them. */
const INLINE = `<p:sp><p:nvSpPr><p:cNvPr id="8" name="TextBox 7"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="600000" y="5000000"/><a:ext cx="6000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Area is </a:t></a:r><a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><m:oMath ${M}><m:r>${RPR}<m:t>π</m:t></m:r><m:sSup><m:e><m:r>${RPR}<m:t>r</m:t></m:r></m:e><m:sup><m:r>${RPR}<m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></a14:m><a:r><a:rPr lang="en-US"/><a:t> here.</a:t></a:r></a:p></p:txBody></p:sp>`;

function withEquations() {
  const pkg = OoxmlPackage.read(buildPptx({ title: 'Maths', slides: [{ layout: 'obj', title: 'Formulae', body: ['Below'] }] }));
  const part = 'ppt/slides/slide1.xml';
  pkg.addPart('ppt/media/image9.png', PNG);
  pkg.write_('ppt/slides/_rels/slide1.xml.rels', pkg.text('ppt/slides/_rels/slide1.xml.rels').replace('</Relationships>', '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image9.png"/></Relationships>'));
  pkg.write_(part, pkg.text(part).replace('</p:spTree>', `${POWERPOINT_EQUATION}${INLINE}</p:spTree>`));
  return pkg.write();
}

test('an equation PowerPoint wrote is one shape, drawn as MathML, its picture kept as the fallback, and saved byte for byte', () => {
  const deck = Deck.open(withEquations());
  const shapes = deck.slide(0).shapes;
  const eq = shapes.filter((s) => s.id === '7');
  assert.equal(eq.length, 1, 'the choice and the fallback are one shape');
  assert.equal(eq[0].alt, true);
  assert.equal(eq[0].fallback.source.part, 'ppt/media/image9.png');
  const run = eq[0].text.paragraphs[0].runs[0];
  assert.equal(run.math.display, true);
  assert.equal(run.math.linear, 'x/2');
  assert.match(run.math.mathml, /<math display="block"><mfrac><mrow><mi>x<\/mi><\/mrow><mrow><mn>2<\/mn><\/mrow><\/mfrac><\/math>/);
  const svg = renderSlide(deck.slide(0));
  assert.match(svg, /<foreignObject x="[\d.]+" y="[\d.]+" width="[\d.]+" height="[\d.]+" overflow="visible"><div xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/);
  assert.match(svg, /<math style="font-family:'Cambria Math'[^"]*" xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML" display="block"><mfrac>/);
  assert.match(sceneText(deck.slide(0)), /x\/2/, 'its linear form stands for its words');
  // Another shape on the slide edited: the equation's XML comes back exactly as it went in.
  const title = shapes.find((s) => s.placeholder?.type === 'title');
  deck.setText(0, title.id, [{ runs: [{ text: 'Formulae, revised' }] }]);
  const saved = OoxmlPackage.read(deck.save()).text('ppt/slides/slide1.xml');
  assert.ok(saved.includes(POWERPOINT_EQUATION), 'byte for byte');
  assert.ok(saved.includes(INLINE));
});

test('words with an inline equation among them are drawn together, the equation inline', () => {
  const deck = Deck.open(withEquations());
  const box = deck.slide(0).shapes.find((s) => s.id === '8');
  assert.deepEqual(box.text.paragraphs[0].runs.map((r) => (r.math ? `[${r.math.linear}]` : r.text)), ['Area is ', '[πr^2]', ' here.']);
  assert.equal(box.text.paragraphs[0].runs[1].math.display, false);
  const svg = renderSlide(deck.slide(0));
  assert.match(svg, /<span[^>]*>Area is <\/span><math style="[^"]*" xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML" display="inline"><mi>π<\/mi><msup>/);
});

test('Insert → Equation writes PowerPoint\'s pair: the Office Math in a14:m with slide run properties, and a PNG fallback', () => {
  const deck = Deck.open(buildPptx({ title: 'Maths', slides: [{ layout: 'obj', title: 'Roots', body: ['a'] }] }));
  const built = linearToOmml('x=(-b±√(b^2-4ac))/2a', { display: true });
  const id = deck.addEquation(0, { omml: built.xml, png: PNG, w: 300, h: 90, linear: 'x=(-b±√(b^2-4ac))/2a' });
  const xml = deck.pkg.text('ppt/slides/slide1.xml');
  const pair = /<mc:AlternateContent[\s\S]*?<\/mc:AlternateContent>/.exec(xml)[0];
  assert.match(pair, /^<mc:AlternateContent xmlns:mc="http:\/\/schemas\.openxmlformats\.org\/markup-compatibility\/2006"><mc:Choice xmlns:a14="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/main" Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="\d+" name="Equation \d+"\/><p:cNvSpPr txBox="1"\/>/);
  assert.match(pair, new RegExp(`<a14:m><m:oMathPara ${M}><m:oMath><m:r><a:rPr lang="en-US"><a:latin typeface="Cambria Math"`));
  assert.doesNotMatch(pair, /<w:/, 'no Word run properties on a slide');
  assert.match(pair, /<mc:Fallback><p:sp>[\s\S]*<a:blipFill><a:blip r:embed="(rId\d+)"\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill>/);
  const rId = /<mc:Fallback>[\s\S]*r:embed="(rId\d+)"/.exec(pair)[1];
  assert.match(deck.pkg.text('ppt/slides/_rels/slide1.xml.rels'), new RegExp(`Id="${rId}" Type="[^"]+/image" Target="\\.\\./media/image\\d+\\.png"`));
  assert.match(deck.pkg.text('[Content_Types].xml'), /<Default Extension="png" ContentType="image\/png"\/>/);
  const again = Deck.open(deck.save());
  const eq = again.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(eq.text.paragraphs[0].runs[0].math.linear, 'x=(-b±√(b^2-4ac))/2a');
  assert.match(eq.text.paragraphs[0].runs[0].math.mathml, /<msqrt>/);
});

test('an equation moves, copies, reorders and deletes as one shape; its words are the editor\'s; editing it again replaces the math and its picture', () => {
  const deck = Deck.open(withEquations());
  deck.setGeometry(0, '7', { x: 10, y: 20, w: 150, h: 60 });
  const xml = () => deck.pkg.text('ppt/slides/slide1.xml');
  assert.equal((xml().match(/<a:off x="95250" y="190500"\/>/g) || []).length, 2, 'the choice and the picture moved together');
  const clip = deck.shapeClip(0, '7');
  const copy = deck.pasteShape(0, clip);
  assert.equal((xml().match(new RegExp(`<p:cNvPr id="${copy}" `, 'g')) || []).length, 2, 'the copy\'s pair share one fresh id');
  assert.equal(deck.slide(0).shapes.filter((s) => s.alt).length, 2);
  deck.reorderShape(0, '7', 'back');
  assert.equal(deck.slide(0).shapes[0].id, '7');
  assert.throws(() => deck.setText(0, '7', [{ runs: [{ text: 'plain' }] }]), /equation editor/);
  const built = linearToOmml('a^2+b^2=c^2', { display: true });
  const pngBefore = deck.pkg.read('ppt/media/image9.png');
  const newPng = Buffer.concat([PNG, Buffer.from([0])]);
  deck.setEquation(0, '7', { omml: built.xml, png: newPng, w: 200, h: 70 });
  assert.equal(deck.slide(0).shapes.find((s) => s.id === '7').text.paragraphs[0].runs[0].math.linear, 'a^2+b^2=c^2');
  assert.notDeepEqual(deck.pkg.read('ppt/media/image9.png'), pngBefore, 'the fallback picture redrawn in place');
  deck.removeShape(0, String(copy));
  deck.removeShape(0, '7');
  assert.equal((xml().match(/mc:AlternateContent/g) || []).length, 0, 'nothing of either pair is left');
  assert.equal(deck.slide(0).shapes.some((s) => s.alt), false);
});

test('an alternate content whose choice this reader does not know draws its fallback, where before it drew nothing', () => {
  const pkg = OoxmlPackage.read(buildPptx({ title: 'Chart ex', slides: [{ layout: 'obj', title: 'A new chart', body: ['a'] }] }));
  const part = 'ppt/slides/slide1.xml';
  pkg.write_(part, pkg.text(part).replace('</p:spTree>', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Chart 8"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"/></a:graphic></p:graphicFrame></mc:Choice>' +
    '<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="9" name="Chart 8"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>This chart needs a newer version</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent></p:spTree>'));
  const deck = Deck.open(pkg.write());
  const s = deck.slide(0).shapes.find((x) => x.id === '9');
  assert.ok(s, 'the fallback is read');
  assert.equal(s.alt, false);
  assert.match(renderSlide(deck.slide(0)), /<rect x="96" y="96" width="192" height="96" fill="#eeeeee"\/>[\s\S]*This chart needs a/);
});
