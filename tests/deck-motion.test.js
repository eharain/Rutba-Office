// Animations → Motion Paths and Effect Options → Sequence, the engine's
// half: a path written as PowerPoint writes one drawn by hand (a custom
// path, `p:animMotion` from where the shape stands, in fractions of the
// slide), kept as it is when it is retimed; and a text shape's effect split
// into one per paragraph — each aimed at its paragraph, each on a click or
// all together, the build list saying `build="p"` — and gathered back.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, readAnimations, motionPath, MOTION_PATHS, shapeParagraphs } from '@rutba/presentation';
import { renderSlide } from '@rutba/presentation/render';

function deck() {
  const d = Deck.open(buildPptx({ title: 'Motion', slides: [{ layout: 'obj', title: 'Moving parts', body: ['One'] }] }));
  const box = d.addShape(0, { preset: 'rect', x: 100, y: 300, w: 200, h: 120 });
  const words = d.addTextBox(0, { x: 500, y: 200, w: 300, h: 200, paragraphs: [{ runs: [{ text: 'First' }] }, { runs: [] }, { runs: [{ text: 'Second' }] }, { runs: [{ text: 'Third' }] }] });
  return { d, box: String(box), words: String(words) };
}
const xmlOf = (d, i = 0) => d.pkg.text(d.slideParts[i].part);

test('the Motion Paths gallery gives each path in PowerPoint\'s own form, a circle round on a wide slide', () => {
  assert.equal(motionPath('down'), 'M 0 0 L 0 0.25 E');
  assert.equal(motionPath('right', 0.5), 'M 0 0 L 0.15 0 E', 'a length in heights, as a fraction of the width');
  const circle = motionPath('circle', 9 / 16);
  assert.match(circle, /^M 0 0 C .* Z E$/);
  // Its widest x is its radius in heights times the aspect: round on the slide.
  const xs = circle.replace(/[A-Z]/g, ' ').trim().split(/\s+/).map(Number).filter((_, i) => i % 2 === 0);
  assert.equal(Math.max(...xs), Number((0.125 * 9 / 16).toFixed(5)));
  for (const [name] of MOTION_PATHS) assert.match(motionPath(name), /^M 0 0 .* E$/, name);
  assert.throws(() => motionPath('zigzag'), /not a motion path/);
});

test('a motion path is written as a custom path from where the shape stands, and read back with its path', () => {
  const { d, box } = deck();
  const at = d.addAnimation(0, box, { kind: 'path', path: motionPath('down') });
  const xml = xmlOf(d);
  assert.match(xml, /<p:cTn [^>]*presetID="0" presetClass="path" presetSubtype="0" accel="50000" decel="50000" fill="hold"/);
  assert.match(xml, new RegExp(`<p:animMotion origin="layout" path="M 0 0 L 0 0.25 E" pathEditMode="relative"><p:cBhvr><p:cTn id="\\d+" dur="2000" fill="hold"/><p:tgtEl><p:spTgt spid="${box}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`));
  const [e] = d.animations(0);
  assert.equal(at, 0);
  assert.deepEqual({ kind: e.kind, effect: e.effect, name: e.name, path: e.path, duration: e.duration, trigger: e.trigger }, { kind: 'path', effect: 'custom', name: 'Motion Path', path: 'M 0 0 L 0 0.25 E', duration: 2, trigger: 'onClick' });
  assert.throws(() => d.addAnimation(0, box, { kind: 'path' }), /needs a path/);
});

test('a path retimed keeps its own path; a new path replaces it; it comes back from the file as it was', () => {
  const { d, box } = deck();
  d.addAnimation(0, box, { kind: 'path', path: motionPath('arcDown') });
  d.setAnimation(0, 0, { duration: 3.5, trigger: 'afterPrevious' });
  let [e] = d.animations(0);
  assert.equal(e.path, motionPath('arcDown'));
  assert.equal(e.duration, 3.5);
  assert.equal(e.trigger, 'afterPrevious');
  d.setAnimation(0, 0, { path: motionPath('loop') });
  [e] = d.animations(0);
  assert.equal(e.path, motionPath('loop'));
  assert.equal(e.duration, 3.5, 'a new path keeps the length it had');
  // From a fade to a path, and back.
  d.setAnimation(0, 0, { kind: 'entr', effect: 'fade' });
  assert.equal(d.animations(0)[0].kind, 'entr');
  d.setAnimation(0, 0, { kind: 'path', path: motionPath('up') });
  assert.equal(d.animations(0)[0].path, motionPath('up'));
  const again = Deck.open(d.save());
  assert.deepEqual(readAnimations(again.pkg.text(again.slideParts[0].part)).map((x) => [x.kind, x.path]), [['path', motionPath('up')]]);
});

test('By Paragraph splits a text shape\'s effect into one per paragraph with words, each on a click, build="p" in the list', () => {
  const { d, words } = deck();
  assert.deepEqual(shapeParagraphs(xmlOf(d), words), [0, 2, 3], 'the empty second paragraph is passed over');
  d.addAnimation(0, words, { kind: 'entr', effect: 'fly', direction: 'left' });
  const first = d.setAnimationSequence(0, 0, 'paragraph');
  assert.equal(first, 0);
  const list = d.animations(0);
  assert.deepEqual(list.map((e) => [e.effect, e.direction, e.paragraph, e.trigger]), [['fly', 'left', 0, 'onClick'], ['fly', 'left', 2, 'onClick'], ['fly', 'left', 3, 'onClick']]);
  const xml = xmlOf(d);
  assert.equal((xml.match(/<p:txEl><p:pRg st="(\d)" end="\1"\/><\/p:txEl>/g) || []).length, 3 * 3, 'every behaviour of each effect aims at its paragraph');
  assert.match(xml, new RegExp(`<p:bldP build="p" spid="${words}" grpId="0"[^>]*/>`));
  // All at Once: the rest start with the first.
  d.setAnimationSequence(0, 1, 'together');
  assert.deepEqual(d.animations(0).map((e) => [e.paragraph, e.trigger]), [[0, 'onClick'], [2, 'withPrevious'], [3, 'withPrevious']]);
  // As One Object, from any of them: one effect on the shape, the build list plain again.
  d.setAnimationSequence(0, 2, 'object');
  const one = d.animations(0);
  assert.deepEqual(one.map((e) => [e.effect, e.direction, e.paragraph, e.trigger]), [['fly', 'left', null, 'onClick']]);
  assert.doesNotMatch(xmlOf(d), /build="p"/);
  assert.doesNotMatch(xmlOf(d), /<p:txEl>/);
});

test('effects split by paragraph survive the file, and the slide draws each paragraph in a group of its own', () => {
  const { d, words } = deck();
  d.addAnimation(0, words, { kind: 'entr', effect: 'fade' });
  d.setAnimationSequence(0, 0, 'paragraph');
  const again = Deck.open(d.save());
  assert.deepEqual(again.animations(0).map((e) => e.paragraph), [0, 2, 3]);
  const scene = again.slide(0);
  const svg = renderSlide(scene, { tagShapes: true });
  const inShape = new RegExp(`data-shape="${words}"[\\s\\S]*?data-para="3"`);
  assert.match(svg, inShape, 'the words\' lines carry their paragraph');
  assert.match(svg, /<g data-para="0">[\s\S]*?First[\s\S]*?<\/g>/);
  assert.throws(() => d.setAnimationSequence(0, 99, 'paragraph'), /no animation/);
});

test('only an effect this writes is split; a shape with no words says so', () => {
  const { d, box } = deck();
  d.addAnimation(0, box, { kind: 'entr', effect: 'zoom' });
  assert.throws(() => d.setAnimationSequence(0, 0, 'paragraph'), /no words/);
  d.addAnimation(0, box, { kind: 'path', path: motionPath('down') });
  assert.throws(() => d.setAnimationSequence(0, 1, 'paragraph'), /entrance, emphasis or exit/);
});
