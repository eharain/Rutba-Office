// Record → narration — the engine's half.
//
// A slide's recorded voice goes in as PowerPoint keeps it: an "Audio
// Recording" speaker in the corner, and in the timing the shape's media
// node (hidden while stopped) and a call that plays it from the start as
// the slide comes in — first in the main sequence, in the group the slide
// begins or a new one ahead of the clicks. Animations edited after keep
// it; Clear takes the speaker, its call and its media node away and leaves
// the rest.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { isNarration } from '../packages/presentation/src/narration.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
const WAV = (() => { const b = Buffer.alloc(44 + 200); b.write('RIFF', 0); b.writeUInt32LE(236, 4); b.write('WAVEfmt ', 8); b.writeUInt32LE(16, 16); b.writeUInt16LE(1, 20); b.writeUInt16LE(1, 22); b.writeUInt32LE(8000, 24); b.writeUInt32LE(16000, 28); b.writeUInt16LE(2, 32); b.writeUInt16LE(16, 34); b.write('data', 36); b.writeUInt32LE(200, 40); return b; })();
const deck = () => Deck.open(buildPptx({ title: 'Talk', slides: [{ layout: 'title', title: 'One' }, { layout: 'obj', title: 'Two', body: ['A', 'B'] }] }));
const narrate = (d, i, ms = 4200) => d.addNarration(i, { data: WAV, contentType: 'audio/wav', durationMs: ms, poster: { data: PNG, contentType: 'image/png' } });

test('a narration goes in as PowerPoint keeps one: a hidden speaker, played from the start as the slide comes in', () => {
  const d = deck();
  const id = narrate(d, 0);
  const shape = d.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.ok(isNarration(shape), 'an Audio Recording speaker');
  assert.ok(shape.geometry.x > d.size.width - 80 && shape.geometry.y > d.size.height - 80, 'in the corner');
  const xml = d.pkg.text('ppt/slides/slide1.xml');
  assert.match(xml, /<p:cTn id="(\d+)" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="\d+" fill="hold"><p:stCondLst><p:cond delay="indefinite"\/><p:cond evt="onBegin" delay="0"><p:tn val="\1"\/><\/p:cond>/);
  assert.match(xml, new RegExp(`presetClass="mediacall"[^>]*nodeType="afterEffect"[\\s\\S]*?<p:cmd type="call" cmd="playFrom\\(0\\.0\\)"><p:cBhvr><p:cTn id="\\d+" dur="4200" fill="hold"/><p:tgtEl><p:spTgt spid="${id}"/>`));
  assert.match(xml, new RegExp(`<p:audio><p:cMediaNode vol="80000" showWhenStopped="0">[\\s\\S]*?<p:spTgt spid="${id}"/></p:tgtEl></p:cMediaNode></p:audio>`));
  const ids = [...xml.matchAll(/<p:cTn\b[^>]*?\sid="(\d+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, ids.length, 'every time node its own id');
  assert.equal(d.animations(0).find((e) => e.kind === 'media')?.shapeId, String(id));
});

test('a slide with clicks gets the narration ahead of them; animations edited after keep it; Clear takes it all away', () => {
  const d = deck();
  const title = d.slide(1).shapes.find((s) => s.placeholder?.type === 'title').id;
  const body = d.slide(1).shapes.find((s) => s.placeholder?.type !== 'title').id;
  d.addAnimation(1, title, { kind: 'entr', effect: 'fade' });
  const id = narrate(d, 1, 3000);
  let list = d.animations(1);
  assert.equal(list[0].kind, 'media');
  assert.equal(list[0].trigger, 'afterPrevious');
  assert.ok(list.some((e) => e.kind === 'entr' && String(e.shapeId) === String(title)), 'the click animation stays');
  // An edit after: the narration is still there, still begun by the slide.
  d.addAnimation(1, body, { kind: 'entr', effect: 'appear' });
  const xml = d.pkg.text('ppt/slides/slide2.xml');
  assert.match(xml, /playFrom\(0\.0\)/);
  assert.match(xml, /evt="onBegin"/);
  assert.match(xml, new RegExp(`<p:audio><p:cMediaNode[\\s\\S]*?spid="${id}"`));
  // Recorded again: one narration, not two.
  narrate(d, 1, 5000);
  assert.equal(d.slide(1).shapes.filter(isNarration).length, 1);
  assert.equal((d.pkg.text('ppt/slides/slide2.xml').match(/<p:audio>/g) || []).length, 1);
  // Cleared: the speaker, its call and its media node go; the animations stay.
  assert.equal(d.clearNarration(1), 1);
  const after = d.pkg.text('ppt/slides/slide2.xml');
  assert.doesNotMatch(after, /playFrom|<p:audio>|Audio Recording/);
  list = d.animations(1);
  assert.equal(list.length, 2);
  assert.equal(d.clearNarration(1), 0);
});

test('a slide whose first effects already begin by themselves takes the call into that group, and the deck saves and opens again', () => {
  const d = deck();
  const title = d.slide(1).shapes.find((s) => s.placeholder?.type === 'title').id;
  d.addAnimation(1, title, { kind: 'entr', effect: 'fade', trigger: 'afterPrevious' });
  narrate(d, 1, 2000);
  const xml = d.pkg.text('ppt/slides/slide2.xml');
  assert.equal((xml.match(/evt="onBegin"/g) || []).length, 1, 'one group begun by the slide');
  const back = Deck.open(d.save());
  assert.equal(back.slide(1).shapes.filter(isNarration).length, 1);
  assert.equal(back.animations(1)[0].kind, 'media');
});
