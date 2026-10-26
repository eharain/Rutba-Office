// Export to Video draws each slide from the window's drawing of it, in which
// every shape is a <g data-shape="id"> of its own. A narration's speaker is
// hidden in the show, so the video leaves it out (2.18.0, A4.7): taken out
// whole, the shapes around it kept as they were.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { isNarration } from '../packages/presentation/src/narration.js';
import { withoutShapes } from '../apps/desktop/renderer/apps/slides/svg-shapes.js';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
const WAV = Buffer.concat([Buffer.from('RIFF'), Buffer.alloc(40)]);

test('the narration\'s speaker is taken out of the slide\'s drawing, and nothing else is', () => {
  const d = Deck.open(buildPptx({ title: 'Talk', slides: [{ layout: 'obj', title: 'Two', body: ['A', 'B'] }] }));
  d.addCameo(0);
  const speaker = d.addNarration(0, { data: WAV, contentType: 'audio/wav', durationMs: 1000, poster: { data: PNG, contentType: 'image/png' } });
  const slide = d.slide(0);
  const svg = renderSlide(slide, { tagShapes: true });
  assert.ok(svg.includes(`<g data-shape="${speaker}"`), 'the speaker is drawn in the editor');
  const ids = slide.shapes.filter(isNarration).map((s) => s.id);
  assert.deepEqual(ids.map(String), [String(speaker)]);
  const out = withoutShapes(svg, ids);
  assert.ok(!out.includes(`<g data-shape="${speaker}"`), 'and not in the video');
  for (const s of slide.shapes.filter((x) => !isNarration(x))) {
    if (svg.includes(`<g data-shape="${s.id}"`)) assert.ok(out.includes(`<g data-shape="${s.id}"`), `shape ${s.id} kept`);
  }
  // What is left is still one drawing: as many groups open as close.
  assert.equal((out.match(/<g[\s>]/g) || []).length, (out.match(/<\/g>/g) || []).length);
  assert.match(out, /<\/svg>\s*$/);
});

test('a shape with groups inside it goes whole; an id not drawn changes nothing', () => {
  const svg = '<svg><g data-shape="1"><rect/></g><g data-shape="2"><g><g/><path/></g><circle/></g><g data-shape="23"><text>x</text></g></svg>';
  assert.equal(withoutShapes(svg, ['2']), '<svg><g data-shape="1"><rect/></g><g data-shape="23"><text>x</text></g></svg>');
  assert.equal(withoutShapes(svg, ['9']), svg);
});
