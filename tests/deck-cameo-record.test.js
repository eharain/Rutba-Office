// Record with the camera, the engine's half: the camera's recording put in a
// cameo's place as a video cut to its shape and marked as the cameo's, a
// second recording replacing the first, Reset to Cameo taking it off again,
// and the slides with a cameo listed for Record to turn the camera on for.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { pngOf } from './fixtures/cube-glb.js';

const poster = pngOf([[[30, 30, 30, 255], [30, 30, 30, 255]], [[30, 30, 30, 255], [30, 30, 30, 255]]]);
const clip = (n) => Buffer.from(`webm-${n}`);

const deckWithCameo = () => {
  const d = Deck.open(buildPptx({ title: 'Talk', slides: [{ layout: 'title', title: 'Hello' }, { layout: 'obj', title: 'Next' }, { layout: 'obj', title: 'Last' }] }));
  d.addCameo(0, { shape: 'ellipse' });
  d.addCameo(2);
  return d;
};

test('the camera\'s recording goes in the cameo\'s place, cut to its shape and marked as its recording', () => {
  const d = deckWithCameo();
  const cameo = d.slide(0).shapes.find((s) => s.cameo);
  const id = d.addCameoRecording(0, { data: clip(1), contentType: 'video/webm', poster });
  const rec = d.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(rec.cameoRecording, true);
  assert.equal(rec.media?.kind, 'video');
  assert.deepEqual(rec.geometry, cameo.geometry);
  assert.equal(rec.preset, 'ellipse');
  assert.ok(d.slide(0).shapes.some((s) => s.cameo), 'the cameo stays under it');
  assert.match(d.pkg.text(d.slideParts[0].part), new RegExp(`<rcam:recording xmlns:rcam="http://schemas\\.rutba\\.io/office/2026/cameo" cameo="${cameo.id}"/>`));
  assert.throws(() => d.addCameoRecording(1, { data: clip(2), poster }), /no cameo to record into/);
});

test('a second recording replaces the first; Reset to Cameo takes it off; the file keeps it', () => {
  const d = deckWithCameo();
  d.addCameoRecording(0, { data: clip(1), poster });
  d.addCameoRecording(0, { data: clip(2), poster });
  assert.equal(d.slide(0).shapes.filter((s) => s.cameoRecording).length, 1);
  const again = Deck.open(d.save());
  assert.equal(again.slide(0).shapes.filter((s) => s.cameoRecording).length, 1, 'kept in the file');
  assert.equal(again.resetCameo(0), 1);
  assert.equal(again.slide(0).shapes.filter((s) => s.cameoRecording).length, 0);
  assert.ok(again.slide(0).shapes.some((s) => s.cameo), 'the cameo is left');
  assert.equal(again.resetCameo(0), 0, 'nothing to take off');
});

test('the slides with a cameo are the ones Record turns the camera on for', () => {
  assert.deepEqual(deckWithCameo().cameoSlides(), [0, 2]);
  assert.deepEqual(Deck.open(buildPptx({ title: 'Plain', slides: [{ layout: 'title', title: 'Hi' }] })).cameoSlides(), []);
});
