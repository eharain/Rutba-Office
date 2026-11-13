// Spoken text layers (D-2026-11-13-31), review 1: where a reading puts its
// clips (a title's band never over another clip, a hidden band left hidden,
// a line taking only the look of its words), the lines a Recorder read kept
// clear of captions and of each other, a line of direction marks alone being
// no line, and a recording with spoken lines keeping its sound.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  linesOf, spokenLayer, placeSpokenLines, placeSpokenCues, soundedTakes, speaksIntoRecording,
} from '../packages/studio/src/spoken.js';
import { emptyProject, addTitle, clipsOf, laneOf } from '../packages/studio/src/project.js';
import { placeCaptions } from '../packages/studio/src/captions.js';
import { emptyStack, addItem, updateItem } from '../packages/studio/src/stack.js';

const line = (text, seconds, n = 1) => ({ text, seconds, url: `C:/p/Spoken lines/sp1-${n}.wav`, name: `sp1-${n}.wav`, size: 100 });

/** Whether two pictures share a lane at the same time. */
function overlaps(doc) {
  const pics = clipsOf(doc).filter((p) => p.type !== 'sound');
  return pics.some((p, i) => pics.some((q, j) => j > i && laneOf(p) === laneOf(q) && p.timing.start < q.timing.end - 0.004 && q.timing.start < p.timing.end - 0.004));
}

test('two lower thirds in a row: the band of the first, stretched over its lines, moves up with its words and covers nothing', () => {
  const a = addTitle(emptyProject(), 'lowerThird', { at: 0, text: 'First' });
  const b = addTitle(a.doc, 'lowerThird', { at: 5.5, text: 'Second' });
  const d = b.doc;
  const aText = a.ids.find((id) => id.startsWith('text'));
  assert.equal(overlaps(d), false);
  const r = placeSpokenLines(d, { from: aText, lines: [line('One', 3, 1), line('Two', 3, 2), line('Three', 3, 3)] });
  assert.equal(overlaps(r.doc), false, 'no two clips share a lane at once');
  const layer = spokenLayer(r.doc, r.layer);
  assert.ok(layer.bands.length, 'a lower third has a band');
  const wordsLane = laneOf(layer.lines[0].text);
  assert.ok(layer.bands.every((band) => laneOf(band) < wordsLane), 'the band stays under its words');
  assert.ok(layer.bands.every((band) => band.timing.start === 0 && band.timing.end === r.end));
  // The second title is where it was.
  for (const id of b.ids) assert.deepEqual(clipsOf(r.doc).find((p) => p.id === id), clipsOf(d).find((p) => p.id === id));
  // Read again from there, it stays clear.
  const again = placeSpokenLines(r.doc, { from: layer.lines[0].text.id, lines: [line('One', 4, 1), line('Two', 4, 2)] });
  assert.equal(overlaps(again.doc), false);
});

test('a band with room on its own lanes stays on them', () => {
  const a = addTitle(emptyProject(), 'lowerThird', { at: 0, text: 'Only' });
  const aText = a.ids.find((id) => id.startsWith('text'));
  const was = Object.fromEntries(clipsOf(a.doc).map((p) => [p.id, laneOf(p)]));
  const r = placeSpokenLines(a.doc, { from: aText, lines: [line('One', 3, 1), line('Two', 3, 2)] });
  const layer = spokenLayer(r.doc, r.layer);
  for (const band of layer.bands) assert.equal(laneOf(band), was[band.id]);
  assert.equal(laneOf(layer.lines[0].text), was[aText]);
});

test('a band hidden before the lines are read stays hidden; one shown stays shown', () => {
  const r0 = addTitle(emptyProject(), 'lowerThird', { at: 0, text: 'Name' });
  const textId = r0.ids.find((id) => id.startsWith('text'));
  const bandId = r0.ids.find((id) => id !== textId);
  const hidden = { ...r0.doc, patches: r0.doc.patches.map((p) => (p.id === bandId ? { ...p, visible: false } : p)) };
  const r = placeSpokenLines(hidden, { from: textId, lines: [line('A', 1, 1)] });
  assert.equal(clipsOf(r.doc).find((p) => p.id === bandId).visible, false);
  const shown = placeSpokenLines(r0.doc, { from: textId, lines: [line('A', 1, 1)] });
  assert.notEqual(clipsOf(shown.doc).find((p) => p.id === bandId).visible, false);
});

test('a line takes the look of its words, not their transition, keyframes, lock or caption mark', () => {
  const t0 = addTitle(emptyProject(), 'title', { at: 0, text: 'A' });
  const textId = t0.ids.find((id) => id.startsWith('text'));
  const marked = { ...t0.doc, patches: t0.doc.patches.map((p) => (p.id === textId ? { ...p, caption: true, transition: { kind: 'fade', seconds: 1 }, keys: { fx: [] }, locked: true, color: '#ffc107' } : p)) };
  const r = placeSpokenLines(marked, { from: textId, lines: [line('A', 1, 1)] });
  const [one] = spokenLayer(r.doc, r.layer).lines;
  for (const k of ['caption', 'transition', 'keys', 'locked']) assert.equal(one.text[k], undefined, k);
  assert.equal(one.text.color, '#ffc107');
});

test('a line of direction marks alone, as pasted Urdu and Arabic carry, is no line', () => {
  const [rlm, lrm, alm, rli, pdi] = [0x200f, 0x200e, 0x061c, 0x2067, 0x2069].map((c) => String.fromCharCode(c));
  const text = ['One', rlm, `${lrm} ${rlm}`, `${alm}${rli}${pdi}`, `${rlm}Two`].join('\n');
  assert.deepEqual(linesOf(text), ['One', `${rlm}Two`]);
});

test('the lines a Recorder read go above every picture of the time, on a lane no caption shares, never over each other', () => {
  let d = emptyProject();
  d = placeCaptions(d, [{ start: 0, end: 30, text: 'A caption all along' }]).doc;
  const r = placeSpokenCues(d, [{ start: 1, end: 3.5, text: 'One' }, { start: 3.2, end: 5, text: 'Two' }], { at: 0 });
  assert.equal(overlaps(r.doc), false);
  const layer = spokenLayer(r.doc, r.layer);
  const capLane = laneOf(clipsOf(d)[0]);
  assert.ok(layer.lines.every((l) => laneOf(l.text) > capLane));
  assert.equal(layer.lines[0].text.timing.end, 3.2, 'a line that ran into the next ends where it begins');
  assert.equal(layer.lines[1].text.line, undefined);
  assert.equal(layer.lines[1].text.spoken.line, 1);
});

test('a take that carries the sound counts as sounded when spoken lines come with it, and only then', () => {
  const takes = [{ role: 'screen', carrier: true, hasAudio: false }, { role: 'camera', carrier: false, hasAudio: false }];
  assert.deepEqual(soundedTakes(takes, []), takes);
  const out = soundedTakes(takes, [{ start: 0, end: 1, text: 'Hi' }]);
  assert.equal(out[0].hasAudio, true);
  assert.equal(out[1].hasAudio, false, 'a take with only a silent track stays silent');
  assert.equal(soundedTakes([{ role: 'voice', hasAudio: true }], [{ start: 0, end: 1, text: 'x' }])[0].hasAudio, true);
});

test('a recording begun with a title ready to read, its voice on and Sound on, is counted as having sound', () => {
  let st = emptyStack();
  ({ stack: st } = addItem(st, { kind: 'title', text: 'T' }));
  const id = st.items.find((x) => x.kind === 'title').id;
  assert.equal(speaksIntoRecording(st, { voices: true }), false, 'no lines');
  st = updateItem(st, id, { lines: 'Hello\nThere' });
  assert.equal(speaksIntoRecording(st, { voices: true }), true);
  assert.equal(speaksIntoRecording(st, { voices: false }), false, 'no voice to read them');
  assert.equal(speaksIntoRecording({ ...st, sound: 'off' }, { voices: true }), false, 'Sound off');
  assert.equal(speaksIntoRecording({ ...st, sound: 'music' }, { voices: true }), false, 'Music instead');
  const quiet = updateItem(st, id, { spoken: { show: true, play: false, voice: '', speed: 1 } });
  assert.equal(speaksIntoRecording(quiet, { voices: true }), false, 'Play the voice off');
});
