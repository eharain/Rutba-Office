// Spoken text layers (D-2026-11-13-31), review 2: a reading never puts or
// moves a clip onto a locked track, finds a lane however many are in use,
// and the cues a Recorder keeps at Stop and at a new Record.
import test from 'node:test';
import assert from 'node:assert/strict';
import { spokenLayer, placeSpokenLines, placeSpokenCues, spokenCueLog } from '../packages/studio/src/spoken.js';
import { emptyProject, addTitle, clipsOf, laneOf, brokenLocks } from '../packages/studio/src/project.js';

const line = (text, seconds, n = 1) => ({ text, seconds, url: `C:/p/Spoken lines/sp1-${n}.wav`, name: `sp1-${n}.wav`, size: 100 });
const lock = (doc, section, lane) => ({ ...doc, lanes: { ...(doc.lanes || {}), [section]: { ...(doc.lanes?.[section] || {}), [lane]: { locked: true } } } });

/** Two lower thirds in a row, the first's band and words to move up past the second. */
function twoTitles() {
  const a = addTitle(emptyProject(), 'lowerThird', { at: 0, text: 'First' });
  const b = addTitle(a.doc, 'lowerThird', { at: 5.5, text: 'Second' });
  return { doc: b.doc, text: a.ids.find((id) => id.startsWith('text')) };
}

test('a band moving up past the next title skips a locked track, and the edit keeps every lock', () => {
  const { doc, text } = twoTitles();
  const open = placeSpokenLines(doc, { from: text, lines: [line('One', 3, 1), line('Two', 3, 2), line('Three', 3, 3)] });
  const openBands = spokenLayer(open.doc, open.layer).bands.map(laneOf);
  // Lock the lowest lane the band went to when nothing was locked: it must go elsewhere.
  const locked = lock(doc, 'v', Math.min(...openBands));
  const r = placeSpokenLines(locked, { from: text, lines: [line('One', 3, 1), line('Two', 3, 2), line('Three', 3, 3)] });
  const layer = spokenLayer(r.doc, r.layer);
  for (const p of [...layer.bands, ...layer.lines.map((l) => l.text)]) assert.notEqual(locked.lanes.v[laneOf(p)]?.locked, true, `${p.id} on a locked track`);
  assert.deepEqual(brokenLocks(locked, r.doc), []);
});

/** A plain title, its words on lane 1 from 0 to 4, with pictures after it on each of `lanes` lanes from its own up. */
function crowded(lanes) {
  const t0 = addTitle(emptyProject(), 'title', { at: 0, text: 'A' });
  const text = t0.ids.find((id) => id.startsWith('text'));
  const from = laneOf(clipsOf(t0.doc).find((p) => p.id === text));
  const pics = [];
  for (let k = 0; k < lanes; k++) pics.push({ id: `img${k}`, type: 'image', url: `C:/p/${k}.png`, timing: { start: 4, end: 60 }, lane: from + k });
  return { doc: { ...t0.doc, patches: [...t0.doc.patches, ...pics] }, text, from };
}

test('the words of a title with no band climb past a locked track', () => {
  const { doc, text, from } = crowded(1);
  const d = lock(doc, 'v', from + 1);
  const r = placeSpokenLines(d, { from: text, lines: [line('One', 5, 1), line('Two', 5, 2)] });
  assert.equal(laneOf(spokenLayer(r.doc, r.layer).lines[0].text), from + 2);
  assert.deepEqual(brokenLocks(d, r.doc), []);
});

test('with 70 lanes in use above the words, a reading still finds the free one past them', () => {
  const { doc, text, from } = crowded(70);
  const r = placeSpokenLines(doc, { from: text, lines: [line('One', 5, 1), line('Two', 5, 2)] });
  assert.equal(laneOf(spokenLayer(r.doc, r.layer).lines[0].text), from + 70);
});

test('the sound of a reading goes on no locked sound track', () => {
  const t0 = addTitle(emptyProject(), 'title', { at: 0, text: 'A' });
  const text = t0.ids.find((id) => id.startsWith('text'));
  const d = lock(lock(t0.doc, 'a', 0), 'a', 1);
  const r = placeSpokenLines(d, { from: text, lines: [line('One', 1, 1)] });
  assert.equal(laneOf(spokenLayer(r.doc, r.layer).lines[0].sound), 2);
  assert.deepEqual(brokenLocks(d, r.doc), []);
});

test('cues at Stop: a line still being read is kept up to the recording\'s length, once', () => {
  const log = spokenCueLog();
  log.clear();
  log.begin(1, 'One', true);
  log.end(2.5);
  log.begin(2.8, 'Two', false);
  assert.deepEqual(log.take(4), [{ start: 1, end: 2.5, text: 'One', show: true }, { start: 2.8, end: 4, text: 'Two', show: false }]);
  // The line read on after Stop ends with no recording running: nothing more is kept.
  log.end(null);
  assert.deepEqual(log.take(9), []);
});

test('cues at a new Record: nothing kept from before, and the line being read counts from the start', () => {
  const log = spokenCueLog();
  log.begin(null, 'Before', true); // begun while nothing recorded
  log.clear(); // Record pressed while it reads
  log.end(1.5);
  assert.deepEqual(log.take(3), [{ start: 0, end: 1.5, text: 'Before', show: true }]);
  // A recording stopped mid-line, then a new one begun while that line still reads.
  log.clear();
  log.begin(0.5, 'Long', true);
  assert.deepEqual(log.take(2), [{ start: 0.5, end: 2, text: 'Long', show: true }]);
  log.clear();
  log.end(1.2);
  assert.deepEqual(log.take(5), [{ start: 0, end: 1.2, text: 'Long', show: true }], 'the rest of it, in the new recording');
  // Lines from an earlier recording never leak into the next.
  log.begin(1, 'Old', true);
  log.end(2);
  log.clear();
  assert.deepEqual(log.take(3), []);
});

test('a line read while nothing records keeps no cue', () => {
  const log = spokenCueLog();
  log.begin(null, 'Preview', true);
  log.end(null);
  assert.deepEqual(log.take(null), []);
});
