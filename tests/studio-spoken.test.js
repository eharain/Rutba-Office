// Spoken text layers (D-2026-11-13-31): a text layer's lines read aloud one
// after another, each shown while its own voice plays, on Video's timeline
// and in the studio's titles, with the person's choice to show the words,
// play the voice, or both.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  linesOf, sentencesOfLine, joinSamples, secondsOf, lineTimes, wrapForBand,
  isSpoken, spokenLayer, spokenLayers, newSpokenLayer, spokenChoice, setSpokenChoice, wordsToRead,
  placeSpokenLines, placeSpokenCues, spokenFileName, normaliseSpoken, normaliseLines,
  LINE_GAP, SENTENCE_GAP, MAX_LINES, MAX_LINES_TEXT,
} from '../packages/studio/src/spoken.js';
import { emptyProject, addTitle, clipsOf, laneClips, laneOf, moveClip } from '../packages/studio/src/project.js';
import { placeCaptions, captionsOf } from '../packages/studio/src/captions.js';
import { normaliseItem, emptyStack, addItem, updateItem } from '../packages/studio/src/stack.js';

const line = (text, seconds, n = 1) => ({ text, seconds, url: `C:/p/Spoken lines/sp1-${n}.wav`, name: `sp1-${n}.wav`, size: 100 });

/** A project with one plain title clip (its words alone) at `at`. */
function withTitle(text, at = 2) {
  const r = addTitle(emptyProject(), 'title', { at, text });
  const textId = r.ids.find((id) => id.startsWith('text'));
  return { doc: r.doc, textId, ids: r.ids };
}

test('a text is read as its lines: each as typed, blank lines and spaces dropped, at most MAX_LINES', () => {
  assert.deepEqual(linesOf('One line.\r\n\n   Two   words  \nThree'), ['One line.', 'Two words', 'Three']);
  assert.deepEqual(linesOf(''), []);
  assert.deepEqual(linesOf(null), []);
  assert.deepEqual(linesOf('a\u2028b'), ['a', 'b']);
  assert.equal(linesOf(Array.from({ length: MAX_LINES + 9 }, (_, i) => `L${i}`).join('\n')).length, MAX_LINES);
});

test('a line of several sentences is spoken a sentence at a time', () => {
  assert.deepEqual(sentencesOfLine('Hello there. How are you? Fine!'), ['Hello there.', 'How are you?', 'Fine!']);
  assert.deepEqual(sentencesOfLine('  '), []);
  assert.deepEqual(sentencesOfLine('No full stop'), ['No full stop']);
});

test('sentences\' samples are joined with quiet between, at the first one\'s rate', () => {
  const a = { samples: new Float32Array([0.5, 0.5]), rate: 10 };
  const b = { samples: new Float32Array([-0.5]), rate: 10 };
  const j = joinSamples([a, b], { gap: 0.2 });
  assert.equal(j.rate, 10);
  assert.deepEqual([...j.samples], [0.5, 0.5, 0, 0, -0.5]);
  assert.equal(joinSamples([]).samples.length, 0);
  assert.equal(joinSamples([{ samples: [], rate: 24000 }, a]).samples.length, 2, 'an empty sentence adds nothing');
  // A part at another rate is stretched to the first's.
  const c = joinSamples([a, { samples: new Float32Array(20), rate: 20 }], { gap: 0 });
  assert.equal(c.samples.length, 12);
  assert.equal(secondsOf(new Float32Array(48000), 24000), 2);
  assert.ok(SENTENCE_GAP > 0 && SENTENCE_GAP < LINE_GAP);
});

test('lines are timed one after another, the gap between', () => {
  assert.deepEqual(lineTimes([1, 2.5], { at: 3, gap: 0.5 }), [{ start: 3, end: 4 }, { start: 4.5, end: 7 }]);
  assert.deepEqual(lineTimes([0], { at: 0 })[0], { start: 0, end: 0.1 }, 'never shorter than the shortest clip');
});

test('a line is wrapped for the band onto two lines at most', () => {
  const fits = (s) => s.length <= 10;
  assert.deepEqual(wrapForBand('short one', fits), ['short one']);
  assert.deepEqual(wrapForBand('one two three four five', fits), ['one two', 'three four five']);
  assert.deepEqual(wrapForBand('', fits), []);
  assert.deepEqual(wrapForBand('enormousword tail', fits), ['enormousword', 'tail']);
});

test('a title read onto the timeline: a text clip and its voice for each line, one group each, line by line from where it stood', () => {
  const { doc, textId } = withTitle('First line\nSecond line', 2);
  const r = placeSpokenLines(doc, { from: textId, lines: [line('First line', 1.5, 1), line('Second line', 2, 2)], voice: 'af_heart', speed: 1.25 });
  assert.equal(r.layer, 'sp1');
  assert.equal(r.start, 2);
  assert.equal(r.end, 2 + 1.5 + LINE_GAP + 2);
  const layer = spokenLayer(r.doc, 'sp1');
  assert.equal(layer.lines.length, 2);
  const [one, two] = layer.lines;
  assert.equal(one.text.text, 'First line');
  assert.deepEqual(one.text.timing, { start: 2, end: 3.5 });
  assert.deepEqual(one.sound.timing, one.text.timing, 'the words are shown exactly while their voice plays');
  assert.deepEqual(two.text.timing, { start: 3.8, end: 5.8 });
  assert.equal(one.text.group, one.sound.group);
  assert.notEqual(one.text.group, two.text.group);
  assert.equal(one.sound.url, 'C:/p/Spoken lines/sp1-1.wav');
  assert.equal(r.doc.sources[one.sound.url].kind, 'audio');
  assert.equal(r.doc.sources[one.sound.url].duration, 1.5);
  assert.equal(laneOf(one.text), laneOf(two.text), 'the lines share a lane');
  assert.equal(layer.voice, 'af_heart');
  assert.equal(layer.speed, 1.25);
  assert.ok(!clipsOf(r.doc).some((p) => p.type === 'text' && !isSpoken(p)), 'the title\'s own words give way to its lines');
  assert.equal(one.text.fx, clipsOf(doc).find((p) => p.id === textId).fx, 'each line keeps the title\'s look');
  assert.equal(spokenLayers(r.doc).length, 1);
  assert.ok(isSpoken(one.text) && isSpoken(one.sound));
});

test('a title\'s band is stretched over all its lines and marked as the layer\'s', () => {
  const r0 = addTitle(emptyProject(), 'lowerThird', { at: 0, text: 'Name' });
  const textId = r0.ids.find((id) => id.startsWith('text'));
  const bandIds = r0.ids.filter((id) => id !== textId);
  assert.ok(bandIds.length, 'a lower third has a band');
  const r = placeSpokenLines(r0.doc, { from: textId, lines: [line('A', 3, 1), line('B', 4, 2)] });
  const band = clipsOf(r.doc).find((p) => p.id === bandIds[0]);
  assert.deepEqual(band.timing, { start: 0, end: r.end });
  assert.equal(band.spoken.band, true);
  assert.equal(band.group, undefined, 'the band leaves the title\'s group, as each line has its own');
  assert.equal(spokenLayer(r.doc, r.layer).bands.length, bandIds.length);
});

test('lines move down a lane when theirs is taken further on', () => {
  const { doc, textId } = withTitle('A\nB', 0);
  const lane = laneOf(clipsOf(doc).find((p) => p.id === textId));
  const blocker = { id: 'text-99', type: 'text', text: 'x', timing: { start: 2, end: 9 }, lane };
  const busy = { ...doc, patches: [...doc.patches, blocker] };
  const r = placeSpokenLines(busy, { from: textId, lines: [line('A', 1.5, 1), line('B', 1.5, 2)] });
  const [one] = spokenLayer(r.doc, r.layer).lines;
  assert.notEqual(laneOf(one.text), lane);
});

test('show the text and play the voice, each on or off, for the whole layer', () => {
  const { doc, textId } = withTitle('A\nB');
  const r = placeSpokenLines(doc, { from: textId, lines: [line('A', 1, 1), line('B', 1, 2)] });
  assert.deepEqual(spokenChoice(r.doc, r.layer), { show: true, play: true });
  const quiet = setSpokenChoice(r.doc, r.layer, { play: false });
  assert.deepEqual(spokenChoice(quiet, r.layer), { show: true, play: false });
  assert.ok(spokenLayer(quiet, r.layer).lines.every((l) => l.sound.visible === false && l.text.visible !== false));
  const unseen = setSpokenChoice(quiet, r.layer, { show: false, play: true });
  assert.deepEqual(spokenChoice(unseen, r.layer), { show: false, play: true });
  // The other clips are untouched.
  const others = (d) => d.patches.filter((p) => !isSpoken(p));
  assert.deepEqual(others(unseen), others(r.doc));
});

test('a caption read aloud becomes a spoken line in its look, from its cue start, no longer a caption', () => {
  const placed = placeCaptions(emptyProject(), [{ start: 1, end: 2, text: 'Hi' }, { start: 2.1, end: 3, text: 'Next' }]);
  const cue = clipsOf(placed.doc).find((p) => p.id === placed.ids[0]);
  const r = placeSpokenLines(placed.doc, { from: placed.ids[0], lines: [line('Hi', 1.6, 1)] });
  const [one] = spokenLayer(r.doc, r.layer).lines;
  assert.equal(one.text.timing.start, 1);
  assert.equal(one.text.caption, undefined, 'a line runs as long as its voice, so it is no cue to overlap the next');
  assert.equal(one.text.fy, cue.fy, 'the look of the caption is kept');
  assert.deepEqual(captionsOf(r.doc).map((c) => c.text), ['Next']);
  // Captions placed again (an import that replaces them) leave the spoken layer whole.
  const again = placeCaptions(r.doc, [{ start: 9, end: 10, text: 'New' }]);
  assert.equal(spokenLayer(again.doc, r.layer).lines[0].text.text, 'Hi');
  assert.ok(spokenLayer(again.doc, r.layer).lines[0].sound);
});

test('read again: the lines made afresh from the words in their clips, from the first line\'s start, the choices kept', () => {
  const { doc, textId } = withTitle('A\nB', 1);
  const r = placeSpokenLines(doc, { from: textId, lines: [line('A', 1, 1), line('B', 1, 2)] });
  let d = setSpokenChoice(r.doc, r.layer, { play: false });
  // A line's words edited, and the layer moved later on.
  const [one] = spokenLayer(d, r.layer).lines;
  d = { ...d, patches: d.patches.map((p) => (p.id === one.text.id ? { ...p, text: 'A changed' } : p)) };
  d = moveClip(d, one.text.id, { start: 4 });
  assert.deepEqual(wordsToRead(d, one.sound.id), ['A changed', 'B']);
  const again = placeSpokenLines(d, { from: one.sound.id, lines: [line('A changed', 2, 1), line('B', 1, 2)], voice: 'bf_emma' });
  assert.equal(again.layer, r.layer);
  const layer = spokenLayer(again.doc, r.layer);
  assert.equal(layer.lines.length, 2);
  assert.equal(layer.lines[0].text.text, 'A changed');
  assert.equal(layer.lines[0].text.timing.start, 4);
  assert.equal(layer.lines[0].text.timing.end, 6);
  assert.deepEqual(spokenChoice(again.doc, r.layer), { show: true, play: false });
  assert.equal(layer.voice, 'bf_emma');
  assert.equal(clipsOf(again.doc).filter(isSpoken).length, 4, 'nothing of the first reading is left over');
});

test('fewer lines when read again leave no clips behind', () => {
  const { doc, textId } = withTitle('A\nB\nC');
  const r = placeSpokenLines(doc, { from: textId, lines: [line('A', 1, 1), line('B', 1, 2), line('C', 1, 3)] });
  const again = placeSpokenLines(r.doc, { from: spokenLayer(r.doc, r.layer).lines[0].text.id, lines: [line('A', 1, 1)] });
  assert.equal(spokenLayer(again.doc, r.layer).lines.length, 1);
  assert.equal(clipsOf(again.doc).filter(isSpoken).length, 2);
});

test('a second layer is a layer of its own; nothing happens for a clip that is not there', () => {
  const { doc, textId } = withTitle('A');
  const r = placeSpokenLines(doc, { from: textId, lines: [line('A', 1, 1)] });
  assert.equal(newSpokenLayer(r.doc), 'sp2');
  const none = placeSpokenLines(r.doc, { from: 'nope', lines: [line('A', 1)] });
  assert.equal(none.doc, r.doc);
  assert.equal(none.layer, null);
  assert.deepEqual(wordsToRead(doc, textId), ['A']);
  assert.deepEqual(wordsToRead(doc, 'nope'), []);
  assert.equal(spokenFileName('sp3', 0), 'sp3-1.wav');
  assert.equal(spokenFileName('sp3', 1, 'k2x'), 'sp3-2-k2x.wav', 'a reading of its own is a file of its own');
});

test('lines with no words or no voice file are left out', () => {
  const { doc, textId } = withTitle('A');
  const r = placeSpokenLines(doc, { from: textId, lines: [line('A', 1, 1), { text: '  ', url: 'x', seconds: 1 }, { text: 'B', seconds: 1 }] });
  assert.equal(spokenLayer(r.doc, r.layer).lines.length, 1);
});

test('the lines a Recorder read land on Video\'s timeline as timed text clips, their voice in the take', () => {
  const cues = [{ start: 0.5, end: 2, text: 'Welcome' }, { start: 2.3, end: 4, text: 'Today' }];
  const r = placeSpokenCues(emptyProject(), cues, { at: 10 });
  const layer = spokenLayer(r.doc, r.layer);
  assert.equal(layer.lines.length, 2);
  assert.deepEqual(layer.lines[0].text.timing, { start: 10.5, end: 12 });
  assert.equal(layer.lines[1].text.text, 'Today');
  assert.equal(layer.lines[0].sound, null, 'no second voice: it is in the recorded sound');
  assert.equal(layer.lines[0].text.caption, undefined);
  assert.equal(layer.lines[0].text.spoken.recorded, true);
  assert.equal(laneClips(r.doc, 'v', laneOf(layer.lines[0].text)).length, 2);
  const hidden = placeSpokenCues(emptyProject(), cues, { show: false });
  assert.ok(spokenLayer(hidden.doc, hidden.layer).lines.every((l) => l.text.visible === false));
  assert.equal(placeSpokenCues(emptyProject(), []).layer, null);
});

test('a studio title keeps its lines and how they are read; other sources have neither', () => {
  const it = normaliseItem({ kind: 'title', id: 'title-1', text: 'Name', lines: 'One\nTwo', spoken: { show: false, play: true, voice: 'af_heart', speed: 1.5 } });
  assert.equal(it.lines, 'One\nTwo');
  assert.deepEqual(it.spoken, { show: false, play: true, voice: 'af_heart', speed: 1.5 });
  const bare = normaliseItem({ kind: 'title', id: 'title-2' });
  assert.equal(bare.lines, '');
  assert.deepEqual(bare.spoken, { show: true, play: true, voice: '', speed: 1 });
  assert.equal(normaliseItem({ kind: 'title', lines: 'x'.repeat(MAX_LINES_TEXT + 50) }).lines.length, MAX_LINES_TEXT);
  assert.equal(normaliseSpoken({ speed: 9 }).speed, 1);
  assert.equal(normaliseLines(5), '');
  const cam = normaliseItem({ kind: 'camera', lines: 'x', spoken: {} });
  assert.equal(cam.lines, undefined);
  assert.equal(cam.spoken, undefined);
  // Kept through an update, as the Properties column makes one.
  let st = emptyStack();
  ({ stack: st } = addItem(st, { kind: 'title', text: 'T' }));
  const id = st.items.find((x) => x.kind === 'title').id;
  st = updateItem(st, id, { lines: 'A\nB', spoken: { ...normaliseSpoken(), play: false } });
  const after = st.items.find((x) => x.id === id);
  assert.equal(after.lines, 'A\nB');
  assert.equal(after.spoken.play, false);
});
