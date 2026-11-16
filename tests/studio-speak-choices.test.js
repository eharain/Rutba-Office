// Spoken words made simpler (D-2026-11-16-01): one choice of three, the same
// in Video, the Recorder and Live (shown and read aloud, read aloud only,
// shown only), words typed in Video landing at the playhead in one step, and
// words typed in the studio said live, their band over the whole frame and
// their voice keeping a recording's sound even when said after Record.
import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import {
  SPEAK_CHOICES, speakChoiceOf, speakSwitches, readingSeconds, MIN_READING, TYPED_LINE,
  placeTypedLines, placeSpokenLines, spokenChoice, spokenLayer, setSpokenChoice, wordsToRead, LINE_GAP,
  typedLinesOf, TYPED_LINE_MOST, spokenCueLog, placeSpokenCues, soundedTakes,
} from '../packages/studio/src/spoken.js';
import { emptyProject, addTitle, clipsOf, laneOf } from '../packages/studio/src/project.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const voiced = (text, seconds, n) => ({ text, seconds, url: `C:/p/Spoken lines/sp1-${n}.wav`, name: `sp1-${n}.wav`, size: 100 });

test('the three choices are the two switches every spoken layer and title already stores', () => {
  assert.deepEqual(SPEAK_CHOICES.map((c) => c.id), ['both', 'read', 'show']);
  assert.equal(speakChoiceOf({ show: true, play: true }), 'both');
  assert.equal(speakChoiceOf({ show: false, play: true }), 'read');
  assert.equal(speakChoiceOf({ show: true, play: false }), 'show');
  // Stored before as neither: no choice is marked, nothing is changed.
  assert.equal(speakChoiceOf({ show: false, play: false }), null);
  // Missing switches are on, as normaliseSpoken reads them.
  assert.equal(speakChoiceOf({}), 'both');
  assert.equal(speakChoiceOf(), 'both');
  for (const c of SPEAK_CHOICES) assert.equal(speakChoiceOf(speakSwitches(c.id)), c.id);
  assert.deepEqual(speakSwitches('nonsense'), { show: true, play: true });
  assert.throws(() => { SPEAK_CHOICES[0].show = false; });
});

test('a line shown without its voice stays up about as long as it takes to read it', () => {
  assert.equal(readingSeconds(''), 0);
  assert.equal(readingSeconds('   '), 0);
  assert.equal(readingSeconds('Hi'), MIN_READING);
  // Ten words at 2.5 a second.
  assert.equal(readingSeconds('one two three four five six seven eight nine ten'), 4);
  assert.equal(readingSeconds('one two three four five six seven eight nine ten', { speed: 2 }), 2);
  assert.equal(readingSeconds('one two three four five six seven eight nine ten', { speed: 0 }), 4, 'a speed of nothing is speed 1');
  // A script written without spaces is counted by its characters, six to a word.
  const long = '\u6587'.repeat(60);
  assert.equal(readingSeconds(long), 4);
  // Urdu with spaces is counted by its words.
  assert.ok(readingSeconds('\u06CC\u06C1 \u0627\u06CC\u06A9 \u0645\u062B\u0627\u0644 \u06C1\u06D2') >= MIN_READING);
});

test('typed words in Video land at the playhead as a spoken layer, its words and its voice, one line after another', () => {
  const base = addTitle(emptyProject(), 'title', { at: 0, text: 'Opening' }).doc;
  const r = placeTypedLines(base, { at: 1, lines: [voiced('Hello there', 1.2, 1), voiced('Second line', 0.8, 2)], voice: 'v1', speed: 1.25 });
  assert.ok(r.layer);
  assert.equal(r.start, 1);
  assert.equal(r.end, 1 + 1.2 + LINE_GAP + 0.8);
  const info = spokenLayer(r.doc, r.layer);
  assert.equal(info.lines.length, 2);
  assert.deepEqual(info.lines.map((l) => l.text.text), ['Hello there', 'Second line']);
  assert.deepEqual(info.lines.map((l) => l.sound.url), ['C:/p/Spoken lines/sp1-1.wav', 'C:/p/Spoken lines/sp1-2.wav']);
  assert.equal(info.voice, 'v1');
  assert.equal(info.speed, 1.25);
  assert.deepEqual(spokenChoice(r.doc, r.layer), { show: true, play: true });
  // Over the title already there, never on its lane.
  const words = info.lines.map((l) => l.text);
  assert.ok(words.every((p) => laneOf(p) > laneOf(clipsOf(base)[0])));
  // Nothing of the seed is left behind.
  assert.equal(clipsOf(r.doc).filter((p) => p.type === 'text' && !p.spoken).length, 1);
  assert.deepEqual(wordsToRead(r.doc, words[0].id), ['Hello there', 'Second line']);
});

test('read aloud only hides the words; show only with a voice keeps the voice made but silent', () => {
  const doc = emptyProject();
  const read = placeTypedLines(doc, { at: 0, lines: [voiced('One', 1, 1)], voice: 'v', show: false, play: true });
  assert.deepEqual(spokenChoice(read.doc, read.layer), { show: false, play: true });
  const shown = placeTypedLines(doc, { at: 0, lines: [voiced('One', 1, 1)], voice: 'v', show: true, play: false });
  assert.deepEqual(spokenChoice(shown.doc, shown.layer), { show: true, play: false });
  // And the person can change their mind afterwards, as with any spoken layer.
  assert.deepEqual(spokenChoice(setSpokenChoice(shown.doc, shown.layer, speakSwitches('both')), shown.layer), { show: true, play: true });
});

test('shown only with no voice: the words alone, each for its reading time, known to have no voice yet', () => {
  const doc = emptyProject();
  const lines = ['A first line to read', 'Then a second'].map((text) => ({ text, seconds: readingSeconds(text) }));
  const r = placeTypedLines(doc, { at: 3, lines, show: true, play: false });
  assert.equal(r.ids.length, 2);
  const clips = clipsOf(r.doc);
  assert.deepEqual(clips.map((p) => p.type), ['text', 'text']);
  assert.equal(clips[0].timing.start, 3);
  assert.equal(clips[1].timing.start, Math.round((3 + lines[0].seconds + LINE_GAP) * 1000) / 1000);
  assert.ok(clips.every((p) => p.spoken.shown && p.spoken.layer === r.layer && p.visible !== false));
  assert.deepEqual(spokenChoice(r.doc, r.layer), { show: true, play: false });
  // Read again later with a voice: the voice is made and kept silent, as chosen, until the person picks read.
  const again = placeSpokenLines(r.doc, { from: r.ids[0], lines: [voiced('A first line to read', 1, 1), voiced('Then a second', 1, 2)], voice: 'v' });
  assert.deepEqual(spokenChoice(again.doc, r.layer), { show: true, play: false });
  assert.equal(spokenLayer(again.doc, r.layer).lines.filter((l) => l.sound).length, 2);
});

test('typed words to be read with no voice made, or no words, place nothing', () => {
  const doc = emptyProject();
  assert.equal(placeTypedLines(doc, { at: 0, lines: [{ text: 'No voice', seconds: 1 }], play: true }).doc, doc);
  assert.equal(placeTypedLines(doc, { at: 0, lines: [{ text: '  ', seconds: 1 }], play: false }).doc, doc);
  assert.equal(placeTypedLines(doc, { at: 0, lines: [] }).ids.length, 0);
});

test('typed words go over every picture there is while they run', () => {
  let doc = emptyProject();
  doc = { ...doc, patches: [...doc.patches, { id: 'image-1', type: 'image', url: 'C:/a.png', timing: { start: 0, end: 10 }, lane: 0 }] };
  doc = addTitle(doc, 'lowerThird', { at: 0, text: 'Name' }).doc;
  const top = Math.max(...clipsOf(doc).filter((p) => p.type !== 'sound').map(laneOf));
  const r = placeTypedLines(doc, { at: 1, lines: [{ text: 'Over all', seconds: 2 }], play: false });
  assert.ok(laneOf(clipsOf(r.doc).find((p) => p.id === r.ids[0])) > top);
});

test('the studio keeps typed words under a name no source can have', () => {
  assert.equal(TYPED_LINE, 'typed');
  assert.doesNotMatch(TYPED_LINE, /^[a-z]+-\d+$/);
});

test('typed words too long for the band are said a sentence at a time, and a long sentence a few words at a time', () => {
  assert.deepEqual(typedLinesOf('Short one\n\nAnother'), ['Short one', 'Another']);
  const para = 'This is the first sentence of a paragraph typed with no line break at all. And this is the second one, which follows it.';
  assert.ok(para.length > TYPED_LINE_MOST);
  const said = typedLinesOf(para, { language: 'en' });
  assert.equal(said.length, 2);
  assert.match(said[0], /^This is the first sentence/);
  assert.match(said[1], /^And this is the second/);
  const run = Array.from({ length: 40 }, (_, i) => `word${i}`).join(' ');
  const parts = typedLinesOf(run);
  assert.ok(parts.length > 1 && parts.every((p) => p.length <= TYPED_LINE_MOST));
  assert.equal(parts.join(' '), run, 'nothing lost or doubled');
  // One word longer than a band stays whole: nothing is cut inside a word.
  assert.deepEqual(typedLinesOf('x'.repeat(100)), ['x'.repeat(100)]);
});

test('a line read while recording says whether its voice is in the sound, and Video takes it at its word', () => {
  const log = spokenCueLog();
  log.begin(1, 'Read aloud', true, true);
  log.end(2);
  log.begin(3, 'Shown alone', true, false);
  log.end(4);
  const cues = log.take(5);
  assert.deepEqual(cues.map((c) => c.play), [true, false]);
  // Shown alone puts nothing in the sound: a take is not made sounded by it.
  const takes = [{ role: 'screen', carrier: true, hasAudio: false }];
  assert.equal(soundedTakes(takes, cues.slice(1))[0].hasAudio, false);
  assert.equal(soundedTakes(takes, cues)[0].hasAudio, true);
  // Old cues, with no word of their voice, read as voiced, as before.
  assert.equal(soundedTakes(takes, [{ start: 0, end: 1, text: 'x', show: true }])[0].hasAudio, true);
  // On Video's timeline, the voiced one is a recorded line; the one shown alone is known to have no voice.
  const a = placeSpokenCues(emptyProject(), cues.slice(0, 1)).doc;
  const b = placeSpokenCues(emptyProject(), cues.slice(1)).doc;
  const la = clipsOf(a)[0].spoken.layer;
  const lb = clipsOf(b)[0].spoken.layer;
  assert.deepEqual(spokenChoice(a, la), { show: true, play: true });
  assert.deepEqual(spokenChoice(b, lb), { show: true, play: false });
  // Each cue keeps its own showing: a line read aloud only comes in hidden.
  const hidden = placeSpokenCues(emptyProject(), [{ start: 0, end: 1, text: 'Heard, not seen', show: false, play: true }]).doc;
  assert.equal(clipsOf(hidden)[0].visible, false);
});

test('a recording with lines read aloud and lines only shown hands them over as two layers, each offering what it can', () => {
  const cues = [
    { start: 0, end: 1, text: 'Read aloud first', show: true, play: true },
    { start: 1.5, end: 2.5, text: 'Only shown', show: true, play: false },
    { start: 3, end: 4, text: 'Read aloud again', show: true, play: true },
  ];
  const r = placeSpokenCues(emptyProject(), cues);
  assert.equal(r.layers.length, 2);
  const [voiced, shown] = r.layers;
  assert.equal(r.layer, voiced, 'the layer of the first line');
  assert.deepEqual(spokenLayer(r.doc, voiced).lines.map((l) => l.text.text), ['Read aloud first', 'Read aloud again']);
  assert.deepEqual(spokenLayer(r.doc, shown).lines.map((l) => l.text.text), ['Only shown']);
  assert.deepEqual(spokenChoice(r.doc, voiced), { show: true, play: true });
  assert.deepEqual(spokenChoice(r.doc, shown), { show: true, play: false });
  // One kind alone is one layer, as before.
  assert.equal(placeSpokenCues(emptyProject(), cues.slice(0, 1)).layers.length, 1);
});

// ── a line's turn ───────────────────────────────────────────────────────────

test('a line made for Show only, read aloud after the choice changed, is held for its reading time and plays nothing', async () => {
  const { sayLine } = await import('../apps/desktop/renderer/studio/say-line.js');
  let spoken = 0;
  const media = { speak: async () => { spoken += 1; return true; } };
  const silent = { samples: new Float32Array(0), rate: 24000, seconds: 0.4 };
  const t0 = performance.now();
  const voiced = await sayLine({ media, sound: silent, play: true, alive: () => true });
  const took = performance.now() - t0;
  assert.equal(voiced, false);
  assert.equal(spoken, 0, 'no empty sound sent to the output (it would count as a voice there)');
  assert.ok(took >= 380, `held for its reading time (${Math.round(took)} ms)`);
  // A line with its voice is played; a line not to be read is only waited out.
  assert.equal(await sayLine({ media, sound: { samples: new Float32Array(10), rate: 24000, seconds: 0.01 }, play: true, alive: () => true }), true);
  assert.equal(spoken, 1);
  assert.equal(await sayLine({ media, sound: { samples: new Float32Array(10), rate: 24000, seconds: 0.01 }, play: false, alive: () => true }), false);
  assert.equal(spoken, 1);
  // Stopped meanwhile: it ends at once.
  const t1 = performance.now();
  await sayLine({ media, sound: { ...silent, seconds: 5 }, play: false, alive: () => false });
  assert.ok(performance.now() - t1 < 200);
});

// ── the studio's Program ────────────────────────────────────────────────────

async function loadProgram() {
  const out = await esbuild.build({
    entryPoints: [path.join(here, '../apps/desktop/renderer/studio/program.js')],
    bundle: true,
    format: 'esm',
    platform: 'node',
    write: false,
    logLevel: 'silent',
    plugins: [{
      name: 'ui-kit',
      setup(b) {
        b.onResolve({ filter: /^@rutba\/office-ui$/ }, () => ({ path: 'ui-kit', namespace: 'stub' }));
        b.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ contents: "export { t } from './messages.js';", resolveDir: path.join(here, '../packages/office-ui/src'), loader: 'js' }));
      },
    }],
  });
  return import(`data:text/javascript;base64,${Buffer.from(out.outputFiles[0].text).toString('base64')}`);
}

/** A canvas whose drawing is written down: every fillText with where it went. */
function fakeCanvas(width, height) {
  const texts = [];
  const rects = [];
  // Text as wide as its font is large, capitals wider than small letters, as a real face draws them.
  const o = { font: '10px sans-serif' };
  const size = () => Number(/(\d+)px/.exec(o.font)?.[1] || 10);
  const widthOf = (s) => [...String(s)].reduce((w, c) => w + (c === ' ' ? 0.28 : c === c.toUpperCase() && c !== c.toLowerCase() ? 0.72 : 0.52), 0) * size();
  Object.assign(o, {
    fillText: (s, x, y, max) => texts.push({ s, x, y, max, width: widthOf(s), size: size() }),
    fillRect: (x, y, w, h) => rects.push({ x, y, w, h }),
    measureText: (s) => ({ width: widthOf(s) }),
  });
  const ctx = new Proxy(o, { get: (t, k) => (k in t ? t[k] : () => {}), set: (t, k, v) => { t[k] = v; return true; } });
  return { width, height, dataset: {}, getContext: () => ctx, texts, rects };
}

function programWith(Program, { canvas, line, stack }) {
  const p = Object.create(Program.prototype);
  p.canvas = canvas;
  p.ctx = canvas.getContext('2d');
  p.stack = stack;
  p.frames = 0;
  p.listeners = new Set();
  p.accent = '#3b82f6';
  p.media = { spokenLine: (id) => (id === TYPED_LINE ? line : null), drawable: () => null };
  return p;
}

const emptyStack = (video = 'on') => ({ video, items: [], sounds: [] });

test('typed words being said are drawn across the foot of the frame, over every source', async () => {
  const { Program } = await loadProgram();
  const canvas = fakeCanvas(1280, 720);
  programWith(Program, { canvas, line: 'Welcome, everyone', stack: emptyStack() }).draw();
  assert.equal(canvas.dataset.typed, '1');
  const said = canvas.texts.find((x) => x.s === 'Welcome, everyone');
  assert.ok(said, 'the words are drawn');
  assert.ok(said.y > 720 * 0.8 && said.y < 720 * 0.93, 'in the band at the foot');
  // The band is wider than a title's (TITLE_BOX is under half the frame).
  assert.ok(canvas.rects.some((r) => r.w > 1280 * 0.9));
  // Nothing said: no band.
  const quiet = fakeCanvas(1280, 720);
  programWith(Program, { canvas: quiet, line: null, stack: emptyStack() }).draw();
  assert.equal(quiet.dataset.typed, '0');
  // A frame turned off shows black only.
  const off = fakeCanvas(1280, 720);
  programWith(Program, { canvas: off, line: 'Hidden', stack: emptyStack('off') }).draw();
  assert.equal(off.dataset.typed, '0');
  assert.ok(!off.texts.some((x) => x.s === 'Hidden'));
  // The tall picture has a band of its own shape.
  const tall = fakeCanvas(1080, 1920);
  programWith(Program, { canvas: tall, line: 'Tall', stack: emptyStack() }).draw();
  assert.equal(tall.dataset.typed, '1');
  assert.ok(tall.rects.some((r) => r.w > 1080 * 0.85 && r.h < 1920 * 0.1));
});

test('the longest typed line, all capitals, fits its two rows on a tall frame and on a wide one, never squeezed', async () => {
  const { Program } = await loadProgram();
  const words = ['WELCOME', 'BACK', 'EVERYONE', 'THE', 'NEXT', 'PART', 'STARTS', 'NOW', 'WITH', 'QUESTIONS', 'FROM', 'YOU', 'ALL', 'TONIGHT', 'AT', 'EIGHT'];
  let line = '';
  for (const w of words) { const next = line ? `${line} ${w}` : w; if (next.length > TYPED_LINE_MOST) break; line = next; }
  line = `${line} ${'W'.repeat(Math.max(1, TYPED_LINE_MOST - line.length - 1))}`;
  assert.equal(line.length, TYPED_LINE_MOST);
  assert.deepEqual(typedLinesOf(line), [line], 'said as one line');
  for (const [w, h] of [[1080, 1920], [720, 1280], [1280, 720], [1920, 1080]]) {
    const canvas = fakeCanvas(w, h);
    programWith(Program, { canvas, line, stack: emptyStack() }).draw();
    const rows = canvas.texts.filter((x) => x.s);
    assert.ok(rows.length >= 1 && rows.length <= 2, `${w}x${h}: one or two rows`);
    assert.equal(rows.map((r) => r.s).join(' '), line, `${w}x${h}: every word drawn, none lost`);
    for (const r of rows) assert.ok(r.width <= r.max + 0.5, `${w}x${h}: a row of ${Math.round(r.width)} px fits its ${Math.round(r.max)} px without squeezing`);
    const band = Math.round(h * (h > w ? 0.09 : 0.13));
    assert.ok(rows.every((r) => r.size >= Math.max(10, Math.round(band * 0.1))), `${w}x${h}: no smaller than the floor`);
  }
});

test('a word wider than the band, a long link or words written without spaces, is made smaller to fit, not squeezed', async () => {
  const { Program } = await loadProgram();
  const link = `https://example.org/${'a'.repeat(67)}`;
  assert.equal(link.length, 87);
  const unspaced = '\u6587'.repeat(40);
  for (const line of [link, unspaced]) {
    const canvas = fakeCanvas(1080, 1920);
    programWith(Program, { canvas, line, stack: emptyStack() }).draw();
    const rows = canvas.texts.filter((x) => x.s);
    assert.equal(rows.length, 1, 'one row: nothing to break it at');
    assert.equal(rows[0].s, line);
    assert.ok(rows[0].width <= rows[0].max + 0.5, `a row of ${Math.round(rows[0].width)} px fits its ${Math.round(rows[0].max)} px`);
    assert.ok(rows[0].size < Math.round(1920 * 0.09 * 0.3), 'smaller than a row that fits at first');
  }
  // Drawn again, the band is not measured again.
  const canvas = fakeCanvas(1080, 1920);
  const p = programWith(Program, { canvas, line: link, stack: emptyStack() });
  p.draw();
  const ctx = canvas.getContext('2d');
  let measured = 0;
  const measure = ctx.measureText;
  ctx.measureText = (s) => { measured += 1; return measure(s); };
  p.draw();
  assert.equal(measured, 0);
});

/** A recorder already stopped, and its writer's answer. */
const partOf = (done) => ({ rec: { state: 'inactive', stop() {} }, stopped: Promise.resolve(), writer: { done: async () => done, written: done.written } });

async function stopWith(Program, { heard, spokeAtRecord, spokenNow, main = true }) {
  const p = Object.create(Program.prototype);
  p.media = { spoken: spokenNow };
  const m = main ? partOf({ written: 100, failed: null }) : null;
  p.recording = { parts: m ? [m] : [], main: m, side: [], clock: { elapsed: () => 5000, paused: false }, path: main ? 'D:/films/p.webm' : null, heard, type: 'video/webm;codecs=vp9', soundOnly: false, spoke: spokeAtRecord };
  return p.stop();
}

test('a voice said into the output after Record keeps the recording\'s sound', async () => {
  const { Program } = await loadProgram();
  assert.equal((await stopWith(Program, { heard: false, spokeAtRecord: 2, spokenNow: 3 })).heard, true, 'said after Record');
  assert.equal((await stopWith(Program, { heard: false, spokeAtRecord: 2, spokenNow: 2 })).heard, false, 'nothing said since');
  assert.equal((await stopWith(Program, { heard: true, spokeAtRecord: 0, spokenNow: 0 })).heard, true, 'heard at Record stays heard');
  // A recording from before this change (no count kept) reads as nothing said since.
  assert.equal((await stopWith(Program, { heard: false, spokeAtRecord: undefined, spokenNow: undefined })).heard, false);
});
