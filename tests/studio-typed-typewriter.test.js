// Words typed in the text to speech box of the Recorder and Live Broadcast
// by the typewriter (S20): read aloud, each line types over its voice, the
// last letter as the voice ends ("With the voice") or the voice a moment
// first ("Voice first"), worked out from Video's typingWithVoice; shown only,
// at the standard typing speed, the line held long enough to type and read.
// The frame draws only the letters typed so far, where the whole line will
// stand, and says how many it drew. The window's program module is bundled
// here as studio-program-draw does.
import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import {
  studioTyping, typedShownSeconds, lettersTyped, rowsTyped, readingSeconds, typingWithVoice, VOICE_FIRST_SECONDS, TYPED_LINE,
} from '../packages/studio/src/spoken.js';
import { CAPTION_SPEED, NEW_CAPTION, captionSeconds } from '../packages/studio/src/project.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const LINE = 'Welcome to the new starters';

test('read aloud with the voice, the last letter lands as the voice ends', () => {
  const typing = studioTyping(LINE, { seconds: 3, play: true, typing: 'with' });
  const { pace } = typingWithVoice({ timing: { start: 0, end: 3 } }, { play: true, typing: 'with' });
  assert.equal(typing.lead, 0);
  assert.ok(Math.abs(typing.cps - LINE.length / pace) < 1e-9);
  assert.equal(lettersTyped(LINE, typing, 0), 0);
  assert.ok(lettersTyped(LINE, typing, 1.5) > 0 && lettersTyped(LINE, typing, 1.5) < LINE.length, 'half way through the voice, part of the line');
  assert.equal(lettersTyped(LINE, typing, 2.95), LINE.length, 'whole a moment before the voice ends');
});

test('voice first: the typing starts a moment after the voice', () => {
  const typing = studioTyping(LINE, { seconds: 3, play: true, typing: 'first' });
  assert.equal(typing.lead, VOICE_FIRST_SECONDS);
  assert.equal(lettersTyped(LINE, typing, VOICE_FIRST_SECONDS - 0.01), 0);
  assert.equal(lettersTyped(LINE, typing, 2.95), LINE.length);
});

test('shown only, the words type at the standard speed and stay long enough to be read', () => {
  const typing = studioTyping(LINE, { seconds: 2, play: false });
  assert.deepEqual(typing, { lead: NEW_CAPTION.leadIn, cps: CAPTION_SPEED.standard });
  // No voice played (none made) types the same way.
  assert.deepEqual(studioTyping(LINE, { seconds: 0, play: true }), typing);
  const long = 'A line long enough that typing it at sixteen letters a second takes longer than reading it';
  assert.ok(typedShownSeconds(long, { typewriter: true }) >= captionSeconds(long, { cps: CAPTION_SPEED.standard, leadIn: NEW_CAPTION.leadIn }));
  assert.ok(typedShownSeconds(long, { typewriter: true }) >= readingSeconds(long));
  assert.equal(typedShownSeconds(long, { typewriter: false }), readingSeconds(long), 'without the typewriter, as before');
});

test('letters are counted as characters, and nothing typed is all of it', () => {
  const urdu = 'خوش آمدید';
  assert.equal(lettersTyped(urdu, null, 0), [...urdu].length);
  assert.equal(lettersTyped('a😀b', { lead: 0, cps: 1 }, 2), 2, 'an emoji is one letter, never half of one');
});

test('rows are cut to the letters typed, the break between rows taking its space', () => {
  const rows = ['Welcome to the', 'new starters'];
  assert.deepEqual(rowsTyped(rows, 0), ['', '']);
  assert.deepEqual(rowsTyped(rows, 7), ['Welcome', '']);
  assert.deepEqual(rowsTyped(rows, 14), ['Welcome to the', '']);
  assert.deepEqual(rowsTyped(rows, 15), ['Welcome to the', '']);
  assert.deepEqual(rowsTyped(rows, 18), ['Welcome to the', 'new']);
  assert.deepEqual(rowsTyped(rows, 99), rows);
});

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

function programWith(Program, media) {
  const texts = [];
  const ctx = new Proxy({ measureText: (s) => ({ width: String(s).length * 10 }), fillText: (s) => texts.push(String(s)) }, {
    get: (o, k) => (k in o ? o[k] : () => {}),
    set: (o, k, v) => { o[k] = v; return true; },
  });
  const p = Object.create(Program.prototype);
  Object.assign(p, { ctx, canvas: { width: 1920, height: 1080, dataset: {} }, media, stack: { video: 'on', items: [] }, accent: '#e03b4c', frames: 0, listeners: new Set(), fade: null });
  return { p, texts };
}

test('the frame draws only the letters typed so far, and says how many', async () => {
  const { Program } = await loadProgram();
  const at = performance.now() - 1000;
  const typing = { at, lead: 0, cps: 7 };
  const media = { drawable: () => null, spokenLine: (id) => (id === TYPED_LINE ? LINE : null), spokenTyping: (id) => (id === TYPED_LINE ? typing : null) };
  const { p, texts } = programWith(Program, media);
  p.draw();
  const drawn = Number(p.canvas.dataset.typedLetters);
  assert.ok(drawn >= 7 && drawn < LINE.length, `about a second's letters drawn, ${drawn}`);
  assert.equal(texts.join(''), [...LINE].slice(0, drawn).join(''));
  assert.equal(p.canvas.dataset.typed, '1');
});

test('shown whole, the frame draws the whole line as before', async () => {
  const { Program } = await loadProgram();
  const media = { drawable: () => null, spokenLine: (id) => (id === TYPED_LINE ? LINE : null), spokenTyping: () => null };
  const { p, texts } = programWith(Program, media);
  p.draw();
  assert.equal(p.canvas.dataset.typedLetters, String(LINE.length));
  assert.deepEqual(texts, [LINE]);
});

test('the studio media keeps how a line types until it is taken down', async () => {
  const out = await esbuild.build({
    stdin: { contents: "export { StudioMedia } from './media.js';", resolveDir: path.join(here, '../apps/desktop/renderer/studio'), loader: 'js' },
    bundle: true, format: 'esm', platform: 'node', write: false, logLevel: 'silent',
    plugins: [{ name: 'stub', setup(b) { b.onResolve({ filter: /^@rutba\/office-ui$/ }, () => ({ path: 'ui-kit', namespace: 'stub' })); b.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ contents: "export { t } from './messages.js'; export const tn = () => ''; export const formatBytes = () => '';", resolveDir: path.join(here, '../packages/office-ui/src'), loader: 'js' })); } }],
  });
  const { StudioMedia: Media } = await import(`data:text/javascript;base64,${Buffer.from(out.outputFiles[0].text).toString('base64')}`);
  const m = Object.create(Media.prototype);
  m.spokenLines = new Map();
  m.spokenTypings = new Map();
  m.setSpokenLine(TYPED_LINE, LINE, { lead: 0, cps: 10 });
  assert.equal(m.spokenLine(TYPED_LINE), LINE);
  assert.equal(m.spokenTyping(TYPED_LINE).cps, 10);
  assert.ok(Number.isFinite(m.spokenTyping(TYPED_LINE).at));
  m.setSpokenLine(TYPED_LINE, LINE);
  assert.equal(m.spokenTyping(TYPED_LINE), null, 'a line shown whole has no typing');
  m.setSpokenLine(TYPED_LINE, LINE, { lead: 0, cps: 10 });
  m.setSpokenLine(TYPED_LINE, null);
  assert.equal(m.spokenTyping(TYPED_LINE), null);
});
