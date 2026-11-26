// Words typed under the frame that type in, letter by letter, in the Recorder
// and in Live Broadcast (2.28.0, S20), said with a real added voice. The
// typed band of the picture carries the count of letters drawn so far
// (data-typed-letters on each canvas): a line said "With the voice" starts
// with few of its letters and reaches all of them as the voice ends; "Voice
// first" starts the voice a moment before the first letter; "Show only"
// types it at the standard speed with no voice. A frame keeps being drawn
// while the voice is made and said (B17: data-frames keeps rising).
// Run as the second half of the real-voice check (RUTBA_VERIFY_ONLY=real-voice
// with RUTBA_GATE_PACKS set), which is given the speech packs.

import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, CHECK_SOUNDS, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { letGoOfStudioWindows } from './verify-let-go.js';
import { GATE_PACKS_ENV } from './gate-packs.js';

const WORDS = 'Welcome to the new starters, and thank you for joining us this morning.';
const LETTERS = [...WORDS].length;

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyTypedTypewriter(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  if (!process.env[GATE_PACKS_ENV]) {
    check('typed typewriter: this run was given the speech packs', false, `${GATE_PACKS_ENV} is not set: name the packages release folder (dist/v<version>)`);
    return;
  }
  for (const app of ['recorder', 'broadcast']) {
    const name = app === 'recorder' ? 'the Recorder' : 'Live Broadcast';
    let kept = null;
    let js = null;
    try {
      const win = await open(app);
      js = (code) => win.webContents.executeJavaScript(code);
      kept = await seedSetup(js, { app, stack: { ...emptyStack(), sounds: CHECK_SOUNDS(false) }, studio: { size: 720, fps: 30, countdown: 0, folder: dir, liveCopy: false, minimise: false } });
      await win.webContents.reload();
      await until(() => js(`Boolean(document.querySelector('.st-canvas'))`), `the ${app} window`, 15000);
      await js(PANEL('main'));
      await until(() => js(IN_PANEL('main')), 'the Record or Live panel', 6000).catch(() => {});
      await js(`document.querySelector('.st-typed-open')?.click(), 1`);
      const box = await ok(until(() => js(`Boolean(document.querySelector('.st-typed .spoken-box-text')) && Boolean(document.querySelector('.st-typed select'))`), 'the Text to speech box with its voice', 15000));
      await js(`(() => { const n = document.querySelector('.st-typed .spoken-box-text'); if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(n, ${JSON.stringify(WORDS)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
      await js(`document.querySelector('.st-typed input.spoken-choice-both')?.click(), 1`);
      await wait(300);
      // The Typewriter box, then its timing.
      await js(`(() => { const c = document.querySelector('.st-typed input.st-typed-typewriter'); if (c && !c.checked) c.click(); return 1; })()`);
      await until(() => js(`Boolean(document.querySelector('.st-typed select.st-typed-timing'))`), 'the typing choice', 5000).catch(() => {});
      const setTiming = (value) => js(`(() => { const s = document.querySelector('.st-typed select.st-typed-timing'); if (!s) return false; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ${JSON.stringify(value)}); s.dispatchEvent(new Event('change', { bubbles: true })); return s.value === ${JSON.stringify(value)}; })()`);
      const canvasState = () => js(`(() => { const c = [...document.querySelectorAll('.st-canvas')].find((x) => x.dataset.typed === '1') || document.querySelector('.st-canvas'); return c ? { typed: c.dataset.typed, letters: Number(c.dataset.typedLetters || 0), frames: Number(c.dataset.frames || 0) } : null; })()`);
      // Says the line once and samples the letters drawn while it is said.
      const sayAndWatch = async () => {
        // Every change of the letters drawn is kept by the page itself: the last letter is on the picture for a tenth
        // of a second before the line leaves, too short for a check that looks four times a second.
        await js(`(() => { window.__typedSeen?.stop?.(); const seen = { counts: [] }; const watch = new MutationObserver(() => { const c = [...document.querySelectorAll('.st-canvas')].find((x) => x.dataset.typed === '1'); if (c) { const n = Number(c.dataset.typedLetters || 0); if (seen.counts[seen.counts.length - 1] !== n) seen.counts.push(n); } }); document.querySelectorAll('.st-canvas').forEach((c) => watch.observe(c, { attributes: true, attributeFilter: ['data-typed-letters', 'data-typed'] })); seen.stop = () => watch.disconnect(); window.__typedSeen = seen; return 1; })()`);
        const pressed = await js(`(() => { const b = document.querySelector('.st-typed .spoken-box-say'); if (!b || b.disabled) return false; b.click(); return true; })()`);
        const samples = [];
        const started = Date.now();
        await until(async () => {
          const st = await canvasState();
          if (st) samples.push(st);
          const done = await js(`[...document.querySelectorAll('.st-canvas')].every((c) => c.dataset.typed === '0') && !document.querySelector('.st-typed-now') && Boolean(document.querySelector('.st-typed .spoken-box-say'))`);
          return samples.some((x) => x.typed === '1') && done;
        }, 'the line to be typed and said', 90000).catch(() => {});
        const drawn = samples.filter((x) => x.typed === '1');
        const counts = await js(`(() => { const s = window.__typedSeen; s?.stop?.(); return s ? s.counts : []; })()`);
        return {
          pressed,
          seconds: Math.round((Date.now() - started) / 100) / 10,
          drawnSamples: drawn.length,
          changes: counts.length,
          first: counts[0] ?? null,
          most: counts.length ? Math.max(...counts) : null,
          rising: counts.every((n, i) => i === 0 || n >= counts[i - 1]),
          framesRose: drawn.length > 1 && drawn[drawn.length - 1].frames > drawn[0].frames,
        };
      };
      const timedWith = await setTiming('with');
      const withVoice = await sayAndWatch();
      check(`typed typewriter: in ${name}, a line said With the voice types in letter by letter on the picture, from a few letters to all ${LETTERS}, while frames keep being drawn`,
        box && timedWith && withVoice.pressed && withVoice.drawnSamples > 2 && withVoice.first < LETTERS && withVoice.most === LETTERS && withVoice.rising && withVoice.framesRose,
        JSON.stringify({ box, timedWith, ...withVoice }));
      const timedFirst = await setTiming('first');
      const voiceFirst = await sayAndWatch();
      check(`typed typewriter: in ${name}, Voice first also types the line in to its last letter`,
        timedFirst && voiceFirst.pressed && voiceFirst.drawnSamples > 2 && voiceFirst.first < LETTERS && voiceFirst.most === LETTERS && voiceFirst.rising,
        JSON.stringify({ timedFirst, ...voiceFirst }));
      // Show only: no voice, the standard speed.
      await js(`document.querySelector('.st-typed input.spoken-choice-show')?.click(), 1`);
      await wait(300);
      const showOnly = await js(`(() => { const s = document.querySelector('.st-typed select.st-typed-timing'); const c = document.querySelector('.st-typed input.spoken-choice-show'); return { chosen: Boolean(c?.checked), timingOff: s ? s.disabled : null }; })()`);
      const shown = await sayAndWatch();
      check(`typed typewriter: in ${name}, Show only types the line in with no voice, its timing choice greyed`,
        showOnly.chosen && showOnly.timingOff === true && shown.pressed && shown.drawnSamples > 1 && shown.first < LETTERS && shown.rising,
        JSON.stringify({ showOnly, ...shown }));
      const complaints = await errorsIn(win);
      check(`typed typewriter: in ${name}, none of that reported an error`, !complaints.length, complaints.join(' | '));
    } catch (err) {
      check(`typed typewriter: the checks in ${name} ran to their end`, false, err?.stack || String(err));
    } finally {
      if (kept && js) await restoreSetup(js, kept).catch(() => {});
      letGoOfStudioWindows();
    }
  }
}
