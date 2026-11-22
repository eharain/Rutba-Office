// A real added voice, used in the windows. No window check had spoken with
// one: the checks of typed words used Shown only, which needs no voice. A
// film take of the released 2.26.1 (17 Nov 2026) reported that in the
// Recorder typed words read aloud were said and never drawn on the picture;
// these checks do what the take did and find the words drawn, with the voice
// cold, with it warm and while recording, so they stand as the guard no
// check was before.
//
// The run is given the release's own speech packs (the engine, the English
// pronunciation pack, the small model and one voice: gate-packs.js,
// GATE_SPEECH_PACKS), added to its add-ons folder before it starts, as the
// language packs are. Voices and languages shows the voice as added, the
// engine offers it and makes a sentence into sound, and in the Recorder words
// typed under the frame and said Shown and read aloud are drawn across the
// foot of the picture while the voice says them, and taken away after. Said
// again while recording with the microphone off, they are drawn again and the
// take carries the voice.
// Run alone with RUTBA_VERIFY_ONLY=real-voice and RUTBA_GATE_PACKS set.

import fs from 'node:fs';
import path from 'node:path';
import { probeMedia } from '@rutba/media/probe';
import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, CHECK_SOUNDS, PANEL, IN_PANEL, RECORDING, PRESS_RECORD } from './verify-studio-seed.js';
import { letGoOfStudioWindows } from './verify-let-go.js';
import { GATE_VOICE, GATE_PACKS_ENV } from './gate-packs.js';

const WORDS = 'Welcome to the new starters.';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyRealVoice(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  // Without the speech packs there is no voice to speak with: said once and plainly, not as seven checks timing out.
  if (!process.env[GATE_PACKS_ENV]) {
    check('real voice: this run was given the speech packs', false, `${GATE_PACKS_ENV} is not set: name the packages release folder (dist/v<version>), whose signed speech packs are added before the run starts`);
    return;
  }
  let kept = null;
  let js = null;
  const folder = path.join(dir, 'real-voice-takes');
  fs.mkdirSync(folder, { recursive: true });
  try {
    // The Recorder first, as a person who opens it to say something: the voice has said nothing yet in this run,
    // so the first line waits on the model being read from the disk.
    const win = await open('recorder');
    js = (code) => win.webContents.executeJavaScript(code);
    kept = await seedSetup(js, { app: 'recorder', stack: { ...emptyStack(), sounds: CHECK_SOUNDS(false) }, studio: { size: 720, fps: 30, countdown: 0, folder, minimise: false } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas'))`), 'the Recorder window', 15000);
    await js(PANEL('main'));
    await until(() => js(IN_PANEL('record')), 'the Record panel', 6000).catch(() => {});
    await js(`document.querySelector('.st-typed-open')?.click(), 1`);
    const box = await ok(until(() => js(`Boolean(document.querySelector('.st-typed .spoken-box-text')) && Boolean(document.querySelector('.st-typed select'))`), 'the Text to speech box with its voice', 15000));
    await js(`(() => { const n = document.querySelector('.st-typed .spoken-box-text'); if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(n, ${JSON.stringify(WORDS)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await js(`document.querySelector('.st-typed input.spoken-choice-both')?.click(), 1`);
    await wait(300);
    const choice = await js(`(() => { const c = document.querySelector('.st-typed input.spoken-choice-both'); return c ? { checked: c.checked, disabled: c.disabled } : null; })()`);
    const pressed = await js(`(() => { const b = document.querySelector('.st-typed .spoken-box-say'); if (!b || b.disabled) return false; b.click(); return true; })()`);
    // The first line waits on the model, which is read from the disk the first time a voice speaks.
    const drawn = await ok(until(() => js(`[...document.querySelectorAll('.st-canvas')].some((c) => c.dataset.typed === '1')`), 'the words on the picture', 90000));
    const during = await js(`({ typed: [...document.querySelectorAll('.st-canvas')].map((c) => c.dataset.typed), now: document.querySelector('.st-typed-now')?.textContent || null, stop: Boolean(document.querySelector('.st-typed .spoken-box-stop')) })`);
    check('real voice: in the Recorder, words typed and said Shown and read aloud are drawn across the foot of the picture while the voice says them',
      box && choice?.checked === true && pressed && drawn && during.stop, JSON.stringify({ box, choice, pressed, drawn, during }));
    const gone = await ok(until(() => js(`[...document.querySelectorAll('.st-canvas')].every((c) => c.dataset.typed === '0') && !document.querySelector('.st-typed-now') && Boolean(document.querySelector('.st-typed .spoken-box-say'))`), 'the words to go', 60000));
    check('real voice: when the words have been said they leave the picture and the button says them again', gone, JSON.stringify(await js(`({ typed: [...document.querySelectorAll('.st-canvas')].map((c) => c.dataset.typed), now: document.querySelector('.st-typed-now')?.textContent || null })`)));
    // Said again while recording, the microphone off: the take has sound only because the voice went into it.
    const TYPED = `[...document.querySelectorAll('.st-canvas')].some((c) => c.dataset.typed === '1')`;
    const SAID = `[...document.querySelectorAll('.st-canvas')].every((c) => c.dataset.typed === '0') && !document.querySelector('.st-typed-now') && Boolean(document.querySelector('.st-typed .spoken-box-say'))`;
    await js(PRESS_RECORD);
    const recording = await ok(until(() => js(RECORDING), 'the recording to start', 12000));
    const again = await js(`(() => { const b = document.querySelector('.st-typed .spoken-box-say'); if (!b || b.disabled) return false; b.click(); return true; })()`);
    const drawnOnAir = await ok(until(() => js(TYPED), 'the words on the picture while recording', 60000));
    const saidOnAir = await ok(until(() => js(SAID), 'the words said while recording', 60000));
    await wait(600);
    await js(`document.querySelector('.st-record')?.click(), 1`);
    const files = () => { try { return fs.readdirSync(folder).map((n) => ({ name: n, size: fs.statSync(path.join(folder, n)).size })); } catch { return []; } };
    await ok(until(() => js(`!(${RECORDING})`), 'the recording to stop', 12000));
    await ok(until(() => files().some((f) => f.size > 2000), 'the take on disk', 12000));
    await wait(800);
    const takes = files();
    const take = takes.find((f) => /\.(webm|mp4|mkv)$/i.test(f.name)) || null;
    const probe = take ? probeMedia(new Uint8Array(fs.readFileSync(path.join(folder, take.name))), { complete: true }) : null;
    check('real voice: said while recording, the words are on the picture and the take carries the voice, the microphone being off',
      recording && again && drawnOnAir && saidOnAir && Boolean(probe?.hasVideo) && Boolean(probe?.hasAudio) && probe.duration > 1,
      JSON.stringify({ recording, again, drawnOnAir, saidOnAir, takes, probe: probe ? { audio: probe.hasAudio, video: probe.hasVideo, duration: probe.duration } : null }));
    // Voices and languages, on Home: the voice is there as one added.
    const home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    const offered = await ok(until(() => hjs(`window.rutbaOffice.speech.voices().then((v) => v.some((x) => x.id === ${JSON.stringify(GATE_VOICE)})).catch(() => false)`), 'the engine to offer the voice', 30000));
    const voices = await hjs(`window.rutbaOffice.speech.voices().then((v) => v.map((x) => x.id)).catch((e) => String(e))`);
    const state = await hjs(`window.rutbaOffice.addons.list().then((l) => ({ here: l.installed.map((p) => p.id + ':' + p.state).sort(), refused: (l.refused || []).map((r) => r.id + ': ' + r.reason) })).catch((e) => ({ here: [], refused: [String(e)] }))`);
    check('real voice: the speech engine loads with this run and offers the added voice', offered, JSON.stringify({ voices, state }));
    await hjs(`location.search = '?voices', 1`).catch(() => {});
    const row = await until(() => home.webContents.executeJavaScript(`(() => { const r = document.querySelector('.voices-row[data-addon=${JSON.stringify(GATE_VOICE)}]'); const b = r?.querySelector('[data-default-voice]'); return r && b ? true : false; })()`).catch(() => false), 'the voice in Voices and languages', 20000).then(() => true).catch(() => false);
    const shown = await home.webContents.executeJavaScript(`(() => { const r = document.querySelector('.voices-row[data-addon=${JSON.stringify(GATE_VOICE)}]'); const b = r?.querySelector('[data-default-voice]'); return { row: Boolean(r), button: b?.textContent.trim() || null, rows: document.querySelectorAll('.voices-row').length }; })()`).catch((e) => ({ row: false, button: String(e) }));
    check('real voice: Voices and languages shows the voice as added, with its default voice button', row && Boolean(shown.button), JSON.stringify(shown));
    const said = await home.webContents.executeJavaScript(`window.rutbaOffice.speech.say({ text: 'Rutba Office reads this aloud.', voice: ${JSON.stringify(GATE_VOICE)}, speed: 1 }).then((o) => ({ samples: o.samples ? o.samples.length : 0, rate: o.rate, cancelled: Boolean(o.cancelled) })).catch((e) => ({ samples: 0, error: String(e?.message || e) }))`);
    check('real voice: the voice makes a sentence into sound, more than half a second of it', said.samples > 12000 && said.rate === 24000, JSON.stringify(said));

    const complaints = [...(await errorsIn(home)), ...(await errorsIn(win))];
    check('real voice: nothing went wrong on the way', !complaints.length, complaints.join(' | '));
  } catch (err) {
    check('real voice: the checks ran to their end', false, err?.stack || String(err));
  } finally {
    if (kept && js) await restoreSetup(js, kept).catch(() => {});
    letGoOfStudioWindows();
  }
}
