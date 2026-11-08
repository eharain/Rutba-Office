// Video: Record Voice, the microphone alone (owner, 2026-10-19: "i want to
// recorder to be able to record the voice only").
//
// A project of one six-second still is opened and the playhead put at two
// seconds. Record Voice on the ribbon opens the Recorder for sound alone: no
// picture, no camera. Recorded
// (with Chromium's stand-in microphone, RUTBA_FAKE_MEDIA) and stopped, the
// take is one sound file in the run's own folder, read back as sound alone,
// and it lands on a sound track at the playhead, marked as the voice. The
// recorder opened from Record Screen afterwards is back on a picture.
// Run alone with RUTBA_VERIFY_ONLY=recorder-voice.

import fs from 'node:fs';
import path from 'node:path';
import { probeMedia } from '@rutba/media/probe';
import { emptyProject, addMedia } from '@rutba/studio/project';
import { writeRvid } from '@rutba/studio/rvid';
import { gradientPng } from './sample-picture.js';
import { withScreenStandIn } from './verify-recorder.js';
import { emptyStack } from '@rutba/studio/stack';
import { seedSetup, restoreSetup, recorderOpened, windowsNow, CHECK_SOUNDS, RECORDING, PRESS_RECORD } from './verify-studio-seed.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export function verifyRecorderVoice(h, args) {
  return withScreenStandIn(() => voiceChecks(h, args));
}

async function voiceChecks(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const folder = path.join(dir, 'voice-recordings');
  const still = path.join(dir, 'voice-still.png');
  const project = path.join(dir, 'voice-project.rvid');
  let win = null;
  let js = null;
  let settingsWere = null;
  try {
    fs.mkdirSync(folder, { recursive: true });
    fs.writeFileSync(still, Buffer.from(gradientPng(64, 36, [40, 90, 160], [200, 120, 60])));
    let doc = emptyProject();
    doc = addMedia(doc, { url: still, name: 'voice-still.png', kind: 'image', width: 64, height: 36, size: fs.statSync(still).size }).doc;
    doc = { ...doc, patches: doc.patches.map((p) => (p.timing ? { ...p, timing: { start: 0, end: 6 } } : p)) };
    fs.writeFileSync(project, Buffer.from(writeRvid(doc, { projectPath: project })));

    win = await open('video', project);
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.vd-clip').length >= 1`), 'the project on the timeline', 10000);
    settingsWere = await seedSetup(js, { app: 'recorder', stack: { ...emptyStack(), sounds: CHECK_SOUNDS() }, studio: { size: 720, fps: 30, countdown: 0, folder, minimise: false } });
    // The playhead at two seconds, as dragging the slider under the preview puts it.
    await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, '2'); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await until(() => js(`Math.abs(Number(document.querySelector('.vd-scrub').value) - 2) < 0.05`), 'the playhead at 2 s', 3000).catch(() => {});

    // Record Voice on the ribbon: the recorder on Voice, with nothing about a picture.
    const known = windowsNow();
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Record Voice')); if (!b || b.disabled) return 'no button'; b.click(); return 'clicked'; })()`);
    const { rw, rjs } = await recorderOpened(until, known);
    const says = await ok(until(() => rjs(`/sound/i.test(document.querySelector('.st-record')?.textContent || '') && !document.querySelector('.st-record').disabled`), 'Record sound', 12000));
    const setup = await rjs(`({ into: document.querySelector('.st-into')?.textContent || '', layers: document.querySelector('.st-canvas')?.dataset.layers || '' })`);
    check('recorder voice: Record Voice opens the Recorder for sound alone, into this project, with nothing about a picture',
      pressed === 'clicked' && says && /Recording into/.test(setup.into) && setup.layers === '', JSON.stringify({ pressed, says, ...setup }));

    // Recorded for a little over two seconds, then stopped.
    await rjs(PRESS_RECORD);
    const recording = await ok(until(() => rjs(RECORDING), 'the recording to start', 8000));
    const files = () => { try { return fs.readdirSync(folder).map((n) => ({ name: n, size: fs.statSync(path.join(folder, n)).size })); } catch { return []; } };
    const grew = await ok(until(() => files().some((f) => f.size > 2000), 'the voice on disk', 6000));
    await wait(1800);
    await rjs(`document.querySelector('.st-record')?.click(), 1`);
    const closed = await ok(until(() => rw.isDestroyed(), 'the Recorder to close', 10000));
    const kept = files();
    const file = kept.length === 1 ? path.join(folder, kept[0].name) : null;
    const probe = file ? probeMedia(new Uint8Array(fs.readFileSync(file)), { complete: true }) : null;
    check('recorder voice: the voice is recorded to one sound file of its own, named Voice, with sound and no picture, as long as it was recorded',
      recording && grew && closed && kept.length === 1 && /^Voice .*\.(weba|ogg)$/.test(kept[0].name) && probe?.hasAudio && !probe?.hasVideo && probe.duration > 1.5,
      JSON.stringify({ recording, grew, closed, kept, probe: probe ? { audio: probe.hasAudio, video: probe.hasVideo, duration: probe.duration } : null }));

    // On the timeline: a sound clip at the playhead, marked as the voice.
    await ok(until(() => js(`document.querySelectorAll('.vd-clip[data-kind="sound"]').length >= 1`), 'the voice on the timeline', 8000));
    const landed = await js(`(() => { const c = document.querySelector('.vd-clip[data-kind="sound"]'); return c ? { start: Number(c.dataset.start), end: Number(c.dataset.end) } : null; })()`);
    const saved = await js(`window.__rutbaVideoDoc ? window.__rutbaVideoDoc().patches.filter((p) => p.type === 'sound').map((p) => ({ take: p.take, lane: p.lane })) : null`).catch(() => null);
    check('recorder voice: the voice lands on a sound track at the playhead (2 s), as long as it was recorded',
      Boolean(landed) && Math.abs(landed.start - 2) < 0.1 && landed.end - landed.start > 1.5 && (saved == null || saved.some((p) => p.take === 'voice')),
      JSON.stringify({ landed, saved }));

    // Record Screen afterwards opens on a picture again.
    const known2 = windowsNow();
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Record Screen')); b?.click(); return 1; })()`);
    const second = await recorderOpened(until, known2);
    const again = await second.rjs(`(() => { const n = document.querySelector('.st-record'); return n ? n.textContent.trim() : ''; })()`);
    second.rw.destroy();
    await ok(until(() => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Record Screen')); return Boolean(b) && !b.disabled; })()`), 'Record Screen free again', 6000));
    check('recorder voice: Record Screen after a voice opens on a picture, not on sound alone', Boolean(again) && !/sound/i.test(again), again);

    const complaints = await errorsIn(win);
    check('recorder voice: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('recorder voice: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js && settingsWere) await restoreSetup(js, settingsWere).catch(() => {});
  }
}
