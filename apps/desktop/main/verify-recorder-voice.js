// Video: Record Voice, the microphone alone (owner, 2026-10-19: "i want to
// recorder to be able to record the voice only").
//
// A project of one six-second still is opened and the playhead put at two
// seconds. Record Voice on the ribbon opens the recorder on Voice: no picture
// to choose, no camera, the microphone's level drawn as it goes. Recorded
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
    settingsWere = await js(`window.rutbaOffice.store.get({ key: 'recorder.settings', fallback: null })`);
    await js(`window.rutbaOffice.store.set({ key: 'recorder.settings', value: ${JSON.stringify({ source: 'window', countdown: 0, camera: true, mic: true, system: false, minimise: false, folder })} }).then(() => 1)`);
    // The playhead at two seconds, as dragging the slider under the preview puts it.
    await js(`(() => { const s = document.querySelector('.vd-scrub'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(s, '2'); s.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await until(() => js(`Math.abs(Number(document.querySelector('.vd-scrub').value) - 2) < 0.05`), 'the playhead at 2 s', 3000).catch(() => {});

    // Record Voice on the ribbon: the recorder on Voice, with nothing about a picture.
    const pressed = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Record Voice')); if (!b || b.disabled) return 'no button'; b.click(); return 'clicked'; })()`);
    await ok(until(() => js(`Boolean(document.querySelector('.rec-dialog'))`), 'the recorder', 6000));
    const heard = await ok(until(() => js(`Number(document.querySelector('.rec-voice-wave')?.dataset.bars || 0) >= 8 && Number(document.querySelector('.rec-meter-mic')?.dataset.peak || 0) > 0`), 'the voice drawn as it goes', 30000));
    const setup = await js(`({
      title: document.querySelector('.rec-dialog .rw-dialog-head')?.textContent || '',
      voice: document.querySelector('.rec-kinds button[data-value="voice"]')?.getAttribute('aria-checked'),
      screen: Boolean(document.querySelector('.rec-screen')), sources: document.querySelectorAll('.rec-source').length,
      camera: Boolean(document.querySelector('.rec-camera-toggle')), inset: Boolean(document.querySelector('.rec-inset')),
      system: Boolean(document.querySelector('.rec-system-toggle')), go: !document.querySelector('.rec-go')?.disabled,
    })`);
    check('recorder voice: Record Voice opens the recorder on Voice, with the microphone\'s level drawn as it goes and no picture, window list, camera or computer sound to set',
      pressed === 'clicked' && heard && /Record Voice/.test(setup.title) && setup.voice === 'true' && !setup.screen && setup.sources === 0 && !setup.camera && !setup.inset && !setup.system && setup.go,
      JSON.stringify({ pressed, heard, ...setup }));

    // Recorded for a little over two seconds, then stopped.
    await js(`document.querySelector('.rec-go')?.click(), 1`);
    const recording = await ok(until(() => js(`(() => { const b = document.querySelector('.rec-stop'); return Boolean(b) && !b.disabled; })()`), 'the recording to start', 8000));
    const files = () => { try { return fs.readdirSync(folder).map((n) => ({ name: n, size: fs.statSync(path.join(folder, n)).size })); } catch { return []; } };
    const grew = await ok(until(() => files().some((f) => f.size > 2000), 'the voice on disk', 6000));
    await wait(1800);
    await js(`document.querySelector('.rec-stop')?.click(), 1`);
    const closed = await ok(until(() => js(`!document.querySelector('.rec-dialog')`), 'the recorder to close', 10000));
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
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Record Screen')); b?.click(); return 1; })()`);
    await ok(until(() => js(`Boolean(document.querySelector('.rec-dialog'))`), 'the recorder again', 6000));
    await wait(600);
    const again = await js(`document.querySelector('.rec-kinds button[aria-checked="true"]')?.dataset.value || ''`);
    await js(`(() => { const b = [...document.querySelectorAll('.rec-dialog button')].find((n) => n.textContent.trim() === 'Cancel'); b?.click(); return 1; })()`);
    await ok(until(() => js(`!document.querySelector('.rec-dialog')`), 'the recorder closed', 6000));
    check('recorder voice: Record Screen after a voice opens on a picture, not on Voice', again && again !== 'voice', again);

    const complaints = await errorsIn(win);
    check('recorder voice: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('recorder voice: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js) await js(`window.rutbaOffice.store.set({ key: 'recorder.settings', value: ${JSON.stringify(settingsWere ?? null)} }).then(() => 1)`).catch(() => {});
  }
}
