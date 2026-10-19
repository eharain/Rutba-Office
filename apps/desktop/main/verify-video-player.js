// Video's play mode (owner, 2026-10-20: "when opening a video it should only
// open in play mode unless user creates new or goes in edit mode").
//
// Two short clips with sound are made in a window (the editor check's own
// fixtures) and the first is given captions beside it. Opened in Video, it
// plays in the player rather than on a timeline: it starts by itself, shows
// its captions, and answers Space, M and the digits. Next File plays the
// second clip of the folder, the speed chosen is kept, and Edit Video puts
// the clip on the editor's timeline, a project not yet changed. A new Video
// window still opens the editor. Run alone with RUTBA_VERIFY_ONLY=video-player.

import fs from 'node:fs';
import path from 'node:path';
import { makeFixtures } from './verify-video-editor.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyVideoPlayer(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  let js = null;
  let kept = null;
  try {
    // A new Video window is the editor, and it makes the clips.
    const maker = await open('video');
    const mjs = (code) => maker.webContents.executeJavaScript(code);
    const fresh = await mjs(`({ editor: Boolean(document.querySelector('.vd-editor')), player: Boolean(document.querySelector('.vp-stage')) })`);
    check('video player: a new Video window opens the editor, not the player', fresh.editor && !fresh.player, JSON.stringify(fresh));
    const clips = await makeFixtures(maker.webContents, dir, 'video-player');
    if (!clips) { check('video player: the clips were made', false, 'this machine could not encode them'); return; }
    fs.writeFileSync(clips.a.replace(/\.webm$/, '.srt'), '1\n00:00:00,000 --> 00:00:01,500\nThe first clip, captioned.\n');

    const win = await open('video', clips.a);
    js = (code) => win.webContents.executeJavaScript(code);
    kept = await js(`window.rutbaOffice.store.get({ key: 'video.player', fallback: null })`);
    // A short clip may have played through before this looks: that it moved is what counts.
    const started = await ok(until(() => js(`(document.querySelector('.vp-video')?.currentTime || 0) > 0.2`), 'the clip playing', 10000));
    const first = await js(`(() => { const s = document.querySelector('.vp-stage'); const v = document.querySelector('.vp-video'); return {
      stage: Boolean(s), editor: Boolean(document.querySelector('.vd-editor')), clips: document.querySelectorAll('.vd-clip').length,
      duration: Number(s?.dataset.duration || 0), title: document.querySelector('.rw-titlebar, .rw-title')?.textContent || document.title,
      cues: v?.textTracks?.[0]?.cues?.length ?? 0, captionsOn: document.querySelector('.vp-captions')?.getAttribute('aria-pressed'),
    }; })()`);
    check('video player: a video opened in Video plays in the player by itself, with no timeline, as long as it is, and with its captions from the .srt beside it',
      started && first.stage && !first.editor && first.clips === 0 && Math.abs(first.duration - 1.6) < 0.25 && first.cues >= 1,
      JSON.stringify({ started, ...first }));

    // A picture to look at, when the run is asked for them.
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await js(`(() => { const v = document.querySelector('.vp-video'); v.pause(); v.currentTime = 0.8; return 1; })()`);
      await wait(800);
      fs.mkdirSync(process.env.RUTBA_VERIFY_CAPTURE, { recursive: true });
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'video-player.png'), (await win.webContents.capturePage()).toPNG());
    }

    // Space pauses; the digit 0 goes to the start; M mutes.
    win.webContents.focus();
    const key = async (keyCode) => { win.webContents.sendInputEvent({ type: 'keyDown', keyCode }); win.webContents.sendInputEvent({ type: 'char', keyCode }); win.webContents.sendInputEvent({ type: 'keyUp', keyCode }); await wait(250); };
    await js(`document.activeElement?.blur?.(), 1`);
    await key('Space');
    const paused = await ok(until(() => js(`document.querySelector('.vp-stage')?.dataset.playing === '0'`), 'the clip paused', 3000));
    await key('0');
    const atStart = await js(`document.querySelector('.vp-video')?.currentTime ?? -1`);
    await key('m');
    const muted = await js(`document.querySelector('.vp-video')?.muted === true`);
    await key('m');
    check('video player: Space pauses, 0 goes back to the start and M mutes', paused && atStart < 0.1 && muted, JSON.stringify({ paused, atStart, muted }));

    // The speed chosen is used and kept for next time.
    await js(`(() => { const s = document.querySelector('.vp-speed'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, '1.5'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await wait(400);
    const rate = await js(`document.querySelector('.vp-video')?.playbackRate`);
    const stored = await js(`window.rutbaOffice.store.get({ key: 'video.player', fallback: null }).then((v) => v?.speed)`);
    check('video player: the speed chosen plays at once and is kept', rate === 1.5 && stored === 1.5, JSON.stringify({ rate, stored }));

    // Next File: the second clip of the folder.
    const pressedNext = await js(`(() => { const b = document.querySelector('.vp-next'); if (!b || b.disabled) return 'no next'; b.click(); return 'clicked'; })()`);
    const moved = await ok(until(() => js(`/clip-b/.test(decodeURIComponent(escape(atob((document.querySelector('.vp-video')?.getAttribute('src') || '').split('/').pop().replace(/-/g, '+').replace(/_/g, '/')))))`), 'the second clip', 6000));
    const prevOn = await js(`document.querySelector('.vp-prev')?.disabled === false`);
    check('video player: Next File plays the next clip in the folder, and Previous File is then offered', pressedNext === 'clicked' && moved && prevOn, JSON.stringify({ pressedNext, moved, prevOn }));

    // Edit Video: the clip on the editor's timeline, nothing yet to save.
    await js(`document.querySelector('.vp-edit')?.click(), 1`);
    const editing = await ok(until(() => js(`Boolean(document.querySelector('.vd-editor')) && document.querySelectorAll('.vd-clip[data-kind="video"]').length === 1`), 'the editor with the clip', 15000));
    const after = await js(`({ player: Boolean(document.querySelector('.vp-stage')), dirty: document.querySelector('.vd-editor')?.dataset.dirty, name: document.querySelector('.vd-clip[data-kind="video"]')?.textContent || '' })`);
    check('video player: Edit Video puts the clip playing on the editor\'s timeline, with nothing yet to save',
      editing && !after.player && after.dirty === '0' && /clip-b/.test(after.name), JSON.stringify({ editing, ...after }));

    const complaints = await errorsIn(win);
    check('video player: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('video player: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js) await js(`window.rutbaOffice.store.set({ key: 'video.player', value: ${JSON.stringify(kept ?? null)} }).then(() => 1)`).catch(() => {});
  }
}
