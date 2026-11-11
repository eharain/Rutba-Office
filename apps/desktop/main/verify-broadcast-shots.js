// Live Broadcast's shots (owner, 2026-11-03, D-2026-11-03-01 and -02: "how
// do we solve the problem of rotating view lets a camera on front and camera
// on side or camera on one speaker or the other").
//
// The window opens on two copies of Chromium's stand-in camera (RUTBA_FAKE_MEDIA),
// one for the front and one for the side, and its stand-in microphone. Each
// camera has a shot of its own and the wide shot has both; the Recorder's one
// Camera shot is not there. A shot put up goes on air faded as the transition
// says. With Preview on, a shot picked is the next one, beside the one on
// air, and goes on air with Cut (at once) or Fade. The timed rotation goes
// through the cameras' shots and the wide one in turn. A shot is paired with
// a microphone from its menu, and Follow the speaker then cuts on its own
// (to the paired shot when the stand-in microphone is heard, to the wide one
// when it is quiet); a shot picked by hand turns it off. Run alone with
// RUTBA_VERIFY_ONLY=broadcast-shots.

import { seedSetup, restoreSetup, readSetup, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { letGoOfStudioWindows } from './verify-let-go.js';

const START_STACK = {
  version: 2,
  items: [
    { id: 'camera-1', kind: 'camera', ref: '', name: 'Front', on: true, live: true, shape: 'rect' },
    { id: 'camera-2', kind: 'camera', ref: '', name: 'Side', on: false, live: true, shape: 'rect' },
  ],
  sounds: [{ id: 'mic-1', kind: 'mic', ref: '', name: '', live: true, muted: false, gain: 1 }, { id: 'system-1', kind: 'system', ref: '', name: '', live: false, muted: false, gain: 1 }],
  scenes: [],
  video: 'on',
  sound: 'on',
  transition: 'fade',
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyBroadcastShots(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  let js = null;
  let kept = null;
  try {
    const home = await open('home');
    await until(() => home.webContents.executeJavaScript(`Boolean(document.querySelector('.home-card[data-app="broadcast"]'))`), 'the launcher\'s Live Broadcast tile', 15000).catch(() => {});
    const tile = await home.webContents.executeJavaScript(`(() => { const c = document.querySelector('.home-card[data-app="broadcast"]'); return c ? c.textContent : ''; })()`);
    check('broadcast shots: the launcher has a Live Broadcast tile of its own', /Live Broadcast/.test(tile), tile.slice(0, 80));

    const win = await open('broadcast');
    js = (code) => win.webContents.executeJavaScript(code);
    kept = await seedSetup(js, { app: 'broadcast', stack: START_STACK, studio: { size: 720, fps: 30, countdown: 0, folder: dir, minimise: false } });
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-livebar'))`), 'the Live Broadcast window', 10000);
    await until(() => js(`(document.querySelector('.st-air .st-canvas')?.dataset.layers || '') === 'camera'`), 'the front camera live', 15000);

    const air = () => js(`(() => { const c = document.querySelector('.st-air .st-canvas'); return { ids: c.dataset.ids || '', fading: c.dataset.fading, scene: document.querySelector('.st-scene.air')?.dataset.scene || null }; })()`);
    const next = () => js(`(() => { const c = document.querySelector('.st-preview-canvas'); return { ids: c?.dataset.ids || '', shown: Boolean(document.querySelector('.st-next-pane')) }; })()`);
    const pick = (key) => js(`document.querySelector('.st-scene[data-scene=${JSON.stringify(key)}]')?.click(), 1`);
    const stack = async () => (await readSetup(js))?.stack;
    const auto = async (re) => {
      await js(`document.querySelector('.st-auto')?.click(), 1`);
      await until(() => js(`document.querySelectorAll('[role="menuitemcheckbox"], [role="menuitem"]').length > 0`), 'the Auto menu', 3000).catch(() => {});
      await js(`[...document.querySelectorAll('[role="menuitemcheckbox"], [role="menuitem"]')].find((n) => ${re}.test(n.textContent))?.click(), 1`);
    };
    const autoSays = () => js(`document.querySelector('.st-auto')?.textContent || ''`);
    const takeWith = (how, ms) => js(`document.querySelector('.st-transition [data-value="${how}"]').click(), new Promise((r) => setTimeout(r, 60))`).then(() => js(`document.querySelector('.st-take-btn').click(), new Promise((r) => setTimeout(r, ${ms}))`));
    const go = async (p) => { await js(PANEL(p)); await until(() => js(IN_PANEL(p)), `the ${p} panel`, 4000).catch(() => {}); };

    // The bar over the panel holds the one Go live button; Back soon is the last scene.
    const line = await js(`[Boolean(document.querySelector('.st-livebar .lv-go')), !document.querySelector('.st-record'), [...document.querySelectorAll('.st-panels [data-panel]')].map((n) => n.textContent).join(' | '), [...document.querySelectorAll('.st-scene[data-scene]')].pop()?.dataset.scene].join(' ')`);
    const keys = await js(`[...document.querySelectorAll('.st-scene[data-scene]')].map((n) => n.dataset.scene + '=' + n.textContent.replace(/^\\d/, '')).join(' | ')`);
    check('broadcast shots: Go live is on the bar over the Live panel, the panels are Set up and Live, Back soon comes last, and each camera has a shot of its own, named, with the wide one of both and no single Camera shot',
      line === 'true true Set up | Live back-soon' && /cam:camera-1=Camera 1/.test(keys) && /cam:camera-2=Camera 2/.test(keys) && /wide=Everyone/.test(keys) && !/(^| )camera=/.test(keys),
      JSON.stringify({ line, keys }));

    // The wide shot put up: both cameras on air, faded in.
    await js(`document.querySelector('.st-scene[data-scene="wide"]').click(), new Promise((r) => setTimeout(r, 120))`);
    const fadingWide = (await air()).fading;
    const wide = await ok(until(async () => { const a = await air(); return a.scene === 'wide' && a.ids.split(' ').sort().join(' ') === 'camera-1 camera-2' && a.fading === '0'; }, 'the wide shot on air', 6000));
    check('broadcast shots: the wide shot puts both cameras on air side by side, faded in as the transition says, and is marked on air',
      wide && fadingWide === '1', JSON.stringify({ fadingWide, air: await air() }));

    // Preview: the shot picked waits beside the one on air, and goes on with Cut or Fade.
    await js(`document.querySelector('.st-preview-btn').click(), 1`);
    await until(async () => (await next()).shown, 'the preview', 3000).catch(() => {});
    await pick('cam:camera-2');
    const waiting = await ok(until(async () => (await next()).ids === 'camera-2', 'the side camera in the preview', 8000));
    const stillOn = await air();
    await takeWith('cut', 80);
    const cutNow = await air();
    const cut = await ok(until(async () => { const a = await air(); return a.ids === 'camera-2' && a.scene === 'cam:camera-2'; }, 'the side camera on air', 4000));
    check('broadcast shots: with Preview first on a shot picked waits beside the one on air, and Take on Cut puts it on air at once, with no fade',
      waiting && stillOn.ids.split(' ').length === 2 && stillOn.scene === 'wide' && cut && cutNow.fading === '0',
      JSON.stringify({ waiting, stillOn, cutNow, next: await next() }));
    await pick('cam:camera-1');
    await until(async () => (await next()).ids === 'camera-1', 'the front camera in the preview', 6000).catch(() => {});
    await takeWith('fade', 100);
    const fadeNow = await air();
    const faded = await ok(until(async () => { const a = await air(); return a.ids === 'camera-1' && a.fading === '0'; }, 'the front camera faded in', 4000));
    check('broadcast shots: Take on Fade dissolves the next shot in over the one on air', fadeNow.fading === '1' && faded, JSON.stringify({ fadeNow, air: await air() }));
    await js(`document.querySelector('.st-preview-btn').click(), 1`);
    const previewOff = await ok(until(async () => !(await next()).shown, 'the preview put away', 3000));

    // The timed rotation, every three seconds: the cameras' shots and the wide one in turn.
    await pick('cam:camera-1');
    await wait(300);
    await auto('/^Rotate every 3 s/');
    const seen = [];
    const at = [];
    const started = Date.now();
    // Twenty seconds to see three shots is a deadline for a loaded machine,
    // and each shot's own gap is held to its three seconds, within what a
    // busy machine's timers and this 150 ms look add (review of 11 Nov: a
    // longer deadline alone let a rotation at half speed pass).
    while (Date.now() - started < 20000 && seen.length < 3) {
      const a = await air();
      if (a.scene && a.scene !== (seen[seen.length - 1] || 'cam:camera-1')) { seen.push(a.scene); at.push(Date.now()); }
      await wait(150);
    }
    const gaps = at.map((t, i) => (t - (i ? at[i - 1] : started)) / 1000);
    const everyThree = gaps.length === 3 && gaps.slice(1).every((g) => g >= 2.5 && g <= 4.5) && gaps[0] <= 4.5;
    const rotating = /Rotate every 3 s/.test(await autoSays()) && (await stack())?.rotate?.every === 3;
    await auto('/^Off/');
    check('broadcast shots: Rotate moves between the side camera, the wide shot and the front camera in turn, a new shot every three seconds',
      previewOff && rotating && everyThree && seen.join(' ') === 'cam:camera-2 wide cam:camera-1', JSON.stringify({ previewOff, rotating, seen, gaps, seconds: (Date.now() - started) / 1000 }));

    // A shot paired with the microphone from its menu in Set up, then Follow the speaker.
    await go('setup');
    await js(`(() => { const b = document.querySelector('.st-scene[data-scene="cam:camera-2"]'); const r = b.getBoundingClientRect(); b.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 4, clientY: r.top + 4 })); return 1; })()`);
    const menu = await ok(until(() => js(`[...document.querySelectorAll('.rw-menu [role="menuitem"], .rw-menu [role="menuitemcheckbox"]')].map((n) => n.textContent).join(' | ').includes('In the rotation')`), 'the shot\'s menu', 3000));
    await js(`[...document.querySelectorAll('.rw-menu [aria-haspopup="menu"]')].find((n) => /microphone/.test(n.textContent))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.rw-submenu [role="menuitemcheckbox"]'))`), 'the microphones', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-submenu [role="menuitemcheckbox"]')].find((n) => !/No microphone/.test(n.textContent))?.click(), 1`);
    const paired = await ok(until(async () => (await stack())?.follow?.['cam:camera-2'] === 'mic-1', 'the side camera paired with the microphone', 4000));
    const tagged = await js(`Boolean(document.querySelector('.st-scene[data-scene="cam:camera-2"] .st-scene-tags'))`);
    await go('main');
    check('broadcast shots: a shot is paired with a microphone from its menu, kept, and shows it is', menu && paired && tagged, JSON.stringify({ menu, paired, tagged, follow: (await stack())?.follow }));

    // Follow on: the side camera's shot when its microphone is heard, the wide one when it is quiet; either way it cuts on its own.
    await pick('cam:camera-1');
    await until(async () => (await air()).scene === 'cam:camera-1', 'the front camera on air', 3000).catch(() => {});
    await auto('/^Follow the speaker/');
    await until(async () => /Follow the speaker/.test(await autoSays()), 'Follow the speaker on', 3000).catch(() => {});
    const following = /Follow the speaker/.test(await autoSays());
    const followed = await ok(until(async () => ['cam:camera-2', 'wide'].includes((await air()).scene), 'Follow the speaker to cut', 9000));
    const followedTo = (await air()).scene;
    const status = await js(`document.querySelector('.st-status')?.textContent || ''`);
    await pick('cam:camera-1');
    const handOff = await ok(until(async () => /Auto: Off/.test(await autoSays()), 'Follow the speaker off', 3000));
    check('broadcast shots: Follow the speaker cuts on its own (to the paired shot, or wide in the quiet) and says so, and a shot picked by hand turns it off',
      following && followed && /Following the speaker/.test(status) && handOff, JSON.stringify({ following, followed, followedTo, status: status.slice(0, 160), handOff }));

    const complaints = await errorsIn(win);
    check('broadcast shots: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('broadcast shots: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js && kept) await restoreSetup(js, kept).catch(() => {});
    letGoOfStudioWindows();
  }
}
