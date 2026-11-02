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
    kept = await js(`Promise.all([window.rutbaOffice.store.get({ key: 'broadcast.studio', fallback: null }), window.rutbaOffice.store.get({ key: 'broadcast.stack', fallback: null })])`);
    await js(`Promise.all([
      window.rutbaOffice.store.set({ key: 'broadcast.studio', value: ${JSON.stringify({ size: 720, fps: 30, countdown: 0, folder: dir, minimise: false })} }),
      window.rutbaOffice.store.set({ key: 'broadcast.stack', value: ${JSON.stringify(START_STACK)} }),
    ]).then(() => 1)`);
    await win.webContents.reload();
    await until(() => js(`Boolean(document.querySelector('.st-canvas')) && Boolean(document.querySelector('.st-switching'))`), 'the Live Broadcast window', 10000);
    await until(() => js(`(document.querySelector('.st-air .st-canvas')?.dataset.layers || '') === 'camera'`), 'the front camera live', 15000);

    const air = () => js(`(() => { const c = document.querySelector('.st-air .st-canvas'); return { ids: c.dataset.ids || '', fading: c.dataset.fading, scene: document.querySelector('.st-scene.air')?.dataset.scene || null }; })()`);
    const next = () => js(`(() => { const c = document.querySelector('.st-preview-canvas'); return { ids: c.dataset.ids || '', shown: !document.querySelector('.st-next-pane')?.classList.contains('off') }; })()`);
    const pick = (key) => js(`document.querySelector('.st-scene[data-scene=${JSON.stringify(key)}]')?.click(), 1`);
    const stack = () => js(`window.rutbaOffice.store.get({ key: 'broadcast.stack', fallback: null })`);

    // The first line: Picture, Go Live and Record, the switches, the gear.
    const line = await js(`[...document.querySelectorAll('.st-top .st-shoot, .st-top .st-golive, .st-top .st-record')].map((n) => n.classList.contains('st-shoot') ? 'picture' : n.classList.contains('st-golive') ? 'live' : 'record').join(' ')`);
    const keys = await js(`[...document.querySelectorAll('.st-scene[data-scene]')].map((n) => n.dataset.scene + '=' + n.textContent.replace(/^\\d/, '')).join(' | ')`);
    check('broadcast shots: Go Live sits between Picture and Record, and each camera has a shot of its own, named, with the wide one of both and no single Camera shot',
      line === 'picture live record' && /cam:camera-1=Camera 1/.test(keys) && /cam:camera-2=Camera 2/.test(keys) && /wide=Everyone/.test(keys) && !/(^| )camera=/.test(keys),
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
    await js(`document.querySelector('.st-cut').click(), new Promise((r) => setTimeout(r, 80))`);
    const cutNow = await air();
    const cut = await ok(until(async () => { const a = await air(); return a.ids === 'camera-2' && a.scene === 'cam:camera-2'; }, 'the side camera on air', 4000));
    check('broadcast shots: with Preview on a shot picked waits beside the one on air, and Cut puts it on air at once, with no fade',
      waiting && stillOn.ids.split(' ').length === 2 && stillOn.scene === 'wide' && cut && cutNow.fading === '0',
      JSON.stringify({ waiting, stillOn, cutNow, next: await next() }));
    await pick('cam:camera-1');
    await until(async () => (await next()).ids === 'camera-1', 'the front camera in the preview', 6000).catch(() => {});
    await js(`document.querySelector('.st-fade').click(), new Promise((r) => setTimeout(r, 100))`);
    const fadeNow = await air();
    const faded = await ok(until(async () => { const a = await air(); return a.ids === 'camera-1' && a.fading === '0'; }, 'the front camera faded in', 4000));
    check('broadcast shots: Fade dissolves the next shot in over the one on air', fadeNow.fading === '1' && faded, JSON.stringify({ fadeNow, air: await air() }));
    await js(`document.querySelector('.st-preview-btn').click(), 1`);
    const previewOff = await ok(until(async () => !(await next()).shown, 'the preview put away', 3000));

    // The timed rotation, every three seconds: the cameras' shots and the wide one in turn.
    await js(`(() => { const s = document.querySelector('.st-rotate-every'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, '3'); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await until(async () => (await stack())?.rotate?.every === 3, 'every three seconds', 3000).catch(() => {});
    await pick('cam:camera-1');
    await wait(300);
    await js(`document.querySelector('.st-rotate').click(), 1`);
    const seen = [];
    const started = Date.now();
    while (Date.now() - started < 10500 && seen.length < 3) {
      const a = await air();
      if (a.scene && a.scene !== (seen[seen.length - 1] || 'cam:camera-1')) seen.push(a.scene);
      await wait(150);
    }
    const rotating = await js(`document.querySelector('.st-rotate').getAttribute('aria-pressed')`);
    await js(`document.querySelector('.st-rotate').click(), 1`);
    check('broadcast shots: Rotate moves between the side camera, the wide shot and the front camera in turn, a new shot every three seconds',
      previewOff && rotating === 'true' && seen.join(' ') === 'cam:camera-2 wide cam:camera-1', JSON.stringify({ previewOff, rotating, seen, seconds: (Date.now() - started) / 1000 }));

    // A shot paired with the microphone from its menu, then Follow the speaker.
    await js(`(() => { const b = document.querySelector('.st-scene[data-scene="cam:camera-2"]'); const r = b.getBoundingClientRect(); b.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 4, clientY: r.top + 4 })); return 1; })()`);
    const menu = await ok(until(() => js(`[...document.querySelectorAll('.rw-menu [role="menuitem"], .rw-menu [role="menuitemcheckbox"]')].map((n) => n.textContent).join(' | ').includes('In the rotation')`), 'the shot\'s menu', 3000));
    await js(`[...document.querySelectorAll('.rw-menu [aria-haspopup="menu"]')].find((n) => /microphone/.test(n.textContent))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.rw-submenu [role="menuitemcheckbox"]'))`), 'the microphones', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-submenu [role="menuitemcheckbox"]')].find((n) => !/No microphone/.test(n.textContent))?.click(), 1`);
    const paired = await ok(until(async () => (await stack())?.follow?.['cam:camera-2'] === 'mic-1', 'the side camera paired with the microphone', 4000));
    const tagged = await js(`Boolean(document.querySelector('.st-scene[data-scene="cam:camera-2"] .st-scene-tags'))`);
    check('broadcast shots: a shot is paired with a microphone from its menu, kept, and shows it is', menu && paired && tagged, JSON.stringify({ menu, paired, tagged, follow: (await stack())?.follow }));

    // Follow on: the side camera's shot when its microphone is heard, the wide one when it is quiet; either way it cuts on its own.
    await pick('cam:camera-1');
    await until(async () => (await air()).scene === 'cam:camera-1', 'the front camera on air', 3000).catch(() => {});
    await js(`document.querySelector('.st-follow').click(), 1`);
    const following = await js(`document.querySelector('.st-follow').getAttribute('aria-pressed')`);
    const followed = await ok(until(async () => ['cam:camera-2', 'wide'].includes((await air()).scene), 'Follow the speaker to cut', 9000));
    const followedTo = (await air()).scene;
    const status = await js(`document.querySelector('.st-status')?.textContent || ''`);
    await pick('cam:camera-1');
    const handOff = await ok(until(() => js(`document.querySelector('.st-follow').getAttribute('aria-pressed') === 'false'`), 'Follow the speaker off', 3000));
    check('broadcast shots: Follow the speaker cuts on its own (to the paired shot, or wide in the quiet) and says so, and a shot picked by hand turns it off',
      following === 'true' && followed && /Following the speaker/.test(status) && handOff, JSON.stringify({ following, followed, followedTo, status: status.slice(0, 160), handOff }));

    const complaints = await errorsIn(win);
    check('broadcast shots: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('broadcast shots: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (js && kept) await js(`Promise.all([window.rutbaOffice.store.set({ key: 'broadcast.studio', value: ${JSON.stringify(kept[0] ?? null)} }), window.rutbaOffice.store.set({ key: 'broadcast.stack', value: ${JSON.stringify(kept[1] ?? null)} })]).then(() => 1)`).catch(() => {});
  }
}
