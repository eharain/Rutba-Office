// A new session in the Recorder, kept or let go (Office 2.27.0).
//
// A new session is written nowhere (D-2026-11-22-11). So a start with setups
// saved says where they are, in a line over the panel with the setups to
// open; and a session the person has arranged is asked about before it is
// left: at the window's close, and before a saved setup takes its place.
// "Keep this session?" has Save as a setup, Don't save and Cancel. Picking a
// scene is not arranging. The Setups menu offers saving the session in a new
// one and New setup with a saved one open, never both. Reset asks before it
// clears a new session, and its warning is read whole in its box.
// Run alone with RUTBA_VERIFY_ONLY=studio-keep.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { emptyStack } from '@rutba/studio/stack';
import { STORE } from '@rutba/studio/setup';
import { withScreenStandIn } from './verify-recorder.js';
import { seedSetup, restoreSetup, CHECK_SETUP, CHECK_SOUNDS, PANEL, IN_PANEL } from './verify-studio-seed.js';
import { letGoOfStudioWindows } from './verify-let-go.js';

export function verifyStudioKeep(h, args) {
  return withScreenStandIn(() => checks(h, args));
}

async function checks(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const folder = path.join(dir, 'studio-keep');
  const docFile = path.join(dir, 'keep-fixture.docx');
  const NAME = 'Kept at the door';
  let wjs = null;
  let kept = null;
  let made = [];
  try {
    fs.mkdirSync(folder, { recursive: true });
    fs.writeFileSync(docFile, buildDocx({ styles: true, paragraphs: [{ text: 'A window that stays while the Recorder closes.' }] }));
    const word = await open('word', docFile);
    wjs = (code) => word.webContents.executeJavaScript(code);
    let win = await open('recorder');
    let js = (code) => win.webContents.executeJavaScript(code);
    const get = (key, by = js) => by(`window.rutbaOffice.store.get({ key: ${JSON.stringify(key)}, fallback: null })`);
    const set = (key, value, by = js) => by(`window.rutbaOffice.store.set({ key: ${JSON.stringify(key)}, value: ${JSON.stringify(value)} }).then(() => 1)`);
    const q = (sel) => `document.querySelector(${JSON.stringify(sel)})`;
    const has = (sel) => js(`Boolean(${q(sel)})`);
    const click = (sel) => js(`(${q(sel)}?.click(), 1)`);
    const seen = (sel, what, ms = 6000) => ok(until(() => has(sel), what, ms));
    const gone = (sel, what, ms = 6000) => ok(until(() => js(`!${q(sel)}`), what, ms));
    // A menu hears its keys on itself.
    const escape = () => js(`(${q('.rw-menu')}?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })), 1)`);
    const BAND = '.st-band.saved';
    const ASK = '.st-keep-dialog';

    // One setup saved before ("Check"), and a start that opens none: a new session.
    kept = await seedSetup(js, { app: 'recorder', stack: { ...emptyStack(), sounds: CHECK_SOUNDS() }, studio: { size: 720, fps: 30, countdown: 0, folder, minimise: false } });
    const newStart = async () => {
      await set(STORE.index, { ...(await get(STORE.index)), current: { recorder: '', broadcast: '' } });
      await win.webContents.reload();
      await until(() => has('.st-panels'), 'the Recorder window', 15000);
      await until(() => has('.st-scene[data-scene]'), 'its scenes', 8000).catch(() => {});
    };
    const current = async (by = js) => (await get(STORE.index, by))?.current?.recorder || '';
    await newStart();

    /* ── the line that says where the saved setups are ── */
    const onMain = await seen(`.st-panel[data-panel="record"] ${BAND}`, 'the line over Record');
    const words = await js(`${q(`${BAND} .st-band-words`)}?.textContent.trim() || ''`);
    await click(`${BAND} .st-band-open`);
    await seen('.rw-menu [role^="menuitem"]', 'the setups to open', 4000);
    const offered = await js(`[...document.querySelectorAll('.rw-menu [role^="menuitem"]')].map((n) => n.textContent.trim())`);
    await escape();
    await gone('.rw-menu', 'the menu closed', 3000);
    await js(PANEL('setup'));
    await until(() => js(IN_PANEL('setup')), 'Set up', 6000).catch(() => {});
    const onSetup = await has(`.st-panel[data-panel="setup"] ${BAND}`);
    // The Setups menu in a new session: saving it under a name, and no New setup beside it.
    const menuKeys = async () => {
      await click('.st-setup-btn');
      await seen('.rw-menu [role^="menuitem"]', 'the Setups menu', 4000);
      const keys = await js(`[...document.querySelectorAll('.rw-menu [role^="menuitem"]')].map((n) => n.dataset.key || '')`);
      await escape();
      await gone('.rw-menu', 'the Setups menu closed', 3000);
      return keys;
    };
    const inNew = await menuKeys();
    check('studio keep: a start with a setup saved opens a new session and says so in a line over Record and over Set up, with the saved setup to open; its Setups menu offers saving the session and no New setup',
      onMain && words.length > 20 && offered.length === 1 && /^Check/.test(offered[0]) && onSetup && inNew.includes('save-this-session-as-a-setup') && !inNew.includes('new-setup') && (await current()) === '',
      JSON.stringify({ onMain, words, offered, onSetup, inNew }));

    /* ── arranging ── */
    const arrange = async () => {
      await js(PANEL('setup'));
      await until(() => js(IN_PANEL('setup')), 'Set up', 6000).catch(() => {});
      await click('.st-scene[data-scene="screen-camera"]');
      await until(() => js(`document.querySelectorAll('.st-box').length >= 2`), 'the frame\'s boxes', 6000).catch(() => {});
      const afterScene = await has(BAND);
      await click('.st-card[data-kind="camera"]');
      await until(() => js(`${q('.st-bar')}?.dataset.kind === 'camera'`), 'the camera\'s properties', 4000).catch(() => {});
      await js(`(() => { const n = ${q('.st-box-w')}; if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, '20'); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
      await wait(300);
      return { afterScene, width: await js(`${q('.st-box-w')}?.value || ''`) };
    };
    const first = await arrange();
    const lineGone = await gone(BAND, 'the line gone once the session is arranged');
    check('studio keep: picking a scene is not arranging, so the line stays; a source resized is, and the line goes',
      first.afterScene === true && Number(first.width) === 20 && lineGone, JSON.stringify({ ...first, lineGone }));

    /* ── asked at the window's close; Cancel stays ── */
    win.close();
    const asked = await seen(ASK, 'Keep this session?');
    const buttons = await js(`['.st-keep-cancel', '.st-keep-drop', '.st-keep-save'].map((s) => Boolean(${q(ASK)}?.querySelector(s)))`);
    const said = await js(`${q(`${ASK} .rw-dialog-body`)}?.textContent.trim().length || 0`);
    await click(`${ASK} .st-keep-cancel`);
    const stayed = (await gone(ASK, 'the question closed')) && !win.isDestroyed();
    await wait(300);
    const stillThere = !win.isDestroyed() && Number(await js(`${q('.st-box-w')}?.value || ''`)) === 20;
    check('studio keep: closing the window on a new session that was arranged asks "Keep this session?" with Save as a setup, Don\'t save and Cancel; Cancel leaves the window and the session as they were',
      asked && buttons.every(Boolean) && said > 40 && stayed && stillThere, JSON.stringify({ asked, buttons, said, stayed, stillThere }));

    /* ── Reset says what it clears, fits its box, and asks before it does it ── */
    const RESET = '.st-confirm-dialog';
    const askReset = async () => {
      await click('.st-setup-btn');
      await seen('.rw-menu [data-key="reset-new-session"]', 'Reset in the Setups menu', 4000);
      await click('.rw-menu [data-key="reset-new-session"]');
      return seen(RESET, 'Reset asks first');
    };
    const resetAsked = await askReset();
    // The warning is read whole: nothing of it is cut off by its box, and the box stands inside the window.
    const fits = await js(`(() => {
      const d = ${q(RESET)}; const b = d?.querySelector('.rw-dialog-body'); const p = b?.querySelector('p');
      if (!d || !b || !p) return null;
      const r = d.getBoundingClientRect(); const pr = p.getBoundingClientRect(); const br = b.getBoundingClientRect();
      return { words: p.textContent.trim().length, cut: b.scrollHeight > b.clientHeight + 1 || b.scrollWidth > b.clientWidth + 1, inBody: pr.bottom <= br.bottom + 1 && pr.right <= br.right + 1,
        inWindow: r.left >= 0 && r.top >= 0 && r.right <= window.innerWidth && r.bottom <= window.innerHeight };
    })()`);
    await js(`([...document.querySelectorAll(${JSON.stringify(`${RESET} .rw-dialog-foot button`)})].find((n) => !n.classList.contains('st-confirm-ok'))?.click(), 1)`);
    const resetLeft = (await gone(RESET, 'the question closed', 4000)) && Number(await js(`${q('.st-box-w')}?.value || ''`)) === 20;
    await askReset();
    await click(`${RESET} .st-confirm-ok`);
    await gone(RESET, 'the question closed', 4000);
    const cleanAgain = await ok(until(() => js(`Number(${q('.st-box-w')}?.value || 0) !== 20`), 'the arrangement gone', 6000));
    const afterReset = await get(STORE.index);
    check('studio keep: Reset in a new session that was arranged asks first, its warning read whole inside its box; Cancel leaves the session, and Reset returns to a new session, what was arranged gone and the saved setup as it was',
      resetAsked && Boolean(fits) && fits.words > 80 && !fits.cut && fits.inBody && fits.inWindow && resetLeft && cleanAgain && (afterReset?.current?.recorder || '') === '' && (afterReset?.setups || []).length === 1,
      JSON.stringify({ resetAsked, fits, resetLeft, cleanAgain, setups: (afterReset?.setups || []).length }));
    await arrange();
    await gone(BAND, 'the line gone again', 4000);

    /* ── asked before a saved setup takes its place; Don't save goes on ── */
    await click('.st-setup-btn');
    await seen('.rw-menu [role="menuitemcheckbox"]', 'the saved setup in the menu', 4000);
    await click('.rw-menu [role="menuitemcheckbox"]');
    const askedAgain = await seen(ASK, 'asked before the setup opens');
    const notYet = (await current()) === '';
    await click(`${ASK} .st-keep-drop`);
    const opened = await ok(until(async () => (await current()) === CHECK_SETUP, 'the saved setup opened', 6000));
    await gone(ASK, 'the question closed', 3000);
    const inSaved = await menuKeys();
    const nothingMade = ((await get(STORE.index))?.setups || []).length === 1;
    check('studio keep: opening a saved setup over an arranged session asks first and opens nothing until answered; Don\'t save opens it and saves nothing, and with a saved setup open the menu offers New setup and not saving the session',
      askedAgain && notYet && opened && nothingMade && inSaved.includes('new-setup') && !inSaved.includes('save-this-session-as-a-setup'), JSON.stringify({ askedAgain, notYet, opened, nothingMade, inSaved }));

    /* ── a saved setup is written as it changes: its window closes unasked. A new one, arranged, is saved at the door ── */
    await newStart();
    const untouched = !(await has(ASK));
    const second = await arrange();
    const complaints = await errorsIn(win);
    win.close();
    const askedLast = await seen(ASK, 'asked at the close');
    await click(`${ASK} .st-keep-save`);
    const naming = await seen('.st-name-input', 'the name box');
    await js(`(() => { const n = ${q('.st-name-input')}; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(NAME)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(100);
    await click('.st-name-ok');
    const closed = await ok(until(() => win.isDestroyed() || null, 'the window closed after the save', 10000));
    // The Recorder is gone: what it kept is read through the window that stayed.
    const index = await get(STORE.index, wjs);
    made = (index?.setups || []).filter((x) => x.id !== CHECK_SETUP).map((x) => x.id);
    const entry = (index?.setups || []).find((x) => x.name === NAME) || null;
    const doc = entry ? await get(STORE.setup(entry.id), wjs) : null;
    const camera = doc?.stack?.items?.find((x) => x.kind === 'camera') || null;
    check('studio keep: a start after a saved setup was open is not asked about; Save as a setup at the close asks for a name, saves the session under it as it was arranged, and then the window closes',
      untouched && Number(second.width) === 20 && askedLast && naming && closed && Boolean(entry) && made.length === 1 && Boolean(camera) && Math.abs(camera.box.w - 0.2) < 0.002,
      JSON.stringify({ untouched, width: second.width, askedLast, naming, closed, saved: entry?.name || null, made: made.length, w: camera?.box?.w ?? null }));

    check('studio keep: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('studio keep: the check ran to the end', false, err?.message || String(err));
  } finally {
    if (wjs) {
      for (const id of made) await wjs(`window.rutbaOffice.store.set({ key: ${JSON.stringify(STORE.setup(id))}, value: null }).then(() => 1)`).catch(() => {});
      if (kept) await restoreSetup(wjs, kept).catch(() => {});
    }
    letGoOfStudioWindows();
  }
}
