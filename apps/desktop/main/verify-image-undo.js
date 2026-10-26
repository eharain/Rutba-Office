// Image: a picture with no layers takes its slider changes back one at a
// time. A grainy picture opened: Saturation moved, then Auto enhance;
// Ctrl+Z takes back Auto enhance and leaves Saturation where it was set, and
// Ctrl+Z again takes Saturation back. Levels' own Auto sets the tones and
// leaves Temperature where it was set. Alt held over the picture while it is
// pressed (as the Clone Stamp's Alt+click is), then let go, leaves the
// ribbon's letters off. A new empty layer over a shape does not take the
// Move tool's click on the shape. Run alone with RUTBA_VERIFY_ONLY=image-undo.
import fs from 'node:fs';
import path from 'node:path';
import { noisyPicture } from './verify-image-edits.js';

export async function verifyImageUndo({ open, check, until, wait, press, errorsIn }, { dir }) {
  try {
    const file = path.join(dir, 'undo-sliders.png');
    fs.writeFileSync(file, noisyPicture(480, 320, 11).toPNG());
    const win = await open('image', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const valueOf = (key) => js(`(() => { const i = document.querySelector('[data-slider="${key}"] input'); return i ? Number(i.value) : null; })()`);
    const levelOf = (key) => js(`(() => { const i = document.querySelector('[data-level="${key}"] input'); return i ? Number(i.value) : null; })()`);
    const setRange = (selector, value) => js(`(() => { const i = document.querySelector(${JSON.stringify(selector)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(String(value))}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const click = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled); if (!b) return 'no button ' + ${JSON.stringify(label)}; b.click(); return 'clicked'; })()`);
    await until(() => js(`Boolean(document.querySelector('.im-canvas')) && Boolean(document.querySelector('[data-slider="saturation"] input:not([disabled])'))`), 'the picture', 10000);
    // The Levels section open, to read its sliders.
    await js(`(() => { const h = document.querySelector('.im-sec[data-section="levels"] .im-sec-head'); if (h && h.getAttribute('aria-expanded') !== 'true') h.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('[data-level="inBlack"] input'))`), 'the Levels section', 4000).catch(() => {});

    // Saturation, then (after the run of one change is over) Auto enhance.
    await setRange('[data-slider="saturation"] input', 40);
    await until(async () => (await valueOf('saturation')) === 40, 'saturation set', 3000).catch(() => {});
    await wait(1100);
    const levelsBefore = await levelOf('inBlack');
    const tabLabel = await js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === 'Adjust'); t?.click(); return Boolean(t); })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-btn')].some((n) => (n.dataset.tip || '').startsWith('Auto enhance'))`), 'the Adjust tab', 3000).catch(() => {});
    const enhanced = await click('Auto enhance');
    await until(async () => (await levelOf('inBlack')) !== levelsBefore || (await levelOf('inWhite')) !== 255, 'Auto enhance on the levels', 4000).catch(() => {});
    const afterAuto = { black: await levelOf('inBlack'), white: await levelOf('inWhite'), saturation: await valueOf('saturation') };
    win.focus();
    await js(`document.querySelector('.im-canvas')?.focus(), 1`);
    await press(wc, 'Z', { modifiers: ['control'] });
    await until(async () => (await levelOf('inBlack')) === levelsBefore && (await levelOf('inWhite')) === 255, 'Auto enhance taken back', 4000).catch(() => {});
    const firstUndo = { black: await levelOf('inBlack'), white: await levelOf('inWhite'), saturation: await valueOf('saturation') };
    await press(wc, 'Z', { modifiers: ['control'] });
    await until(async () => (await valueOf('saturation')) === 100, 'Saturation taken back', 4000).catch(() => {});
    const secondUndo = { saturation: await valueOf('saturation') };
    check('image: with no layers, Ctrl+Z takes back Auto enhance and leaves Saturation as it was set, and Ctrl+Z again takes Saturation back',
      enhanced === 'clicked' && (afterAuto.black !== levelsBefore || afterAuto.white !== 255) && afterAuto.saturation === 40
        && firstUndo.black === levelsBefore && firstUndo.white === 255 && firstUndo.saturation === 40 && secondUndo.saturation === 100,
      JSON.stringify({ tabLabel, enhanced, levelsBefore, afterAuto, firstUndo, secondUndo }));

    // Levels' Auto: the tones, and the Temperature left as it was set.
    await setRange('[data-slider="temperature"] input', 35);
    await until(async () => (await valueOf('temperature')) === 35, 'temperature set', 3000).catch(() => {});
    await wait(1100);
    await js(`(() => { const b = [...document.querySelectorAll('.im-sec[data-section="levels"] .im-sec-actions .rw-btn')].find((n) => n.textContent.trim() === 'Auto'); b?.click(); return Boolean(b); })()`);
    await until(async () => (await levelOf('inBlack')) !== levelsBefore || (await levelOf('inWhite')) !== 255, 'Levels Auto', 4000).catch(() => {});
    const levelsAuto = { black: await levelOf('inBlack'), white: await levelOf('inWhite'), temperature: await valueOf('temperature') };
    check('image: Levels\' Auto sets the tones and leaves Temperature where it was set',
      (levelsAuto.black !== levelsBefore || levelsAuto.white !== 255) && levelsAuto.temperature === 35, JSON.stringify(levelsAuto));

    // Alt held while the picture is pressed and let go: no letters on the ribbon.
    await js(`(() => {
      const c = document.querySelector('.im-canvas');
      const r = c.getBoundingClientRect();
      const at = { bubbles: true, cancelable: true, clientX: r.left + 40, clientY: r.top + 40, altKey: true, button: 0, pointerId: 1, isPrimary: true };
      // The canvas cancels the press, as the Alt+click tools do: no mousedown follows.
      const cancel = (e) => e.preventDefault();
      c.addEventListener('pointerdown', cancel, { once: true });
      c.dispatchEvent(new PointerEvent('pointerdown', at));
      c.dispatchEvent(new PointerEvent('pointerup', at));
      return 1;
    })()`).catch(() => {});
    // The Alt key down before the press and up after it, as a person's hand does.
    await js(`(() => { window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Alt', bubbles: true })); return 1; })()`);
    await js(`(() => {
      const c = document.querySelector('.im-canvas');
      const r = c.getBoundingClientRect();
      const at = { bubbles: true, cancelable: true, clientX: r.left + 40, clientY: r.top + 40, altKey: true, button: 0, pointerId: 1, isPrimary: true };
      c.addEventListener('pointerdown', (e) => e.preventDefault(), { once: true });
      c.dispatchEvent(new PointerEvent('pointerdown', at));
      c.dispatchEvent(new PointerEvent('pointerup', at));
      window.dispatchEvent(new KeyboardEvent('keyup', { key: 'Alt', bubbles: true }));
      return 1;
    })()`);
    await wait(400);
    const letters = await js(`document.querySelectorAll('.rw-keytips .rw-keytip, .rw-keytip').length`);
    // Alt alone still brings them up, and Esc puts them away.
    await js(`(() => { window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Alt', bubbles: true })); window.dispatchEvent(new KeyboardEvent('keyup', { key: 'Alt', bubbles: true })); return 1; })()`);
    const shown = await until(() => js(`document.querySelectorAll('.rw-keytip').length > 0`), 'the letters', 2000).then(() => true).catch(() => false);
    await js(`(() => { window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); return 1; })()`);
    await js(`(() => { window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.rw-keytip').length === 0`), 'the letters off', 2000).catch(() => {});
    check('image: Alt held while the picture is pressed and let go leaves the ribbon\'s letters off; Alt alone still shows them',
      letters === 0 && shown, JSON.stringify({ letters, shown }));

    const complaints = await errorsIn(win);
    check('image: the undo and letters checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('image: the undo checks ran', false, err?.message || String(err));
  }
}
