// Feature screenshots for the website: the second kind of capture
// `build/make-screens.js` takes.
//
// The app screenshots open each app on a sample and photograph it as it
// first draws. A feature does not show on first draw: tracked changes need
// All Markup, a PivotTable needs its field list, the Slide Master needs its
// view. So each feature in screen-features.js opens its app on a sample the
// screens build made, and its `setup` drives the window into the state that
// shows the feature, with the same moves the window checks make (a tab, a
// button found by its tip, a field set the way typing sets it, a key). Then
// the window is photographed and closed.
//
// Runs inside the application's main process, started by make-screens.js
// with RUTBA_OFFICE_SMOKE and RUTBA_SCREENS_FEATURES (the names to take).

import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow } from 'electron';
import { readImage, writeImage, readText, writeText } from '@rutba/office-shell/electron/clipboard';
import { consoleMessage } from './console-message.js';
import { FEATURES } from './screen-features.js';

const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function until(condition, what, timeout = 8000) {
  const deadline = Date.now() + timeout;
  for (;;) {
    let held = false;
    try {
      held = await condition();
    } catch {
      held = false;
    }
    if (held) return held;
    if (Date.now() > deadline) throw new Error(`waited ${timeout} ms for ${what} and it never happened`);
    await wait(80);
  }
}

/** What each app draws first, as the window checks wait for it. */
const READY_IN = {
  home: '.home-card',
  word: '.wd-block',
  sheets: '.sh-cells',
  slides: '.sl-svg, .sl-thumb',
  pictures: '.pv-image, .pv-pdf, .pv-video, .pv-audio, .pv-tile, .pv-empty',
  image: '.im-canvas',
  video: '.vd-clip, .vd-empty, .vp-stage',
  recorder: '.st-canvas',
  pdf: '.pdf-thumb, .pdf-empty',
  mail: '.ml-row, .ml-list, .ml-compose-cta',
  calendar: '.rw-app',
  contacts: '.rw-app',
  talk: '.tk-signin, .tk-contact',
  base: '.bs-start, .bs-app',
};

/**
 * The moves a setup makes, bound to one window. Each answers rather than
 * throws where a check would record a detail: a setup that could not do a
 * step says which, and the capture is reported as failed.
 */
function harness(win, { samples, doc, stores, mail, broadcast, windows }) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const h = {
    win,
    wc,
    js,
    wait,
    doc,
    stores,
    mail,
    broadcast,
    windows,
    samples,
    sample: (name) => path.join(samples, name),
    /** Waits for a condition in the page (an expression) to be truthy, and answers it. */
    until: (expr, what = expr, ms = 8000) => until(() => js(expr), what, ms),
    /** Waits for a condition evaluated here, in the main process. */
    untilHere: (fn, what, ms = 8000) => until(fn, what, ms),
    /** A ribbon tab by its label. */
    tab: async (label) => {
      const done = await js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!t) return false; t.click(); return true; })()`);
      if (!done) throw new Error(`no tab ${label}`);
      await wait(300);
    },
    /** A button whose tip, title or words begin with the label. */
    button: async (label, scope = '.rw-btn, button') => {
      const find = `[...document.querySelectorAll(${JSON.stringify(scope)})].find((n) => [n.dataset.tip, n.title, n.getAttribute('aria-label'), n.textContent].some((w) => (w || '').trim().startsWith(${JSON.stringify(label)})) && !n.disabled && n.offsetParent !== null)`;
      await until(() => js(`Boolean(${find})`), `the ${label} button`, 5000);
      await js(`(() => { const b = ${find}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return true; })()`);
      await wait(300);
    },
    /** An item of an open menu, by its words. */
    menu: async (label) => {
      const find = `[...document.querySelectorAll('.rw-menu button, [role="menuitem"], [role="menuitemradio"], [role="option"]')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(label)}))`;
      await until(() => js(`Boolean(${find})`), `the ${label} menu item`, 5000);
      await js(`(${find}).click(), true`);
      await wait(300);
    },
    /** Clicks the first element matching a selector. */
    click: async (selector, what = selector) => {
      await until(() => js(`Boolean(document.querySelector(${JSON.stringify(selector)}))`), what, 6000);
      await js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return true; })()`);
      await wait(250);
    },
    /** Sets an input, textarea or select the way typing or choosing does. */
    field: async (selector, value) => {
      await until(() => js(`Boolean(document.querySelector(${JSON.stringify(selector)}))`), selector, 5000);
      await js(`(() => {
        const el = document.querySelector(${JSON.stringify(selector)});
        const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
        Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
      })()`);
      await wait(150);
    },
    /** A key press, with modifiers ('control', 'shift', 'alt'). */
    key: async (keyCode, modifiers = []) => {
      wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
      wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
      await wait(150);
    },
    /** Types words into whatever has the keyboard. */
    type: async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(35);
      }
    },
    /** A real click at a point of the page. */
    clickAt: async (x, y, clickCount = 1) => {
      wc.sendInputEvent({ type: 'mouseMove', x, y });
      wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount });
      wc.sendInputEvent({ type: 'mouseUp', x, y, button: 'left', clickCount });
      await wait(200);
    },
    /** A real drag between two points of the page. */
    drag: async (from, to, steps = 12) => {
      wc.sendInputEvent({ type: 'mouseMove', x: from.x, y: from.y });
      wc.sendInputEvent({ type: 'mouseDown', x: from.x, y: from.y, button: 'left', clickCount: 1 });
      for (let i = 1; i <= steps; i++) {
        const x = Math.round(from.x + ((to.x - from.x) * i) / steps);
        const y = Math.round(from.y + ((to.y - from.y) * i) / steps);
        wc.sendInputEvent({ type: 'mouseMove', x, y, button: 'left' });
        await wait(16);
      }
      wc.sendInputEvent({ type: 'mouseUp', x: to.x, y: to.y, button: 'left', clickCount: 1 });
      await wait(250);
    },
    /** The middle of the first element matching a selector, in page pixels. */
    centre: (selector) => js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), left: r.left, top: r.top, width: r.width, height: r.height } : null; })()`),
    /** Queues the answer a person would give the next file dialog of a kind. */
    answer: (kind, value) => {
      globalThis.__rutbaCheckDialogAnswers = globalThis.__rutbaCheckDialogAnswers || [];
      globalThis.__rutbaCheckDialogAnswers.push({ kind, answer: value });
    },
    /** Every window open now, to tell a new one from. */
    allWindows: () => BrowserWindow.getAllWindows(),
    /** A window the app opened of its own (a slide show, the presenter view). */
    newWindow: async (before, what) => until(() => BrowserWindow.getAllWindows().find((w) => !before.includes(w) && !w.isDestroyed()), what, 10000),
  };
  return h;
}

/**
 * Takes the toasts off before the picture: a toast says what just happened,
 * and in a still it reads as a notice the product left lying about.
 */
const clean = (blur) => `(() => {
  // Hidden rather than removed: the toasts are React's to take away.
  document.querySelectorAll('.rw-toast-host').forEach((n) => { n.style.visibility = 'hidden'; });
  document.querySelectorAll('.rw-tip.on').forEach((n) => n.classList.remove('on'));
  ${blur ? 'document.activeElement?.blur?.();' : ''}
  return true;
})()`;

async function photograph(win, file) {
  win.webContents.invalidate();
  await wait(700);
  const png = (await win.webContents.capturePage()).toPNG();
  fs.writeFileSync(file, png);
  return png.length;
}

export async function runFeatureScreens({ windows, outDir, samples, doc, stores, mail, broadcast }) {
  const asked = (process.env.RUTBA_SCREENS_FEATURES || '').split(',').filter(Boolean);
  const list = FEATURES.filter((f) => asked.includes(f.name));
  fs.mkdirSync(outDir, { recursive: true });
  // Mail's features read the screens build's invented mailbox, as the app
  // capture does.
  if (process.env.RUTBA_SMOKE_SEED && mail) {
    const { seedMail } = await import('./seed-mail.js');
    await seedMail({ stores, mail, name: process.env.RUTBA_SMOKE_SEED_NAME || undefined }).catch((e) => console.log(`seeding failed: ${e.message}`));
  }
  // Copy Text puts what it read on the clipboard, which is the person's:
  // what was on it goes back at the end.
  const held = { text: await readText(), image: await readImage() };
  const lines = [];
  // The first window stays to the end: closing the last window quits the
  // application, and a run of thirty would end after its first.
  let keeper = null;
  for (const feature of list) {
    const started = Date.now();
    const messages = [];
    let ok = false;
    let size = 0;
    const before = BrowserWindow.getAllWindows();
    let win = null;
    try {
      const file = feature.file ? path.join(samples, feature.file) : null;
      // What has to be in place before the window opens: a message
      // delivered, a rating stored, a setting the window reads as it starts.
      if (feature.before) await feature.before({ stores, mail, samples, sample: (name) => path.join(samples, name), wait, profile: app.getPath('userData') });
      win = windows.create({ app: feature.app, file, ...(feature.query ? { query: feature.query } : {}) });
      win.webContents.on('console-message', (...args) => {
        const m = consoleMessage(args);
        if (m.level >= 3) messages.push(m.text.split('\n')[0].slice(0, 200));
      });
      await new Promise((resolve) => win.webContents.once('did-finish-load', resolve));
      const ready = `${READY_IN[feature.app] || '.rw-app'}, .rw-empty`;
      await until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector(${JSON.stringify(ready)}))`), `the ${feature.app} window to draw`, 15000).catch(() => {});
      await wait(feature.settle ?? 1500);
      win.webContents.focus();
      const h = harness(win, { samples, doc, stores, mail, broadcast, windows });
      // A setup may answer another window to photograph (the presenter view).
      const shown = (await feature.setup(h)) || win;
      // A setup whose subject has the keyboard (an open list, KeyTips, Find)
      // keeps it; the rest lose the caret's blink and the focus ring.
      await shown.webContents.executeJavaScript(clean(!feature.keepFocus)).catch(() => {});
      await wait(feature.after ?? 600);
      size = await photograph(shown, path.join(outDir, `${feature.name}.png`));
      ok = size > 5000;
    } catch (err) {
      messages.push(`setup: ${err?.message || err}`);
    }
    // What the feature started for its picture (an AI app connected) goes.
    if (feature.done) await feature.done().catch((err) => messages.push(`done: ${err?.message || err}`));
    lines.push(`${ok ? 'ok  ' : 'FAIL'} ${feature.name.padEnd(26)} ${String(Date.now() - started).padStart(6)} ms ${String(Math.round(size / 1024)).padStart(5)} KB`);
    for (const m of messages) lines.push(`       ${m}`);
    console.log(lines.slice(-1 - messages.length).join('\n'));
    // Every window this feature opened goes, but the first of the run.
    for (const w of BrowserWindow.getAllWindows()) {
      if (before.includes(w) || w.isDestroyed()) continue;
      if (!keeper) {
        keeper = w;
        w.hide();
        continue;
      }
      try {
        windows.forceClose ? windows.forceClose(w) : w.destroy();
      } catch {
        /* already going */
      }
    }
    await wait(300);
  }
  if (!held.image.isEmpty()) await writeImage(held.image);
  else await writeText(held.text);
  const report = lines.join('\n');
  fs.writeFileSync(path.join(outDir, 'report.txt'), `${report}\n`);
  return lines.every((l) => !l.startsWith('FAIL'));
}
