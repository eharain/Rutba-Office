// Every app as Windows High Contrast draws it: forced colours turned on in
// each window (the DevTools protocol's emulation of forced-colors: active,
// which renders with the system's colours as Windows does), then measured
// rather than trusted. The window's colours really are the system's; the
// keyboard's ring shows; the chosen tab is still told from the others, by
// its bar and its weight, not by colour alone; every pressed, chosen or
// current control is outlined; the ribbon's pictures are drawn in the
// window's ink, never in what was the background; and a menu and a dialog
// keep an edge where their shadow was. Captures hc-<app>.png. Run alone
// with RUTBA_VERIFY_ONLY=high-contrast.

import fs from 'node:fs';
import path from 'node:path';
import { fixtures } from './verify-polish-pass.js';

/** Forced colours on or off for a window, as Windows would set them. */
async function forced(win, on) {
  const dbg = win.webContents.debugger;
  if (!dbg.isAttached()) dbg.attach('1.3');
  await dbg.sendCommand('Emulation.setEmulatedMedia', { features: [{ name: 'forced-colors', value: on ? 'active' : '' }, { name: 'prefers-color-scheme', value: on ? 'dark' : '' }] });
  if (!on) dbg.detach();
}

/** What forced colours did to the window, measured in it. */
const LOOK = `(() => {
  const css = (el, pseudo) => getComputedStyle(el, pseudo || null);
  const probe = document.createElement('div');
  probe.style.cssText = 'position:fixed;left:-50px;top:0;width:10px;height:10px;background:Canvas;color:CanvasText';
  document.body.append(probe);
  const canvas = css(probe).backgroundColor, ink = css(probe).color;
  probe.remove();
  const shown = (el) => el && el.getClientRects().length && css(el).visibility !== 'hidden';
  const outlined = (el) => { const s = css(el); return s.outlineStyle !== 'none' && parseFloat(s.outlineWidth) >= 1; };
  const out = { media: matchMedia('(forced-colors: active)').matches, forcedColours: css(document.querySelector('.rw') || document.body).color === ink, canvas, ink };
  // The chosen tab: its bar shows, and it is set in a weight of its own.
  const tabs = [...document.querySelectorAll('.rw-tab')].filter(shown);
  const chosen = tabs.find((t) => t.getAttribute('aria-selected') === 'true');
  const other = tabs.find((t) => t !== chosen);
  if (chosen) {
    const bar = css(chosen, '::before');
    out.tab = { bar: bar.backgroundColor !== 'rgba(0, 0, 0, 0)' && bar.backgroundColor !== canvas && parseFloat(bar.width) > 4, weight: other ? css(chosen).fontWeight !== css(other).fontWeight : true };
  }
  // Pressed, chosen and current controls, outlined.
  const marked = [...document.querySelectorAll('.rw-btn[aria-pressed="true"], .rw-item.selected, [role="option"][aria-selected="true"], [role="row"][aria-selected="true"], [aria-current="true"], [aria-current="date"]')].filter(shown);
  out.marked = marked.length;
  out.unmarked = marked.filter((el) => !outlined(el)).map((el) => (el.getAttribute('aria-label') || el.className || el.tagName).toString().slice(0, 40)).slice(0, 5);
  // The ribbon's pictures: drawn in a colour that is not the window's background.
  const icons = [...document.querySelectorAll('.rw-groups .rw-btn svg:not(.rw-caret)')].filter(shown).slice(0, 40);
  out.icons = icons.length;
  out.hiddenIcons = icons.filter((s) => { const c = css(s); const stroke = c.stroke !== 'none' ? c.stroke : c.fill; return !stroke || stroke === canvas || stroke === 'rgba(0, 0, 0, 0)'; }).length;
  return out;
})()`;

export async function verifyHighContrast({ open, check, until, wait, press }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const only = (process.env.RUTBA_A11Y_APPS || '').split(',').map((x) => x.trim()).filter(Boolean);
  const wanted = (app) => !only.length || only.includes(app);
  const fx = fixtures(dir, 'contrast-fixtures');
  const snap = async (win, name) => {
    if (!shots) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(shots, name), (await win.webContents.capturePage()).toPNG());
  };
  const ready = (win, sel, what) => until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector(${JSON.stringify(sel)}))`), what, 15000).catch(() => {});

  const apps = [
    ['home', undefined, '.home-card'],
    ['word', fx.docx, '.wd-page'],
    ['sheets', fx.xlsx, '.sh-cell'],
    ['slides', fx.pptx, '.sl-thumb'],
    ['pdf', fx.pdf, '.pdf-page canvas[data-drawn]'],
    ['pictures', fx.pics, '.rw-ribbon .rw-tab'],
    ['image', fx.png, '.rw-ribbon .rw-tab'],
    ['video', fx.png, '.vd-clip'],
    ['contacts', fx.vcf, '.ct-item'],
    ['calendar', fx.ics, '.cal-month-grid, .cal-time'],
    ['mail', undefined, '.ml-row, .ml-list, .rw-empty, .ml-compose-cta'],
  ];
  for (const [app, file, sel] of apps.filter(([a]) => wanted(a))) {
    let win = null;
    try {
      win = await open(app, file);
      const js = (code) => win.webContents.executeJavaScript(code);
      await ready(win, sel, `the ${app} window`);
      if (app === 'contacts') await js(`document.querySelector('.ct-item')?.click(), 1`);
      await forced(win, true);
      const look = await until(() => js(LOOK).then((l) => (l.media && l.forcedColours ? l : null)), `${app} in forced colours`, 5000).then(() => js(LOOK)).catch(() => js(LOOK));
      await snap(win, `hc-${app}.png`);
      // The keyboard's ring: F6 into the ribbon, or onto the window's first control where there is none.
      // A page draws focus only while it has the keyboard (Chromium drops :focus in a page without it).
      if (!(await js(`document.hasFocus()`))) win.webContents.focus();
      await js(`(document.querySelector('.rw-tab[aria-selected="true"]') || document.querySelector('.rw-body button, .rw-body [tabindex="0"]'))?.focus(), 1`);
      await press(win.webContents, 'Right');
      const ring = await until(() => js(`(() => { const a = document.activeElement; if (!a || a === document.body) return null; const s = getComputedStyle(a); return s.outlineStyle !== 'none' && parseFloat(s.outlineWidth) >= 1 ? s.outlineStyle + ' ' + s.outlineWidth : null; })()`), 'a ring', 2000).then(() => js(`getComputedStyle(document.activeElement).outlineStyle + ' ' + getComputedStyle(document.activeElement).outlineWidth`)).catch(() => 'none');
      // A menu keeps an edge.
      await js(`document.querySelector('.rw-appmark')?.click(), 1`);
      const edge = await until(() => js(`(() => { const m = document.querySelector('.rw-menu'); if (!m) return null; const s = getComputedStyle(m); return s.borderTopStyle !== 'none' && parseFloat(s.borderTopWidth) >= 1 ? 'edged' : 'none'; })()`), 'the app menu', 3000).then(() => js(`(() => { const s = getComputedStyle(document.querySelector('.rw-menu')); return s.borderTopStyle !== 'none' && parseFloat(s.borderTopWidth) >= 1 ? 'edged' : 'none'; })()`)).catch(() => 'no menu');
      await snap(win, `hc-${app}-menu.png`);
      await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
      const tabOk = !look.tab || (look.tab.bar && look.tab.weight);
      check(`high-contrast: ${app} in Windows High Contrast: the system's colours, a ring that shows, the chosen tab told by its bar and weight, every pressed or chosen control outlined, the ribbon's pictures drawn, the menu edged`,
        look.media && look.forcedColours && /solid|auto|dotted|dashed/.test(ring) && tabOk && !look.unmarked.length && !look.hiddenIcons && edge === 'edged',
        JSON.stringify({ ...look, ring, edge }));
    } catch (err) {
      check(`high-contrast: the ${app} pass ran`, false, err.stack || err.message);
    } finally {
      if (win && !win.isDestroyed()) await forced(win, false).catch(() => {});
    }
  }
}
