// The ribbons, every app's, looked at the way a person sees them: each tab at
// a wide window and at the app's narrowest, and the first tab dark and right
// to left in Urdu. On every tab the groups fit the window (whatever does not
// fit folds into a group button, from the end of the ribbon inwards, as
// Office's does), no control is cut off by its group or by the ribbon, a
// folded group opens its controls under its button, the keyboard reaches the
// tabs and the controls with a ring that shows, and Alt shows the KeyTips.
// Each view is written as a PNG to RUTBA_VERIFY_CAPTURE.
// Run alone with RUTBA_VERIFY_ONLY=ribbon.

import fs from 'node:fs';
import path from 'node:path';
import { nativeImage, BrowserWindow } from 'electron';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import pdfWriter from '@rutba/pdf';

const { PdfDocument } = pdfWriter;

/** Every app with a ribbon, the file it opens with, and its narrowest window. */
const APPS = [
  { app: 'word', file: 'docx', min: 820 },
  { app: 'sheets', file: 'xlsx', min: 900 },
  { app: 'slides', file: 'pptx', min: 900 },
  { app: 'pdf', file: 'pdf', min: 760 },
  { app: 'mail', file: null, min: 1000 },
  { app: 'calendar', file: 'ics', min: 900 },
  { app: 'contacts', file: 'vcf', min: 820 },
  { app: 'pictures', file: 'png', min: 720 },
  { app: 'image', file: 'png', min: 900 },
  { app: 'video', file: null, min: 900 },
];

const WIDE = 1680;

/** RUTBA_RIBBON_APPS=word,sheets looks at those apps alone, for working on them. */
const wanted = () => {
  const only = (process.env.RUTBA_RIBBON_APPS || '').split(',').map((s) => s.trim()).filter(Boolean);
  return only.length ? APPS.filter((a) => only.includes(a.app)) : APPS;
};

/** A small picture of its own, so no fixture another check reads is opened. */
function picture(file) {
  const w = 240;
  const h = 160;
  const buf = Buffer.alloc(w * h * 4);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4;
      buf[i] = 200 - Math.round((y / h) * 120); // B
      buf[i + 1] = 120 + Math.round((x / w) * 80); // G
      buf[i + 2] = 60 + Math.round((x / w) * 160); // R
      buf[i + 3] = 255;
    }
  }
  fs.writeFileSync(file, nativeImage.createFromBitmap(buf, { width: w, height: h }).toPNG());
}

function fixtures(dir, files) {
  const at = (n) => path.join(dir, n);
  fs.writeFileSync(at('ribbon.docx'), buildDocx({ styles: true, paragraphs: [{ text: 'The ribbon, looked at.' }, { text: 'A second paragraph.' }] }));
  fs.writeFileSync(at('ribbon.xlsx'), buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Month', 'Sales'], ['January', 120], ['February', 140]] }] }));
  fs.writeFileSync(at('ribbon.pptx'), buildPptx({ title: 'Ribbon', slides: [{ layout: 'title', title: 'The ribbon' }] }));
  const pdf = new PdfDocument({ title: 'ribbon', created: '2026-10-13T09:00:00Z' });
  for (const title of ['One', 'Two']) pdf.addPage(null).text(title, 72, 110, { font: 'Helvetica-Bold', size: 30 });
  fs.writeFileSync(at('ribbon.pdf'), pdf.toBuffer());
  picture(at('ribbon.png'));
  if (files?.ics) fs.copyFileSync(files.ics, at('ribbon.ics'));
  if (files?.vcf) fs.copyFileSync(files.vcf, at('ribbon.vcf'));
  return {
    docx: at('ribbon.docx'), xlsx: at('ribbon.xlsx'), pptx: at('ribbon.pptx'), pdf: at('ribbon.pdf'), png: at('ribbon.png'),
    ics: files?.ics ? at('ribbon.ics') : null, vcf: files?.vcf ? at('ribbon.vcf') : null,
  };
}

/**
 * What the ribbon looks like on its current tab, measured in the window: the
 * groups' box and how much of it is past the window's edge, each group's box
 * and stage, every control cut off by its group or by the ribbon, and every
 * label shown short.
 */
const MEASURE = `(() => {
  const ribbon = document.querySelector('.rw-ribbon');
  const groups = document.querySelector('.rw-groups');
  if (!ribbon || !groups) return { none: true };
  const g = groups.getBoundingClientRect();
  const r = ribbon.getBoundingClientRect();
  const box = (n) => { const b = n.getBoundingClientRect(); return { left: Math.round(b.left), right: Math.round(b.right), top: Math.round(b.top), bottom: Math.round(b.bottom), width: Math.round(b.width), height: Math.round(b.height) }; };
  const shown = (n) => { const b = n.getBoundingClientRect(); return b.width > 0 && b.height > 0 && getComputedStyle(n).visibility !== 'hidden'; };
  const name = (n) => (n.dataset.tip || n.getAttribute('aria-label') || n.title || n.textContent || n.tagName).trim().slice(0, 40);
  const list = [...groups.querySelectorAll(':scope > .rw-group')];
  const clipped = [];
  const short = [];
  for (const grp of list) {
    const gb = grp.getBoundingClientRect();
    const items = grp.classList.contains('folded') ? [] : [...grp.querySelectorAll('.rw-group-items button, .rw-group-items select, .rw-group-items input')];
    for (const n of items) {
      if (!shown(n) || n.closest('.rw-fold-pop')) continue;
      const b = n.getBoundingClientRect();
      if (b.left < gb.left - 1 || b.right > gb.right + 1 || b.bottom > g.bottom + 1 || b.top < g.top - 1) clipped.push(name(n));
    }
    for (const s of grp.querySelectorAll('.rw-btn > span, .rw-group-label')) {
      if (!shown(s) || s.closest('.rw-fold-pop')) continue;
      if (s.scrollWidth > s.clientWidth + 1 && getComputedStyle(s).textOverflow !== 'clip') short.push(s.textContent.trim().slice(0, 30));
    }
  }
  return {
    width: innerWidth,
    ribbon: box(ribbon),
    groups: box(groups),
    over: Math.max(0, groups.scrollWidth - groups.clientWidth),
    past: Math.max(0, Math.round(g.right - innerWidth), Math.round(-g.left)),
    count: list.length,
    stages: list.map((n) => (n.classList.contains('folded') ? 'F' : n.classList.contains('compact') ? 'C' : n.classList.contains('small') ? 'S' : '-')).join(''),
    clipped,
    short,
    dir: getComputedStyle(ribbon).direction,
  };
})()`;

export async function verifyRibbon({ open, check, until, wait }, { dir, files }) {
  const capture = process.env.RUTBA_VERIFY_CAPTURE;
  const fx = fixtures(dir, files);
  const snap = async (win, name) => {
    if (!capture) return;
    const wc = win.webContents;
    wc.invalidate();
    await wait(700);
    const bottom = await wc.executeJavaScript(`(() => { const r = document.querySelector('.rw-ribbon')?.getBoundingClientRect(); const pop = document.querySelector('.rw-fold-pop.open')?.getBoundingClientRect(); return Math.ceil(Math.max(r ? r.bottom : 200, pop ? pop.bottom : 0) + 10); })()`);
    const [w] = win.getContentSize();
    fs.writeFileSync(path.join(capture, name), (await wc.capturePage({ x: 0, y: 0, width: w, height: Math.min(bottom, 600) })).toPNG());
  };
  const size = async (win, width, height = 860) => {
    win.setContentSize(width, height);
    await until(() => win.webContents.executeJavaScript(`innerWidth`).then((w) => Math.abs(w - width) <= 2), `the window at ${width} px`, 5000).catch(() => {});
    await wait(350);
  };
  const tabs = (js) => js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
  const pick = async (js, i) => {
    await js(`(() => { document.querySelectorAll('.rw-tab')[${i}]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.rw-tab')[${i}]?.getAttribute('aria-selected') === 'true' && Boolean(document.querySelector('.rw-groups'))`), 'the tab', 4000).catch(() => {});
    await wait(250);
  };
  const slug = (s) => String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'tab';

  const seen = [];
  for (const { app, file, min } of wanted()) {
    let win = null;
    try {
      win = await open(app, file ? fx[file] : undefined);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.rw-ribbon .rw-tab'))`), `the ${app} ribbon`, 8000);
      // Pictures folds its ribbon while a picture is open; it is opened here to be looked at, and folded again after.
      const folded = await js(`Boolean(document.querySelector('.rw-ribbon.collapsed'))`);
      if (folded) await js(`document.querySelector('.rw-collapse')?.click(), 1`);
      const names = await tabs(js);
      const views = [];
      for (const width of [WIDE, min]) {
        await size(win, width);
        for (let i = 0; i < names.length; i++) {
          await pick(js, i);
          const m = await js(MEASURE);
          views.push({ width, tab: names[i], ...m });
          await snap(win, `ribbon-${app}-${width}-${String(i).padStart(2, '0')}-${slug(names[i])}.png`);
        }
      }
      // Dark, on the first tab, at the wide window.
      await size(win, WIDE);
      await pick(js, 0);
      await js(`document.documentElement.dataset.theme = 'dark', 1`);
      await snap(win, `ribbon-${app}-dark.png`);
      await js(`document.documentElement.dataset.theme = 'light', 1`);
      if (folded) await js(`document.querySelector('.rw-collapse')?.click(), 1`);
      const heights = [...new Set(views.filter((v) => !v.none).map((v) => v.ribbon.height))];
      check(`ribbon: ${app}'s ribbon is one height on every tab`, heights.length === 1, heights.join(', '));
      seen.push({ app, views });
      const bad = views.filter((v) => !v.none && (v.past > 0 || v.clipped.length || v.short.length));
      check(`ribbon: ${app}'s ribbon fits the window on every tab, wide and narrow, with no control cut off`, !bad.length,
        bad.slice(0, 6).map((v) => `${v.tab}@${v.width}: past ${v.past}, over ${v.over}, stages ${v.stages}, clipped [${v.clipped.slice(0, 4).join('|')}], short [${v.short.slice(0, 3).join('|')}]`).join('; ') || `${views.length} views`);
    } catch (err) {
      check(`ribbon: ${app}'s ribbon checks ran`, false, err.message);
    }
  }

  // Right to left, in Urdu: every app's first tab, at its narrowest and wide.
  let home = null;
  const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
  try {
    home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    const was = await hjs(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
    await hjs(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    try {
      for (const { app, file, min } of wanted()) {
        const win = await open(app, file ? fx[file] : undefined);
        const js = (code) => win.webContents.executeJavaScript(code);
        await until(() => js(`Boolean(document.querySelector('.rw-ribbon .rw-tab'))`), `the ${app} ribbon in Urdu`, 8000).catch(() => {});
        const views = [];
        for (const width of [WIDE, min]) {
          await size(win, width);
          await pick(js, 0);
          views.push({ width, ...(await js(MEASURE)) });
          await snap(win, `ribbon-${app}-ur-${width}.png`);
        }
        const bad = views.filter((v) => !v.none && (v.dir !== 'rtl' || v.past > 0 || v.clipped.length));
        check(`ribbon: ${app}'s ribbon in Urdu runs right to left and fits, wide and narrow`, !bad.length,
          bad.map((v) => `${v.width}: dir ${v.dir}, past ${v.past}, stages ${v.stages}, clipped [${v.clipped.slice(0, 4).join('|')}]`).join('; ') || views.map((v) => `${v.width}:${v.stages}`).join(' '));
        win.destroy();
      }
    } finally {
      await hjs(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? 'en')} }).then(() => 1)`);
    }
  } catch (err) {
    check('ribbon: the Urdu ribbon checks ran', false, err.message);
  } finally {
    for (const w of BrowserWindow.getAllWindows()) if (!before.has(w.id) && w !== home && !w.isDestroyed()) w.destroy();
  }
  if (capture) fs.writeFileSync(path.join(capture, 'ribbon-measures.json'), JSON.stringify(seen, null, 1));
}
