// A gallery in the ribbon shrinks before its group folds, as Office's do.
// Word's Home tab at 1,680 px shows its Styles gallery at its full length
// with the gallery's up, down and more arrows; at 1,300 px the gallery is
// shorter, the group still open, and the arrows page through the styles it
// no longer shows. Narrowing the window step by step, Word's Styles and
// Presentations' Themes and Variants never fold before their gallery is at
// its fewest items, the count only falls as the window narrows, and the
// ribbon's own rules hold at every step: nothing cut off, nothing past the
// window, the folded groups the last ones. Captures go to
// RUTBA_VERIFY_CAPTURE. Run alone with RUTBA_VERIFY_ONLY=gallery-fold.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import { MEASURE } from './verify-ribbon.js';

/** Every gallery on the tab: the group it is in, how many items it holds and shows, whether that group is folded, and its arrows. */
const GALLERIES = `(() => [...document.querySelectorAll('.rw-groups > .rw-group')].flatMap((g) => [...g.querySelectorAll('.rw-gallery')].map((gal) => {
  const items = gal.querySelector('.rw-gallery-items');
  const box = items?.getBoundingClientRect();
  const kids = items ? [...items.children] : [];
  const seen = kids.filter((k) => { const r = k.getBoundingClientRect(); return box && r.width > 0 && r.left >= box.left - 1 && r.right <= box.right + 1; }).length;
  return {
    group: g.querySelector('.rw-group-label')?.textContent.trim() || '',
    folded: g.classList.contains('folded'),
    items: kids.length,
    shown: Number(gal.dataset.shown || 0),
    seen,
    max: Number(gal.dataset.max || 0),
    min: Number(gal.dataset.min || 0),
    arrows: ['.rw-gallery-up', '.rw-gallery-down', '.rw-gallery-more'].every((s) => gal.querySelector(s)),
  };
})))()`;

export async function verifyGalleryFold({ open, check, until, wait }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const own = path.join(dir, 'gallery-fold-fixtures');
  fs.mkdirSync(own, { recursive: true });
  const docx = path.join(own, 'gallery.docx');
  const pptx = path.join(own, 'gallery.pptx');
  fs.writeFileSync(docx, buildDocx({ styles: true, paragraphs: [{ text: 'A gallery shrinks before it folds.' }, { text: 'A second paragraph.' }] }));
  fs.writeFileSync(pptx, buildPptx({ title: 'Gallery', slides: [{ layout: 'title', title: 'Galleries' }] }));

  const size = async (win, width, height = 860) => {
    win.setContentSize(width, height);
    await until(() => win.webContents.executeJavaScript('innerWidth').then((w) => Math.abs(w - width) <= 2), `the window at ${width} px`, 5000).catch(() => {});
    await wait(350);
  };
  const snap = async (win, name, full = false) => {
    if (!shots) return;
    const wc = win.webContents;
    wc.invalidate();
    await wait(700);
    const [w, h] = win.getContentSize();
    const bottom = full ? h : await wc.executeJavaScript(`Math.ceil((document.querySelector('.rw-ribbon')?.getBoundingClientRect().bottom || 200) + 10)`);
    fs.writeFileSync(path.join(shots, name), (await wc.capturePage({ x: 0, y: 0, width: w, height: Math.min(bottom, h) })).toPNG());
  };
  const pick = async (js, label) => {
    await js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); t?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-tab')].find((n) => n.textContent.trim() === ${JSON.stringify(label)})?.getAttribute('aria-selected') === 'true' && Boolean(document.querySelector('.rw-groups .rw-gallery'))`), `the ${label} tab with its gallery`, 5000).catch(() => {});
    await wait(250);
  };
  const ok = (m) => !m.none && m.past === 0 && m.over <= 1 && !m.clipped.length && /^[-CS]*F*$/.test(m.stages);

  /** Widths from wide to narrow: each gallery's count never rises, and its group folds only once it is at its fewest. */
  const sweep = async (win, js, app, tab, groups) => {
    const steps = [];
    for (let width = 1680; width >= 820; width -= 60) {
      await size(win, width);
      await until(() => js(GALLERIES).then((l) => l.every((g) => g.folded || g.seen === g.shown)), `${app} ${tab} galleries settled at ${width}`, 3000).catch(() => {});
      steps.push({ width, galleries: await js(GALLERIES), m: await js(MEASURE) });
    }
    const bad = [];
    for (const name of groups) {
      let last = Infinity;
      let lastShown = null;
      let stagesBefore = null;
      for (const s of steps) {
        // A gallery may take back room another group gave up at this width, by shrinking its buttons or folding.
        const roomFreed = stagesBefore !== null && s.m.stages !== stagesBefore;
        stagesBefore = s.m.stages;
        const g = s.galleries.find((x) => x.group === name);
        if (!g) { bad.push(`${name} missing at ${s.width}`); break; }
        if (g.folded) {
          if (lastShown !== null && lastShown > g.min) bad.push(`${name} folded at ${s.width} with ${lastShown} shown the step before (fewest ${g.min})`);
          break;
        }
        if (g.shown > last && !roomFreed) bad.push(`${name} grew to ${g.shown} at ${s.width} with no other group shrunk there`);
        if (g.seen !== g.shown) bad.push(`${name} shows ${g.seen} of the ${g.shown} it should at ${s.width}`);
        last = g.shown;
        lastShown = g.shown;
      }
    }
    const rules = steps.filter((s) => !ok(s.m)).map((s) => `${s.width}: past ${s.m.past}, over ${s.m.over}, stages ${s.m.stages}, clipped [${s.m.clipped.slice(0, 3).join('|')}]`);
    const trail = groups.map((name) => `${name} ${steps.map((s) => { const g = s.galleries.find((x) => x.group === name); return !g ? '?' : g.folded ? 'F' : g.shown; }).join(',')}`).join('; ');
    check(`gallery-fold: ${app} ${tab}: ${groups.join(' and ')} shrink item by item as the window narrows (growing back only into room another group gave up) and fold only at their fewest, the ribbon fitting at every width from 1680 to 820`, !bad.length && !rules.length,
      `${bad.slice(0, 4).join('; ') || 'monotonic'}; rules ${rules.slice(0, 3).join('; ') || 'kept'}; ${trail}`);
  };

  // Word, Home: the Styles gallery at 1680 and 1300.
  try {
    const win = await open('word', docx);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.rw-ribbon .rw-tab'))`), 'the Word ribbon', 8000);
    await pick(js, 'Home');
    await size(win, 1680);
    const wide = (await js(GALLERIES)).find((g) => g.group === 'Styles');
    const wideM = await js(MEASURE);
    await snap(win, 'gallery-word-home-1680.png');
    await snap(win, 'gallery-word-home-1680-window.png', true);
    check('gallery-fold: Word Home at 1680 px shows the Styles gallery at its full length, with its up, down and more arrows',
      Boolean(wide && !wide.folded && wide.shown === Math.min(wide.max, wide.items) && wide.seen === wide.shown && wide.arrows) && ok(wideM), JSON.stringify({ wide, stages: wideM.stages }));

    await size(win, 1300);
    await until(() => js(GALLERIES).then((l) => l.some((g) => g.group === 'Styles' && g.seen === g.shown)), 'the gallery settled', 3000).catch(() => {});
    const mid = (await js(GALLERIES)).find((g) => g.group === 'Styles');
    const midM = await js(MEASURE);
    await snap(win, 'gallery-word-home-1300.png');
    await snap(win, 'gallery-word-home-1300-window.png', true);
    check('gallery-fold: Word Home at 1300 px keeps the Styles group open with a shorter gallery, not folded into a button, and the ribbon fits',
      Boolean(mid && !mid.folded && mid.shown < (wide?.shown || 0) && mid.shown >= mid.min && mid.seen === mid.shown) && ok(midM), JSON.stringify({ mid, stages: midM.stages, clipped: midM.clipped }));

    // The arrows: down shows the styles after these, up the ones before, more all of them.
    const first = () => js(`(() => { const gal = [...document.querySelectorAll('.rw-groups .rw-gallery')][0]; const items = gal.querySelector('.rw-gallery-items'); const box = items.getBoundingClientRect(); return [...items.children].findIndex((k) => k.getBoundingClientRect().left >= box.left - 1); })()`);
    const press = (sel) => js(`(() => { const b = [...document.querySelectorAll('.rw-groups .rw-gallery')][0]?.querySelector(${JSON.stringify(sel)}); if (!b) return 'none'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const upAtStart = await js(`[...document.querySelectorAll('.rw-groups .rw-gallery')][0]?.querySelector('.rw-gallery-up')?.disabled`);
    const down = await press('.rw-gallery-down');
    const paged = await until(async () => (await first()) > 0, 'the next styles', 3000).catch(() => false);
    await snap(win, 'gallery-word-home-1300-paged.png');
    const up = await press('.rw-gallery-up');
    const backTo = await until(async () => (await first()) === 0, 'the first styles again', 3000).catch(() => false);
    const more = await press('.rw-gallery-more');
    await until(() => js(`document.querySelectorAll('.rw-menu button').length`).then((n) => n >= (mid?.items || 1)), 'every style in a menu', 3000).catch(() => false);
    const menu = await js(`document.querySelectorAll('.rw-menu button').length`);
    await snap(win, 'gallery-word-home-1300-more.png', true);
    await js(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
    await js(`(() => { document.querySelector('.wd-page')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); return 1; })()`);
    check('gallery-fold: the Styles gallery\'s arrows page through it (up greyed at the start, down to the styles after, up back) and more lists every style',
      upAtStart === true && down === 'clicked' && paged === true && up === 'clicked' && backTo === true && more === 'clicked' && menu >= (mid?.items || 1),
      JSON.stringify({ upAtStart, down, paged, up, backTo, more, menu }));

    // A style pressed in the gallery is applied to the caret's paragraph, as the old style box did; Normal puts it back.
    const style = () => js(`(async () => { const all = await window.rutbaOffice.doc.sessions({}); const mine = all.filter((s) => s.kind === 'doc').pop(); return (await window.rutbaOffice.doc.model({ id: mine.id })).format?.paragraphStyle || 'Normal'; })()`);
    const pressStyle = (id) => js(`(() => { const b = document.querySelector('.rw-ribbon .wd-styles [data-style="${id}"]'); if (!b) return 'no ' + ${JSON.stringify(id)}; b.click(); return 'clicked'; })()`);
    const h1 = await pressStyle('Heading1');
    const applied = await until(async () => (await style()) === 'Heading1', 'Heading 1 applied', 4000).catch(() => false);
    const marked = await until(() => js(`Boolean(document.querySelector('.rw-ribbon .wd-styles [data-style="Heading1"].on'))`), 'Heading 1 marked in the gallery', 3000).catch(() => false);
    await pressStyle('Normal');
    const normal = await until(async () => (await style()) === 'Normal', 'Normal again', 4000).catch(() => false);
    check('gallery-fold: a style pressed in the Styles gallery is applied to the paragraph and marked there, though scrolled out of sight, and Normal puts it back',
      h1 === 'clicked' && applied === true && marked === true && normal === true, JSON.stringify({ h1, applied, marked, normal }));
    // Saved, so the window closes clean and leaves no unsaved copy for the launcher to offer back.
    await js(`(() => { document.querySelector('.wd-page')?.focus(); return 1; })()`);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S', modifiers: ['control'] });
    await until(() => js(`!document.querySelector('.rw-doctitle .dirty')`), 'the document saved', 4000).catch(() => {});

    await sweep(win, js, 'Word', 'Home', ['Styles']);
    await pick(js, 'Design');
    await size(win, 1680);
    await snap(win, 'gallery-word-design-1680.png');
    await sweep(win, js, 'Word', 'Design', [(await js(GALLERIES))[0]?.group || 'Document Formatting']);
    await size(win, 1300);
    await snap(win, 'gallery-word-design-1300.png');
  } catch (err) {
    check('gallery-fold: the Word checks ran', false, err.stack || err.message);
  }

  // Presentations, Design: Themes and Variants.
  try {
    const win = await open('slides', pptx);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.rw-ribbon .rw-tab'))`), 'the Presentations ribbon', 8000);
    await pick(js, 'Design');
    await until(() => js(`document.querySelectorAll('.sl-rs-themes .sl-rs-item svg').length >= 2`), 'the theme pictures', 8000).catch(() => {});
    await size(win, 1680);
    await snap(win, 'gallery-slides-design-1680.png');
    await size(win, 1300);
    await snap(win, 'gallery-slides-design-1300.png');
    await size(win, 1000);
    await snap(win, 'gallery-slides-design-1000.png');
    await sweep(win, js, 'Presentations', 'Design', ['Themes', 'Variants']);
  } catch (err) {
    check('gallery-fold: the Presentations checks ran', false, err.stack || err.message);
  }
}
