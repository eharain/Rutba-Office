// No greyed "not built yet" button is left in Documents, Worksheets or
// Presentations. Every tab of the three is walked: no ribbon control says
// it is not built, and none of the buttons for somebody's online service
// (Translate, Dictate, Forms, Presenter Coach, Online Videos, Automate a
// Task, From Web, Keep Slides Updated, live subtitles), nor Macros or Ink
// to Math, is drawn. The two that work on this computer are driven:
// Review → Mark All as Read on a deck with other people's comments (unread
// marks on the cards and the markers, a thread opened is read, the button
// reads the rest, and what was read is still known when the deck is opened
// again, a comment added meanwhile unread), and Design → Slide Size
// (Standard asks Maximize or Ensure Fit, Widescreen goes straight back, the
// dialog turns the slides to portrait, and the file keeps the size).
// Its files are written in a folder of its own. Run alone with
// RUTBA_VERIFY_ONLY=soon.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { Deck, buildPptx } from '@rutba/presentation';
import { consoleMessage } from './console-message.js';

/** The buttons that are gone, by the words their label or tip begins with. */
const GONE = ['Translate', 'Dictate', 'Forms', 'Rehearse with Coach', 'Online Videos', 'Automate a Task', 'From Web', 'Keep Slides Updated', 'Always Use Subtitles', 'Subtitle Settings', 'Macros', 'Ink to Math', 'Ink to Maths'];

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 * @param {{ dir: string }} args the run's folder; the block writes in its own folder inside it
 */
export async function verifySoon(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const own = path.join(dir, 'soon-fixtures');
  fs.mkdirSync(own, { recursive: true });
  const at = (n) => path.join(own, n);
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const capture = async (win, name) => {
    if (!shots) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(shots, name), (await win.webContents.capturePage()).toPNG());
  };

  // 1. Every tab of the three document apps: nothing greyed as not built, and none of the gone buttons.
  fs.writeFileSync(at('soon.docx'), buildDocx({ styles: true, paragraphs: [{ text: 'Nothing greyed out.' }] }));
  fs.writeFileSync(at('soon.xlsx'), buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Month', 'Sales'], ['January', 120]] }] }));
  fs.writeFileSync(at('soon.pptx'), buildPptx({ title: 'Soon', slides: [{ layout: 'title', title: 'Nothing greyed out' }] }));
  for (const [app, file] of [['word', 'soon.docx'], ['sheets', 'soon.xlsx'], ['slides', 'soon.pptx']]) {
    try {
      const win = await open(app, at(file));
      win.setContentSize(1680, 900);
      const js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`Boolean(document.querySelector('.rw-ribbon .rw-tab'))`), `the ${app} ribbon`, 8000);
      const names = await js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
      const found = { notBuilt: [], gone: [], tabs: 0 };
      for (let i = 0; i < names.length; i++) {
        await js(`(() => { document.querySelectorAll('.rw-tab')[${i}]?.click(); return 1; })()`);
        const ready = await until(() => js(`document.querySelectorAll('.rw-tab')[${i}]?.getAttribute('aria-selected') === 'true' && document.querySelectorAll('.rw-groups > .rw-group').length > 0`), `${app} ${names[i]}`, 4000).catch(() => false);
        if (ready === true) found.tabs++;
        const seen = await js(`(() => {
          const gone = ${JSON.stringify(GONE)};
          const words = (b) => [b.dataset.tip || b.title || '', b.querySelector('.rw-btn-label, span')?.textContent || b.textContent || ''].map((s) => s.trim());
          const all = [...document.querySelectorAll('.rw-ribbon .rw-groups button')];
          return {
            notBuilt: all.filter((b) => /not built yet/i.test(b.dataset.tip || b.title || '')).map((b) => (b.dataset.tip || b.title).slice(0, 60)),
            gone: all.filter((b) => words(b).some((w) => gone.some((g) => w === g || w.startsWith(g + ':') || w.startsWith(g + ' \u2014')))).map((b) => words(b)[1] || words(b)[0]),
          };
        })()`);
        for (const n of seen.notBuilt) found.notBuilt.push(`${names[i]}: ${n}`);
        for (const n of seen.gone) found.gone.push(`${names[i]}: ${n}`);
      }
      check(`soon: ${app} draws every tab with no control greyed as not built yet, and none of the online-only buttons, Macros or Ink to Math`,
        found.tabs === names.length && found.notBuilt.length === 0 && found.gone.length === 0,
        `${found.tabs}/${names.length} tabs; not built ${JSON.stringify(found.notBuilt.slice(0, 4))}; gone but drawn ${JSON.stringify(found.gone.slice(0, 4))}`);
    } catch (err) {
      check(`soon: the ${app} tabs were walked`, false, err.stack || err.message);
    }
  }

  // 2. Review → Mark All as Read, on a deck with other people's comments.
  try {
    const file = at('read.pptx');
    const made = Deck.open(buildPptx({ title: 'Read checks', slides: [{ layout: 'title', title: 'Budget' }, { layout: 'obj', title: 'Spending', body: ['Travel'] }] }));
    const first = made.addComment(0, { text: 'Is this the final figure?', author: 'Sara Khan' });
    const second = made.addComment(1, { text: 'Travel looks high.', author: 'Omar Malik' });
    made.replyComment(1, second, { text: 'It includes the conference.', author: 'Sara Khan' });
    fs.writeFileSync(file, made.save());
    let win = await open('slides', file);
    let wc = win.webContents;
    let js = (code) => wc.executeJavaScript(code);
    const errors = [];
    const listen = (w) => w.webContents.on('console-message', (...args) => { const m = consoleMessage(args); if (m.level >= 3) errors.push(m.text.split('\n')[0].slice(0, 160)); });
    listen(win);
    const tab = (label) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const markTip = () => js(`(() => { const b = document.querySelector('.rw-ribbon .sl-mark-read'); return b ? { tip: b.dataset.tip || b.title || '', disabled: b.disabled } : null; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    await tab('Review');
    const two = await until(async () => { const m = await markTip(); return Boolean(m && !m.disabled && /^Mark All as Read: 2 comment threads/.test(m.tip)); }, 'Mark All as Read counting two unread threads', 6000).catch(() => false);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Show Comments')); b?.click(); return 1; })()`);
    const marked = await until(() => js(`Boolean(document.querySelector('.sl-cm-card.unread[data-comment="${first}"] .sl-cm-dot') && document.querySelector('.sl-cm-marker.unread[data-comment="${first}"]'))`), 'the unread card and marker', 5000).catch(() => false);
    await capture(win, 'soon-unread.png');
    check('soon: Presentations marks other people\'s comment threads unread (a stripe and a dot on the card, a dot on the marker), and Mark All as Read counts them',
      two === true && marked === true, `counted ${two}; marked ${marked}; ${JSON.stringify(await markTip())}`);

    // Opening a thread reads it.
    await js(`(() => { document.querySelector('.sl-cm-card[data-comment="${first}"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); return 1; })()`);
    const opened = await until(() => js(`!document.querySelector('.sl-cm-card.unread[data-comment="${first}"]') && !document.querySelector('.sl-cm-marker.unread')`), 'the opened thread read', 5000).catch(() => false);
    const one = await until(async () => /^Mark All as Read: 1 comment thread has/.test((await markTip())?.tip || ''), 'one thread left unread', 4000).catch(() => false);
    check('soon: a thread opened in the pane is read, its marks go, and the button counts the one left', opened === true && one === true, `opened ${opened}; one ${one}; ${JSON.stringify(await markTip())}`);

    // Mark All as Read reads the rest, across the slides.
    await js(`(() => { document.querySelector('.rw-ribbon .sl-mark-read')?.click(); return 1; })()`);
    const none = await until(async () => { const m = await markTip(); return Boolean(m?.disabled && /has been read on this computer/.test(m.tip)); }, 'nothing unread', 4000).catch(() => false);
    const toast = await until(() => js(`[...document.querySelectorAll('.rw-toast')].some((t) => /1 comment thread marked as read/.test(t.textContent))`), 'the toast', 3000).catch(() => false);
    const stored = await js(`(async () => { const all = await window.rutbaOffice.store.get({ key: 'slides.commentsRead', fallback: {} }); const e = Object.entries(all || {}).find(([k]) => k.endsWith('read.pptx')); return e ? e[1].ids.length : -1; })()`).catch((e) => String(e.message));
    check('soon: Mark All as Read reads every thread in the deck, greys itself with a reason, and keeps the three ids on this computer', none === true && toast === true && stored === 3, `greyed ${none}; toast ${toast}; stored ${stored}`);
    await capture(win, 'soon-read.png');

    // Opened again with a comment added meanwhile: only the new one is unread.
    win.close();
    await until(() => win.isDestroyed(), 'the deck closed', 5000).catch(() => {});
    const later = Deck.open(fs.readFileSync(file));
    later.addComment(1, { text: 'Can we see last year too?', author: 'Lena Weiss' });
    fs.writeFileSync(file, later.save());
    win = await open('slides', file);
    wc = win.webContents;
    js = (code) => wc.executeJavaScript(code);
    listen(win);
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides again', 8000);
    await tab('Review');
    const fresh = await until(async () => /^Mark All as Read: 1 comment thread has/.test((await markTip())?.tip || ''), 'the one new thread unread', 6000).catch(() => false);
    const strip = await js(`document.querySelectorAll('.sl-thumb .sl-thumb-cm').length`);
    check('soon: opened again, what was read stays read, and the comment added meanwhile is the one unread', fresh === true && strip === 2, `fresh ${fresh}; ${JSON.stringify(await markTip())}; slides with comments ${strip}`);
    check('soon: the read checks raised no errors in the window', errors.length === 0, errors.slice(0, 3).join(' | '));
  } catch (err) {
    check('soon: the Mark All as Read checks ran', false, err.stack || err.message);
  }

  // 3. Design → Slide Size.
  try {
    const file = at('size.pptx');
    fs.writeFileSync(file, buildPptx({ title: 'Size', slides: [{ layout: 'title', title: 'Slide Size', body: 'Scaled to fit' }, { layout: 'obj', title: 'Second', body: ['One', 'Two'] }] }));
    const win = await open('slides', file);
    win.setContentSize(1300, 760);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
    const tab = (label) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const menuItem = async (label) => {
      await js(`(() => { const b = document.querySelector('.rw-ribbon .sl-slide-size'); if (!b) return 'none'; b.click(); return 'clicked'; })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `the "${label}" item`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'missing'; b.click(); return 'clicked'; })()`);
    };
    const sizeIs = (cx, cy) => until(() => { const s = model().size; return s.cx === cx && s.cy === cy; }, `the slides at ${cx} by ${cy}`, 6000).catch(() => false);
    const inside = () => model().slide.shapes.every((s) => !s.geometry || (s.geometry.x >= -0.5 && s.geometry.y >= -0.5 && s.geometry.x + s.geometry.w <= model().size.width + 0.5 && s.geometry.y + s.geometry.h <= model().size.height + 0.5));
    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slides', 8000);
    await tab('Design');
    const live = await until(() => js(`(() => { const b = document.querySelector('.rw-ribbon .sl-slide-size'); return Boolean(b && !b.disabled && /^Slide Size: the slides are 33\\.87 by 19\\.05 cm, 16:9 widescreen/.test(b.dataset.tip || '')); })()`), 'Slide Size live', 4000).catch(() => false);
    const standard = await menuItem('Standard (4:3)');
    const asked = await until(() => js(`Boolean(document.querySelector('.sl-size-fit') && document.querySelector('.sl-size-maximize'))`), 'Maximize or Ensure Fit', 4000).catch(() => false);
    await capture(win, 'soon-size-ask.png');
    await js(`(() => { document.querySelector('.sl-size-fit')?.click(); return 1; })()`);
    const four = await sizeIs(9144000, 6858000);
    const stage = await until(() => js(`(() => { const s = document.querySelector('.sl-stage .sl-slide'); const r = s?.getBoundingClientRect(); return r && r.height > 0 ? Math.abs(r.width / r.height - 4 / 3) < 0.04 : false; })()`), 'the stage at 4:3', 4000).catch(() => false);
    check('soon: Design → Slide Size → Standard (4:3) asks Maximize or Ensure Fit, and Ensure Fit makes the slides 4:3 with everything on them', live === true && standard === 'clicked' && asked === true && four === true && inside(),
      `live ${live}; ${standard}; asked ${asked}; size ${JSON.stringify(model().size)}; stage 4:3 ${stage}; inside ${inside()}`);
    await capture(win, 'soon-size-43.png');

    const wide = await menuItem('Widescreen (16:9)');
    const back = await sizeIs(12192000, 6858000);
    check('soon: Widescreen (16:9) from 4:3 needs no question and makes the slides wide again', wide === 'clicked' && back === true && inside(), `${wide}; ${JSON.stringify(model().size)}`);

    const custom = await menuItem('Custom Slide Size…');
    const dialog = await until(() => js(`Boolean(document.querySelector('.sl-size-preset') && document.querySelector('.sl-size-w')?.value === '33.87' && document.querySelector('.sl-size-h')?.value === '19.05')`), 'the dialog with the size now', 4000).catch(() => false);
    await js(`(() => { document.querySelector('.sl-size-portrait')?.click(); return 1; })()`);
    const turned = await until(() => js(`document.querySelector('.sl-size-w')?.value === '19.05' && document.querySelector('.sl-size-h')?.value === '33.87' && Boolean(document.querySelector('.sl-size-fit'))`), 'portrait, with the question', 4000).catch(() => false);
    await capture(win, 'soon-size-dialog.png');
    await js(`(() => { document.querySelector('.sl-size-fit')?.click(); return 1; })()`);
    const tall = await sizeIs(6858000, 12192000);
    check('soon: Custom Slide Size opens on the size now, Portrait turns it, and Ensure Fit makes the slides portrait', custom === 'clicked' && dialog === true && turned === true && tall === true && inside(), `${custom}; dialog ${dialog}; turned ${turned}; ${JSON.stringify(model().size)}`);
    await capture(win, 'soon-size-portrait.png');

    await js(`(() => { document.querySelector('.sl-stage')?.focus(); return 1; })()`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'S', modifiers: ['control'] });
    const saved = await until(() => { try { const d = Deck.open(fs.readFileSync(file)); return d.size.cx === 6858000 && d.size.cy === 12192000; } catch { return false; } }, 'the saved size', 8000).catch(() => false);
    check('soon: saved, the file carries the portrait size', saved === true, `saved ${saved}`);
  } catch (err) {
    check('soon: the Slide Size checks ran', false, err.stack || err.message);
  }
}
