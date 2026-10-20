// The template gallery: File > New's templates on the launcher (More
// templates) and in each document app's menu (New from Template), each card
// the picture the suite drew of its template; choosing one opens a new,
// untitled copy with the template's content, the template left as it was;
// the gallery in Urdu, right to left. Run alone with
// RUTBA_VERIFY_ONLY=templates (with the Personal checks in
// verify-personal-templates.js).
//
// Also the template pictures themselves: RUTBA_VERIFY_ONLY=template-thumbs
// with RUTBA_TEMPLATE_THUMBS=<folder> opens each template in its own app and
// photographs its first page, its grid or its first slide
// (build/make-template-thumbs.js runs it).
import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { TEMPLATES, templatePath } from '@rutba/templates';
import { SheetView } from '@rutba/sheet-view';

/** until(), answering with what the condition found rather than true. */
async function got(condition, what, timeout = 8000) {
  const deadline = Date.now() + timeout;
  for (;;) {
    let held = null;
    try {
      held = await condition();
    } catch {
      held = null;
    }
    if (held) return held;
    if (Date.now() > deadline) throw new Error(`waited ${timeout} ms for ${what} and it never happened`);
    await new Promise((r) => setTimeout(r, 80));
  }
}

/** What each app draws its document in, and so what its picture is of. */
const SURFACE = {
  word: `.wd-page`,
  sheets: `.sh-cells`,
  slides: `.sl-stage svg.sl-svg, .sl-svg`,
};

/** A window of an app on a gallery template, made the way a card makes one. */
async function openTemplate({ windows, until, wait }, app, id) {
  const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
  windows.create({ app, query: { template: id } });
  const win = await got(() => BrowserWindow.getAllWindows().find((w) => !before.has(w.id)) || false, `the ${app} window for ${id}`, 8000);
  await new Promise((resolve) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', resolve) : resolve()));
  await until(() => win.webContents.executeJavaScript(`Boolean(document.querySelector(${JSON.stringify(SURFACE[app])}))`), `the ${app} surface`, 12000);
  await wait(400);
  return win;
}

/**
 * The pictures: each template opened in its app, at a size its first page,
 * its grid or its first slide fits in, photographed and scaled to the
 * gallery's size.
 */
export async function makeTemplateThumbs(h, { windows, outDir }) {
  const { open, check, until, wait } = h;
  fs.mkdirSync(outDir, { recursive: true });
  // A window of the run's own first: the last window closing would end it.
  await open('home');
  const made = [];
  for (const entry of TEMPLATES) {
    let win = null;
    try {
      win = await openTemplate({ windows, until, wait }, entry.app, entry.id);
      const js = (code) => win.webContents.executeJavaScript(code);
      // Room for a whole page or a slide at its own size.
      win.setContentSize(entry.app === 'word' ? 1180 : 1500, 1000);
      // A page is taller than a window the screen allows: the page drawn smaller.
      const zoom = entry.app === 'word' ? 0.6 : 1;
      win.webContents.setZoomFactor(zoom);
      await wait(900);
      if (entry.app === 'word') await until(() => js(`(document.querySelectorAll('.wd-page').length > 0) && document.fonts.status === 'loaded'`), 'the page', 8000);
      if (entry.app === 'sheets') await until(() => js(`document.querySelectorAll('.sh-cell').length > 6`), 'the cells', 8000);
      await wait(entry.app === 'sheets' ? 1200 : 700);
      const rect = await js(`(() => {
        const n = document.querySelector(${JSON.stringify(entry.app === 'word' ? '.wd-sheet' : SURFACE[entry.app])});
        const r = n.getBoundingClientRect();
        const app = ${JSON.stringify(entry.app)};
        if (app === 'sheets') return { x: r.left, y: r.top, width: Math.min(r.width, 1040), height: Math.min(r.height, 640) };
        return { x: r.left, y: r.top, width: r.width, height: r.height };
      })()`).then((r) => ({ x: Math.round(r.x * zoom), y: Math.round(r.y * zoom), width: Math.round(r.width * zoom), height: Math.round(r.height * zoom) }));
      // A borderless table's dashed guides are the screen's, not the page's.
      if (entry.app === 'word') await js();
      // A borderless table's dashed guides are the screen's, not the page's.
      if (entry.app === 'word') await js(`document.querySelector('.wd-page')?.classList.add('wd-no-table-gridlines'), 1`);
      win.webContents.invalidate();
      await wait(700);
      const shot = await win.webContents.capturePage(rect);
      const width = entry.app === 'word' ? 300 : 420;
      const scaled = shot.resize({ width, quality: 'best' });
      fs.writeFileSync(path.join(outDir, entry.thumb), scaled.toPNG());
      made.push(`${entry.id} ${scaled.getSize().width}x${scaled.getSize().height}`);
    } catch (err) {
      check(`template-thumbs: ${entry.id} was pictured`, false, err?.message || String(err));
    } finally {
      if (win && !win.isDestroyed()) windows.forceClose ? windows.forceClose(win) : win.destroy();
      await wait(300);
    }
  }
  check('template-thumbs: every template was pictured by the app it opens in', made.length === TEMPLATES.length, made.join(', '));
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc
 * @param {{ dir: string, windows: object }} args
 */
export async function verifyTemplates(h, { dir, windows }) {
  const { open, check, until, wait, errorsIn, doc } = h;
  const capture = async (win, name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
  };
  // The windows the gallery's cards open, which the harness does not know of.
  const spawned = [];
  const stamps = Object.fromEntries(TEMPLATES.map((t) => [t.id, fs.statSync(templatePath(t)).mtimeMs]));
  let home = null;
  try {
    /* -- the launcher: Start something > More templates ------------------- */
    home = await open('home');
    const hjs = (code) => home.webContents.executeJavaScript(code);
    await until(() => hjs(`Boolean(document.querySelector('.home-more'))`), 'the More templates tile', 6000);
    await hjs(`document.querySelector('.home-more').click(), 1`);
    const shown = await got(() => hjs(`(() => {
      const cards = [...document.querySelectorAll('.tg-card[data-template]')];
      const pictured = cards.filter((c) => { const i = c.querySelector('img'); return i && i.complete && i.naturalWidth > 100; });
      return cards.length >= ${TEMPLATES.length} && pictured.length >= ${TEMPLATES.length} ? { cards: cards.length, pictured: pictured.length } : false;
    })()`), 'the gallery with its pictures', 8000).catch(() => null);
    const rail = await hjs(`[...document.querySelectorAll('.tg-rail-item')].map((b) => b.dataset.view)`);
    check('templates: Start something > More templates opens the gallery, every template pictured by the suite, by app and by category',
      Boolean(shown) && ['word', 'sheets', 'slides', 'business', 'finance', 'personal'].every((v) => rail.includes(v)),
      JSON.stringify({ shown, rail }));
    await capture(home, 'templates-gallery.png');

    // A search narrows it; the rail's Presentations shows the seven decks.
    const setSearch = (value) => hjs(`(() => { const i = document.querySelector('.tg-head .rw-search input'); const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(i, ${JSON.stringify(value)}); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await setSearch('budget');
    const found = await got(() => hjs(`(() => { const c = [...document.querySelectorAll('.tg-card[data-template]')].map((n) => n.dataset.template); return c.length === 1 ? c : false; })()`), 'one card for budget', 4000).catch(() => null);
    await setSearch('');
    await hjs(`document.querySelector('.tg-rail-item[data-view="slides"]').click(), 1`);
    const decks = await got(() => hjs(`(() => { const c = [...document.querySelectorAll('.tg-card[data-template]')].map((n) => n.dataset.template); return c.length === 7 && c.every((x) => x.startsWith('deck-')) ? c : false; })()`), 'the seven decks', 4000).catch(() => null);
    check('templates: a search finds the budget alone, and Presentations lists the seven decks', JSON.stringify(found) === '["sheet-budget"]' && Boolean(decks), JSON.stringify({ found, decks }));
    await hjs(`document.querySelector('.tg-rail-item[data-view="all"]').click(), 1`);

    /* -- choosing one: an untitled copy with the template's content ------- */
    const choose = async (id) => {
      const before = new Set(doc.sessions().map((s) => s.id));
      const winsBefore = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
      await until(() => hjs(`Boolean(document.querySelector('.tg-card[data-template="${id}"]'))`), `the ${id} card`, 4000);
      await hjs(`document.querySelector('.tg-card[data-template="${id}"]').click(), 1`);
      const session = await got(() => doc.sessions().find((s) => !before.has(s.id)) || false, `the ${id} copy`, 10000).catch(() => null);
      const win = await got(() => BrowserWindow.getAllWindows().find((w) => !winsBefore.has(w.id)) || false, `the ${id} window`, 8000).catch(() => null);
      if (win) spawned.push(win);
      return { session, win };
    };
    const reopenGallery = async () => {
      if (!(await hjs(`Boolean(document.querySelector('.tg-card'))`))) await hjs(`document.querySelector('.home-more').click(), 1`);
      await until(() => hjs(`Boolean(document.querySelector('.tg-card[data-template]'))`), 'the gallery again', 5000);
    };

    // A workbook: the budget, its total following a figure typed in.
    const book = await choose('sheet-budget');
    let recalculated = null;
    if (book.session) {
      const file = path.join(dir, 'from-template.xlsx');
      await doc.apply({ id: book.session.id, ops: [{ op: 'setCell', row: 4, col: 2, value: '3400' }] });
      await doc.save({ id: book.session.id, path: file });
      const v = SheetView.open(fs.readFileSync(file));
      recalculated = { C7: v.displayValue(6, 2).text, C22: v.displayValue(21, 2).text, charts: (v.drawings.get(v.activeSheet) || []).filter((d) => d.kind === 'chart').length };
    }
    if (book.win) await until(() => book.win.webContents.executeJavaScript(`document.querySelectorAll('.sh-cell').length > 6`), 'the budget grid', 8000).catch(() => {});
    check('templates: choosing Monthly budget opens a new, untitled workbook from it whose totals recalculate (Salary 3,400 makes the income £3,650)',
      Boolean(book.session) && !book.session.path && Boolean(book.win) && recalculated?.C7 === '£3,650' && recalculated?.C22 === '£974' && recalculated?.charts === 2,
      JSON.stringify({ session: book.session ? { name: book.session.name, path: book.session.path } : null, recalculated }));
    if (book.win) await capture(book.win, 'templates-budget.png');

    // A document: the report, its headings and its table of contents.
    await reopenGallery();
    const report = await choose('word-report');
    const model = report.session ? doc.model({ id: report.session.id }) : null;
    const headings = (model?.blocks || []).filter((b) => /^Heading[12]$/.test(b.style || '')).map((b) => b.text);
    const toc = model?.tableOfContents?.entries?.length || 0;
    if (report.win) await until(() => report.win.webContents.executeJavaScript(`document.querySelectorAll('.wd-sheet').length >= 3`), 'the report pages', 10000).catch(() => {});
    const pages = report.win ? await report.win.webContents.executeJavaScript(`document.querySelectorAll('.wd-sheet').length`) : 0;
    check('templates: choosing Report opens a new, untitled document with its cover, its headings and a table of contents listing them',
      Boolean(report.session) && !report.session.path && headings.length === 8 && headings[0] === 'Introduction' && toc === 8 && pages >= 3,
      JSON.stringify({ name: report.session?.name, path: report.session?.path, headings: headings.length, toc, pages }));
    if (report.win) await capture(report.win, 'templates-report.png');

    // A deck: the pitch, its eight slides on the theme's layouts.
    await reopenGallery();
    const pitch = await choose('deck-pitch');
    const deckModel = pitch.session ? doc.model({ id: pitch.session.id }) : null;
    if (pitch.win) await until(() => pitch.win.webContents.executeJavaScript(`document.querySelectorAll('.sl-thumb').length >= 8`), 'the pitch slides', 8000).catch(() => {});
    const thumbs = pitch.win ? await pitch.win.webContents.executeJavaScript(`document.querySelectorAll('.sl-thumb').length`) : 0;
    check('templates: choosing Pitch deck opens a new, untitled presentation of eight slides',
      Boolean(pitch.session) && !pitch.session.path && (deckModel?.outline || []).length === 8 && thumbs === 8,
      JSON.stringify({ name: pitch.session?.name, slides: (deckModel?.outline || []).length, thumbs }));
    if (pitch.win) await capture(pitch.win, 'templates-pitch.png');
    const untouched = TEMPLATES.every((t) => fs.statSync(templatePath(t)).mtimeMs === stamps[t.id]);
    check('templates: the templates themselves are left as they were', untouched, untouched ? 'unchanged' : 'a template file changed');

    /* -- an app's own File > New ---------------------------------------- */
    const sheets = book.win;
    if (sheets) {
      const sjs = (code) => sheets.webContents.executeJavaScript(code);
      await sjs(`document.querySelector('.rw-appmark').click(), 1`);
      await until(() => sjs(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'New from Template…')`), 'New from Template in the menu', 4000).catch(() => {});
      await sjs(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'New from Template…')?.click(), 1`);
      const own = await got(() => sjs(`(() => { const c = [...document.querySelectorAll('.tg-card[data-template]')].map((n) => n.dataset.template); return c.length === 11 ? c : false; })()`), 'the Worksheets gallery', 6000).catch(() => null);
      check('templates: Worksheets\' File > New from Template shows a blank workbook first, then its ten templates, and no other app\'s',
        Boolean(own) && own[0] === 'blank' && own.slice(1).every((x) => x.startsWith('sheet-')),
        JSON.stringify(own));
      await capture(sheets, 'templates-sheets-new.png');
      await sjs(`document.querySelector('.tg-close')?.click(), 1`);
    }
    const complaints = [...await errorsIn(home), ...(book.win ? await errorsIn(book.win) : []), ...(report.win ? await errorsIn(report.win) : []), ...(pitch.win ? await errorsIn(pitch.win) : [])];
    check('templates: the gallery and the windows it opened report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    // The windows the cards opened are not the harness's: closed here.
    for (const w of [book.win, report.win, pitch.win]) if (w && !w.isDestroyed()) w.destroy();

    /* -- in Urdu, right to left ------------------------------------------ */
    await hjs(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    const urdu = await open('home');
    try {
      const ujs = (code) => urdu.webContents.executeJavaScript(code);
      await until(() => ujs(`Boolean(document.querySelector('.home-more'))`), 'the Urdu launcher', 6000);
      await ujs(`document.querySelector('.home-more').click(), 1`);
      await until(() => ujs(`document.querySelectorAll('.tg-card[data-template]').length >= ${TEMPLATES.length}`), 'the Urdu gallery', 8000).catch(() => {});
      const look = await ujs(`(() => {
        const rail = document.querySelector('.tg-rail').getBoundingClientRect();
        const grid = document.querySelector('.tg-grid').getBoundingClientRect();
        return { dir: document.documentElement.dir, lang: document.documentElement.lang, railRight: rail.left > grid.left,
          all: document.querySelector('.tg-rail-item[data-view="all"]')?.textContent.trim(), name: document.querySelector('.tg-card[data-template="word-letter"] .tg-name')?.textContent.trim(),
          head: document.querySelector('.tg-head h2')?.textContent.trim() };
      })()`);
      check('templates: in Urdu the gallery runs right to left, its rail on the right, its words and the templates\' names in Urdu',
        look.dir === 'rtl' && look.lang === 'ur' && look.railRight && /[؀-ۿ]/.test(look.all || '') && /[؀-ۿ]/.test(look.name || '') && /[؀-ۿ]/.test(look.head || ''),
        JSON.stringify(look));
      await capture(urdu, 'templates-gallery-ur.png');
    } finally {
      await hjs(`window.rutbaOffice.store.set({ key: 'language', value: null }).then(() => 1)`);
    }
  } catch (err) {
    check('templates: the gallery checks ran', false, err?.message || String(err));
  } finally {
    for (const w of spawned) if (!w.isDestroyed()) w.destroy();
  }
}
