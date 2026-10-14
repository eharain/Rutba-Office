// PDF: bookmarks, page labels and links, in the window and through every
// page operation that writes a file. A five-page book made here has nested
// bookmarks (one through a named destination, one a web link), page labels
// (i, ii, then A-1 onward), a link on its first page to its fourth through a
// named destination, and a form field on two pages. The thumbnails and the
// status bar show the labels; the Bookmarks pane lists the outline and a
// bookmark goes to its page; the link goes to its page; moving the last
// page to the front takes its label with it and the bookmark to it follows
// it. Then the file written by Save As, by Extract, by Split and by Merge
// Files is read back: bookmarks pointing at their pages where they now
// stand (a bookmark to a page left out giving way to its own), each page
// labelled as it was, the link made explicit, and the form's field kept,
// a merge's second field of the same name renamed.
// Run alone with RUTBA_VERIFY_ONLY=pdf-outline.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import pdfReader from '@rutba/pdf/reader';
import pdfNavigation from '@rutba/pdf/navigation';
import pdfForms from '@rutba/pdf/forms';

const require = createRequire(import.meta.url);
const { book } = require('../../../packages/pdf/test/nav-fixture.js');
const { openPdf, Ref } = pdfReader;
const { readOutline, pageLabels, labelText } = pdfNavigation;
const { readFields } = pdfForms;

/** A file's outline as "title->first words of its page" lines, indented, and its labels. */
function readBack(file) {
  const pdf = openPdf(fs.readFileSync(file));
  const pageOf = (dest) => {
    if (!dest) return 'link';
    const idx = pdf.pages().findIndex((p) => p.ref && dest[0] instanceof Ref && p.ref.num === dest[0].num);
    return idx < 0 ? '?' : pdf.pageText(idx).split('\n')[0];
  };
  const lines = [];
  const walk = (nodes, depth) => nodes.forEach((n) => { lines.push(`${'  '.repeat(depth)}${n.title}->${pageOf(n.dest)}`); walk(n.kids, depth + 1); });
  walk(readOutline(pdf), 0);
  const firstLink = (() => {
    for (let i = 0; i < pdf.pageCount; i++) {
      const list = pdf.resolve(pdf.pages()[i].dict.get('Annots')) || [];
      for (const a of list) {
        const d = pdf.resolve(a);
        if (d?.get('Subtype')?.name === 'Link') return { on: pdf.pageText(i).split('\n')[0], to: pageOf(d.get('Dest')) };
      }
    }
    return null;
  })();
  return {
    pages: Array.from({ length: pdf.pageCount }, (_, i) => pdf.pageText(i).split('\n')[0]),
    outline: lines,
    labels: (pageLabels(pdf) || []).map(labelText),
    link: firstLink,
    fields: readFields(pdf).map((f) => `${f.name}=${f.value}`),
  };
}

function driver(win, { until, wait }) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
  const role = async (r) => {
    const sel = `[data-role="${r}"]`;
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}) && !document.querySelector(${JSON.stringify(sel)}).disabled)`), `the ${r} button`, 5000).catch(() => {});
    return js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'no button ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  };
  const tip = (label) => js(`(() => {
    const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(label)}));
    if (!b) return 'no button ${label}';
    if (b.disabled) return 'disabled ${label}';
    b.click();
    return 'clicked';
  })()`);
  const labels = () => js(`[...document.querySelectorAll('.pdf-thumb .pdf-thumb-label')].map((n) => n.textContent)`);
  const status = () => js(`document.querySelector('[data-role="page-status"]')?.textContent || ''`);
  const currentPage = () => js(`Number(document.querySelector('.pdf-thumb.current')?.dataset.index ?? -1)`);
  const bookmarks = () => js(`[...document.querySelectorAll('.pdf-bookmark')].map((n) => ({ title: n.dataset.title, page: n.dataset.page, gone: n.classList.contains('gone') }))`);
  const pick = (i, { ctrl = false } = {}) => js(`(() => {
    const el = document.querySelectorAll('.pdf-thumb')[${i}];
    if (!el) return 'no thumbnail';
    const r = el.getBoundingClientRect();
    const at = { clientX: r.left + r.width / 2, clientY: r.top + r.height / 2, bubbles: true, cancelable: true, button: 0, ctrlKey: ${ctrl} };
    el.dispatchEvent(new MouseEvent('mousedown', at));
    el.dispatchEvent(new MouseEvent('mouseup', at));
    return 'picked';
  })()`);
  const drag = (from, to) => js(`(async () => {
    const all = document.querySelectorAll('.pdf-thumb');
    const a = all[${from}].getBoundingClientRect();
    const b = all[${to}].getBoundingClientRect();
    const x = a.left + a.width / 2;
    const ev = (y) => ({ clientX: x, clientY: y, bubbles: true, cancelable: true, button: 0, buttons: 1 });
    all[${from}].dispatchEvent(new MouseEvent('mousedown', ev(a.top + a.height / 2)));
    for (let i = 1; i <= 8; i++) {
      const y = a.top + a.height / 2 + ((b.top + 4) - (a.top + a.height / 2)) * (i / 8);
      window.dispatchEvent(new MouseEvent('mousemove', ev(y)));
      await new Promise((r) => setTimeout(r, 30));
    }
    window.dispatchEvent(new MouseEvent('mouseup', ev(b.top + 4)));
    return 'dragged';
  })()`);
  const setInput = (sel, value) => js(`(() => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return false;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  })()`);
  return { wc, js, capture, tab, role, tip, labels, status, currentPage, bookmarks, pick, drag, setInput };
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (a folder of this block's own is made in it)
 */
export async function verifyPdfOutline(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'pdf-outline');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'handbook.pdf');
  const other = path.join(root, 'appendix.pdf');
  fs.writeFileSync(file, book('Page'));
  fs.writeFileSync(other, book('Extra'));

  const win = await open('pdf', file);
  const d = driver(win, { until, wait });
  await until(async () => (await d.labels()).length === 5, 'five thumbnails', 20000).catch(() => {});

  // Labels on the thumbnails and in the status bar.
  const labelled = await until(async () => (await d.labels()).join() === 'i,ii,A-1,A-2,A-3', 'the labels', 8000).then(() => true).catch(() => false);
  const status = await d.status();
  check('pdf-outline: the thumbnails show each page\'s own label and the status bar names it beside the page\'s place', labelled && /Page i \(1 of 5\)/.test(status), JSON.stringify({ labels: await d.labels(), status }));

  // The Bookmarks pane, and a bookmark that goes through a named destination.
  await d.role('side-bookmarks');
  await until(async () => (await d.bookmarks()).length === 6, 'six bookmarks', 6000).catch(() => {});
  const listed = await d.bookmarks();
  await d.capture('pdf-outline-01-bookmarks.png');
  await d.js(`[...document.querySelectorAll('.pdf-bookmark')].find((n) => n.dataset.title === 'Detail')?.click(), 1`);
  const toDetail = await until(async () => /A-2/.test(await d.status()), 'Detail’s page', 6000).then(() => true).catch(() => false);
  check('pdf-outline: the Bookmarks pane lists the outline, nested, and a bookmark through a named destination goes to its page', listed.map((b) => `${b.title}:${b.page}`).join() === 'Contents:1,Intro:1,Body:3,Detail:4,End:5,Web:' && toDetail,
    JSON.stringify({ listed, toDetail, status: await d.status() }));

  // The link on the first page, through its named destination, to page 4: from the first page, in the strip.
  await d.role('side-pages');
  await until(async () => (await d.labels()).length === 5, 'the strip', 5000).catch(() => {});
  await d.pick(0);
  await until(async () => (await d.currentPage()) === 0, 'the first page', 5000).catch(() => {});
  await until(() => d.js(`Boolean(document.querySelector('.pdf-page[data-index="0"] .pdf-link'))`), 'the link', 8000).catch(() => {});
  await wait(300);
  const linkTo = await d.js(`document.querySelector('.pdf-page[data-index="0"] .pdf-link')?.dataset.target ?? null`);
  await d.js(`(document.querySelector('.pdf-page[data-index="0"] .pdf-link')?.dispatchEvent(new MouseEvent('click', { bubbles: true, button: 0 })), 1)`);
  const followed = await until(async () => (await d.currentPage()) === 3, 'the linked page', 6000).then(() => true).catch(() => false);
  check('pdf-outline: a link on a page goes to the page it names', linkTo === '3' && followed, JSON.stringify({ linkTo, followed, current: await d.currentPage() }));

  // The last page moved to the front: its label goes with it, and End's bookmark follows it.
  await d.pick(4);
  await d.drag(4, 0);
  const moved = await until(async () => (await d.labels()).join() === 'A-3,i,ii,A-1,A-2', 'the labels moved', 6000).then(() => true).catch(() => false);
  await d.role('side-bookmarks');
  const after = await until(async () => (await d.bookmarks()).find((b) => b.title === 'End')?.page === '1', 'End at the front', 6000).then(async () => d.bookmarks()).catch(async () => d.bookmarks());
  await d.capture('pdf-outline-02-moved.png');
  check('pdf-outline: a page moved takes its label with it, and the bookmark to it follows it', moved && after.find((b) => b.title === 'End')?.page === '1' && after.find((b) => b.title === 'Contents')?.page === '2',
    JSON.stringify({ labels: await d.labels(), after }));

  // Save As: read back.
  const saved = path.join(root, 'handbook-moved.pdf');
  answer('save', saved);
  await d.tip('Save As');
  const wrote = await until(() => fs.existsSync(saved), 'the saved file', 10000).then(() => true).catch(() => false);
  const s = wrote ? readBack(saved) : {};
  check('pdf-outline: Save As keeps the bookmarks on their pages where they now stand, each page\'s label, the link and the field', wrote
    && s.pages.join() === 'Page 5,Page 1,Page 2,Page 3,Page 4'
    && s.outline.join('|') === 'Contents->Page 1|  Intro->Page 1|  Body->Page 3|    Detail->Page 4|    End->Page 5|Web->link'
    && s.labels.join() === 'A-3,i,ii,A-1,A-2' && s.link?.on === 'Page 1' && s.link?.to === 'Page 4' && s.fields.join() === 'reviewer=Page reviewer', JSON.stringify(s));

  // Extract the last two pages of the list (Page 3, Page 4): Body's page goes, Detail stays and moves up.
  await d.role('side-pages');
  await until(async () => (await d.labels()).length === 5, 'the strip', 5000).catch(() => {});
  await d.pick(3);
  await d.pick(4, { ctrl: true });
  const extracted = path.join(root, 'handbook-extract.pdf');
  answer('save', extracted);
  await d.tab('Pages');
  await d.role('extract-pages');
  const ex = await until(() => fs.existsSync(extracted), 'the extract', 10000).then(() => readBack(extracted)).catch(() => null);
  check('pdf-outline: Extract keeps the bookmarks of the pages taken, one to a page left out giving way to its own, and their labels', ex
    && ex.pages.join() === 'Page 3,Page 4' && ex.outline.join('|') === 'Body->Page 3|  Detail->Page 4|Web->link' && ex.labels.join() === 'A-1,A-2' && ex.fields.join() === 'reviewer=Page reviewer',
  JSON.stringify(ex));

  // Split into files of two.
  const splitDir = path.join(root, 'split');
  fs.mkdirSync(splitDir, { recursive: true });
  await d.role('split');
  await until(() => d.js(`Boolean(document.querySelector('[data-role="split-every"]'))`), 'the Split dialog', 4000).catch(() => {});
  await d.setInput('[data-role="split-every"]', '2');
  answer('open', [splitDir]);
  await d.role('split-go');
  const parts = [1, 2, 3].map((n) => path.join(splitDir, `handbook-moved_Part${n}.pdf`));
  const split = await until(() => parts.every((p) => fs.existsSync(p)), 'three parts', 10000).then(() => parts.map(readBack)).catch(() => null);
  check('pdf-outline: Split gives each part the bookmarks and labels of its own pages', split
    && split.map((p) => p.labels.join('/')).join(' ') === 'A-3/i ii/A-1 A-2'
    && split[0].outline.join('|') === 'Contents->Page 1|  Intro->Page 1|  End->Page 5|Web->link' && split[1].outline.join('|') === 'Body->Page 3|Web->link'
    && split[2].outline.join('|') === 'Detail->Page 4|Web->link', JSON.stringify(split && split.map((p) => ({ labels: p.labels, outline: p.outline }))));

  // Merge another book at the end: its bookmarks follow, its labels kept, its field of the same name renamed.
  answer('open', [other]);
  await d.role('merge');
  await until(async () => (await d.labels()).length === 10, 'ten pages', 15000).catch(() => {});
  const merged = path.join(root, 'handbook-merged.pdf');
  await d.tab('Home');
  answer('save', merged);
  await d.tip('Save As');
  const m = await until(() => fs.existsSync(merged), 'the merged file', 10000).then(() => readBack(merged)).catch(() => null);
  check('pdf-outline: Merge Files follows one file\'s bookmarks with the other\'s, keeps both files\' labels and keeps two fields of one name apart', m
    && m.outline.length === 12 && m.outline[6] === 'Contents->Extra 1' && m.outline[9] === '    Detail->Extra 4'
    && m.labels.join() === 'A-3,i,ii,A-1,A-2,i,ii,A-1,A-2,A-3' && m.fields.join() === 'reviewer=Page reviewer,reviewer_2=Extra reviewer', JSON.stringify(m));
  await d.capture('pdf-outline-03-merged.png');
  const errors = await errorsIn(win);
  check('pdf-outline: no error was shown', !errors.length, errors.join(' | '));
}
