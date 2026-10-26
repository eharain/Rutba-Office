// PDF: a file protected with a password. A two-page PDF made here and
// protected with AES-256 (the user password and a different owner password)
// opens on the Password dialog; a wrong password is said to be wrong and
// asked again; the right one draws the pages, with their words to find,
// and Export as Pictures writes them.
// Cancel leaves the window on the protected screen, nothing shown as an
// error, with Enter Password to try again.
// Run alone with RUTBA_VERIFY_ONLY=pdf-password.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import pdfWriter from '@rutba/pdf';

const require = createRequire(import.meta.url);
const { protectPdf } = require('../../../packages/pdf/test/protect-fixture.js');
const { PdfDocument } = pdfWriter;

const USER = 'open sesame';
const OWNER = 'owner key 26';

function makeProtected(file) {
  const doc = new PdfDocument({ title: 'Board figures', created: '2026-10-26T09:00:00Z' });
  doc.addPage().text('Quarterly figures for the board', 72, 90, { font: 'Helvetica-Bold', size: 20 });
  doc.addPage().text('The second page lists the risks.', 72, 90, { font: 'Helvetica', size: 14 });
  fs.writeFileSync(file, protectPdf(doc.toBuffer(), { kind: 'aes-256', user: USER, owner: OWNER }));
}

const answer = (kind, value) => {
  (globalThis.__rutbaCheckDialogAnswers ||= []).push({ kind, answer: value });
};

function driver(win, { until, wait }) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const asked = () => until(() => js(`Boolean(document.querySelector('.pw-open-password')) && !/Opening/.test(document.querySelector('.pw-open-ok')?.textContent || '')`), 'the Password dialog', 15000).then(() => true).catch(() => false);
  // The password typed into the dialog's box and OK pressed, as a person would.
  const type = (words) => js(`(() => {
    const el = document.querySelector('.pw-open-password');
    if (!el) return 'no password box';
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(words)});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return 'typed';
  })()`).then(async (said) => {
    await wait(100);
    await js(`document.querySelector('.pw-open-ok')?.click(), 1`);
    return said;
  });
  const setInput = (sel, value) => js(`(() => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return false;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  })()`);
  return { wc, js, capture, asked, type, setInput };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made (a folder of this block's own is made in it)
 */
export async function verifyPdfPassword(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'pdf-password');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'board.pdf');
  makeProtected(file);

  // The Password dialog, a wrong password, then the right one.
  const win = await open('pdf', file);
  const a = driver(win, { until, wait });
  const first = await a.asked();
  await a.capture('pdf-password-01-asked.png');
  const nothingYet = await a.js(`document.querySelectorAll('.pdf-thumb').length`);
  check('pdf-password: an AES-256 file opens on the Password dialog, its name in it, no page shown yet', first && nothingYet === 0
    && await a.js(`(document.querySelector('.pw-file')?.textContent || '') === 'board.pdf'`), JSON.stringify({ first, nothingYet }));
  await a.type('not the password');
  const wrong = await until(() => a.js(`Boolean(document.querySelector('.pw-error')) && Boolean(document.querySelector('.pw-open-password')) && document.querySelector('.pw-open-password').value === ''`), 'the dialog to say the password is wrong', 15000).then(() => true).catch(() => false);
  await a.capture('pdf-password-02-wrong.png');
  check('pdf-password: a wrong password is said to be wrong, and the dialog asks again', wrong, JSON.stringify({ wrong, error: await a.js(`document.querySelector('.pw-error')?.textContent || ''`) }));
  await a.type(USER);
  const drawn = await until(() => a.js(`!document.querySelector('.pw-open-password') && document.querySelectorAll('.pdf-thumb').length === 2 && document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1`), 'the pages drawn', 20000).then(() => true).catch(() => false);
  const words = await until(() => a.js(`/Quarterly figures/.test(document.querySelector('.pdf-page .textLayer[data-ready="yes"]')?.textContent || '')`), 'the page\'s words', 10000).then(() => true).catch(() => false);
  await a.capture('pdf-password-03-open.png');
  check('pdf-password: the right password draws both pages, their words laid over them to select and copy', drawn && words, JSON.stringify({ drawn, words }));

  // Find reads the protected file's words.
  await a.js(`[...document.querySelectorAll('.rw-groups .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Find'))?.click(), 1`);
  await until(() => a.js(`Boolean(document.querySelector('[data-role="find-input"]'))`), 'the find box', 4000).catch(() => {});
  await a.setInput('[data-role="find-input"]', 'risks');
  const found = await until(() => a.js(`(() => { const c = document.querySelector('.pdf-find-count'); return c?.dataset.for === 'risks' && c.dataset.count === '1'; })()`), 'one match', 10000).then(() => true).catch(() => false);
  check('pdf-password: Find finds a word on the protected file\'s second page', found, JSON.stringify({ found, count: await a.js(`document.querySelector('.pdf-find-count')?.textContent || ''`) }));
  await a.js(`document.querySelector('.pdf-find .rw-btn[data-tip^="Close Find"]')?.click(), 1`);

  // Export as Pictures draws the protected pages too.
  const pictures = path.join(root, 'pictures');
  fs.mkdirSync(pictures, { recursive: true });
  await a.js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Pages')?.click(), 1`);
  await until(() => a.js(`Boolean(document.querySelector('[data-role="export-pictures"]'))`), 'the Export as Pictures button', 4000).catch(() => {});
  await a.js(`document.querySelector('[data-role="export-pictures"]')?.click(), 1`);
  await until(() => a.js(`Boolean(document.querySelector('[data-role="export-dialog"]'))`), 'the export dialog', 4000).catch(() => {});
  answer('open', [pictures]);
  await a.js(`document.querySelector('[data-role="export-go"]')?.click(), 1`);
  const png = [1, 2].map((n) => path.join(pictures, `board_Page_${n}.png`));
  const exported = await until(() => png.every((p) => fs.existsSync(p) && fs.statSync(p).size > 1000), 'the pictures', 20000).then(() => true).catch(() => false);
  check('pdf-password: Export as Pictures writes the protected file\'s pages', exported, JSON.stringify({ exported, files: fs.readdirSync(pictures) }));
  const errors = await errorsIn(win);
  check('pdf-password: no error was shown while it was opened', !errors.length, errors.join(' | '));

  // Cancel: the protected screen, no error, and Enter Password to try again.
  const other = await open('pdf', file);
  const b = driver(other, { until, wait });
  await b.asked();
  await b.js(`document.querySelector('.pw-open-cancel')?.click(), 1`);
  const closed = await until(() => b.js(`!document.querySelector('.pw-open-password') && Boolean(document.querySelector('[data-role="pdf-password-retry"]'))`), 'the protected screen', 8000).then(() => true).catch(() => false);
  const bad = await errorsIn(other);
  await b.capture('pdf-password-04-cancelled.png');
  check('pdf-password: Cancel closes the dialog onto the protected screen, nothing shown as an error, with Enter Password to try again', closed && !bad.length
    && /This PDF is protected/.test(await b.js(`document.body.textContent`)), JSON.stringify({ closed, bad }));
  await b.js(`document.querySelector('[data-role="pdf-password-retry"]')?.click(), 1`);
  const again = await b.asked();
  await b.type(OWNER);
  const byOwner = await until(() => b.js(`document.querySelectorAll('.pdf-thumb').length === 2 && document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1`), 'the pages drawn', 20000).then(() => true).catch(() => false);
  check('pdf-password: Enter Password asks again, and the owner password opens the file too', again && byOwner, JSON.stringify({ again, byOwner }));
}
