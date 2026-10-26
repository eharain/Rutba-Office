// PDF: a file protected with a password. A two-page PDF made here and
// protected with AES-256 (the user password and a different owner password)
// opens on the Password dialog; a wrong password is said to be wrong and
// asked again; the right one draws the pages, with their words to find,
// and Export as Pictures writes them. A sticky note is written into it and
// a page turned, and Save As writes a file the same passwords open, the note
// and the turn in it, its words not in the file plain. Cancel leaves the
// window on the protected screen, nothing shown as an error, with Enter
// Password to try again. A file with only an owner password, whose
// permissions keep changes back, opens without asking, says so, refuses a
// change in a sentence, and allows it once its owner password is given.
// Run alone with RUTBA_VERIFY_ONLY=pdf-password.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import pdfWriter from '@rutba/pdf';
import pdfReader from '@rutba/pdf/reader';
import pdfAnnotations from '@rutba/pdf/annotations';

const require = createRequire(import.meta.url);
const { protectPdf } = require('../../../packages/pdf/test/protect-fixture.js');
const { PdfDocument } = pdfWriter;
const { openPdf } = pdfReader;
const { readAnnotations } = pdfAnnotations;

const USER = 'open sesame';
const OWNER = 'owner key 26';

function makeProtected(file, options = { kind: 'aes-256', user: USER, owner: OWNER }) {
  const doc = new PdfDocument({ title: 'Board figures', created: '2026-10-26T09:00:00Z' });
  doc.addPage().text('Quarterly figures for the board', 72, 90, { font: 'Helvetica-Bold', size: 20 });
  doc.addPage().text('The second page lists the risks.', 72, 90, { font: 'Helvetica', size: 14 });
  fs.writeFileSync(file, protectPdf(doc.toBuffer(), options));
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
  const role = async (r) => {
    const sel = `[data-role="${r}"]`;
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(sel)}) && !document.querySelector(${JSON.stringify(sel)}).disabled)`), `the ${r} button`, 5000).catch(() => {});
    return js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'no button ${r}'; if (b.disabled) return 'disabled ${r}'; b.click(); return 'clicked'; })()`);
  };
  const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
  const tip = (label) => js(`(() => {
    const b = [...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(label)}));
    if (!b) return 'no button ${label}';
    if (b.disabled) return 'disabled ${label}';
    b.click();
    return 'clicked';
  })()`);
  const setText = (sel, value) => js(`(() => {
    const el = document.querySelector(${JSON.stringify(sel)});
    if (!el) return false;
    Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(String(value))});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return true;
  })()`);
  /** A press and let go on page `index`'s comment layer, at fractions of the page. */
  const clickOn = (index, [fx, fy]) => js(`(async () => {
    const svg = document.querySelector('.pdf-page[data-index="${index}"] .pdf-annots');
    if (!svg) return 'no layer';
    const r = svg.getBoundingClientRect();
    const at = { clientX: r.left + r.width * ${fx}, clientY: r.top + r.height * ${fy}, bubbles: true, cancelable: true, button: 0, buttons: 1 };
    svg.dispatchEvent(new MouseEvent('mousedown', at));
    window.dispatchEvent(new MouseEvent('mousemove', at));
    await new Promise((res) => setTimeout(res, 20));
    window.dispatchEvent(new MouseEvent('mouseup', at));
    return 'clicked';
  })()`);
  const turns = () => js(`[...document.querySelectorAll('.pdf-thumb')].map((t) => Number(t.dataset.rotate || 0))`);
  const toasts = () => js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent)`);
  return { wc, js, capture, asked, type, setInput, role, tab, tip, setText, clickOn, toasts, turns };
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

  // A sticky note written into the protected file, a page turned, and Save As.
  await a.tab('Comment');
  await until(() => a.js(`Boolean(document.querySelector('.pdf-page[data-index="0"] .pdf-annots'))`), 'the comment layer', 8000).catch(() => {});
  await a.role('tool-note');
  await a.clickOn(0, [0.8, 0.15]);
  const editor = await until(() => a.js(`Boolean(document.querySelector('[data-role="note-editor"] textarea'))`), 'the note editor', 5000).then(() => true).catch(() => false);
  await a.setText('[data-role="note-text"]', 'Check the second quarter');
  await a.role('note-post');
  const noted = await until(() => a.js(`[...document.querySelectorAll('.pdf-comments .pdf-comment')].some((c) => /Check the second quarter/.test(c.textContent))`), 'the note listed', 12000).then(() => true).catch(() => false);
  check('pdf-password: a sticky note is written into the protected file and listed', editor && noted, JSON.stringify({ editor, noted, toasts: await a.toasts() }));
  await a.tab('Pages');
  await wait(200);
  await a.role('rotate-right');
  const turnedHere = await until(async () => (await a.turns()).includes(90), 'a page turned', 6000).then(() => true).catch(() => false);
  check('pdf-password: Rotate Right turns a page of the protected file', turnedHere, JSON.stringify({ turns: await a.turns(), toasts: await a.toasts() }));
  await a.tab('Home');
  const saved = path.join(root, 'board-saved.pdf');
  answer('save', saved);
  await a.tip('Save As');
  const wrote = await until(() => fs.existsSync(saved) && fs.statSync(saved).size > 0, 'the saved file', 15000).then(() => true).catch(() => false);
  let back = {};
  if (wrote) {
    await wait(300);
    const bytes = fs.readFileSync(saved);
    const plain = openPdf(bytes);
    const read = openPdf(bytes, { password: USER });
    back = {
      locked: plain.locked,
      access: read.access,
      owner: openPdf(bytes, { password: OWNER }).access,
      pages: read.locked ? 0 : read.pageCount,
      turned: read.locked ? [] : [0, 1].map((i) => read.pageInfo(i).rotate),
      text: read.locked ? '' : read.pageText(1),
      notes: read.locked ? [] : readAnnotations(read).map((n) => n.contents),
      plainWords: bytes.includes('Check the second quarter') || bytes.includes('second page lists'),
    };
  }
  check('pdf-password: Save As writes a file the same passwords open, with the note and the turned page in it and its words encrypted', wrote && back.locked === true
    && back.access === 'user' && back.owner === 'owner' && back.pages === 2 && back.turned.includes(90) && /second page lists the risks/.test(back.text)
    && back.notes.includes('Check the second quarter') && back.plainWords === false, JSON.stringify({ wrote, back, toasts: await a.toasts() }));

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

  // Only an owner password: permissions that keep changes, comments and page assembly back.
  const limited = path.join(root, 'limited.pdf');
  makeProtected(limited, { kind: 'rc4-128', user: '', owner: OWNER, permissions: -4 & ~(1 << 3) & ~(1 << 5) & ~(1 << 10) });
  const lim = await open('pdf', limited);
  const c = driver(lim, { until, wait });
  const shown = await until(() => c.js(`document.querySelectorAll('.pdf-thumb').length === 2 && document.querySelectorAll('.pdf-page canvas[data-drawn]').length >= 1`), 'the pages drawn', 20000).then(() => true).catch(() => false);
  const noAsk = !(await c.js(`Boolean(document.querySelector('.pw-open-password'))`));
  const notice = await c.js(`Boolean(document.querySelector('[data-role="pdf-restricted"]'))`);
  await c.tab('Pages');
  await c.role('rotate-right');
  const refused = await until(async () => (await c.toasts()).some((m) => /does not allow it to be changed/.test(m)), 'the refusal', 5000).then(() => true).catch(() => false);
  await wait(300);
  const unturned = await c.turns();
  await c.capture('pdf-password-05-limited.png');
  check('pdf-password: a file with only an owner password opens without asking, says its permissions keep changes back, and refuses a turn in a sentence', shown && noAsk && notice && refused && !unturned.includes(90),
    JSON.stringify({ shown, noAsk, notice, refused, unturned, toasts: await c.toasts() }));
  await c.tab('Comment');
  await until(() => c.js(`Boolean(document.querySelector('.pdf-page[data-index="0"] .pdf-annots'))`), 'the comment layer', 8000).catch(() => {});
  await c.role('tool-note');
  const noComment = await until(async () => (await c.toasts()).some((m) => /does not allow comments/.test(m)), 'the refusal', 6000).then(() => true).catch(() => false);
  await c.clickOn(0, [0.8, 0.15]);
  const noEditor = !(await until(() => c.js(`Boolean(document.querySelector('[data-role="note-editor"] textarea'))`), 'the note editor', 1500).then(() => true).catch(() => false));
  const listed = await c.js(`document.querySelectorAll('.pdf-comments .pdf-comment').length`);
  check('pdf-password: the Sticky Note tool is refused in a sentence, and no comment is added', noComment && noEditor && listed === 0, JSON.stringify({ noComment, noEditor, listed, toasts: await c.toasts() }));
  await c.role('pdf-owner-password');
  const ownerAsked = await c.asked();
  const askedFor = await c.js(`document.querySelector('.pw-sub')?.textContent || ''`);
  await c.type('not the owner');
  const ownerWrong = await until(() => c.js(`Boolean(document.querySelector('.pw-error'))`), 'the dialog to say it is wrong', 10000).then(() => true).catch(() => false);
  await c.type(OWNER);
  const lifted = await until(() => c.js(`!document.querySelector('.pw-open-password') && !document.querySelector('[data-role="pdf-restricted"]')`), 'the permissions lifted', 10000).then(() => true).catch(() => false);
  await c.tab('Pages');
  await c.role('rotate-right');
  await until(async () => (await c.turns()).includes(90), 'a page turned', 6000).catch(() => {});
  const turnedNow = await c.turns();
  check('pdf-password: Enter Owner Password asks for it, says when it is wrong, and the right one allows the change', ownerAsked && /owner password/.test(askedFor) && ownerWrong && lifted && turnedNow.includes(90),
    JSON.stringify({ ownerAsked, askedFor, ownerWrong, lifted, turnedNow }));
}
