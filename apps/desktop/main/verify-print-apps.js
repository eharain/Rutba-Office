// File → Print in Mail, Calendar, Contacts, Image and Pictures (2.27.0): in
// each, Ctrl+P opens the shared print dialog, the dialog counts the sheets,
// and Save as PDF writes those sheets with what the window showed on them.
// The pages themselves are unit-tested (tests/print-pages.test.js); this is
// the window: the command reaching the dialog, the page laid out by the
// hidden window, and the file. Run alone with RUTBA_VERIFY_ONLY=printapps.
import fs from 'node:fs';
import path from 'node:path';
import pdfReader from '@rutba/pdf/reader';
import pdfRedact from '@rutba/pdf/redact';
import { labelKeyed } from './verify-keys.js';
import { noisyPicture } from './verify-image-edits.js';

const { openPdf } = pdfReader;
const { pageGlyphs } = pdfRedact;

const pad = (n) => String(n).padStart(2, '0');
const icsDay = (d) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyPrintApps({ open, check, until, wait, press, errorsIn }, { dir }) {
  const root = path.join(dir, 'print-apps');
  fs.mkdirSync(root, { recursive: true });
  const stamp = Date.now();

  const save = `((list) => list.find((n) => ${labelKeyed('Save as PDF…')}(n)) || list.find((n) => n.textContent.trim() === 'Save as PDF…'))([...document.querySelectorAll('.rw-dialog button')])`;

  /**
   * Ctrl+P in a window, the count the dialog gives read, and its sheets saved
   * as a PDF: { opened, said, wrote, pages, text, pictures } with the PDF read back.
   */
  const printed = async (win, name) => {
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const file = path.join(root, `${name}.pdf`);
    fs.rmSync(file, { force: true });
    const out = { opened: false, said: null, wrote: false, pages: 0, text: '', pictures: 0 };
    win.focus();
    wc.focus();
    await press(wc, 'p', { modifiers: ['control'] });
    out.opened = await until(() => js(`Boolean(${save})`), `${name}: the print dialog`, 8000).then(() => true).catch(() => false);
    if (!out.opened) return out;
    // The count is the hidden window's own: it comes a moment after the dialog.
    await until(() => js(`(() => { const b = ${save}; const c = [...document.querySelectorAll('.rw-dialog .chip')].pop(); return Boolean(b && !b.disabled && c && /\\d/.test(c.textContent)); })()`), `${name}: the sheets counted`, 20000).catch(() => {});
    const chip = await js(`[...document.querySelectorAll('.rw-dialog .chip')].pop()?.textContent || ''`);
    out.said = /\d+/.test(chip) ? Number(chip.match(/\d+/)[0]) : null;
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: file });
    await js(`(() => { const b = ${save}; b?.click(); return Boolean(b); })()`);
    out.wrote = await until(() => fs.existsSync(file) && fs.statSync(file).size > 500, `${name}: the PDF`, 30000).then(() => true).catch(() => false);
    await until(() => js(`!(${save})`), `${name}: the dialog to close`, 6000).catch(() => {});
    if (!out.wrote) {
      // Left open, the dialog would take the next check's keys.
      await press(wc, 'Escape');
      return out;
    }
    await wait(300);
    const bytes = fs.readFileSync(file);
    const pdf = openPdf(bytes);
    out.pages = pdf.pageCount;
    // The words as the page's fonts say them (the printer writes glyph numbers, each font with its own map back to text).
    out.text = Array.from({ length: pdf.pageCount }, (_, i) => pageGlyphs(bytes, i).text).join('\n');
    out.pictures = pdf.pages().filter((page) => {
      const res = pdf.resolve(page.dict.get('Resources') ?? page.inherited?.Resources);
      const xo = res?.get ? pdf.resolve(res.get('XObject')) : null;
      return Boolean(xo?.size);
    }).length;
    return out;
  };
  // The printer sets each letter in its place, and a gap between two reads back as a space: words are
  // compared letter by letter, spaces left out on both sides.
  const has = (r, words) => r.text.replace(/\s+/g, '').includes(String(words).replace(/\s+/g, ''));
  const brief = (r) => JSON.stringify({ opened: r.opened, said: r.said, wrote: r.wrote, pages: r.pages, pictures: r.pictures, text: r.text.replace(/\s+/g, ' ').slice(0, 160) });
  const quiet = async (app, win) => {
    const complaints = await errorsIn(win);
    check(`print-apps: ${app} reports nothing after its print`, complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  };

  // Mail: the open message.
  try {
    const subject = `Print check ${stamp}`;
    const raw = [
      'From: Amina Yusuf <amina@work.example>', 'To: You <you@example.com>', `Subject: ${subject}`,
      `Date: ${new Date(stamp + 2 * 86_400_000).toUTCString()}`,
      `Message-ID: <print-apps-${stamp}@example.org>`, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '',
      'The quarterly figures are attached to the earlier message.', 'Kestrel is the word this check looks for.', '',
    ].join('\r\n');
    const win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    const accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);
    await mail('deliverTest', { accountId, folder: 'Inbox', raw });
    const folders = await mail('folders', { accountId });
    const inbox = folders.find((f) => /^inbox$/i.test(f.name || f.path))?.path || 'Inbox';
    const openInbox = () => js(`(() => { const f = document.querySelector('.ml-drop[data-folder="${inbox}"] .rw-item, .ml-drop[data-folder="${inbox}"] button') || [...document.querySelectorAll('[title="${inbox}"]')][0]; f?.click(); return Boolean(f); })()`);
    const row = `[...document.querySelectorAll('.ml-list > .ml-row')].find((r) => r.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(subject)}))`;
    await openInbox();
    await until(async () => (await js(`Boolean(${row})`)) || (await openInbox(), await wait(400), false), 'the message in the list', 12000).catch(() => {});
    await js(`${row}?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-reader'))`), 'the message open', 8000).catch(() => {});
    await wait(600);
    const r = await printed(win, 'mail');
    check('print-apps: Mail prints the open message: the dialog counts its sheets and Save as PDF writes them with the subject and the words',
      r.opened && r.wrote && r.pages >= 1 && r.said === r.pages && has(r, subject) && has(r, 'Kestrel is the word this check looks for.') && has(r, 'Amina Yusuf'), brief(r));
    await quiet('Mail', win);
  } catch (err) {
    check('print-apps: the Mail print check ran', false, err?.message || String(err));
  }

  // Calendar: the view shown, a month with three events in it.
  try {
    const now = new Date();
    const day = (n) => new Date(now.getFullYear(), now.getMonth(), n);
    const events = [[day(4), 'Heron budget review'], [day(12), 'Osprey supplier call'], [day(20), 'Plover team lunch']];
    const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba Office//print check//EN'];
    for (const [i, [d, title]] of events.entries()) {
      lines.push('BEGIN:VEVENT', `UID:print-apps-${stamp}-${i}@example.org`, `DTSTAMP:${icsDay(now)}T000000Z`, `DTSTART:${icsDay(d)}T100000`, `DTEND:${icsDay(d)}T110000`, `SUMMARY:${title}`, 'END:VEVENT');
    }
    lines.push('END:VCALENDAR');
    const file = path.join(root, 'print-check.ics');
    fs.writeFileSync(file, lines.join('\r\n'));
    const win = await open('calendar', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.body.textContent.includes('Heron budget review')`), 'the events on the grid', 12000).catch(() => {});
    const r = await printed(win, 'calendar');
    check('print-apps: Calendar prints the view shown: the dialog counts its sheets and Save as PDF writes them with every event',
      r.opened && r.wrote && r.pages >= 1 && r.said === r.pages && events.every(([, title]) => has(r, title)), brief(r));
    await quiet('Calendar', win);
  } catch (err) {
    check('print-apps: the Calendar print check ran', false, err?.message || String(err));
  }

  // Contacts: the people shown, two added for the check and taken out after it.
  try {
    const win = await open('contacts');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('html[data-ready]'))`), 'the contacts window', 12000).catch(() => {});
    const family = `Printcheck${stamp}`;
    const names = [`Wren ${family}`, `Finch ${family}`];
    const ids = [];
    let r = { opened: false, said: null, wrote: false, pages: 0, text: '', pictures: 0 };
    try {
      for (const [i, full] of names.entries()) {
        const contact = {
          id: null, name: { full, given: full.split(' ')[0], family, middle: '', prefix: '', suffix: '' },
          nickname: '', org: '', department: '', title: '', role: '',
          emails: [{ value: `print${i}-${stamp}@example.org`, type: 'work', pref: true, label: null }], phones: [], addresses: [],
          birthday: '', note: '', urls: [], categories: [], kind: 'individual', extra: [],
        };
        ids.push(await js(`window.rutbaOffice.contacts.save(${JSON.stringify({ contact, book: 'local' })}).then((c) => c.id)`));
      }
      // Searched for, the list is asked again and shows the two alone; no card is chosen, so the list prints.
      await js(`(() => { const i = document.querySelector('.rw-search input'); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(family)}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
      await until(() => js(`${JSON.stringify(names)}.every((n) => document.body.textContent.includes(n))`), 'the two people in the list', 10000).catch(() => {});
      await wait(300);
      r = await printed(win, 'contacts');
    } finally {
      if (ids.length) await js(`window.rutbaOffice.contacts.remove(${JSON.stringify({ ids })})`).catch(() => {});
    }
    check('print-apps: Contacts prints the people shown: the dialog counts its sheets and Save as PDF writes them with each name',
      r.opened && r.wrote && r.pages >= 1 && r.said === r.pages && names.every((n) => has(r, n)), brief(r));
    await quiet('Contacts', win);
  } catch (err) {
    check('print-apps: the Contacts print check ran', false, err?.message || String(err));
  }

  // Image: the picture as it stands, fitted to one sheet.
  try {
    const file = path.join(root, 'print-check.png');
    fs.writeFileSync(file, noisyPicture(480, 320, 7).toPNG());
    const win = await open('image', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('html[data-ready]')) && Boolean(document.querySelector('canvas'))`), 'the picture drawn', 12000).catch(() => {});
    await wait(800);
    const r = await printed(win, 'image');
    check('print-apps: Image prints the picture: the dialog counts one sheet and Save as PDF writes it with the picture on it',
      r.opened && r.wrote && r.pages === 1 && r.said === 1 && r.pictures === 1, brief(r));
    await quiet('Image', win);
  } catch (err) {
    check('print-apps: the Image print check ran', false, err?.message || String(err));
  }

  // Pictures: the picture it was opened with.
  try {
    const pics = path.join(root, 'Photos');
    fs.mkdirSync(pics, { recursive: true });
    for (const [i, name] of ['a.png', 'b.png', 'c.png'].entries()) fs.writeFileSync(path.join(pics, name), noisyPicture(240 + i * 40, 160 + i * 20, 20 + i).toPNG());
    const win = await open('pictures', path.join(pics, 'b.png'));
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`[...document.querySelectorAll('.chip')].some((c) => / of 3$/.test(c.textContent.trim()))`), 'the picture shown among its three', 12000).catch(() => {});
    await wait(500);
    const r = await printed(win, 'pictures');
    check('print-apps: Pictures prints: the dialog counts its sheets and Save as PDF writes them, a picture on every sheet',
      r.opened && r.wrote && r.pages >= 1 && r.said === r.pages && r.pictures === r.pages, brief(r));
    await quiet('Pictures', win);
  } catch (err) {
    check('print-apps: the Pictures print check ran', false, err?.message || String(err));
  }
}
