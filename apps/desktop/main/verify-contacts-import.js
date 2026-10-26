// Contacts and the accounts Calendar and Contacts share, and Mail's
// invitation strip: the faults the how-to scripts found (2026-10-24, A6.8,
// A6.9, A6.14, A6.15). A CSV Outlook wrote in ANSI comes in with its accents,
// after a preview that shows each column, what it was read as and lets one
// be changed, and says what it leaves out; today's Google export keeps its
// companies, job titles and kinds, and a second import of it adds nobody
// again. An account is called what the person names it and renamed later,
// and new cards go to its address book. Mail's invitation strip remembers
// the answer when the message is opened again, and the reply's file shows
// its size. The account's server is the suite's own small one, on this
// computer. Run alone with RUTBA_VERIFY_ONLY=contacts-import.

import fs from 'node:fs';
import path from 'node:path';
import { startDavServer } from '../../../tests/fixtures/dav-server.js';

const latin1 = (s) => Buffer.from(s, 'latin1');

function fixtures(dir) {
  const own = path.join(dir, 'contacts-import-fixtures');
  fs.mkdirSync(own, { recursive: true });
  // Outlook's "Comma Separated Values" export: ANSI, so "ë" and "ü" are one byte each.
  const outlook = path.join(own, 'outlook.csv');
  fs.writeFileSync(outlook, latin1([
    'First Name,Last Name,Company,E-mail Address,Mobile Phone,Notes',
    'Zoë,Müller,Café Nord,zoe@example.com,07700 900111,Met at the fair',
    ',,,,,Only a note: nothing a card keeps',
    'Sam,Patel,,sam@example.org,,',
    '',
  ].join('\r\n')));
  const google = path.join(own, 'google.csv');
  fs.writeFileSync(google, [
    'First Name,Middle Name,Last Name,Organization Name,Organization Title,Labels,E-mail 1 - Label,E-mail 1 - Value,Phone 1 - Label,Phone 1 - Value',
    'Kim,,Lee,Tech Style Ltd,Head of Product,* myContacts ::: Suppliers,* Work,kim@example.com,Mobile,+44 7700 900123',
    'Bo,,Lund,,,* myContacts,,,Home,0113 496 0000',
    '',
  ].join('\r\n'));
  return { outlook, google };
}

export async function verifyContactsImport({ open, check, until, wait, errorsIn }, { dir }) {
  const fx = fixtures(dir);
  let server = null;
  // The cards the address book held before: the ones the imports add go at the end, so the
  // Contacts checks after this one find the book as the run made it (their .vcf holds Kim Lee
  // and Sam Patel too, which a second import rightly adds nobody for).
  let win = null;
  let held = null;
  try {
    win = await open('contacts', fx.outlook);
    const js = (code) => win.webContents.executeJavaScript(code);
    held = await js(`window.rutbaOffice.contacts.list({}).then((l) => l.map((c) => c.id))`).catch(() => null);
    const click = (sel, text) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => x.textContent.trim() === ${JSON.stringify(text)}); b?.click(); return Boolean(b); })()`);
    const clickLike = (sel, re) => js(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => ${re}.test(x.textContent.trim())); b?.click(); return Boolean(b); })()`);
    const type = (sel, value) => js(`(() => { const i = document.querySelector(${JSON.stringify(sel)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(value)}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    const toasts = () => js(`[...document.querySelectorAll('.rw-toast')].map((t) => t.textContent).join(' | ')`);

    // Opened with the window: shown, then "Add all" checks it in the preview first.
    await until(() => js(`Boolean(document.querySelector('.ct-offer'))`), 'the file\'s cards on show', 8000).catch(() => {});
    await clickLike('.ct-offer .rw-btn', '/^Add/');
    await until(() => js(`Boolean(document.querySelector('.ct-csv .ct-csv-row[data-column]'))`), 'the preview', 6000).catch(() => {});
    const preview = await js(`(() => {
      const rows = [...document.querySelectorAll('.ct-csv .ct-csv-row[data-column]')].map((r) => ({ h: r.dataset.column, v: r.querySelector('.v')?.textContent, as: r.querySelector('select')?.selectedOptions[0]?.textContent }));
      return { rows, cards: [...document.querySelectorAll('.ct-csv-card b')].map((b) => b.textContent), left: document.querySelector('.ct-csv-left')?.textContent || '', go: document.querySelector('.ct-csv-go')?.textContent };
    })()`);
    const row = (h) => preview.rows.find((r) => r.h === h) || {};
    check('contacts: a CSV Outlook wrote in ANSI is read with its accents, in a preview of each column and what it was read as',
      row('First Name').v === 'Zoë' && row('First Name').as === 'First name' && row('Company').as === 'Company' && preview.cards[0] === 'Zoë Müller' && /Import 2 cards/.test(preview.go || ''), JSON.stringify(preview));
    check('contacts: the preview says what it leaves out', /Left out 1 row/.test(preview.left), preview.left);
    // Notes left out: then said as not read.
    await js(`(() => { const s = document.querySelector('.ct-csv-row[data-column="Notes"] select'); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ''); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
    await until(() => js(`/Not read: the Notes column/.test(document.querySelector('.ct-csv-left')?.textContent || '')`), 'Notes left out', 4000).catch(() => {});
    const changed = await js(`document.querySelector('.ct-csv-left')?.textContent || ''`);
    check('contacts: a column can be changed in the preview, and one left out is named', /Not read: the Notes column/.test(changed), changed);
    await js(`document.querySelector('.ct-csv-go')?.click(), 'import'`);
    await until(() => js(`!document.querySelector('.ct-csv') && document.body.innerText.includes('Zoë Müller')`), 'the cards in', 6000).catch(() => {});
    const said = await toasts();
    const zoe = await js(`(async () => { const all = await window.rutbaOffice.contacts.list({ query: 'zoe@' }); return all[0] ? { name: all[0].display, org: all[0].org, note: all[0].note, phone: all[0].phones?.[0] } : null; })()`);
    check('contacts: the import keeps them as previewed, and says what it left out',
      zoe?.name === 'Zoë Müller' && zoe.org === 'Café Nord' && !zoe.note && zoe.phone?.type === 'mobile' && /2 added/.test(said) && /Left out 1 row/.test(said) && /Not read: the Notes column/.test(said), JSON.stringify({ zoe, said }));

    // Today's Google export, twice.
    for (const round of [1, 2]) {
      const people = await open('contacts', fx.google);
      const pjs = (code) => people.webContents.executeJavaScript(code);
      await until(() => pjs(`Boolean(document.querySelector('.ct-offer'))`), 'the Google file on show', 8000).catch(() => {});
      await pjs(`[...document.querySelectorAll('.ct-offer .rw-btn')].find((b) => /^Add/.test(b.textContent.trim()))?.click(), 'add'`);
      await until(() => pjs(`Boolean(document.querySelector('.ct-csv-go:not([disabled])'))`), 'the preview', 6000).catch(() => {});
      await pjs(`document.querySelector('.ct-csv-go')?.click(), 'import'`);
      await until(() => pjs(`!document.querySelector('.ct-csv') && document.querySelectorAll('.rw-toast').length > 0`), 'the import', 6000).catch(() => {});
      const done = await pjs(`[...document.querySelectorAll('.rw-toast')].map((t) => t.textContent).join(' | ')`);
      if (round === 1) {
        const kim = await pjs(`(async () => { const [c] = await window.rutbaOffice.contacts.list({ query: 'kim@example.com' }); return c ? { org: c.org, title: c.title, tags: c.categories, email: c.emails?.[0]?.type, phone: c.phones?.[0]?.type } : null; })()`);
        check('contacts: today\'s Google export keeps the company, job title and kinds, and no myContacts tag',
          kim?.org === 'Tech Style Ltd' && kim.title === 'Head of Product' && kim.email === 'work' && kim.phone === 'mobile' && JSON.stringify(kim.tags) === '["Suppliers"]', JSON.stringify(kim));
      } else {
        const bos = await pjs(`(async () => (await window.rutbaOffice.contacts.list({ query: 'Bo Lund' })).length)()`);
        check('contacts: importing the same Google export again adds nobody, a card with no email address among them', /0 added/.test(done) && /2 already here/.test(done) && bos === 1, JSON.stringify({ done, bos }));
      }
    }

    // Accounts: named here, renamed, and new cards go to its book.
    server = await startDavServer();
    await click('.rw-ribbon .rw-btn', 'Accounts');
    await until(() => js(`Boolean(document.querySelector('input.dav-url'))`), 'the Accounts dialog', 4000).catch(() => {});
    await type('input.dav-url', server.url);
    await type('input.dav-user', 'ann');
    await type('input.dav-password', 'secret');
    await type('input.dav-name-new', 'Family');
    await wait(150);
    await js(`document.querySelector('.dav-add').click(), 'add'`);
    await until(() => js(`Boolean(document.querySelector('.dav-account .dav-name'))`), 'the account listed', 10000).catch(() => {});
    await until(() => js(`Boolean(document.querySelector('.dav-default-where'))`), 'where new cards go', 4000).catch(() => {});
    const added = await js(`({ name: document.querySelector('.dav-account .dav-name')?.textContent, book: document.querySelector('.dav-default-book')?.selectedOptions[0]?.textContent, where: document.querySelector('.dav-default-where')?.textContent || '', error: document.querySelector('.dav-error')?.textContent || '' })`);
    check('contacts: an account added in a new profile is called what the person names it, and new cards go to its address book, said plainly',
      added.name === 'Family' && added.book === 'Family: Friends' && /sent to Friends on the server of Family/.test(added.where), JSON.stringify(added));
    await js(`document.querySelector('.dav-rename-start')?.click(), 'rename'`);
    await until(() => js(`Boolean(document.querySelector('input.dav-rename'))`), 'the name to change', 3000).catch(() => {});
    await type('input.dav-rename', 'Home');
    await wait(100);
    await js(`document.querySelector('.dav-rename-save')?.click(), 'saved'`);
    await until(() => js(`document.querySelector('.dav-account .dav-name')?.textContent === 'Home'`), 'the new name', 4000).catch(() => {});
    const renamed = await js(`(async () => ({ shown: document.querySelector('.dav-account .dav-name')?.textContent, kept: (await window.rutbaOffice.dav.accounts()).map((a) => a.name) }))()`);
    check('contacts: Rename in Accounts changes the account\'s name', renamed.shown === 'Home' && JSON.stringify(renamed.kept) === '["Home"]', JSON.stringify(renamed));
    await js(`(() => { const s = document.querySelector('.dav-default-book'); if (!s) return false; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ''); s.dispatchEvent(new Event('change', { bubbles: true })); return true; })()`);
    await until(() => js(`/stay on this computer/.test(document.querySelector('.dav-default-where')?.textContent || '')`), 'new cards kept here', 4000).catch(() => {});
    const here = await js(`(async () => ({ where: document.querySelector('.dav-default-where')?.textContent || '', label: document.querySelector('.dav-default-book option:not([value=""])')?.textContent, kept: (await window.rutbaOffice.dav.books()).defaultBook }))()`);
    check('contacts: "This computer only" chosen in Accounts says new cards stay here, and books carry their account\'s name',
      /stay on this computer and are not sent/.test(here.where) && here.label === 'Home: Friends' && here.kept === null, JSON.stringify(here));
    await js(`[...document.querySelectorAll('.dav-account .rw-btn')].find((b) => b.textContent.trim() === 'Remove')?.click(), 'removed'`);
    await until(() => js(`!document.querySelector('.dav-account')`), 'the account gone', 4000).catch(() => {});
    await click('.rw-dialog .rw-btn', 'Close');

    // Mail's invitation strip: answered, opened again, it says so; the reply's file has its size.
    const mail = await open('mail');
    const mjs = (code) => mail.webContents.executeJavaScript(code);
    await until(() => mjs(`document.querySelectorAll('.ml-row').length > 0`), 'the seeded messages', 8000).catch(() => {});
    const openInvite = () => mjs(`[...document.querySelectorAll('.ml-row')].find((r) => /Invitation: Quarterly numbers/.test(r.textContent))?.click(), 'opened'`);
    await openInvite();
    await until(() => mjs(`Boolean(document.querySelector('.ml-invite .rw-btn'))`), 'the invitation card', 8000).catch(() => {});
    await mjs(`[...document.querySelectorAll('.ml-invite .rw-btn')].find((b) => b.textContent.trim() === 'Tentative')?.click(), 'tentative'`);
    await until(() => mjs(`Boolean(document.querySelector('.ml-attachment .size'))`), 'the reply with its file', 8000).catch(() => {});
    const size = await mjs(`document.querySelector('.ml-attachment .size')?.textContent || ''`);
    check('mail: the reply to an invitation carries reply.ics with its size, not 0 B', size && !/^0\s*B$/.test(size.trim()), size);
    await mjs(`[...document.querySelectorAll('.rw-btn, button')].find((b) => /^(Close|Discard|Cancel)$/.test(b.textContent.trim()))?.click(), 'closed'`);
    await until(() => mjs(`!document.querySelector('.ml-attachment .size')`), 'the reply to close', 4000).catch(() => {});
    await mjs(`[...document.querySelectorAll('.ml-row')].find((r) => !/Invitation: Quarterly numbers/.test(r.textContent))?.click(), 'another'`);
    await wait(500);
    await openInvite();
    await until(() => mjs(`/you answered/.test(document.querySelector('.ml-invite-who')?.textContent || '')`), 'the answer remembered', 6000).catch(() => {});
    const who = await mjs(`document.querySelector('.ml-invite-who')?.textContent || ''`);
    check('mail: an invitation answered and opened again says what was answered', /you answered tentative/.test(who), who);

    const complaints = [...(await errorsIn(win)), ...(await errorsIn(mail))];
    check('contacts and mail: the import, accounts and invitation checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('contacts: the import and accounts checks ran', false, err.message);
  } finally {
    await server?.close();
    if (win && !win.isDestroyed() && held) {
      await win.webContents.executeJavaScript(`(async () => {
        const held = new Set(${JSON.stringify(held)});
        const added = (await window.rutbaOffice.contacts.list({})).map((c) => c.id).filter((id) => !held.has(id));
        if (added.length) await window.rutbaOffice.contacts.remove({ ids: added });
        return added.length;
      })()`).catch(() => {});
    }
  }
}
