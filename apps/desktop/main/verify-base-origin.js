// Data Files' ready-made databases linked to their pages online, and a newer
// design offered to a database made from one (owner, 2026-10-22,
// D-2026-10-22-03 and -04): the start screen's sample and template cards
// each have a small link to their page on office.rutba.io; a database made
// from a template has "About this database online"; one made from an older
// design of it is offered the newer one, and Update the design brings it in
// with its records kept. Run alone with RUTBA_VERIFY_ONLY=base-origin.

import fs from 'node:fs';
import path from 'node:path';
import { Base, buildTemplate, designPrint } from '@rutba/database';

export async function verifyBaseOrigin({ open, check, until, wait }, { dir }) {
  const current = path.join(dir, 'Contacts now.rbase');
  const older = path.join(dir, 'Contacts older.rbase');
  try {
    for (const f of [current, older]) for (const s of ['', '-wal', '-shm']) fs.rmSync(f + s, { force: true });
    let b = Base.open(current);
    buildTemplate(b, 'contacts');
    b.close();
    // As a database made from an earlier set would be: without Notes, and remembering so.
    b = Base.open(older);
    buildTemplate(b, 'contacts');
    b.insert('Contacts', { 'First Name': 'Amira', 'Last Name': 'Khan' });
    const print = designPrint(b);
    b.removeField('Contacts', 'Notes');
    print.tables.Contacts = print.tables.Contacts.filter((f) => f !== 'Notes');
    b.setMeta('origin.design', JSON.stringify(print));
    b.close();

    // The start screen: each card's page.
    let win = await open('base');
    let js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-sample[data-sample="cars"] .bs-page-link'))`), "the cards' links", 8000);
    const links = await js(`({ cars: document.querySelector('.bs-sample[data-sample="cars"] .bs-page-link')?.dataset.page, contacts: document.querySelector('.bs-template[data-template="contacts"]')?.parentElement.querySelector('.bs-page-link')?.dataset.page, samples: document.querySelectorAll('.bs-sample .bs-page-link').length, templates: document.querySelectorAll('.bs-card-wrap .bs-page-link').length })`);
    globalThis.__rutbaCheckOpened = [];
    await js(`document.querySelector('.bs-sample[data-sample="cars"] .bs-page-link').click(), 1`);
    await wait(300);
    const opened = (globalThis.__rutbaCheckOpened || []).slice();
    check('base: each sample and ready-made database on the start screen links to its page online, and the link opens it in the browser',
      links.cars === 'https://office.rutba.io/samples/copperbeech-motors' && links.contacts === 'https://office.rutba.io/samples/contacts' && links.samples === 25 && links.templates === 5 && opened.includes(links.cars),
      JSON.stringify({ links, opened }));

    // A database made from the template: its page, and nothing to update.
    win = await open('base', current);
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-about-online'))`), 'the link to its page', 8000).catch(() => {});
    await wait(800);
    const made = await js(`({ page: document.querySelector('.bs-about-online')?.dataset.page || null, banner: Boolean(document.querySelector('.bs-update')) })`);
    check('base: a database made from a ready-made one links to its page online, and offers nothing while it is up to date',
      made.page === 'https://office.rutba.io/samples/contacts' && !made.banner, JSON.stringify(made));
    // Owner, 2026-10-23: "little world / web icon next to the database name to link to outside world",
    // and "fix panel becomes part of the scrolled panel".
    const placed = await js(`(() => {
      const globe = document.querySelector('.bs-about-online');
      const foot = document.querySelector('.bs-side-foot');
      return { inTitle: Boolean(globe?.closest('.rw-titlebar .rw-doctitle')), named: globe?.getAttribute('aria-label') || '', inPanel: Boolean(globe?.closest('.bs-side')), footPosition: foot ? getComputedStyle(foot).position : null };
    })()`);
    check('base: About this database online is a globe beside the database name in the title bar, and the side panel foot scrolls with the panel',
      placed.inTitle && placed.named === 'About this database online' && !placed.inPanel && placed.footPosition !== 'sticky', JSON.stringify(placed));

    // One made from an older design: offered the newer, and brought up to date, its record kept.
    win = await open('base', older);
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.bs-update'))`), 'the newer design offered', 8000).catch(() => {});
    const offered = await js(`({ shown: Boolean(document.querySelector('.bs-update')), changes: document.querySelector('.bs-update')?.dataset.changes, words: document.querySelector('.bs-update')?.textContent || '' })`);
    if (process.env.RUTBA_VERIFY_CAPTURE) { win.webContents.invalidate(); await wait(700); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'base-origin-update.png'), (await win.webContents.capturePage()).toPNG()); }
    await js(`document.querySelector('.bs-update-now').click(), 1`);
    await until(() => js(`!document.querySelector('.bs-update')`), 'the banner gone', 8000).catch(() => {});
    const b2 = Base.open(older);
    const brought = { notes: b2.fields('Contacts').some((f) => f.name === 'Notes'), records: b2.count('Contacts'), name: b2.rows('Contacts')[0]?.['Last Name'] };
    b2.close();
    const again = await js(`!document.querySelector('.bs-update')`);
    check('base: a database made from an older design is offered the newer one, and Update the design brings it in, its records as they were',
      offered.shown && offered.changes === '1' && /records stay as they are/.test(offered.words) && brought.notes && brought.records === 1 && brought.name === 'Khan' && again,
      JSON.stringify({ offered, brought, again }));
  } catch (err) {
    check('base: the ready-made links and newer design checks ran to the end', false, err?.message || String(err));
  }
}
