// Templates that need a newer Office (the owner, 2026-10-21: "the templates
// should have min version check so that the outdate software or templates
// can easily be upgraded to new set"), checked in the window: an Office
// older than the set's Data Files part (manifest.json's minOffice) shows
// each ready-made database as needing that version, with no way to make it
// but an update; at its own version the same cards make databases again.
// And the template gallery says what the update check found of a newer set:
// ready for the next start, or needing a newer Office.

import { broadcast } from '@rutba/office-shell/electron/main';

export async function verifyTemplateVersions(h) {
  const { open, check, until, wait } = h;
  const own = process.env.RUTBA_APP_VERSION;
  try {
    // An Office one release behind the set's Data Files part.
    process.env.RUTBA_APP_VERSION = '2.12.0';
    let win = await open('base');
    let js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.bs-sample[data-needs]').length > 0`), 'the cards marked', 8000).catch(() => {});
    const old = await js(`({
      samples: document.querySelectorAll('.bs-sample').length,
      marked: [...document.querySelectorAll('.bs-sample[data-needs]')].map((n) => n.dataset.needs),
      buttons: document.querySelectorAll('.bs-sample .bs-sample-data').length,
      update: document.querySelector('.bs-sample-update')?.textContent || '',
      templates: [...document.querySelectorAll('.bs-template[data-needs]')].length,
    })`);
    check('base: an Office older than the templates\' Data Files part shows each ready-made database as needing that version, offering the update rather than the database',
      old.samples > 0 && old.marked.length === old.samples && old.marked.every((v) => v === '2.13.0') && old.buttons === 0 && /Needs Rutba Office 2\.13\.0/.test(old.update) && old.templates > 0,
      JSON.stringify(old));
    win.close();

    // This Office, as it is: nothing marked.
    process.env.RUTBA_APP_VERSION = own;
    win = await open('base');
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.bs-sample .bs-sample-data').length > 0`), 'the cards', 8000).catch(() => {});
    const now = await js(`({ marked: document.querySelectorAll('[data-needs]').length, buttons: document.querySelectorAll('.bs-sample .bs-sample-data').length })`);
    check('base: at its own version Office marks none of its ready-made databases, and each can be made', now.marked === 0 && now.buttons > 0, JSON.stringify(now));
    win.close();

    // The gallery, told by the update check of a newer set.
    win = await open('home');
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.home-more'))`), 'the launcher', 8000);
    await js(`document.querySelector('.home-more').click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.tg-dialog'))`), 'the gallery', 8000);
    // Each template links to its page online (D-2026-10-22-03).
    await until(() => js(`document.querySelectorAll('.tg-card-wrap .tg-page-link').length > 10`), 'the templates’ links', 6000).catch(() => {});
    const pages = await js(`({ cards: document.querySelectorAll('.tg-card-wrap').length, links: [...document.querySelectorAll('.tg-card-wrap .tg-page-link')].map((n) => n.dataset.page) })`);
    check('templates: each template in the gallery links to its page online, at office.rutba.io/templates/<its name>',
      pages.cards > 10 && pages.links.length === pages.cards && pages.links.every((u) => /^https:\/\/office\.rutba\.io\/templates\/[a-z0-9-]+$/.test(u)) && pages.links.includes('https://office.rutba.io/templates/word-letter'),
      JSON.stringify({ cards: pages.cards, first: pages.links.slice(0, 3) }));
    // A document made from one links to that template's page, in its status bar; a blank one does not.
    const made = await open('word', null, { template: 'word-letter' });
    const mjs = (code) => made.webContents.executeJavaScript(code);
    await until(() => mjs(`Boolean(document.querySelector('.rw-template-help'))`), 'the template link', 10000).catch(() => {});
    const help = await mjs(`({ page: document.querySelector('.rw-template-help')?.dataset.page || null, words: document.querySelector('.rw-template-help')?.textContent || '' })`);
    const blank = await open('word', null, { template: 'blank' });
    await wait(1500);
    const none = await blank.webContents.executeJavaScript(`!document.querySelector('.rw-template-help')`);
    made.close();
    blank.close();
    check("templates: a document made from a template links to that template's page from its status bar, and a blank one has no such link",
      help.page === 'https://office.rutba.io/templates/word-letter' && /Letter template/.test(help.words) && none, JSON.stringify({ help, none }));
    const said = {};
    for (const [phase, extra] of [['ready', {}], ['needs-office', { needs: '9.0.0' }]]) {
      broadcast('update:state', { state: 'idle', templates: { phase, set: '9.9.0', needs: null, ...extra } });
      await until(() => js(`document.querySelector('.tg-set')?.dataset.setPhase === ${JSON.stringify(phase)}`), `the ${phase} line`, 4000).catch(() => {});
      said[phase] = await js(`document.querySelector('.tg-set')?.textContent || ''`);
    }
    broadcast('update:state', { state: 'idle', templates: null });
    await until(() => js(`!document.querySelector('.tg-set')`), 'the line gone', 4000).catch(() => {});
    const gone = await js(`!document.querySelector('.tg-set')`);
    check('templates: the gallery says a newer template set is ready for the next start, or that it needs a newer Rutba Office, and nothing when there is none',
      /9\.9\.0.*next time Rutba Office starts/.test(said.ready) && /9\.9\.0.*need Rutba Office 9\.0\.0/.test(said['needs-office']) && gone, JSON.stringify({ said, gone }));
  } catch (err) {
    check('base: the template version checks ran', false, err.stack || err.message);
  } finally {
    process.env.RUTBA_APP_VERSION = own;
  }
}
