// Templates that need a newer Office (the owner, 2026-10-21: "the templates
// should have min version check so that the outdate software or templates
// can easily be upgraded to new set"), checked in the window: an Office
// older than the set's Data Files part (manifest.json's minOffice) shows
// each ready-made database as needing that version, with no way to make it
// but an update; at its own version the same cards make databases again.

export async function verifyTemplateVersions(h) {
  const { open, check, until } = h;
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
  } catch (err) {
    check('base: the template version checks ran', false, err.stack || err.message);
  } finally {
    process.env.RUTBA_APP_VERSION = own;
  }
}
