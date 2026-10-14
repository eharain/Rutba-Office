// Documents: File → Info → Check for Issues → Inspect Document. The card
// is on the Info page; the Document Inspector lists the kinds Word's does,
// Inspect says what the letter carries (comments, tracked changes, personal
// details, properties, headers and footers, hidden words), Remove All takes
// a kind out of the open document and the page shows it gone, and the file
// saved afterwards holds none of it. Run alone with RUTBA_VERIFY_ONLY=inspect.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { secretLetter } from './verify-review-fixtures.js';

/** The letter, with a hidden line and a setting part to say Word's flags in. */
function letter() {
  const pkg = OoxmlPackage.read(secretLetter());
  pkg.write_('word/document.xml', pkg.text('word/document.xml').replace('<w:sectPr>', '<w:p><w:r><w:t xml:space="preserve">Shown. </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden salary note</w:t></w:r></w:p><w:sectPr>'));
  pkg.addPart('word/settings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  pkg.addRelationshipTo('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
  return pkg.write();
}

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyInspect(h, { dir }) {
  const { open, check, wait, capture, doc } = h;
  // Waits as the harness's until does, and answers what the condition found once it held.
  const until = async (fn, what, ms) => { let found; await h.until(async () => { found = await fn(); return found; }, what, ms); return found; };
  const own = path.join(dir, 'inspect');
  const file = path.join(own, 'letter.docx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, letter());
  } catch (err) {
    return check('inspect: the fixture was made', false, err.message);
  }
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const fresh = async (name) => {
      wc.invalidate();
      await wait(700);
      await capture(win, name);
    };
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="3"]'))`), 'the letter', 10000);

    // 1. File → Info: the Inspect Document card; Check for Issues → Inspect Document.
    await js(`document.querySelector('.rw-appmark')?.click(), 1`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info')))`), 'the app menu', 4000).catch(() => {});
    const viaMenu = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim().startsWith('Info')); b?.click(); return Boolean(b); })()`);
    const info = await until(() => js(`Boolean(document.querySelector('.pw-card[data-card="inspect"]'))`), 'the Inspect Document card', 5000).catch(() => false);
    await fresh('inspect-info.png');
    await js(`document.querySelector('.pw-card[data-card="inspect"] .pw-card-tile')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.pw-menu [data-item="inspect"]'))`), 'Check for Issues', 3000).catch(() => {});
    await js(`document.querySelector('.pw-menu [data-item="inspect"]')?.click(), 1`);
    const chooser = await until(() => js(`document.querySelectorAll('.di-dialog .di-row input[type="checkbox"]').length >= 8`), 'the Document Inspector', 8000).catch(() => false);
    const kinds = await js(`[...document.querySelectorAll('.di-dialog .di-row h5')].map((n) => n.textContent)`);
    check(
      'inspect: File → Info has the Inspect Document card, and Check for Issues → Inspect Document opens the Document Inspector with Word\'s kinds ticked',
      info === true && chooser === true && ['Comments', 'Revisions', 'Personal Information', 'Document Properties', 'Custom XML Data', 'Headers, Footers, and Watermarks', 'Invisible Content', 'Hidden Text'].every((k) => kinds.includes(k)),
      JSON.stringify({ viaMenu, info, chooser, kinds }),
    );

    // 2. Inspect: what the letter carries.
    await js(`document.querySelector('.di-inspect')?.click(), 1`);
    const results = await until(() => js(`(() => { const rows = [...document.querySelectorAll('.di-dialog .di-row[data-found]')]; return rows.length >= 8 ? Object.fromEntries(rows.map((r) => [r.dataset.kind, Number(r.dataset.found)])) : null; })()`), 'the results', 8000).catch(() => null);
    await fresh('inspect-results.png');
    check(
      'inspect: Inspect finds the comment, the tracked changes, the personal details, the properties, the header and footer and the hidden words, each with Remove All',
      Boolean(results) && results.comments === 1 && results.revisions >= 2 && results.personal >= 2 && results.properties >= 1 && results.headers === 2 && results.hidden === 1
        && (await js(`['comments', 'revisions', 'personal', 'headers', 'hidden'].every((k) => Boolean(document.querySelector('.di-remove-' + k)))`)),
      JSON.stringify(results),
    );

    // 3. Remove All, kind by kind, and the page shows it gone.
    const before = (doc.model({ id: session.id }).blocks || []).map((b) => b.text).join('|');
    for (const k of ['comments', 'revisions', 'hidden', 'headers', 'personal']) {
      await js(`document.querySelector('.di-remove-${k}')?.click(), 1`);
      await until(() => js(`document.querySelector('.di-row[data-kind="${k}"]')?.dataset.found === '0'`), `${k} removed`, 8000).catch(() => {});
    }
    const after = await js(`Object.fromEntries([...document.querySelectorAll('.di-dialog .di-row[data-found]')].map((r) => [r.dataset.kind, Number(r.dataset.found)]))`);
    await fresh('inspect-removed.png');
    await js(`document.querySelector('.di-close')?.click(), 1`);
    const page = await until(() => js(`(() => { const t = document.querySelector('.wd-page')?.textContent || ''; return !/Hidden salary note/.test(t) && /by Ayesha Khan/.test(t) && !/old\\.example/.test(t) && !document.querySelector('.wd-comment-mark, .wd-ins, .wd-del') ? t.slice(0, 80) : null; })()`), 'the page as it now is', 8000).catch(() => null);
    const model = doc.model({ id: session.id });
    check(
      'inspect: Remove All takes each kind out of the open document, comments gone, tracked changes accepted, hidden words and headers removed, the authors blanked, and the page shows it',
      after.comments === 0 && after.revisions === 0 && after.hidden === 0 && after.headers === 0 && after.personal === 0 && Boolean(page) && doc.meta({ id: session.id }).dirty === true && (model.comments || []).length === 0,
      JSON.stringify({ after, page, before: before.slice(0, 80) }),
    );

    // 4. Saved: none of it in the file.
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn, .rw-quick .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Save')); b?.click(); return Boolean(b); })()`);
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => {});
    const pkg = OoxmlPackage.read(fs.readFileSync(file));
    const body = pkg.text('word/document.xml');
    const saved = {
      comments: !pkg.has('word/comments.xml') && !/commentReference/.test(body),
      revisions: !/<w:(?:ins|del)\b/.test(body),
      hidden: !/Hidden salary note/.test(body),
      headers: !pkg.has('word/header1.xml') && !/headerReference|footerReference/.test(body),
      personal: !/Bilal Ahmed/.test(pkg.text('docProps/core.xml')) && /<w:removePersonalInformation\/>/.test(pkg.text('word/settings.xml')),
    };
    check('inspect: the saved document holds none of what was removed, and tells Word to keep personal information out', Object.values(saved).every(Boolean), JSON.stringify(saved));
  } catch (err) {
    check('inspect: the Inspect checks ran', false, err.stack || err.message);
  }
}
