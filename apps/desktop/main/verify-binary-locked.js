// Password-protected Word, Excel and PowerPoint 97-2003 files.
//
// locked.xls, locked.doc and locked.ppt were saved with the password
// "Rutba-1" by Excel, Word and PowerPoint themselves
// (tools/make-locked-fixtures.ps1). Each opens on the Password dialog in its
// window; the workbook's refuses a wrong password in place, and the right
// one opens each file in full — the cells, the words, the slides — marked
// as saved under its password. Run alone with RUTBA_VERIFY_ONLY=lockedbinary.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const FIXTURES = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'tests', 'fixtures', 'binary');
const PASSWORD = 'Rutba-1';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the files are copied
 */
export async function verifyBinaryLocked(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc } = h;
  // Open a file on its Password dialog, and type into it as a person does.
  const opening = async (app, name) => {
    const file = path.join(dir, name);
    fs.copyFileSync(path.join(FIXTURES, name), file);
    const win = await open(app, file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const asked = await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Password"] .pw-open-password'))`), `the Password dialog for ${name}`, 12000).then(() => true, () => false);
    const typeIn = async (text) => {
      await js(`(() => { const i = document.querySelector('.pw-open-password'); if (!i) return false; i.focus(); i.select?.(); return true; })()`);
      await wait(80);
      win.webContents.insertText(text);
      await wait(150);
      return js(`(() => { const b = document.querySelector('.pw-open-ok'); if (!b || b.disabled) return 'no'; b.click(); return 'clicked'; })()`);
    };
    const meta = () => {
      const session = doc.sessions().filter((s) => s.path === file).pop();
      return session ? doc.meta({ id: session.id }) : null;
    };
    return { win, js, asked, typeIn, meta, file };
  };

  try {
    // The workbook: a wrong password first, then the right one.
    const xls = await opening('sheets', 'locked.xls');
    const name = await xls.js(`document.querySelector('.pw-file')?.textContent || ''`);
    await xls.typeIn('rutba-1');
    const refused = await until(() => xls.js(`Boolean(document.querySelector('.pw-error')) && document.querySelector('.pw-open-password')?.value === ''`), 'the wrong password refused', 15000).then(() => true, () => false);
    check('locked 97-2003: a password-protected .xls opens on the Password dialog, which refuses a wrong password in place',
      xls.asked && name === 'locked.xls' && refused, JSON.stringify({ asked: xls.asked, name, refused }));
    await xls.typeIn(PASSWORD);
    const cells = await until(() => xls.js(`document.querySelectorAll('.sh-cells .sh-cell').length > 10`), 'the cells', 20000).then(() => true, () => false);
    const words = await xls.js(`[...document.querySelectorAll('.sh-cells .sh-cell')].map((c) => c.textContent.trim()).filter(Boolean).slice(0, 6)`);
    check('locked 97-2003: with its password the workbook opens in full, saved under the password again',
      cells && xls.meta()?.encrypted === true, `cells ${JSON.stringify(words)}; encrypted ${xls.meta()?.encrypted}`);
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      xls.win.webContents.invalidate();
      await wait(400);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'locked-xls.png'), (await xls.win.webContents.capturePage()).toPNG());
    }

    // The document and the presentation, with the right password at once.
    const word = await opening('word', 'locked.doc');
    await word.typeIn(PASSWORD);
    const page = await until(() => word.js(`[...document.querySelectorAll('.wd-page [data-block]')].some((b) => b.textContent.includes('A locked document'))`), 'the words', 20000).then(() => true, () => false);
    check('locked 97-2003: a password-protected .doc opens to its words with its password', word.asked && page && word.meta()?.encrypted === true,
      JSON.stringify({ asked: word.asked, page, encrypted: word.meta()?.encrypted }));

    const deck = await opening('slides', 'locked.ppt');
    await deck.typeIn(PASSWORD);
    const slides = await until(() => deck.js(`document.querySelectorAll('.sl-thumb').length === 2 && (document.querySelector('.sl-stage')?.textContent || '').includes('A locked deck')`), 'the slides', 20000).then(() => true, () => false);
    check('locked 97-2003: a password-protected .ppt opens to its slides with its password', deck.asked && slides && deck.meta()?.encrypted === true,
      JSON.stringify({ asked: deck.asked, slides, encrypted: deck.meta()?.encrypted }));

    const complaints = [...(await errorsIn(xls.win)), ...(await errorsIn(word.win)), ...(await errorsIn(deck.win))];
    check('locked 97-2003: opening them reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('locked 97-2003: the checks ran', false, err.message);
  }
}
