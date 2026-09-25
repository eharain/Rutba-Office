// Password-protected files: opened through the Password dialog, protected
// through File → Info → Encrypt with Password, saved encrypted, opened again.
//
// The fixture is made here with the engine — a document encrypted as Office
// 2013 onwards encrypts one — so nothing on disk is trusted but what this
// suite writes. A wrong password is typed first and refused in the dialog;
// the right one opens the words; Info shows the file as protected; a new
// password is set through the Encrypt Document dialog (a mismatched confirm
// refused first); Save writes it; and the file opens again only with the
// new one. Worksheets and Presentation open their own protected fixtures
// through the same dialog. Run alone with RUTBA_VERIFY_ONLY=encrypted.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import { encryptPackage, decryptPackage, isEncryptedPackage } from '@rutba/office-formats/crypt';

const FIRST = 'Opal-7 lantern';
const SECOND = 'Basalt 42!';
const WORDS = 'Board minutes — not for circulation.';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, capture, doc, sessionFor
 */
export async function verifyEncrypted(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const file = path.join(dir, 'protected.docx');
  try {
    fs.writeFileSync(file, encryptPackage(buildDocx({ paragraphs: [WORDS, 'Second item: the budget.'], styles: true }), FIRST));
  } catch (err) {
    return check('encrypted: the protected fixture was made with the engine', false, err.message);
  }

  let win;
  try {
    win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const fresh = async (name) => {
      win.webContents.invalidate();
      await wait(700);
      await capture(win, name);
    };
    /** Type into a field the way a person does: focus it, then the characters. */
    const typeInto = async (selector, text) => {
      await js(`(() => { const i = document.querySelector(${JSON.stringify(selector)}); if (!i) return false; i.focus(); i.select?.(); return true; })()`);
      await wait(80);
      win.webContents.insertText(text);
      await wait(150);
    };
    const click = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'none'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);

    // The Password dialog, inside the window, naming the file.
    const asked = await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Password"] .pw-open-password'))`), 'the Password dialog', 12000).catch(() => false);
    const prompt = await js(`(() => ({
      file: document.querySelector('.pw-file')?.textContent || '',
      masked: document.querySelector('.pw-open-password')?.type || '',
      ok: document.querySelector('.pw-open-ok')?.disabled,
      cancel: Boolean(document.querySelector('.pw-open-cancel')),
      reveal: Boolean(document.querySelector('.pw-reveal')),
      page: Boolean(document.querySelector('.wd-block')),
    }))()`);
    check(
      'encrypted: a protected document opens on a Password dialog in the window, naming the file, masked, with Cancel and OK',
      asked === true && prompt.file === 'protected.docx' && prompt.masked === 'password' && prompt.ok === true && prompt.cancel && prompt.reveal && !prompt.page,
      JSON.stringify(prompt)
    );
    await wait(500);
    await fresh('encrypted-prompt.png');

    // A wrong password: refused in place, the box emptied and marked.
    await typeInto('.pw-open-password', 'opal-7 lantern');
    const typed = await js(`document.querySelector('.pw-open-password')?.value || ''`);
    await click('.pw-open-ok');
    const refused = await until(() => js(`Boolean(document.querySelector('.pw-error')) && document.querySelector('.pw-open-password')?.value === '' && !document.querySelector('.pw-open-ok .rw-spinner')`), 'the wrong password to be refused', 15000).catch(() => false);
    const said = await js(`document.querySelector('.pw-error')?.textContent || ''`);
    check(
      'encrypted: a wrong password is refused in the dialog — "That password is not right", the box emptied for another try',
      typed === 'opal-7 lantern' && refused === true && /That password is not right/.test(said) && !(await js(`Boolean(document.querySelector('.wd-block'))`)),
      JSON.stringify({ typed, said })
    );
    await fresh('encrypted-wrong.png');

    // Show/hide: the eye shows what was typed and hides it again.
    await typeInto('.pw-open-password', FIRST);
    await click('.pw-reveal');
    await wait(150);
    const shown = await js(`document.querySelector('.pw-open-password')?.type`);
    await fresh('encrypted-reveal.png');
    await click('.pw-reveal');
    await wait(150);
    const hidden = await js(`document.querySelector('.pw-open-password')?.type`);
    check('encrypted: the eye shows the password and hides it again', shown === 'text' && hidden === 'password', `${shown} → ${hidden}`);

    // The right one: the words, as any document opens.
    await click('.pw-open-ok');
    const opened = await until(() => js(`[...document.querySelectorAll('.wd-page [data-block]')].some((b) => b.textContent.includes(${JSON.stringify(WORDS)}))`), 'the document to open', 20000).catch(() => false);
    const dialogGone = await js(`!document.querySelector('.rw-dialog[aria-label="Password"]')`);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const meta = session ? doc.meta({ id: session.id }) : null;
    check(
      'encrypted: the right password opens the document — its words on the page, the dialog gone, the session marked encrypted',
      opened === true && dialogGone && meta?.encrypted === true,
      JSON.stringify({ opened, dialogGone, encrypted: meta?.encrypted })
    );

    // File → Info: the Protect Document card says a password opens it.
    const openInfo = async () => {
      await click('.rw-appmark');
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info')))`), 'the app menu', 4000).catch(() => {});
      await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info'))?.click()`);
      return until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Info"] .pw-card'))`), 'the Info page', 5000).catch(() => false);
    };
    const info = await openInfo();
    await wait(400);
    const card = await js(`(() => ({
      on: document.querySelector('.pw-card')?.classList.contains('on'),
      heading: document.querySelector('.pw-card h4')?.textContent || '',
      said: document.querySelector('.pw-card-notes')?.textContent || '',
      encryption: document.querySelector('.pw-info-encryption')?.textContent || '',
      name: document.querySelector('.pw-info-name')?.textContent || '',
    }))()`);
    check(
      'encrypted: File → Info shows Protect Document marked — "A password is required to open this document." — and the encryption in its properties',
      info === true && card.on && card.heading === 'Protect Document' && /A password is required to open this document\./.test(card.said) && /AES-256/.test(card.encryption) && card.name === 'protected.docx',
      JSON.stringify(card)
    );
    await fresh('encrypted-info.png');

    // Protect Document → Encrypt with Password.
    await click('.pw-card-tile');
    await until(() => js(`Boolean(document.querySelector('.pw-menu-item[data-item="encrypt"]'))`), 'the Protect menu', 3000).catch(() => {});
    await fresh('encrypted-menu.png');
    await click('.pw-menu-item[data-item="encrypt"]');
    const setOpen = await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Encrypt Document"] .pw-set-password'))`), 'the Encrypt Document dialog', 4000).catch(() => false);
    const setWords = await js(`(() => ({
      current: document.querySelector('.pw-set-current')?.textContent || '',
      caution: document.querySelector('.pw-caution')?.textContent || '',
      confirm: Boolean(document.querySelector('.pw-set-again')),
    }))()`);
    check(
      'encrypted: Encrypt with Password asks for the password twice, warns that a lost one cannot be recovered, and says this document has one already',
      setOpen === true && setWords.confirm && /cannot be recovered/.test(setWords.caution) && /already needs a password/.test(setWords.current),
      JSON.stringify(setWords)
    );
    // A confirm that does not match is refused.
    await typeInto('.pw-set-password', SECOND);
    await typeInto('.pw-set-again', 'Basalt 43!');
    const mismatch = await until(() => js(`Boolean(document.querySelector('.pw-set-mismatch')) && document.querySelector('.pw-set-ok')?.disabled === true`), 'the mismatch', 3000).catch(() => false);
    await fresh('encrypted-set-mismatch.png');
    check('encrypted: a confirm that does not match keeps OK greyed and says so', mismatch === true);
    await typeInto('.pw-set-again', SECOND);
    const ready = await until(() => js(`document.querySelector('.pw-set-ok')?.disabled === false && !document.querySelector('.pw-set-mismatch')`), 'OK to wake', 3000).catch(() => false);
    await fresh('encrypted-set.png');
    await click('.pw-set-ok');
    const back = await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Info"] .pw-card.on'))`), 'Info again', 8000).catch(() => false);
    const dirty = session ? doc.meta({ id: session.id }).dirty : null;
    check('encrypted: OK takes the new password and returns to Info, the document marked as changed', ready === true && back === true && dirty === true, JSON.stringify({ ready, back, dirty }));
    await click('.pw-info-close');
    await wait(300);

    // Save: the file on disk opens with the new password and not the old.
    await js(`[...document.querySelectorAll('.rw-btn')].find((b) => (b.dataset.tip || b.title || '').startsWith('Save'))?.click()`);
    const saved = await until(() => session && doc.meta({ id: session.id }).dirty === false, 'the save', 20000).catch(() => false);
    const disk = fs.readFileSync(file);
    let newOpens = false;
    let oldRefused = false;
    try {
      newOpens = decryptPackage(disk, SECOND).bytes.length > 0;
    } catch {
      newOpens = false;
    }
    try {
      decryptPackage(disk, FIRST);
    } catch (err) {
      oldRefused = err.code === 'password';
    }
    check(
      'encrypted: Save writes the file encrypted with the new password — Agile Encryption in a compound file the old password no longer opens',
      saved === true && isEncryptedPackage(disk) && newOpens && oldRefused,
      JSON.stringify({ saved, encrypted: isEncryptedPackage(disk), newOpens, oldRefused })
    );
  } catch (err) {
    check('encrypted: the Word checks ran', false, err.message);
  }

  // Opened again, in a new window, with the new password.
  try {
    const again = await open('word', file);
    const js = (code) => again.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.pw-open-password'))`), 'the Password dialog again', 12000);
    await js(`document.querySelector('.pw-open-password').focus()`);
    again.webContents.insertText(SECOND);
    await wait(150);
    await js(`document.querySelector('.pw-open-ok').click()`);
    const reopened = await until(() => js(`[...document.querySelectorAll('.wd-page [data-block]')].some((b) => b.textContent.includes(${JSON.stringify(WORDS)}))`), 'the document to open again', 20000).catch(() => false);
    check('encrypted: the saved file opens again in a new window with the new password', reopened === true);
    again.webContents.invalidate();
    await wait(700);
    await capture(again, 'encrypted-reopened.png');
  } catch (err) {
    check('encrypted: the reopen check ran', false, err.message);
  }

  // Worksheets and Presentation ask through the same dialog.
  const others = [
    ['sheets', 'protected.xlsx', () => buildXlsx({ sheets: [{ name: 'Pay', rows: [['Salary', 91000]] }] }), `[...document.querySelectorAll('.sh-cell')].some((c) => c.textContent.includes('Salary'))`],
    ['slides', 'protected.pptx', () => buildPptx({ title: 'Plan', slides: [{ layout: 'title', title: 'Secret plan', body: 'Q4' }] }), `Boolean(document.querySelector('.sl-svg, .sl-thumb'))`],
  ];
  for (const [app, name, build, shownJs] of others) {
    try {
      const target = path.join(dir, name);
      fs.writeFileSync(target, encryptPackage(build(), FIRST, { spinCount: 20000 }));
      const w = await open(app, target);
      const js = (code) => w.webContents.executeJavaScript(code);
      const asked = await until(() => js(`document.querySelector('.pw-file')?.textContent === ${JSON.stringify(name)}`), `the ${app} Password dialog`, 12000).catch(() => false);
      if (app === 'slides') {
        // The dark theme, for the capture only: the page's own attribute,
        // put back before anything else is read.
        await js(`document.documentElement.dataset.theme = 'dark'`);
        w.webContents.invalidate();
        await wait(900);
        await capture(w, 'encrypted-prompt-dark.png');
        await js(`document.documentElement.dataset.theme = 'light'`);
      }
      await js(`document.querySelector('.pw-open-password').focus()`);
      w.webContents.insertText(FIRST);
      await wait(150);
      await js(`document.querySelector('.pw-open-ok').click()`);
      const shown = await until(() => js(shownJs), `the ${app} content`, 20000).catch(() => false);
      const meta = doc.sessions().filter((s) => s.path === target).pop();
      check(`encrypted: ${app === 'sheets' ? 'Worksheets' : 'Presentation'} asks for the password in its window and opens the file with it`, asked === true && shown === true && Boolean(meta), JSON.stringify({ asked, shown }));
      if (app === 'sheets') {
        // Info's Protect Workbook menu, as Excel lists it.
        await js(`document.querySelector('.rw-appmark').click()`);
        await wait(250);
        await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info'))?.click()`);
        await until(() => js(`Boolean(document.querySelector('.pw-card-tile'))`), 'Info', 4000).catch(() => {});
        await js(`document.querySelector('.pw-card-tile').click()`);
        await wait(250);
        const items = await js(`[...document.querySelectorAll('.pw-menu-item .pw-menu-label')].map((n) => n.textContent)`);
        check('encrypted: Worksheets\' Info → Protect Workbook offers Encrypt with Password beside sheet and structure protection', items.join('|') === 'Encrypt with Password|Protect Current Sheet|Protect Workbook Structure', JSON.stringify(items));
        w.webContents.invalidate();
        await wait(700);
        await capture(w, 'encrypted-sheets-info.png');
        await js(`document.documentElement.dataset.theme = 'dark'`);
        w.webContents.invalidate();
        await wait(900);
        await capture(w, 'encrypted-sheets-info-dark.png');
        await js(`document.documentElement.dataset.theme = 'light'`);
        await js(`document.querySelector('.pw-info-close')?.click()`);
      }
    } catch (err) {
      check(`encrypted: the ${app} check ran`, false, err.message);
    }
  }
}
