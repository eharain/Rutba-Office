// Passwords checked and set without holding the other windows.
//
// A password is hashed 100,000 times or more. The hashing ran on the
// document service's one thread, so while Unprotect Workbook checked a
// password every other window's typing waited for it (the window checks
// once logged the main process 5.2 s late). It runs on the password
// threads now (password-host.js). Here a workbook whose structure is
// locked with a password spun 500,000 times (a count Excel reads, and long
// enough to type through) is unprotected with a wrong password while a
// document in another window is typed into: each key is on its page within
// 200 ms, the dialog's OK shows it is working, and the refusal still comes
// in Excel's words. The same while Restrict Editing's Stop Protection
// checks a wrong password (Word's words), while Encrypt with Password's key
// is spun and the workbook saved encrypted, and while the encrypted file
// is opened with a wrong password. Run alone with RUTBA_VERIFY_ONLY=pwfast.
import fs from 'node:fs';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { buildXlsx, buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { isEncryptedPackage } from '@rutba/office-formats/crypt';
import { wordLegacyHex } from '@rutba/ooxml/protection';
import { createPasswordHost } from './password-host.js';

const PASSWORD = 'Budget2026';
const SPIN = 500000;
const ENCRYPT = 'Vault 7 Open';
const RESTRICT = 'Stamp 9';
/** How long a key may take to reach the page while a password is hashed elsewhere. */
const KEY_MS = 200;

/**
 * @param {object} h the harness: open, check, until, wait, capture
 * @param {{ dir: string }} where the fixtures are written
 */
export async function verifyPwFast(h, { dir }) {
  const { open, check, until, wait, capture } = h;
  const book = path.join(dir, 'pwfast.xlsx');
  const note = path.join(dir, 'pwfast-note.docx');
  const restricted = path.join(dir, 'pwfast-restricted.docx');
  // The fixture's hash, made on a password thread of the main process's own.
  const host = createPasswordHost({ threads: 1 });
  try {
    const salt = randomBytes(16).toString('base64');
    const hash = await host.hash({ password: PASSWORD, salt, spinCount: SPIN, algorithm: 'SHA-512' });
    const pkg = OoxmlPackage.read(buildXlsx({ sheets: [{ name: 'Budget', rows: [['Item', 'Amount'], ['Rent', 1200], ['Ink', 40]] }, { name: 'Notes', rows: [['n']] }] }));
    pkg.write_('xl/workbook.xml', pkg.text('xl/workbook.xml').replace('<sheets>', `<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="${hash}" workbookSaltValue="${salt}" workbookSpinCount="${SPIN}" lockStructure="1"/><sheets>`));
    fs.writeFileSync(book, pkg.write());
    fs.writeFileSync(note, buildDocx({ paragraphs: ['Minutes of the meeting', 'Typed while a password is checked: '], styles: true }));
    // A document as Word leaves it under Restrict Editing: read only, a password (Word's key, hashed 100,000 times).
    const wsalt = randomBytes(16).toString('base64');
    const whash = await host.hash({ password: wordLegacyHex(RESTRICT), salt: wsalt, spinCount: 100000, algorithm: 'SHA-512' });
    const dpkg = OoxmlPackage.read(buildDocx({ paragraphs: ['Service agreement: the terms below are fixed.', 'Standard clauses follow.'], styles: true }));
    const settings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/>'
      + `<w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="${whash}" w:salt="${wsalt}"/>`
      + '<w:defaultTabStop w:val="720"/></w:settings>';
    if (dpkg.has('word/settings.xml')) dpkg.write_('word/settings.xml', settings);
    else {
      dpkg.addPart('word/settings.xml', Buffer.from(settings, 'utf8'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
      dpkg.addRelationshipTo('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
    }
    fs.writeFileSync(restricted, dpkg.write());
  } catch (err) {
    await host.stop();
    return check('pwfast: the fixtures were made', false, err.message);
  }

  try {
    const sheet = await open('sheets', book);
    const word = await open('word', note);
    const sjs = (code) => sheet.webContents.executeJavaScript(code);
    const wjs = (code) => word.webContents.executeJavaScript(code);
    const wc = word.webContents;
    // A read at once, not after the window's requests are answered (the harness's
    // executeJavaScript waits for those): what the window shows while one is pending.
    const now = (win, code) => win.webContents.mainFrame.executeJavaScript(code);
    const shot = async (win, name) => {
      win.webContents.invalidate();
      await wait(700);
      await capture(win, name);
    };
    const setField = (js, selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      setter.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'set';
    })()`);
    const click = (js, selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'none'; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);

    await until(() => sjs(`Boolean(document.querySelector('.sh-cell[data-ref="B3"]'))`), 'the workbook', 10000);
    await until(() => wjs(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the document', 10000);

    // The caret at the end of the document's second paragraph, by a real click.
    const caret = async () => {
      const at = await wjs(`(() => { const b = document.querySelector('.wd-page [data-block="1"]'); if (!b) return null; b.scrollIntoView({ block: 'center' }); const q = b.getBoundingClientRect(); const r = document.createRange(); r.selectNodeContents(b); const t = r.getBoundingClientRect(); return { x: Math.round(t.right + 6), y: Math.round(q.top + q.height / 2) }; })()`);
      if (!at) return false;
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      await wait(350);
      return true;
    };

    /**
     * The document window times its own keys: when each "q" went down and
     * when the page, which shows only what the engine sends back, first
     * held it.
     */
    const probe = () => wjs(`(() => {
      const count = () => ((document.querySelector('.wd-page')?.textContent || '').match(/q/g) || []).length;
      window.__pw = { base: count(), downs: [], seen: [] };
      if (!window.__pwHooked) {
        window.__pwHooked = true;
        window.addEventListener('keydown', (e) => { if (e.key === 'q' && window.__pw) window.__pw.downs.push(performance.now()); }, true);
        new MutationObserver(() => {
          const p = window.__pw;
          if (!p) return;
          const n = count() - p.base;
          while (p.seen.length < n) p.seen.push(performance.now());
        }).observe(document.body, { subtree: true, childList: true, characterData: true });
      }
      return window.__pw.base;
    })()`);
    const key = () => {
      wc.sendInputEvent({ type: 'keyDown', keyCode: 'q' });
      wc.sendInputEvent({ type: 'char', keyCode: 'q' });
      wc.sendInputEvent({ type: 'keyUp', keyCode: 'q' });
    };
    /** Keys typed every `gap` ms while `busy()` holds (at least `least`, at most `most`); then each key's time to the page. */
    const typeThrough = async (busy, { gap = 100, least = 2, most = 30, shotAt = null } = {}) => {
      let typed = 0;
      let during = 0;
      let shooting = null;
      while (typed < most) {
        const still = await busy();
        if (!still && typed >= least) break;
        key();
        typed++;
        if (still) during++;
        // The capture of the working dialog goes on beside the keys, not in their way.
        if (shotAt && typed === 2) shooting = shotAt();
        await wait(gap);
      }
      await shooting;
      await until(() => wjs(`window.__pw.seen.length >= ${typed}`), 'every key on the page', 8000).catch(() => {});
      const p = await wjs(`({ downs: window.__pw.downs, seen: window.__pw.seen })`);
      const lags = p.downs.map((d, i) => (p.seen[i] == null ? null : Math.round(p.seen[i] - d)));
      return { typed, during, lags, worst: lags.some((l) => l == null) ? null : Math.max(...lags) };
    };
    const quick = (r) => r.during >= 2 && r.lags.length === r.typed && r.worst != null && r.worst < KEY_MS;

    // ── Unprotect Workbook with a wrong password, the document typed into meanwhile.
    await sjs(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === 'Review')?.click(); return 1; })()`);
    const pressed = await sjs(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || n.textContent || '').trim().startsWith('Unprotect Workbook') && !n.disabled);
      if (!b) return 'no Unprotect Workbook';
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const asked = await until(() => sjs(`Boolean(document.querySelector('.sh-password'))`), 'the password dialog', 5000).catch(() => false);
    await setField(sjs, '.sh-password', 'budget2026');
    await caret();
    await probe();
    let busyLook = null;
    const sheetBusy = () => now(sheet, `document.querySelector('.sh-password-ok')?.getAttribute('aria-busy') === 'true'`);
    await click(sjs, '.sh-password-ok');
    const unprotect = await typeThrough(sheetBusy, {
      shotAt: async () => {
        busyLook = await now(sheet, `(() => { const b = document.querySelector('.sh-password-ok'); return b ? { busy: b.getAttribute('aria-busy'), disabled: b.disabled, label: b.textContent.trim(), spinner: Boolean(b.querySelector('.rw-spinner')), cancel: [...document.querySelectorAll('.rw-dialog-foot .rw-btn')].find((x) => x.textContent.trim() === 'Cancel')?.disabled } : null; })()`);
        await shot(sheet, 'pwfast-unprotect-checking.png');
      },
    });
    await until(() => sjs(`Boolean(document.querySelector('.sh-password-error'))`), 'the refusal', 20000).catch(() => false);
    const refusal = await sjs(`document.querySelector('.sh-password-error')?.textContent || ''`);
    await shot(sheet, 'pwfast-unprotect-refused.png');
    await shot(word, 'pwfast-document-typed.png');
    check(`pwfast: while Unprotect Workbook checks a wrong password (${SPIN.toLocaleString('en-GB')} rounds), another window's typed keys reach its page within ${KEY_MS} ms, and the refusal comes in Excel's words`,
      pressed === 'clicked' && asked === true && quick(unprotect) && /^The password you supplied is not correct\./.test(refusal),
      `${pressed}; ${unprotect.typed} keys, ${unprotect.during} while checking, ms to the page ${JSON.stringify(unprotect.lags)}; refusal ${JSON.stringify(refusal)}`);
    check('pwfast: while the password is checked the dialog says so: OK held with a spinner and "Checking…", Cancel held',
      Boolean(busyLook) && busyLook.busy === 'true' && busyLook.disabled === true && busyLook.spinner === true && busyLook.label === 'Checking…' && busyLook.cancel === true,
      JSON.stringify(busyLook));

    // The right password takes it off, as before.
    await setField(sjs, '.sh-password', PASSWORD);
    await click(sjs, '.sh-password-ok');
    const off = await until(() => sjs(`!document.querySelector('.sh-password') && Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Protect Workbook'))`), 'the workbook unprotected', 20000).catch(() => false);
    check('pwfast: the right password still unprotects the workbook', off === true);

    // ── Restrict Editing's Stop Protection with a wrong password, the other document typed into.
    const locked = await open('word', restricted);
    const ljs = (code) => locked.webContents.executeJavaScript(code);
    await until(() => ljs(`Boolean(document.querySelector('.wd-restrict-stop'))`), 'the Restrict Editing pane', 10000).catch(() => {});
    await click(ljs, '.wd-restrict-stop');
    const unprotectAsked = await until(() => ljs(`Boolean(document.querySelector('.wd-unprotect-password'))`), 'Unprotect Document', 5000).catch(() => false);
    await setField(ljs, '.wd-unprotect-password', 'stamp 9');
    await caret();
    await probe();
    let wordLook = null;
    await click(ljs, '.wd-unprotect-ok');
    const stopping = await typeThrough(() => now(locked, `document.querySelector('.wd-unprotect-ok')?.getAttribute('aria-busy') === 'true'`), {
      shotAt: async () => {
        wordLook = await now(locked, `(() => { const b = document.querySelector('.wd-unprotect-ok'); return b ? { busy: b.getAttribute('aria-busy'), disabled: b.disabled, label: b.textContent.trim(), spinner: Boolean(b.querySelector('.rw-spinner')), cancel: document.querySelector('.wd-unprotect-cancel')?.disabled } : null; })()`);
        await shot(locked, 'pwfast-restrict-checking.png');
      },
    });
    const notRight = await until(() => ljs(`/That password is not right, so the protection stays on\\./.test(document.querySelector('.wd-unprotect-error')?.textContent || '')`), 'Word\'s refusal', 20000).catch(() => false);
    check(`pwfast: while Stop Protection checks a wrong password, another window's typed keys reach its page within ${KEY_MS} ms; the dialog's OK is held with a spinner, and the refusal comes in Word's words`,
      unprotectAsked === true && quick(stopping) && notRight === true && wordLook?.busy === 'true' && wordLook.spinner === true && wordLook.label === 'Checking…' && wordLook.cancel === true,
      `${stopping.typed} keys, ${stopping.during} while checking, ms to the page ${JSON.stringify(stopping.lags)}; dialog ${JSON.stringify(wordLook)}; refused ${notRight}`);
    await setField(ljs, '.wd-unprotect-password', RESTRICT);
    await click(ljs, '.wd-unprotect-ok');
    const stopped = await until(() => ljs(`!document.querySelector('.wd-unprotect-password') && !document.querySelector('.wd-restrict-stop')`), 'the protection off', 20000).catch(() => false);
    check('pwfast: the right password still stops the protection', stopped === true);

    // ── Encrypt with Password, then Save: the key spun while the document is typed into.
    await click(sjs, '.rw-appmark');
    await until(() => sjs(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info')))`), 'the app menu', 4000).catch(() => {});
    await sjs(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim().startsWith('Info'))?.click()`);
    await until(() => sjs(`Boolean(document.querySelector('.pw-card-tile'))`), 'Info', 5000).catch(() => {});
    await click(sjs, '.pw-card-tile');
    await until(() => sjs(`Boolean(document.querySelector('.pw-menu-item[data-item="encrypt"]'))`), 'the Protect menu', 4000).catch(() => {});
    await click(sjs, '.pw-menu-item[data-item="encrypt"]');
    const setOpen = await until(() => sjs(`Boolean(document.querySelector('.pw-set-password'))`), 'Encrypt Document', 5000).catch(() => false);
    await setField(sjs, '.pw-set-password', ENCRYPT);
    await setField(sjs, '.pw-set-again', ENCRYPT);
    await until(() => sjs(`document.querySelector('.pw-set-ok')?.disabled === false`), 'OK to wake', 3000).catch(() => {});
    await caret();
    await probe();
    let saved = false;
    let saving = null;
    const encrypt = await typeThrough(async () => {
      // OK, then Info closed and Save pressed, as a person would, while the keys go on.
      if (!saving) {
        saving = (async () => {
          await click(sjs, '.pw-set-ok');
          await until(() => sjs(`Boolean(document.querySelector('.pw-info-close'))`), 'Info again', 5000).catch(() => {});
          await click(sjs, '.pw-info-close');
          await sjs(`[...document.querySelectorAll('.rw-btn')].find((b) => (b.dataset.tip || b.title || '').startsWith('Save'))?.click()`);
          // While keys are timed, only the file's first eight bytes are read (an encrypted
          // workbook is a compound file): reading and parsing the whole of it here, on the
          // main process the keys pass through, held the last key past its budget.
          saved = await until(async () => {
            const fh = await fs.promises.open(book, 'r');
            try { const head = Buffer.alloc(8); await fh.read(head, 0, 8, 0); return head.readUInt32LE(0) === 0xe011cfd0; } finally { await fh.close(); }
          }, 'the encrypted save', 20000).catch(() => false);
        })();
      }
      return !saved;
    }, { gap: 60 });
    await saving;
    // The keys done: now the whole file, checked as an encrypted package.
    saved = saved === true && isEncryptedPackage(fs.readFileSync(book));
    let opens = false;
    try {
      opens = (await host.decrypt(fs.readFileSync(book), ENCRYPT)).bytes.length > 0;
    } catch {
      opens = false;
    }
    check(`pwfast: while Encrypt with Password's key is spun and the workbook saved encrypted, another window's typed keys reach its page within ${KEY_MS} ms`,
      setOpen === true && saved === true && opens && quick(encrypt),
      `${encrypt.typed} keys, ${encrypt.during} while encrypting, ms to the page ${JSON.stringify(encrypt.lags)}; saved ${saved}, opens with its password ${opens}`);

    // ── The encrypted workbook opened with a wrong password, the document typed into meanwhile.
    const again = await open('sheets', book);
    const ajs = (code) => again.webContents.executeJavaScript(code);
    await until(() => ajs(`Boolean(document.querySelector('.pw-open-password'))`), 'the Password dialog', 12000).catch(() => {});
    await setField(ajs, '.pw-open-password', 'vault 7 open');
    await caret();
    await probe();
    await click(ajs, '.pw-open-ok');
    const opening = await typeThrough(() => now(again, `document.querySelector('.pw-open-ok')?.disabled === true && !document.querySelector('.pw-error')`), { gap: 25, shotAt: () => shot(again, 'pwfast-open-checking.png') });
    const wrongSaid = await until(() => ajs(`Boolean(document.querySelector('.pw-error'))`), 'the wrong password said', 15000).catch(() => false);
    check(`pwfast: while an encrypted workbook's wrong password is tried, another window's typed keys reach its page within ${KEY_MS} ms, and the Password dialog says it is not right`,
      wrongSaid === true && quick(opening),
      `${opening.typed} keys, ${opening.during} while trying, ms to the page ${JSON.stringify(opening.lags)}`);
  } catch (err) {
    check('pwfast: the checks ran', false, err.message);
  } finally {
    await host.stop();
  }
}
