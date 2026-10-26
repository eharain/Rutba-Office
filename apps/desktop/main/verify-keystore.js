// Saved passwords that will not open on this computer (D-2026-10-26-09), in
// the window: against the test IMAP server, run in this process on
// 127.0.0.1, never a real account.
//
// A mail account's password is saved, and is in the profile's keystore, not
// in the clear. The keystore's key is then made unopenable, as it is on
// another computer or Windows account, and the keystore opened again as the
// next start would. Office says once, in plain words, that saved passwords
// must be typed again, and not again after; Mail asks for the account's
// password rather than failing, and with it typed the account signs in.
// Everything is put back as it was. Run alone with RUTBA_VERIFY_ONLY=keystore.

import { app } from 'electron';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';

const isSealed = (text) => typeof text === 'string' && text.startsWith('{"rutbaSealed":');

/** Every file under a folder, as one text, to look for a password in the clear. */
function allText(dir) {
  const out = [];
  const walk = (d) => {
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) walk(p);
      else out.push(fs.readFileSync(p, 'latin1'));
    }
  };
  try { walk(dir); } catch { /* none */ }
  return out.join('\n');
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyKeystore(h) {
  const { open, check, until, wait, errorsIn } = h;
  const secrets = globalThis.__rutbaCheckSecrets;
  if (!secrets?.refresh) {
    check('keystore: the check can open the keystore again', false, 'no check handle on the saved passwords');
    return;
  }
  const { imapServer } = await import('../../../packages/mailbox/test/imap-server.js');
  const address = 'keys@checks.example';
  const password = `Kestrel-${crypto.randomBytes(6).toString('hex')}`;
  const server = await imapServer({ users: { [address]: password } });
  const dir = app.getPath('userData');
  const keyFile = path.join(dir, 'keystore-key.json');
  const storeDir = path.join(dir, 'keystore');
  const backup = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-keystore-check-'));
  const before = new Set(fs.readdirSync(dir));
  let win = null;
  let accountId = null;
  let broken = false;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    const reload = () => new Promise((resolve) => { win.webContents.once('did-finish-load', resolve); win.webContents.reload(); });
    await until(() => js(`Boolean(document.querySelector('.ml-accounts'))`), 'the Mail window', 8000);
    const account = { ...server.account(), name: 'Keystore Check' };
    accountId = (await js(`window.rutbaOffice.mail.addAccount(${JSON.stringify({ account, password })})`))?.id || null;

    // Saved in the keystore, and nowhere in the clear.
    const keyText = (() => { try { return fs.readFileSync(keyFile, 'utf8'); } catch { return ''; } })();
    const rec = (() => { try { return isSealed(keyText) ? null : JSON.parse(keyText); } catch { return null; } })();
    const inClear = [allText(storeDir), keyText, (() => { try { return fs.readFileSync(path.join(dir, 'settings.json'), 'utf8'); } catch { return ''; } })()].some((t) => t.includes(password));
    const kept = secrets.get(`mail:${accountId}`) === password;
    check('keystore: a mail account\'s password is kept in the profile\'s keystore, its key wrapped by the system with the profile\'s id, and the password nowhere in the clear',
      Boolean(accountId) && kept && rec?.by === 'system' && /^[0-9a-f]{64}$/.test(rec?.uid || '') && !inClear && !fs.existsSync(path.join(dir, 'secrets.json')),
      JSON.stringify({ accountId: Boolean(accountId), kept, by: rec?.by, sealed: isSealed(keyText), inClear }));
    if (!rec?.wrapped) return;

    // The key made unopenable, as on another computer, and the keystore opened as the next start would.
    fs.cpSync(storeDir, path.join(backup, 'keystore'), { recursive: true });
    fs.copyFileSync(keyFile, path.join(backup, 'keystore-key.json'));
    broken = true;
    fs.writeFileSync(keyFile, JSON.stringify({ ...rec, wrapped: crypto.randomBytes(120).toString('base64') }, null, 2));
    const status = secrets.refresh();
    const asideNow = fs.readdirSync(dir).filter((n) => !before.has(n) && /set-aside/.test(n));
    check('keystore: a keystore whose key will not unwrap here is put aside, renamed and not deleted, and a fresh one is started',
      status?.status === 'open' && status?.aside?.why === 'other-computer' && asideNow.some((n) => n.startsWith('keystore-set-aside-')) && asideNow.some((n) => n.startsWith('keystore-key-set-aside-')) && secrets.get(`mail:${accountId}`) === null,
      JSON.stringify({ status, asideNow }));

    // Said once.
    await reload();
    await until(() => js(`Boolean(document.querySelector('.pf-keystore-notice'))`), 'the notice', 8000).catch(() => {});
    const notice = await js(`document.querySelector('.pf-keystore-notice')?.textContent || ''`);
    await js(`document.querySelector('.pf-keystore-notice .pf-keystore-ok')?.click(), 1`);
    await reload();
    await until(() => js(`Boolean(document.querySelector('.ml-accounts'))`), 'the Mail window again', 8000).catch(() => {});
    await wait(1500);
    const again = await js(`Boolean(document.querySelector('.pf-keystore-notice'))`);
    check('keystore: Office says once, in plain words, that the saved passwords could not be opened on this computer and must be typed again',
      /could not be opened on this computer/.test(notice) && /another computer/.test(notice) && !/[\u2013\u2014]/.test(notice) && again === false,
      JSON.stringify({ notice, again }));

    // Mail asks for the password, rather than failing at every fetch.
    await until(() => js(`Boolean(document.querySelector('.ml-password-again'))`), 'Mail asking for the password', 8000).catch(() => {});
    const asked = await js(`document.querySelector('.ml-password-again')?.textContent || ''`);
    const askedFor = asked.includes(address);
    const heardBefore = server.heard.length;
    await js(`(() => { const i = document.querySelector('.ml-password-again .ml-password-input'); const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(i, ${JSON.stringify(password)}); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(200);
    await js(`document.querySelector('.ml-password-again .ml-password-save')?.click(), 1`);
    const closed = await until(() => js(`!document.querySelector('.ml-password-again')`), 'the password saved', 8000).then(() => true).catch(() => false);
    const signedIn = await until(() => server.heard.slice(heardBefore).some((l) => /LOGIN "?keys@checks\.example/i.test(l)), 'a sign-in at the server', 15000).then(() => true).catch(() => false);
    const hasPassword = (await js(`window.rutbaOffice.mail.accounts()`)).find((a) => a.id === accountId)?.hasPassword === true;
    check('keystore: Mail asks for the account\'s password as it opens, and with it typed again the account is kept in the new keystore and signs in at the server',
      askedFor && closed && hasPassword && signedIn && secrets.get(`mail:${accountId}`) === password,
      JSON.stringify({ askedFor, closed, hasPassword, signedIn, heard: signedIn ? undefined : server.heard.slice(heardBefore, heardBefore + 6) }));
    const complaints = [...(await errorsIn(win))];
    check('keystore: no errors in the Mail window', complaints.length === 0, complaints.join(' | ') || 'none');
  } finally {
    if (accountId && win && !win.isDestroyed()) await win.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)}, keepMail: false })`).catch(() => {});
    // Put back as it was: the keystore and its key from before, what the check put aside gone.
    if (broken) {
      try {
        fs.rmSync(storeDir, { recursive: true, force: true });
        fs.cpSync(path.join(backup, 'keystore'), storeDir, { recursive: true });
        fs.copyFileSync(path.join(backup, 'keystore-key.json'), keyFile);
        for (const n of fs.readdirSync(dir)) if (!before.has(n) && /set-aside/.test(n)) fs.rmSync(path.join(dir, n), { recursive: true, force: true });
        secrets.refresh();
        secrets.forgetNotice?.();
        if (accountId) secrets.delete(`mail:${accountId}`);
      } catch (err) {
        check('keystore: the keystore put back as it was', false, String(err?.message || err));
      }
    }
    fs.rmSync(backup, { recursive: true, force: true });
    await server.close().catch(() => {});
  }
}
