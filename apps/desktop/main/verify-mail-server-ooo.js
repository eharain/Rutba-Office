// Mail's Out of office at the mail server (owner, 2026-10-22,
// D-2026-10-22-01), checked on a stand-in ManageSieve server: the dialog
// finds that the account's server can send the replies and chooses it; Save
// writes the reply with its message and makes it active, the person's own
// filters still run after it; Rutba Office then sends none itself; turning
// it off gives the filters back. Run alone with
// RUTBA_VERIFY_ONLY=mail-server-ooo.

import { standInSieve } from '@rutba/mailbox/managesieve-standin';
import { SCRIPT_NAME, readScript } from '@rutba/mailbox/managesieve';

export async function verifyMailServerOoo({ open, check, until, wait, errorsIn }) {
  const sieve = standInSieve({ user: 'amira@checks.example', pass: 'check', scripts: { 'my-filters': 'require "fileinto";\r\nif header :contains "subject" "invoice" { fileinto "Invoices"; }\r\n' } });
  sieve.setActive('my-filters');
  let win = null;
  let accountId = null;
  try {
    const port = await sieve.listen();
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the accounts', 8000);
    // An account of its own, its server's filters at the stand-in.
    const made = await js(`window.rutbaOffice.mail.addAccount({ account: { email: 'amira@checks.example', name: 'Amira Khan', imap: { host: '127.0.0.1', port: 993, secure: true }, smtp: { host: '127.0.0.1', port: 587 } }, password: 'check' })`);
    accountId = made.id;
    await js(`window.rutbaOffice.mail.updateAccount({ id: ${JSON.stringify(accountId)}, patch: { sieve: { host: '127.0.0.1', port: ${port}, plain: true } } }).then(() => 1)`);
    // The window reads its accounts as it opens: opened again, it has this one.
    win.webContents.reload();
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 1`), 'the account in the window', 10000);

    const scope = `document.querySelector('.rw-dialog[aria-label="Automatic replies"]')`;
    const openDialog = async () => {
      await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Folder')?.click(), 1`);
      await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /Out of office/.test(b.textContent) && !b.disabled))`), 'Out of office', 5000);
      await js(`[...document.querySelectorAll('button')].find((b) => /Out of office/.test(b.textContent) && !b.disabled)?.click(), 1`);
      await until(() => js(`Boolean(${scope})`), 'the Automatic replies dialog', 5000);
      // This account, chosen in the dialog's own list.
      await js(`(() => { const s = ${scope}.querySelector('.ml-signature-account'); if (!s) return 0; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ${JSON.stringify(accountId)}); s.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      // Chosen, once the server has answered (not while the dialog is still asking it).
      await until(() => js(`Boolean(${scope}.querySelector('[data-where="server"] input:checked')) && !/Asking/.test(${scope}.querySelector('[data-where="server"]').textContent)`), 'the mail server chosen', 8000).catch(async (err) => {
        const seen = await js(`({ account: ${scope}.querySelector('.ml-signature-account')?.value, server: ${scope}.querySelector('[data-where="server"]')?.textContent })`);
        throw new Error(`${err.message}: ${JSON.stringify(seen)}`);
      });
    };
    const save = async () => {
      await js(`[...${scope}.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Save')?.click(), 1`);
      await until(() => js(`!${scope}`), 'the dialog to close', 8000);
    };

    // On, at the server, with its own words.
    await openDialog();
    const offered = await js(`({ server: !${scope}.querySelector('[data-where="server"] input').disabled, words: ${scope}.querySelector('[data-where="server"]').textContent, contactsOff: ${scope}.querySelector('.ml-ooo-contacts input').disabled })`);
    await js(`${scope}.querySelector('.ml-ooo-enabled input')?.click(), 1`);
    await js(`(() => { const el = ${scope}.querySelector('.ml-ooo-message'); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, 'Away at the trade fair until Monday.'); el.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await save();
    await until(() => Promise.resolve(sieve.active() === SCRIPT_NAME), 'the reply active at the server', 6000).catch(() => {});
    const script = sieve.scripts.get(SCRIPT_NAME) || '';
    const said = readScript(script);
    const kept = await js(`window.rutbaOffice.mail.accounts().then((l) => l.find((a) => a.id === ${JSON.stringify(accountId)})?.autoReply || null)`);
    check('mail: Out of office finds the account\'s mail server can send the replies, chooses it, and Save writes the reply with its words there and makes it active, the person\'s own filters still run after it',
      offered.server && offered.contactsOff && sieve.active() === SCRIPT_NAME && /vacation :days 4/.test(script) && /Away at the trade fair until Monday\./.test(script) && /include :personal "my-filters"/.test(script) && said?.previous === 'my-filters' && kept?.enabled === true && kept?.where === 'server',
      JSON.stringify({ offered, active: sieve.active(), said, kept }));

    // A message arrives: the server answers it, so Rutba Office sends nothing itself.
    const SENDER = `server-ooo-${Date.now()}@example.org`;
    const raw = [`From: Check <${SENDER}>`, 'To: Amira <amira@checks.example>', 'Subject: Hello', `Date: ${new Date().toUTCString()}`, `Message-ID: <server-ooo-${Date.now()}@example.org>`, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '', 'Hello.', ''].join('\r\n');
    await js(`window.rutbaOffice.mail.deliverTest({ accountId: ${JSON.stringify(accountId)}, folder: 'Inbox', raw: ${JSON.stringify(raw)} }).then(() => 1)`);
    await wait(1500);
    const answered = await js(`window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: 'Sent', limit: 50 }).then((r) => r.rows.filter((m) => JSON.stringify(m.to || '').includes(${JSON.stringify(SENDER)})).length, () => 0)`);
    check('mail: with the replies at the mail server, Rutba Office sends none of its own', answered === 0, `${answered} sent from here`);

    // Off: the person's filters are the active script again.
    await openDialog();
    const shown = await js(`${scope}.querySelector('.ml-ooo-message')?.value || ''`);
    await js(`${scope}.querySelector('.ml-ooo-enabled input')?.click(), 1`);
    await save();
    await until(() => Promise.resolve(sieve.active() === 'my-filters'), 'the filters back', 6000).catch(() => {});
    check('mail: opened again it shows the server\'s reply, and turning it off at the server makes the person\'s own filters active again',
      shown === 'Away at the trade fair until Monday.' && sieve.active() === 'my-filters', JSON.stringify({ shown, active: sieve.active() }));

    const complaints = await errorsIn(win);
    check('mail: the server\'s Out of office reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the server Out of office checks ran to the end', false, err?.message || String(err));
  } finally {
    if (win && accountId) await win.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(accountId)} }).then(() => 1)`).catch(() => {});
    await sieve.close().catch(() => {});
  }
}
