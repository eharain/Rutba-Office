// Mail's account manager, Add account's counterpart (owner, 2026-10-22,
// D-2026-10-22-09: "add email account should have counter remove account on
// an account manager"): Manage accounts beside Add account lists every
// account with its Signature, Out of office and Remove; Remove asks first and
// takes the account off; an account's own menu in the side panel offers
// Remove account too. Run alone with RUTBA_VERIFY_ONLY=mail-accounts-manager.

export async function verifyMailAccountsManager({ open, check, until, wait, errorsIn }) {
  let win = null;
  const made = [];
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the accounts', 8000);
    for (const [email, name] of [['ana@checks.example', 'Ana Silva'], ['raj@checks.example', 'Raj Patel']]) {
      made.push((await js(`window.rutbaOffice.mail.addAccount({ account: { email: ${JSON.stringify(email)}, name: ${JSON.stringify(name)}, imap: { host: 'mail.checks.example', port: 993, secure: true }, smtp: { host: 'mail.checks.example', port: 587 } }, password: 'check' })`)).id);
    }
    win.webContents.reload();
    await until(() => js(`[...document.querySelectorAll('.ml-accounts button')].some((b) => /Raj Patel/.test(b.textContent))`), 'the two accounts', 10000);

    // The manager, beside Add account.
    await js(`[...document.querySelectorAll('.ml-add-account button')].find((b) => /Manage accounts/.test(b.textContent)).click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-account-list'))`), 'the account manager', 5000);
    const listed = await js(`[...document.querySelectorAll('.ml-account-row')].map((r) => ({ email: r.dataset.account, remove: Boolean(r.querySelector('.ml-account-remove')), signature: Boolean(r.querySelector('.ml-account-signature')) }))`);
    // Remove, asked first: Remove.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 0 } });
    await js(`document.querySelector('.ml-account-row[data-account="ana@checks.example"] .ml-account-remove').click(), 1`);
    await until(() => js(`!document.querySelector('.ml-account-row[data-account="ana@checks.example"]')`), 'Ana removed', 6000).catch(() => {});
    const afterRemove = await js(`({ rows: [...document.querySelectorAll('.ml-account-row')].map((r) => r.dataset.account), side: [...document.querySelectorAll('.ml-accounts button')].map((b) => b.textContent.trim()) })`);
    await js(`[...document.querySelectorAll('.rw-dialog button')].find((b) => b.textContent.trim() === 'Done')?.click(), 1`);
    check('mail: Manage accounts beside Add account lists every account with its Signature and Remove, and Remove, once agreed, takes the account off the list and the side panel',
      listed.some((a) => a.email === 'ana@checks.example' && a.remove && a.signature) && listed.some((a) => a.email === 'raj@checks.example')
        && !afterRemove.rows.includes('ana@checks.example') && afterRemove.rows.includes('raj@checks.example') && !afterRemove.side.some((s) => /Ana Silva/.test(s)),
      JSON.stringify({ listed, afterRemove }));

    // An account's own menu: Remove account; Cancel leaves it.
    await js(`[...document.querySelectorAll('.ml-accounts button')].find((b) => /Raj Patel/.test(b.textContent)).dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 100, clientY: 200 })), 1`);
    await until(() => js(`[...document.querySelectorAll('[role="menuitem"]')].some((n) => /Remove account/.test(n.textContent))`), 'the account menu', 4000);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'message', answer: { response: 2 } });
    await js(`[...document.querySelectorAll('[role="menuitem"]')].find((n) => /Remove account/.test(n.textContent)).click(), 1`);
    await wait(800);
    const kept = await js(`[...document.querySelectorAll('.ml-accounts button')].some((b) => /Raj Patel/.test(b.textContent))`);
    check('mail: an account\'s own menu offers Remove account, and Cancel leaves the account where it is', kept, String(kept));
    const complaints = await errorsIn(win);
    check('mail: the account manager reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the account manager checks ran to the end', false, err?.message || String(err));
  } finally {
    for (const id of made) await win?.webContents.executeJavaScript(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(id)} }).then(() => 1, () => 1)`).catch(() => {});
  }
}
