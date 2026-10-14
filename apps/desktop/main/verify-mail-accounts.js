// Several mail accounts added at once.
//
// From Add account, Several accounts opens a box that takes rows of
// addresses and passwords however they were written. Pasted into it: a
// comment, a Gmail address and a password with spaces in it, a Yahoo one
// separated by a comma with commas in its password, an Outlook.com address
// with no password, a Fastmail one separated by a colon, a line that is not
// an address, the Gmail address again, and an address at a domain nothing
// knows. The box is read the moment the rows are pasted and cleared, so no
// password stays on screen (nor in a capture). Each row says at once what it
// is; Add sets them up together (in a
// check run from the provider table alone, never the network): three added
// as their providers, the Outlook.com row offered Microsoft's own sign-in,
// the unknown domain offered setting up by hand, and the bad and repeated
// lines left out. A table with a header (a spreadsheet's CSV) adds an account
// with the servers its columns name. The line that is not an address is
// shown by its number only, never its text. A domain whose records name
// servers somewhere else (staged for the check, never the network) waits
// for the person to agree to them by name, in the list and in Add account,
// before its password goes there; a line typed into the box goes into the
// list when Enter ends it. Then the box in Urdu, right to left. Every
// account the check added is removed. Run alone with
// RUTBA_VERIFY_ONLY=mail-accounts.

const ROWS = [
  '# the office, set up on Monday',
  'ali.bulk@gmail.com  app pass word one',
  'bo.bulk@yahoo.com,secret,with,commas',
  'cy.bulk@outlook.com',
  'dee.bulk@fastmail.com:colon:pass',
  'not an address',
  'ali.bulk@gmail.com again',
  'ed.bulk@nowhere.example password',
].join('\n');

/**
 * @param {object} h the harness: open, check, until, wait, capture, errorsIn
 */
export async function verifyMailAccounts(h) {
  const { open, check, until, wait, capture, errorsIn } = h;
  let win = null;
  let js = null;
  let language = null;
  try {
    win = await open('mail');
    js = (code) => win.webContents.executeJavaScript(code);
    const shot = async (w, name) => {
      w.webContents.invalidate();
      await wait(700);
      await capture(w, name);
    };
    const waitFor = async (fn, what, ms) => {
      let got = null;
      await until(async () => {
        got = await fn();
        return Boolean(got);
      }, what, ms).catch(() => {});
      return got || null;
    };
    language = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
    // Pasted as a person pastes: a paste event carrying the text.
    const paste = (value, w = win) => w.webContents.executeJavaScript(`(() => { const el = document.querySelector('.ml-bulk-text'); if (!el) return false; el.focus(); const dt = new DataTransfer(); dt.setData('text/plain', ${JSON.stringify(value)}); el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); return true; })()`);
    const openBulk = async (w) => {
      const run = (code) => w.webContents.executeJavaScript(code);
      await until(() => run(`Boolean([...document.querySelectorAll('button')].find((b) => /Add account|اکاؤنٹ شامل کریں/.test(b.textContent)))`), 'an Add account button', 10000);
      await run(`[...document.querySelectorAll('.ml-accounts button, button')].find((b) => /^(Add account|اکاؤنٹ شامل کریں)$/.test(b.textContent.trim()))?.click(), 1`);
      await until(() => run(`Boolean(document.querySelector('.ml-several-accounts'))`), 'the account dialog', 6000).catch(() => {});
      await run(`document.querySelector('.ml-several-accounts')?.click(), 1`);
      return until(() => run(`Boolean(document.querySelector('.ml-bulk-text'))`), 'the Several accounts box', 6000).then(() => true, () => false);
    };

    const opened = await openBulk(win);
    await paste(ROWS);
    const preview = await waitFor(async () => {
      const s = await js(`(() => ({ rows: [...document.querySelectorAll('.ml-bulk-row')].map((r) => [r.dataset.email, r.dataset.outcome]), add: document.querySelector('.ml-bulk-add')?.textContent.trim() || '', shown: document.querySelector('.rw-dialog')?.textContent || '', box: document.querySelector('.ml-bulk-text')?.value ?? null }))()`);
      return s.rows.length === 7 ? s : false;
    }, 'the rows read', 6000);
    const want = [
      ['ali.bulk@gmail.com', 'ready-to-add'],
      ['bo.bulk@yahoo.com', 'ready-to-add'],
      ['cy.bulk@outlook.com', 'no-password'],
      ['dee.bulk@fastmail.com', 'ready-to-add'],
      ['', 'unreadable'],
      ['ali.bulk@gmail.com', 'repeated'],
      ['ed.bulk@nowhere.example', 'ready-to-add'],
    ];
    check('mail accounts: Several accounts reads pasted rows of addresses and passwords at once, whatever separates them, clears the box, says what each row is, and shows no password anywhere',
      opened && Boolean(preview) && JSON.stringify(preview.rows) === JSON.stringify(want) && preview.add === 'Add 5 accounts' && preview.box === '' && !/app pass word|secret,with|colon:pass|not an address/.test(preview.shown) && /Line 6 could not be read: it does not start with an email address\./.test(preview.shown),
      JSON.stringify(preview));
    await shot(win, 'mail-accounts-1-rows.png');

    await js(`document.querySelector('.ml-bulk-add')?.click(), 1`);
    const result = await waitFor(async () => {
      const s = await js(`(() => ({ done: Boolean(document.querySelector('.ml-bulk-done')), rows: [...document.querySelectorAll('.ml-bulk-row')].map((r) => ({ email: r.dataset.email, outcome: r.dataset.outcome, state: r.querySelector('.ml-bulk-state')?.textContent || '', signIn: Boolean(r.querySelector('.ml-bulk-signin')), byHand: Boolean(r.querySelector('.ml-bulk-byhand')) })), lead: document.querySelector('.ml-bulk-lead')?.textContent || '' }))()`);
      return s.done ? s : false;
    }, 'the accounts set up', 20000);
    const row = (email, outcome) => result?.rows.find((r) => r.email === email && r.outcome === outcome);
    const accounts = await js(`window.rutbaOffice.mail.accounts()`);
    const of = (email) => accounts.find((a) => a.email === email);
    check('mail accounts: Add sets them up together: Gmail, Yahoo and Fastmail added with their own servers, Outlook.com offered Microsoft\'s sign-in, the unknown domain offered setting up by hand, the bad and repeated lines left out',
      Boolean(result) && /^Added 3 accounts\./.test(result.lead)
        && /Added: Google/.test(row('ali.bulk@gmail.com', 'added')?.state || '') && Boolean(row('bo.bulk@yahoo.com', 'added')) && Boolean(row('dee.bulk@fastmail.com', 'added'))
        && row('cy.bulk@outlook.com', 'sign-in')?.signIn && row('ed.bulk@nowhere.example', 'not-found')?.byHand && Boolean(row('', 'unreadable')) && Boolean(row('ali.bulk@gmail.com', 'repeated'))
        && of('ali.bulk@gmail.com')?.imap?.host === 'imap.gmail.com' && of('bo.bulk@yahoo.com')?.imap?.host === 'imap.mail.yahoo.com' && of('dee.bulk@fastmail.com')?.smtp?.host === 'smtp.fastmail.com' && !of('cy.bulk@outlook.com') && !of('ed.bulk@nowhere.example')
        && ['ali.bulk@gmail.com', 'bo.bulk@yahoo.com', 'dee.bulk@fastmail.com'].every((e) => of(e)?.hasPassword),
      JSON.stringify({ result, accounts: accounts.filter((a) => /\.bulk@/.test(a.email)).map((a) => [a.email, a.imap?.host, a.hasPassword]) }));
    await shot(win, 'mail-accounts-2-added.png');

    // Sign in from a row: the Add account dialog for that address.
    await js(`[...document.querySelectorAll('.ml-bulk-row')].find((r) => r.dataset.email === 'cy.bulk@outlook.com')?.querySelector('.ml-bulk-signin')?.click(), 1`);
    const signIn = await waitFor(() => js(`(() => { const i = document.querySelector('#ml-address'); const card = document.querySelector('.ml-found-item'); return i && i.value === 'cy.bulk@outlook.com' && card ? { offered: true, card: card.textContent.trim().slice(0, 60) } : null; })()`), 'Add account for the Outlook.com row, its sign-in offered', 8000);
    check('mail accounts: a row that signs in with its provider\'s page opens Add account for that address, the sign-in offered', Boolean(signIn) && signIn.offered, JSON.stringify(signIn));
    await js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`);
    const sidebar = await waitFor(() => js(`(() => { const names = [...document.querySelectorAll('.ml-accounts button')].map((b) => b.textContent.trim()); return ['ali.bulk@gmail.com', 'bo.bulk@yahoo.com', 'dee.bulk@fastmail.com'].every((e) => names.some((n) => n.includes(e))) ? names : null; })()`), 'the accounts in the sidebar', 6000);
    check('mail accounts: the accounts added are in the sidebar', Boolean(sidebar), JSON.stringify(sidebar));

    // A spreadsheet's table, its servers named in its columns.
    const table = await js(`window.rutbaOffice.mail.addAccounts({ text: 'Email,Password,Name,IMAP server,IMAP port,SMTP server,SMTP port\\r\\nzed.bulk@checks.example,"p,w",Zed Table,mail.checks.example,993,mail.checks.example,587\\r\\n' })`);
    const zed = (await js(`window.rutbaOffice.mail.accounts()`)).find((a) => a.email === 'zed.bulk@checks.example');
    check('mail accounts: a table with a header adds an account with the servers its columns name',
      table.results[0]?.outcome === 'added' && zed?.name === 'Zed Table' && zed?.imap?.host === 'mail.checks.example' && zed?.imap?.port === 993 && zed?.smtp?.port === 587,
      JSON.stringify({ table, zed: zed && { name: zed.name, imap: zed.imap, smtp: zed.smtp } }));

    // A domain whose records name servers somewhere else: staged for the check, so nothing is looked up.
    const elsewhere = { provider: null, oauth: null, imap: { host: 'mail.collector.example', port: 993, secure: true, foreign: true }, smtp: { host: 'smtp.collector.example', port: 465, secure: true, foreign: true }, foreign: ['mail.collector.example', 'smtp.collector.example'], candidates: [] };
    await js(`window.rutbaOffice.mail.discoveryForTest({ domain: 'elsewhere.checks.example', found: ${JSON.stringify(elsewhere)} })`);
    await js(`document.querySelector('.ml-bulk-done')?.click(), 1`);
    await until(() => js(`!document.querySelector('.ml-bulk-text') && !document.querySelector('.ml-bulk-done')`), 'the list closed', 4000).catch(() => {});
    await openBulk(win);
    // Typed, not pasted: Enter ends the line and it goes into the list.
    await js(`(() => { const el = document.querySelector('.ml-bulk-text'); el.focus(); Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, 'fay.bulk@elsewhere.checks.example  typed pass'); el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })); return 1; })()`);
    const typed = await waitFor(() => js(`(() => { const rows = [...document.querySelectorAll('.ml-bulk-row')].map((r) => r.dataset.email); const box = document.querySelector('.ml-bulk-text')?.value; return rows.includes('fay.bulk@elsewhere.checks.example') ? { rows, box } : null; })()`), 'the typed line in the list', 4000);
    check('mail accounts: a line typed into the box goes into the list when Enter ends it, and the box is cleared', Boolean(typed) && typed.box === '', JSON.stringify(typed));
    await js(`document.querySelector('.ml-bulk-add')?.click(), 1`);
    const asks = await waitFor(() => js(`(() => { const r = [...document.querySelectorAll('.ml-bulk-row')].find((x) => x.dataset.email === 'fay.bulk@elsewhere.checks.example'); return r && r.dataset.outcome === 'confirm' ? { state: r.querySelector('.ml-bulk-state')?.textContent || '', agree: Boolean(r.querySelector('.ml-bulk-agree')) } : null; })()`), 'the row waiting for agreement', 10000);
    const notYet = (await js(`window.rutbaOffice.mail.accounts()`)).some((a) => a.email === 'fay.bulk@elsewhere.checks.example');
    check('mail accounts: a row whose domain names servers elsewhere waits, naming them, and is not added until the person agrees',
      Boolean(asks) && asks.agree && /Its mail servers are not at elsewhere\.checks\.example: mail\.collector\.example, smtp\.collector\.example\. The password goes to them only if you agree\./.test(asks.state) && !notYet,
      JSON.stringify({ asks, notYet }));
    await shot(win, 'mail-accounts-4-agree-row.png');
    await js(`[...document.querySelectorAll('.ml-bulk-row')].find((x) => x.dataset.email === 'fay.bulk@elsewhere.checks.example')?.querySelector('.ml-bulk-agree')?.click(), 1`);
    const agreed = await waitFor(() => js(`(() => { const r = [...document.querySelectorAll('.ml-bulk-row')].find((x) => x.dataset.email === 'fay.bulk@elsewhere.checks.example'); return r && r.dataset.outcome === 'added' ? r.dataset.outcome : null; })()`), 'the row added once agreed', 10000);
    const fay = (await js(`window.rutbaOffice.mail.accounts()`)).find((a) => a.email === 'fay.bulk@elsewhere.checks.example');
    check('mail accounts: once agreed to, the row is set up with those servers', agreed === 'added' && fay?.imap?.host === 'mail.collector.example' && fay?.hasPassword, JSON.stringify({ agreed, imap: fay?.imap }));
    await js(`document.querySelector('.ml-bulk-done')?.click(), 1`);
    await until(() => js(`!document.querySelector('.ml-bulk-done')`), 'the list closed', 4000).catch(() => {});

    // The same in Add account: the servers named, Use these servers, and only then the account.
    await until(() => js(`Boolean([...document.querySelectorAll('button')].find((b) => /^Add account$/.test(b.textContent.trim())))`), 'an Add account button', 10000).catch(() => {});
    await js(`[...document.querySelectorAll('button')].find((b) => /^Add account$/.test(b.textContent.trim()))?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('#ml-address'))`), 'the account dialog', 6000).catch(() => {});
    const typeInto = (selector, value) => js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(value)}); el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
    await typeInto('#ml-address', 'gus.bulk@elsewhere.checks.example');
    await typeInto('.rw-dialog input[type="password"]', 'dialog pass');
    await until(() => js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Add account')?.disabled === false`), 'Add to wake', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-dialog-foot button')].find((b) => b.textContent.trim() === 'Add account')?.click(), 1`);
    const panel = await waitFor(() => js(`(() => { const p = document.querySelector('.ml-agree'); return p ? { text: p.textContent, hosts: [...p.querySelectorAll('.ml-agree-list b')].map((b) => b.textContent) } : null; })()`), 'the servers to agree to', 8000);
    const gusBefore = (await js(`window.rutbaOffice.mail.accounts()`)).some((a) => a.email === 'gus.bulk@elsewhere.checks.example');
    check('mail accounts: Add account names servers outside the address\'s domain and asks before the password goes to them',
      Boolean(panel) && /elsewhere\.checks\.example says its mail is on servers somewhere else/.test(panel.text) && JSON.stringify(panel.hosts) === JSON.stringify(['mail.collector.example', 'smtp.collector.example']) && !gusBefore,
      JSON.stringify({ panel, gusBefore }));
    await shot(win, 'mail-accounts-5-agree-dialog.png');
    await js(`document.querySelector('.ml-agree-yes')?.click(), 1`);
    const gus = await waitFor(async () => (await js(`window.rutbaOffice.mail.accounts()`)).find((a) => a.email === 'gus.bulk@elsewhere.checks.example') || null, 'the account added once agreed', 8000);
    check('mail accounts: Use these servers adds the account with them', gus?.imap?.host === 'mail.collector.example' && gus?.smtp?.host === 'smtp.collector.example', JSON.stringify(gus && { imap: gus.imap, smtp: gus.smtp }));
    await js(`window.rutbaOffice.mail.discoveryForTest({ domain: 'elsewhere.checks.example', found: null })`);

    const complaints = await errorsIn(win);
    check('mail accounts: the window reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');

    // The box in Urdu, right to left.
    await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    const w = await open('mail');
    const urOpened = await openBulk(w);
    await paste('ali.bulk@gmail.com  pw\nnot an address', w);
    const ur = await waitFor(async () => {
      const s = await w.webContents.executeJavaScript(`(() => ({ dir: document.documentElement.dir, lead: document.querySelector('.ml-bulk-lead')?.textContent || '', textDir: document.querySelector('.ml-bulk-text')?.getAttribute('dir'), states: [...document.querySelectorAll('.ml-bulk-state')].map((n) => n.textContent), add: document.querySelector('.ml-bulk-add')?.textContent || '' }))()`);
      return s.states.length === 2 ? s : false;
    }, 'the Urdu box', 6000);
    await shot(w, 'mail-accounts-3-urdu.png');
    check('mail accounts: in Urdu the box reads right to left in Urdu, the addresses kept left to right',
      urOpened && Boolean(ur) && ur.dir === 'rtl' && /[؀-ۿ]/.test(ur.lead) && ur.textDir === 'ltr' && ur.states.every((x) => /[؀-ۿ]/.test(x)) && /[؀-ۿ]/.test(ur.add),
      JSON.stringify(ur));
    w.destroy();
  } catch (err) {
    check('mail accounts: the checks ran', false, err.stack || err.message);
  } finally {
    if (js) {
      await js(`window.rutbaOffice.mail.discoveryForTest({ domain: 'elsewhere.checks.example', found: null })`).catch(() => {});
      await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(language ?? null)} }).then(() => 1)`).catch(() => {});
      const list = await js(`window.rutbaOffice.mail.accounts()`).catch(() => []);
      for (const a of list.filter((x) => /\.bulk@/.test(x.email))) await js(`window.rutbaOffice.mail.removeAccount({ id: ${JSON.stringify(a.id)}, keepMail: false })`).catch(() => {});
    }
  }
}
