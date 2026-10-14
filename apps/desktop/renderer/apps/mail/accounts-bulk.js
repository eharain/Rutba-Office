// Add several accounts at once: rows of addresses and passwords, pasted or
// loaded from a file, read as they are typed and set up together.
//
// The rows are read here as they are typed (@rutba/mailbox/account-rows), so
// each line shows at once whether it is an address with a password, and the
// passwords are never shown back. Adding sends the text to the main
// process's addAccounts, which sets each up as the single Add account does -
// finding the servers, trying them until one signs in - three at a time, and
// says how each went as it goes.

import React, { useEffect, useMemo, useState } from 'react';
import { Button, Dialog, Icon, Spinner, t, tn } from '@rutba/office-ui';
import { parseAccountRows } from '@rutba/mailbox/account-rows';

const domainOf = (email) => String(email || '').split('@')[1] || '';

/** What a row's outcome says, in a sentence. */
function sentence(r) {
  const who = r.provider || r.server;
  switch (r.outcome) {
    case 'ready-to-add':
      return t('Ready');
    case 'not-an-address':
      return t('Not an email address');
    case 'repeated':
      return t('Already in the list above');
    case 'no-password':
      return t('No password: it will be offered its provider\'s own sign-in');
    case 'exists':
      return t('Already set up');
    case 'waiting':
      return t('Waiting');
    case 'trying':
      return t('Finding its servers and signing in…');
    case 'added':
      return who ? t('Added: {provider}', { provider: who }) : t('Added');
    case 'added-no-outgoing':
      return t('Added; the outgoing server did not take the password, so sending needs its settings checked');
    case 'sign-in':
      return t('Signs in with {provider}\'s own page', { provider: who || t('its provider') });
    case 'app-password':
      return t('{provider} wants an app password rather than the account\'s own', { provider: who || t('The provider') });
    case 'password':
      return t('The password was refused');
    case 'not-found':
      return t('No mail server found for {domain}', { domain: domainOf(r.email) });
    case 'unreachable':
      return t('The mail server did not answer');
    case 'certificate':
      return t('The mail server\'s certificate is not valid');
    default:
      return r.message || t('Could not be set up');
  }
}

const GOOD = new Set(['added', 'ready-to-add', 'exists']);
const BUSY = new Set(['trying', 'waiting']);

/**
 * @param {{ shell, onClose: (added: number) => void, onSignIn: (email: string) => void, onByHand: (email: string) => void, onImportFile?: () => void, toast }} props
 */
export function BulkAccountsDialog({ shell, onClose, onSignIn, onByHand, onImportFile, toast }) {
  const [text, setText] = useState('');
  const [have, setHave] = useState([]);
  const [run, setRun] = useState(null); // { job, results: [...] } once started
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    shell.mail
      .accounts()
      .then((list) => setHave((list || []).map((a) => String(a.email).toLowerCase())))
      .catch(() => {});
  }, [shell]);

  const parsed = useMemo(() => parseAccountRows(text), [text]);
  const preview = parsed.rows.map((r) => ({
    line: r.line,
    email: r.email,
    outcome: r.problem === 'not-an-address' || r.problem === 'repeated' ? r.problem : have.includes(r.email.toLowerCase()) ? 'exists' : r.problem === 'no-password' ? 'no-password' : 'ready-to-add',
  }));
  const toAdd = preview.filter((r) => r.outcome === 'ready-to-add' || r.outcome === 'no-password').length;
  const rows = run ? run.results : preview;

  useEffect(() => {
    if (!run) return undefined;
    return shell.on('mail:accountsProgress', (p) => {
      if (p.job !== run.job) return;
      setRun((cur) => (cur && cur.job === p.job ? { ...cur, results: cur.results.map((r, i) => (i === p.index ? { ...r, ...p } : r)) } : cur));
    });
  }, [shell, run?.job]);

  const load = async () => {
    const paths = await shell.dialog.open({
      title: t('Accounts from a file'),
      filters: [
        { name: t('Lists of accounts (.csv, .tsv, .txt)'), extensions: ['csv', 'tsv', 'txt'] },
        { name: t('Accounts (JSON)'), extensions: ['json'] },
        { name: t('All files'), extensions: ['*'] },
      ],
    });
    if (!paths?.[0]) return;
    if (/\.json$/i.test(paths[0]) && onImportFile) return onImportFile(paths[0]);
    try {
      const { text: read } = await shell.fs.readText({ path: paths[0], encoding: 'utf8' });
      setText(read);
    } catch (err) {
      toast?.(err.message, { tone: 'bad' });
    }
  };

  const start = async () => {
    const job = `bulk-${Date.now()}`;
    setRun({ job, results: preview.map((r) => ({ ...r, outcome: r.outcome === 'ready-to-add' || r.outcome === 'no-password' ? 'waiting' : r.outcome })) });
    setBusy(true);
    try {
      const r = await shell.mail.addAccounts({ text, job });
      setRun({ job, results: r.results });
    } catch (err) {
      toast?.(err.message, { tone: 'bad' });
    } finally {
      setBusy(false);
    }
  };

  const added = run ? run.results.filter((r) => r.outcome === 'added' || r.outcome === 'added-no-outgoing').length : 0;
  const done = run && !busy;
  return (
    <Dialog
      title={t('Add several accounts')}
      width={680}
      onClose={() => onClose(added)}
      actions={
        done ? (
          <Button primary className="ml-bulk-done" label={t('Done')} onClick={() => onClose(added)} />
        ) : (
          <>
            <Button label={t('Cancel')} disabled={busy} onClick={() => onClose(added)} />
            <Button primary className="ml-bulk-add" label={busy ? t('Adding…') : tn(toAdd, 'Add {count} account', 'Add {count} accounts')} disabled={busy || !toAdd} onClick={start} />
          </>
        )
      }
    >
      <div className="ml-bulk-accounts">
        {!run ? (
          <>
            <p className="ml-bulk-lead">
              {t('One account a line: the address, then its password, separated by a space, a tab, a comma, a semicolon, a colon or a bar. A first line naming the columns (email, password, name, IMAP server...) is read as a table, as a spreadsheet saves one. A line starting # is left out.')}
            </p>
            <textarea
              className="ml-bulk-text"
              spellCheck={false}
              dir="ltr"
              value={text}
              onChange={(e) => setText(e.target.value)}
              placeholder={'ali@example.com  password\nbo@example.com,password\ncy@example.com:password'}
            />
            <div className="ml-bulk-tools">
              <Button icon="open" label={t('Load from a file…')} onClick={load} />
              <span className="ml-bulk-note">{t('Each is set up as Add account does: its servers found, and tried until one signs in. Passwords go to this computer\'s keystore.')}</span>
            </div>
          </>
        ) : (
          <p className="ml-bulk-lead" role="status">
            {busy ? t('Setting up the accounts, three at a time…') : tn(added, 'Added {count} account.', 'Added {count} accounts.')}
          </p>
        )}
        {rows.length ? (
          <ul className="ml-bulk-rows">
            {rows.map((r, i) => (
              <li key={`${r.line}-${i}`} className={`ml-bulk-row ${GOOD.has(r.outcome) ? 'good' : BUSY.has(r.outcome) ? 'busy' : r.outcome === 'no-password' || r.outcome === 'sign-in' ? 'plain' : 'bad'}`} data-outcome={r.outcome} data-email={r.email}>
                <span className="ml-bulk-mark">{r.outcome === 'trying' ? <Spinner /> : <Icon name={GOOD.has(r.outcome) ? 'check' : BUSY.has(r.outcome) ? 'clock' : r.outcome === 'sign-in' || r.outcome === 'no-password' ? 'lock' : 'info'} size={14} />}</span>
                <span className="ml-bulk-email" dir="ltr">{r.email || t('(empty line)')}</span>
                <span className="ml-bulk-state">{sentence(r)}</span>
                {done && r.outcome === 'sign-in' ? (
                  <button type="button" className="ml-bulk-signin" onClick={() => onSignIn(r.email)}>
                    {t('Sign in')}
                  </button>
                ) : null}
                {done && ['not-found', 'unreachable', 'password', 'app-password', 'certificate', 'failed'].includes(r.outcome) ? (
                  <button type="button" className="ml-bulk-byhand" onClick={() => onByHand(r.email)}>
                    {t('Set up by hand')}
                  </button>
                ) : null}
              </li>
            ))}
          </ul>
        ) : null}
      </div>
    </Dialog>
  );
}

export const BULK_CSS = `
.ml-bulk-accounts { display: flex; flex-direction: column; gap: 10px; }
.ml-bulk-lead { margin: 0; color: var(--ink-2); font-size: 12.5px; line-height: 1.45; }
.ml-bulk-text { width: 100%; box-sizing: border-box; min-height: 140px; resize: vertical; font-family: var(--mono); font-size: 12px; line-height: 1.5; padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); color: var(--ink); }
.ml-bulk-text:focus { outline: 2px solid color-mix(in srgb, var(--accent) 45%, transparent); outline-offset: -1px; }
.ml-bulk-tools { display: flex; align-items: center; gap: 12px; }
.ml-bulk-note { color: var(--ink-3); font-size: 11.5px; line-height: 1.4; }
.ml-bulk-rows { list-style: none; margin: 0; padding: 0; border: 1px solid var(--line); border-radius: var(--r-2); max-height: 260px; overflow: auto; }
.ml-bulk-row { display: grid; grid-template-columns: 20px minmax(160px, 38%) 1fr auto auto; gap: 8px; align-items: center; padding: 5px 10px; font-size: 12.5px; }
.ml-bulk-row + .ml-bulk-row { border-top: 1px solid var(--line-soft); }
.ml-bulk-mark { display: inline-flex; color: var(--ink-3); }
.ml-bulk-mark .rw-spinner { width: 12px; height: 12px; border-width: 2px; }
.ml-bulk-row.good .ml-bulk-mark { color: var(--good, #2e7d32); }
.ml-bulk-row.bad .ml-bulk-mark { color: var(--bad, #c62828); }
.ml-bulk-row.plain .ml-bulk-mark { color: var(--accent); }
.ml-bulk-email { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink); }
.ml-bulk-state { color: var(--ink-2); }
.ml-bulk-row.bad .ml-bulk-state { color: var(--bad, #c62828); }
.ml-bulk-row button { border: 0; background: none; color: var(--accent); font: inherit; font-weight: 600; cursor: pointer; padding: 0; white-space: nowrap; }
`;

let installed = false;

/** The dialog's styles, once. */
export function installBulkStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-mail-bulk-css';
  style.textContent = BULK_CSS;
  document.head.appendChild(style);
}
