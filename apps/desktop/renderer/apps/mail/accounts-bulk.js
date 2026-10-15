// Add several accounts at once: rows of addresses and passwords, pasted or
// loaded from a file, read as they are typed and set up together.
//
// The rows are read here (@rutba/mailbox/account-rows) the moment they are
// pasted, or when the box is left or Add is pressed, and the box is cleared:
// the passwords are held only in this window's memory until Add hands them
// to the main process, which keeps them where Mail keeps every password (the
// system's protected store) and never writes them anywhere else. Each line
// shows at once whether it is an address with a password, and a password is
// never shown back. Add sends the rows to addAccounts, which sets each up as
// the single Add account does - the built-in table, or the domain's own
// records looked up only now, then its servers tried until one signs in -
// three at a time, and says how each went as it goes. A row whose domain
// names servers somewhere else waits for the person to agree to them by name
// before its password goes there; the main process holds that row (in
// memory only) until they do, or the dialog closes. A line that does not
// start with an address is shown by its number only, never its text.

import React, { useEffect, useMemo, useState } from 'react';
import { Button, Dialog, Icon, Spinner, t, tn } from '@rutba/office-ui';
import { parseAccountRows, checkRows } from '@rutba/mailbox/account-rows';

const domainOf = (email) => String(email || '').split('@')[1] || '';

/** What a row's outcome says, in a sentence. */
function sentence(r) {
  const who = r.provider || r.server;
  switch (r.outcome) {
    case 'ready-to-add':
      return t('Ready');
    case 'not-an-address':
      return t('Not an email address');
    case 'unreadable':
      return t('Line {line} could not be read: it does not start with an email address.', { line: r.line });
    case 'confirm':
      return t('Its mail servers are not at {domain}: {hosts}. The password goes to them only if you agree.', { domain: domainOf(r.email), hosts: (r.foreign || []).map((x) => x.host).join(', ') });
    case 'repeated':
      return t('Already in the list above');
    case 'no-password':
      return t('No password: it will be offered its provider\'s own sign-in');
    case 'exists':
      return t('Already set up');
    case 'waiting':
      return t('Waiting');
    case 'trying':
      return r.known
        ? t('Signing in to {provider}\'s servers…', { provider: r.provider || domainOf(r.email) })
        : t('Looking up {domain}: its DNS records (SRV and MX), autoconfig.{domain} and autodiscover.{domain}, then trying the servers they name…', { domain: domainOf(r.email) });
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
  // Rows read from the box, which is then cleared.
  const [held, setHeld] = useState([]);
  const [have, setHave] = useState([]);
  const [run, setRun] = useState(null); // { job, results: [...] } once started
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    shell.mail
      .accounts()
      .then((list) => setHave((list || []).map((a) => String(a.email).toLowerCase())))
      .catch(() => {});
  }, [shell]);

  const parsed = useMemo(() => checkRows([...held, ...parseAccountRows(text).rows]), [held, text]);
  /** Read the box, or what was pasted into it, into the rows, and clear it. */
  const take = (more = text) => {
    const read = parseAccountRows(more).rows;
    if (read.length) setHeld((h) => [...h, ...read]);
    setText('');
    return read;
  };
  const preview = parsed.map((r) => ({
    line: r.line,
    email: r.email,
    outcome: r.problem === 'not-an-address' || r.problem === 'repeated' || r.problem === 'unreadable' ? r.problem : have.includes(r.email.toLowerCase()) ? 'exists' : r.problem === 'no-password' ? 'no-password' : 'ready-to-add',
  }));
  const toAdd = preview.filter((r) => r.outcome === 'ready-to-add' || r.outcome === 'no-password').length;
  const remove = (i) => setHeld((h) => h.filter((_, n) => n !== i));
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
      // Read into the list at once, as a paste is: the passwords never sit in the box.
      const { text: read } = await shell.fs.readText({ path: paths[0], encoding: 'utf8' });
      take(read);
    } catch (err) {
      toast?.(err.message, { tone: 'bad' });
    }
  };

  const start = async () => {
    const job = `bulk-${Date.now()}`;
    const rows = checkRows([...held, ...parseAccountRows(text).rows]);
    setText('');
    setHeld([]);
    setRun({ job, results: preview.map((r) => ({ ...r, outcome: r.outcome === 'ready-to-add' || r.outcome === 'no-password' ? 'waiting' : r.outcome })) });
    setBusy(true);
    try {
      const r = await shell.mail.addAccounts({ rows: rows.map(({ problem, ...row }) => row), job });
      setRun({ job, results: r.results });
    } catch (err) {
      toast?.(err.message, { tone: 'bad' });
    } finally {
      setBusy(false);
    }
  };

  /** The person agreed to a row's servers by name: the main process tries them now with the password it holds. */
  const agree = async (i) => {
    const r = run.results[i];
    setRun((cur) => ({ ...cur, results: cur.results.map((x, n) => (n === i ? { ...x, outcome: 'trying', known: false } : x)) }));
    try {
      await shell.mail.confirmAccounts({ job: run.job, index: i, hosts: (r.foreign || []).map((x) => x.host) });
    } catch (err) {
      toast?.(err.message, { tone: 'bad' });
    }
  };
  /** Closing forgets every row still held for agreement, with its password. */
  const close = (n) => {
    if (run) shell.mail.forgetAccountsJob({ job: run.job }).catch(() => {});
    onClose(n);
  };

  const added = run ? run.results.filter((r) => r.outcome === 'added' || r.outcome === 'added-no-outgoing').length : 0;
  const done = run && !busy;
  return (
    <Dialog
      title={t('Add several accounts')}
      width={680}
      onClose={() => close(added)}
      actions={
        done ? (
          <Button primary className="ml-bulk-done" label={t('Done')} onClick={() => close(added)} />
        ) : (
          <>
            <Button label={t('Cancel')} disabled={busy} onClick={() => close(added)} />
            <Button primary className="ml-bulk-add" label={busy ? t('Adding…') : tn(toAdd, 'Add {count} account', 'Add {count} accounts')} disabled={busy || !toAdd} onClick={start} />
          </>
        )
      }
    >
      <div className="ml-bulk-accounts">
        {!run ? (
          <>
            <p className="ml-bulk-lead">
              {t('One account a line: the address, then its password, separated by a space, a tab, a comma, a semicolon, a colon or a bar. A first line naming the columns (email, password, name, IMAP server...) is read as a table, as a spreadsheet saves one. A line starting # is left out.')}{' '}
              {t('What is pasted is read into the list below at once and the box is cleared; the passwords go only to this computer\'s protected store.')}
            </p>
            <textarea
              className="ml-bulk-text"
              spellCheck={false}
              dir="ltr"
              value={text}
              onChange={(e) => setText(e.target.value)}
              onPaste={(e) => {
                const pasted = e.clipboardData?.getData('text') || '';
                if (!pasted) return;
                e.preventDefault();
                take(`${text}${text && !text.endsWith('\n') ? '\n' : ''}${pasted}`);
              }}
              onKeyDown={(e) => {
                // A typed line goes into the list when Enter ends it, rather than sitting in the box.
                if (e.key === 'Enter' && !e.shiftKey && text.trim()) {
                  e.preventDefault();
                  take();
                }
              }}
              onBlur={() => text.trim() && take()}
              placeholder={'ali@example.com  password\nbo@example.com,password\ncy@example.com:password'}
            />
            <div className="ml-bulk-tools">
              <Button icon="open" label={t('Load from a file…')} onClick={load} />
              <span className="ml-bulk-note">{t('Each is set up as Add account does: a provider in the built-in list straight away; for any other domain, its own records are looked up when you press Add (its DNS, autoconfig and autodiscover addresses), and nothing else is asked.')}</span>
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
              <li key={`${r.line}-${i}`} className={`ml-bulk-row ${GOOD.has(r.outcome) ? 'good' : BUSY.has(r.outcome) ? 'busy' : r.outcome === 'no-password' || r.outcome === 'sign-in' || r.outcome === 'confirm' ? 'plain' : 'bad'}`} data-outcome={r.outcome} data-email={r.email}>
                <span className="ml-bulk-mark">{r.outcome === 'trying' ? <Spinner /> : <Icon name={GOOD.has(r.outcome) ? 'check' : BUSY.has(r.outcome) ? 'clock' : r.outcome === 'sign-in' || r.outcome === 'no-password' ? 'lock' : 'info'} size={14} />}</span>
                <span className="ml-bulk-email" dir="ltr">{r.email || t('Line {line}', { line: r.line })}</span>
                <span className="ml-bulk-state">{sentence(r)}</span>
                {!run && i < held.length ? (
                  <button type="button" className="ml-bulk-remove" aria-label={t('Remove')} data-tip={t('Remove')} onClick={() => remove(i)}>
                    <Icon name="close" size={12} />
                  </button>
                ) : null}
                {done && r.outcome === 'sign-in' ? (
                  <button type="button" className="ml-bulk-signin" onClick={() => onSignIn(r.email)}>
                    {t('Sign in')}
                  </button>
                ) : null}
                {r.outcome === 'confirm' ? (
                  <button type="button" className="ml-bulk-agree" onClick={() => agree(i)}>
                    {t('Use these servers')}
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
.ml-bulk-row.plain .ml-bulk-mark { color: var(--accent-text); }
.ml-bulk-email { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink); }
.ml-bulk-state { color: var(--ink-2); }
.ml-bulk-row.bad .ml-bulk-state { color: var(--bad, #c62828); }
.ml-bulk-row button { border: 0; background: none; color: var(--accent-text); font: inherit; font-weight: 600; cursor: pointer; padding: 0; white-space: nowrap; }
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
