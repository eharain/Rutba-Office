// The question a database asks before it reads its sources (A2.18): a
// database whose tables are read from places this computer has not allowed
// (one someone else made and sent, or one from before 2.18) names each
// place, by its kind and where it is, and reads nothing until the person
// allows it. Not now keeps the copies the tables have, and Refresh on one
// asks again. main/base-trust.js keeps the answer, on this computer, never
// in the file.

import React, { useState } from 'react';
import { Button, Icon, Dialog, t } from '@rutba/office-ui';

/** What a place is, in words: a web address, a file or database on a share, in the person's own folders or outside them, or an engine. */
function kindWords(a) {
  const db = a.kind === 'rbase';
  if (a.place === 'engine') return t('A {engine} engine at {host}', { engine: a.engine || '', host: a.host });
  if (a.place === 'web') return a.kind === 'json' ? t('A web address answering JSON, on {host}', { host: a.host }) : t('A file at a web address, on {host}', { host: a.host });
  if (a.place === 'share') return db ? t('A database on a network share, on {server}', { server: a.server }) : t('A file on a network share, on {server}', { server: a.server });
  if (a.place === 'mine') return db ? t('A database in your own folders') : t('A file in your own folders');
  return db ? t('A database outside your own folders') : t('A file outside your own folders');
}

/** What reading it does, said plainly. */
function noteWords(a) {
  if (a.place === 'engine') return t('Connecting tells {host} that this database was opened, and from where.', { host: a.host });
  if (a.place === 'web') return t('Reading it tells {host} that this database was opened, and from where.', { host: a.host });
  if (a.place === 'share') return t('Opening a network share makes Windows send your sign-in to {server}.', { server: a.server });
  if (a.place === 'mine') return t('What is read goes into this database, and with it if it is sent on.');
  return t('It is not in your own folders: it may be another person\'s file or one of the computer\'s own. What is read goes into this database, and with it if it is sent on.');
}

/** Where it is, in full: the whole web address, the path, or the engine's account and host. */
const whereWords = (a) => (a.place === 'engine'
  ? t('{user} at {host}, database {database}.', { user: a.user, host: a.port ? `${a.host}:${a.port}` : a.host, database: a.database })
  : a.place === 'web' ? a.url : a.path);

const ICONS = { web: 'globe', share: 'folderOpen', mine: 'folder', outside: 'shield', engine: 'link' };

/** Read from these places? Each named; Allow for this database, or Not now. */
export function SourcesQuestion({ shell, db, onDone, onClose, toast }) {
  const [busy, setBusy] = useState(false);
  const ask = db.ask || [];
  const allow = async () => {
    setBusy(true);
    try {
      onDone(await shell.base.allowSources({ id: db.id, addresses: ask.map((a) => a.address) }));
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog title={t('Read from these places?')} width={620} className="bs-ask" onClose={onClose} actions={(
      <>
        <Button label={t('Not now')} className="bs-ask-later" onClick={onClose} />
        <Button primary label={t('Allow for this database')} className="bs-ask-allow" disabled={busy || !ask.length} onClick={allow} />
      </>
    )}>
      <p className="bs-hint">{t('This database reads records from the places below when it opens and on a schedule. They were not linked on this computer, so nothing has been read from them yet. Allow them only if you know where this database came from.')}</p>
      <ul className="bs-ask-list">
        {ask.map((a) => (
          <li key={a.id} className="bs-ask-item" data-ask={a.kind} data-place={a.place}>
            <Icon name={ICONS[a.place] || 'info'} size={16} />
            <div>
              <strong className="bs-ask-kind">{kindWords(a)}</strong>
              <code className="bs-ask-where">{whereWords(a)}</code>
              <small className="bs-ask-note">{noteWords(a)}</small>
              <small className="bs-ask-tables">{t('Read into {tables}', { tables: a.tables.join(', ') })}</small>
            </div>
          </li>
        ))}
      </ul>
      <p className="bs-hint">{t('Your answer is kept for this file on this computer. Until you allow them, its tables keep the copies they have and are not read.')}</p>
    </Dialog>
  );
}

export const TRUST_CSS = `
.bs-ask-list { list-style: none; margin: 8px 0; padding: 0; max-height: 320px; overflow: auto; }
.bs-ask-item { display: flex; gap: 10px; align-items: flex-start; padding: 8px 0; border-bottom: 1px solid var(--line-soft); }
.bs-ask-item > div { display: flex; flex-direction: column; gap: 3px; min-width: 0; }
.bs-ask-item[data-place="share"] > svg, .bs-ask-item[data-place="outside"] > svg { color: var(--bad, #b42318); }
.bs-ask-where { font-size: 12px; overflow-wrap: anywhere; user-select: text; }
.bs-ask-note, .bs-ask-tables { color: var(--ink-2); }
.bs-linked-bar.blocked { background: rgba(181, 120, 0, .08); }
`;
