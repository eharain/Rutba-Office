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
  if (a.place === 'engine') return t('base.trust.engine-engine-host', { engine: a.engine || '', host: a.host });
  if (a.place === 'web') return a.kind === 'json' ? t('base.trust.web-address-answering', { host: a.host }) : t('base.trust.file-web-address', { host: a.host });
  if (a.place === 'share') return db ? t('base.trust.database-network-share', { server: a.server }) : t('base.trust.file-network-share', { server: a.server });
  if (a.place === 'mine') return db ? t('base.trust.database-own-folders') : t('base.trust.file-own-folders');
  return db ? t('base.trust.database-outside-own') : t('base.trust.file-outside-own');
}

/** What reading it does, said plainly. */
function noteWords(a) {
  if (a.place === 'engine') return t('base.trust.connecting-tells-host', { host: a.host });
  if (a.place === 'web') return t('base.trust.reading-tells-host', { host: a.host });
  if (a.place === 'share') return t('base.trust.opening-network-share', { server: a.server });
  if (a.place === 'mine') return t('base.trust.read-goes-database');
  return t('base.trust.not-own-folders');
}

/** Where it is, in full: the whole web address, the path, or the engine's account and host. */
const whereWords = (a) => (a.place === 'engine'
  ? t('base.user-host-database', { user: a.user, host: a.port ? `${a.host}:${a.port}` : a.host, database: a.database })
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
    <Dialog title={t('base.trust.read-places')} width={620} className="bs-ask" onClose={onClose} actions={(
      <>
        <Button label={t('common.not-now')} className="bs-ask-later" onClick={onClose} />
        <Button primary label={t('base.trust.allow')} className="bs-ask-allow" disabled={busy || !ask.length} onClick={allow} />
      </>
    )}>
      <p className="bs-hint">{t('base.trust.database-reads-records')}</p>
      <ul className="bs-ask-list">
        {ask.map((a) => (
          <li key={a.id} className="bs-ask-item" data-ask={a.kind} data-place={a.place}>
            <Icon name={ICONS[a.place] || 'info'} size={16} />
            <div>
              <strong className="bs-ask-kind">{kindWords(a)}</strong>
              <code className="bs-ask-where">{whereWords(a)}</code>
              <small className="bs-ask-note">{noteWords(a)}</small>
              <small className="bs-ask-tables">{t('base.trust.read-into-tables', { tables: a.tables.join(', ') })}</small>
            </div>
          </li>
        ))}
      </ul>
      <p className="bs-hint">{t('base.trust.answer-kept-file')}</p>
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
