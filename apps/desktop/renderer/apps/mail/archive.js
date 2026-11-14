// Mail as the archive tool: the export for legal hold, the search language's
// help, and the strip that says an archive is open where it lies.
//
// The engine is @rutba/mailbox (archive.js, search-index.js, legal-export.js)
// behind the main process's mail-archive.js; this is only what the window
// shows of it.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Field, Input, Icon, Chip, Progress, Select, formatBytes, t, tn } from '@rutba/office-ui';

export const isArchiveId = (id) => typeof id === 'string' && id.startsWith('archive:');

/** The words of an exported PDF, in the window's language. */
export function pdfLabels() {
  return {
    from: t('common.from'),
    to: t('common.to'),
    cc: t('common.cc'),
    bcc: t('common.bcc'),
    date: t('common.date'),
    messageId: t('mail.archive.message-id'),
    source: t('common.source'),
    folder: t('common.folder'),
    attachments: t('common.attachments'),
    name: t('common.name'),
    type: t('common.type'),
    size: t('common.size'),
    sha256: t('mail.archive.sha-256'),
    noSubject: t('mail.no-subject'),
    noBody: t('mail.archive.message-no-text'),
    summaryTitle: t('mail.archive.export-summary'),
    exported: t('mail.archive.exported'),
    selection: t('mail.archive.what-was-exported'),
    messages: t('mail.messages'),
    dates: t('mail.archive.dates'),
    sources: t('common.sources'),
    files: t('common.files'),
    untouched: t('mail.archive.sources-opened-reading'),
    cancelled: t('mail.archive.cancelled-before-end'),
    footer: t('mail.archive.exported-rutba-office'),
  };
}

/** The search language, as the help lists it: what to type, and what it finds. */
export function searchSyntax() {
  return [
    ['from:amina', t('mail.archive.sent-someone-name')],
    ['to:bo  cc:legal  bcc:', t('mail.archive.sent-someone-copied')],
    ['subject:contract', t('mail.archive.word-subject')],
    ['"indemnity clause"', t('mail.archive.words-together-order')],
    ['has:attachment', t('mail.archive.carries-a-file')],
    ['filename:pdf', t('mail.archive.file-name-holds')],
    ['is:unread  is:read  is:starred', t('mail.archive.read-not-starred')],
    ['after:2019-01-01  before:2020-01-01', t('mail.archive.sent-after-day')],
    ['received:2019-03  received:>=2019-03', t('mail.archive.month-year')],
    ['received:2019-01-01..2019-06-30', t('mail.archive.between-two-days')],
    ['older_than:1y  newer_than:30d', t('mail.archive.older-newer-many')],
    ['size>1mb  larger:5m  smaller:100kb', t('mail.archive.larger-smaller-size')],
    ['in:projects', t('mail.archive.filed-folder')],
    ['account:northwind', t('mail.archive.one-account-archive')],
    ['invoice OR receipt', t('mail.archive.either')],
    ['-from:newsletter', t('mail.archive.leave-out-follows')],
  ];
}

export function SearchHelp({ onClose, onInsert }) {
  return (
    <Dialog
      title={t('mail.search-syntax')}
      width={620}
      onClose={onClose}
      actions={<Button primary label={t('common.close')} onClick={onClose} />}
    >
      <p className="ml-help-lead">{t('mail.archive.type-words-find')}</p>
      <div className="ml-help-table" role="list">
        {searchSyntax().map(([code, meaning]) => (
          <button key={code} type="button" role="listitem" className="ml-help-row" onClick={() => onInsert(code.split('  ')[0])}>
            <code dir="ltr">{code}</code>
            <span>{meaning}</span>
          </button>
        ))}
      </div>
    </Dialog>
  );
}

/** What kind of archive, in the window's language. */
export function archiveKind(kind) {
  if (kind === 'pst') return t('mail.archive.outlook-data-file');
  if (kind === 'ost') return t('mail.archive.outlook-offline-file');
  if (kind === 'olm') return t('mail.archive.outlook-mac-archive');
  if (kind === 'mbox') return t('mail.archive.mbox-archive');
  return kind || '';
}

/** How far an archive's indexing has got, in words: "40,000 of 100,000". */
export const indexedCount = (index) => t('mail.archive.done-of-total', { done: (index?.done || 0).toLocaleString(), total: (index?.total || 0).toLocaleString() });

/**
 * The strip at the top of an archive's message list: what it is, that it is
 * read only (or how far its indexing has got), and Export, Stop or Resume
 * indexing, and Close.
 */
export function ArchiveBanner({ archive, onClose, onExport, onStopIndexing, onResumeIndexing }) {
  const indexing = archive.index?.state === 'indexing';
  const paused = archive.index?.state === 'paused';
  const note = archive.missing
    ? t('mail.archive.file-no-longer')
    : indexing
      ? t('mail.archive.read-only-opened-lies', { count: indexedCount(archive.index) })
      : paused
        ? t('mail.archive.read-only-opened-lies-indexing', { count: indexedCount(archive.index) })
        : t('mail.archive.read-only-opened');
  return (
    <div className="ml-archive-banner" data-archive={archive.id}>
      <Icon name="archive" size={16} />
      <strong className="ml-archive-name" title={archive.path}>
        {archive.name}
      </strong>
      <span className="ml-archive-size" dir="ltr">{formatBytes(archive.bytes || 0)}</span>
      <span className="ml-archive-note">
        {archiveKind(archive.kind)} · {note}
      </span>
      <span className="ml-archive-actions">
        <button type="button" onClick={onExport}>
          {t('common.export-2')}
        </button>
        {indexing && onStopIndexing ? (
          <button type="button" className="ml-archive-stop" onClick={onStopIndexing}>
            {t('mail.archive.stop')}
          </button>
        ) : null}
        {paused && onResumeIndexing ? (
          <button type="button" className="ml-archive-resume" onClick={onResumeIndexing}>
            {t('mail.archive.resume')}
          </button>
        ) : null}
        <button type="button" onClick={onClose}>
          {t('mail.close')}
        </button>
      </span>
    </div>
  );
}

/**
 * Export for legal hold: the search's results, this folder, or folders
 * chosen; to an mbox and PDFs, with a manifest and a summary. Progress and
 * Cancel while it runs, and a sentence at the end saying what was written.
 */
export function ExportHoldDialog({ shell, query, resultCount, scopeSources, folderOf, folders, sourceName, onClose }) {
  const [what, setWhat] = useState(query && resultCount ? 'results' : 'folder');
  const [chosen, setChosen] = useState(() => new Set(folderOf ? [folderOf.folder] : []));
  const [mbox, setMbox] = useState(true);
  const [pdf, setPdf] = useState('combined');
  const [attachments, setAttachments] = useState(false);
  const [dir, setDir] = useState('');
  const [state, setState] = useState('choose');
  const [progress, setProgress] = useState({ done: 0, total: 0 });
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const job = useRef(`hold-${Date.now().toString(36)}`);

  useEffect(() => {
    shell.store.get({ key: 'mail.exportDir', fallback: '' }).then((d) => d && setDir((x) => x || d)).catch(() => {});
    const off = shell.on('mail:exportProgress', (p) => {
      if (p.job === job.current) setProgress({ done: p.done, total: p.total });
    });
    return () => off?.();
  }, [shell]);

  const count = useMemo(() => {
    if (what === 'results') return resultCount || 0;
    const list = what === 'folder' ? [folderOf?.folder] : [...chosen];
    return (folders || []).filter((f) => list.includes(f.path)).reduce((n, f) => n + (f.total || 0), 0);
  }, [what, resultCount, chosen, folders, folderOf]);

  const browse = async () => {
    const [picked] = await shell.dialog.open({ title: t('mail.archive.export-to'), directory: true, defaultPath: dir || undefined });
    if (picked) setDir(picked);
  };

  const run = async () => {
    setState('running');
    setError(null);
    setProgress({ done: 0, total: count });
    shell.store.set({ key: 'mail.exportDir', value: dir }).catch(() => {});
    const selection =
      what === 'results'
        ? t('mail.archive.the-search-query', { query })
        : what === 'folder'
          ? t('mail.archive.folder-folder-source', { folder: folderOf?.folder || '', source: sourceName || '' })
          : tn(chosen.size, 'mail.archive.count-folder-source', { source: sourceName || '' });
    try {
      const r = await shell.mail.exportHold({
        job: job.current,
        dir,
        mbox,
        pdf: pdf === 'none' ? null : pdf,
        attachments,
        labels: pdfLabels(),
        selection,
        ...(what === 'results'
          ? { query, sources: scopeSources || null }
          : { folders: (what === 'folder' ? [folderOf?.folder] : [...chosen]).map((folder) => ({ accountId: folderOf.accountId, folder })) }),
      });
      setResult(r);
      setState('done');
    } catch (err) {
      setError(err.message);
      setState('choose');
    }
  };

  const sentence = (r) => {
    const parts = [];
    if (r.mbox) parts.push(t('mail.archive.an-mbox'));
    if (r.combinedPages) parts.push(tn(r.combinedPages, 'mail.archive.pdf-count-page'));
    if (r.pdfs) parts.push(tn(r.pdfs, 'mail.archive.count-pdf-one'));
    if (r.attachments) parts.push(tn(r.attachments, 'mail.archive.count-attachment'));
    parts.push(t('mail.archive.manifest-every-hash'));
    return [
      r.cancelled
        ? t('mail.archive.stopped-after-done', { done: r.exported.toLocaleString(), total: r.requested.toLocaleString() })
        : tn(r.exported, 'mail.archive.exported-count-message', {
            sources: tn(r.sources, 'mail.archive.count-mailbox-archive'),
            dir: r.dir,
          }),
      `${parts.join(', ')}.`,
      t('mail.archive.nothing-mailboxes'),
    ].join(' ');
  };

  if (state === 'done' && result) {
    return (
      <Dialog
        title={t('mail.archive.export-legal-hold')}
        width={560}
        onClose={onClose}
        actions={
          <>
            <Button label={t('mail.archive.show-the')} icon="folderOpen" onClick={() => shell.shell.openPath({ path: result.dir })} />
            <Button primary label={t('common.close')} onClick={onClose} />
          </>
        }
      >
        <div className="ml-hold-done">
          <Icon name={result.cancelled ? 'info' : 'check'} size={18} />
          <p className="ml-hold-sentence">{sentence(result)}</p>
        </div>
        <div className="ml-hold-files" dir="ltr">
          {(result.files || []).map((f) => (
            <span key={f}>{f}</span>
          ))}
        </div>
      </Dialog>
    );
  }

  const running = state === 'running';
  return (
    <Dialog
      title={t('mail.archive.export-legal-hold')}
      width={580}
      className="ml-hold"
      onClose={running ? undefined : onClose}
      actions={
        running ? (
          <Button label={t('common.cancel')} onClick={() => shell.mail.cancelExport({ job: job.current })} />
        ) : (
          <>
            <Button label={t('common.cancel')} onClick={onClose} />
            <Button primary label={tn(count, 'mail.archive.export-count')} disabled={!count || !dir || (!mbox && pdf === 'none')} onClick={run} />
          </>
        )
      }
    >
      {running ? (
        <div className="ml-hold-progress">
          <p>{t('mail.archive.exported-done-total', { done: progress.done.toLocaleString(), total: (progress.total || count).toLocaleString() })}</p>
          <Progress value={progress.done} max={progress.total || count || 1} />
        </div>
      ) : (
        <>
          <div className="ml-hold-section">{t('common.what-to-export')}</div>
          <div className="ml-hold-choices" role="radiogroup">
            {query && resultCount ? (
              <label className="ml-hold-choice">
                <input type="radio" name="hold-what" checked={what === 'results'} onChange={() => setWhat('results')} />
                <span>{tn(resultCount, 'mail.archive.count-result-search', { query })}</span>
              </label>
            ) : null}
            {folderOf?.folder ? (
              <label className="ml-hold-choice">
                <input type="radio" name="hold-what" checked={what === 'folder'} onChange={() => setWhat('folder')} />
                <span>{t('mail.archive.this-folder-folder', { folder: folderOf.folder })}</span>
              </label>
            ) : null}
            {folders?.length > 1 ? (
              <label className="ml-hold-choice">
                <input type="radio" name="hold-what" checked={what === 'folders'} onChange={() => setWhat('folders')} />
                <span>{t('mail.archive.folders-of-source', { source: sourceName || '' })}</span>
              </label>
            ) : null}
          </div>
          {what === 'folders' ? (
            <div className="ml-import-folders ml-hold-folders">
              {folders.map((f) => {
                const on = chosen.has(f.path);
                return (
                  <label key={f.path} className="ml-import-folder" style={{ paddingInlineStart: 8 + Math.min(f.depth || 0, 5) * 13 }}>
                    <input
                      type="checkbox"
                      className="ml-check"
                      checked={on}
                      onChange={() =>
                        setChosen((s) => {
                          const next = new Set(s);
                          if (on) next.delete(f.path);
                          else next.add(f.path);
                          return next;
                        })
                      }
                    />
                    <Icon name="folder" size={14} />
                    <span className="name">{f.name || f.path}</span>
                    <span className="count">{(f.total || 0).toLocaleString()}</span>
                  </label>
                );
              })}
            </div>
          ) : null}

          <div className="ml-hold-section">{t('mail.archive.write')}</div>
          <label className="ml-hold-check">
            <input type="checkbox" className="ml-check" checked={mbox} onChange={(e) => setMbox(e.target.checked)} />
            <span>{t('mail.archive.mbox-file-messages')}</span>
          </label>
          <div className="ml-hold-row">
            <span>{t('mail.archive.pdf')}</span>
            <Select value={pdf} onChange={(e) => setPdf(e.target.value)} aria-label={t('mail.archive.pdf')}>
              <option value="combined">{t('mail.archive.one-pdf-all')}</option>
              <option value="each">{t('mail.archive.one-pdf-each')}</option>
              <option value="both">{t('mail.archive.both')}</option>
              <option value="none">{t('mail.archive.no-pdf')}</option>
            </Select>
          </div>
          <label className="ml-hold-check">
            <input type="checkbox" className="ml-check" checked={attachments} onChange={(e) => setAttachments(e.target.checked)} />
            <span>{t('mail.archive.write-each-messages')}</span>
          </label>

          <Field label={t('mail.archive.new-folder')}>
            <div className="ml-hold-dir">
              <Input value={dir} onChange={(e) => setDir(e.target.value)} placeholder={t('mail.archive.choose-a-folder')} dir="ltr" />
              <Button label={t('common.browse')} icon="folderOpen" onClick={browse} />
            </div>
          </Field>
          <p className="rw-hint ml-hold-note">
            {t('mail.archive.manifest-csv-json')}
          </p>
          {error ? <p className="ml-hold-error">{error}</p> : null}
        </>
      )}
    </Dialog>
  );
}

export const ARCHIVE_CSS = `
.ml-archive-banner {
  display: grid; grid-template-columns: auto 1fr auto; align-items: center; column-gap: 8px; row-gap: 2px;
  padding: 8px 12px 7px; background: color-mix(in srgb, var(--accent) 7%, var(--surface));
  border-bottom: 1px solid var(--line-soft); font-size: 12px; color: var(--ink-2);
}
.ml-archive-banner > svg { color: var(--accent-text); grid-row: 1 / span 2; align-self: start; margin-top: 1px; }
.ml-archive-name { color: var(--ink); font-size: 12.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; min-width: 0; }
.ml-archive-size { font-size: 11px; color: var(--ink-3); white-space: nowrap; font-variant-numeric: tabular-nums; }
.ml-archive-note { grid-column: 2 / 4; color: var(--ink-3); font-size: 11.5px; line-height: 1.4; }
.ml-archive-actions { grid-column: 2 / 4; display: flex; gap: 14px; margin-top: 3px; }
.ml-archive-actions button { border: 0; background: none; color: var(--accent-text); font: inherit; font-weight: 600; cursor: pointer; padding: 0; white-space: nowrap; }
.ml-archive-actions button:hover { text-decoration: underline; }

.ml-where {
  display: flex; align-items: center; gap: 4px; font-size: 10.5px; color: var(--ink-3);
  overflow: hidden; white-space: nowrap; margin-top: 1px;
}
.ml-where svg { flex: none; }
.ml-where span { overflow: hidden; text-overflow: ellipsis; }
.ml-where.archive { color: color-mix(in srgb, var(--accent) 70%, var(--ink-3)); }

.ml-scope {
  display: flex; align-items: center; gap: 6px; padding: 5px 10px; font-size: 11.5px; color: var(--ink-3);
  border-bottom: 1px solid var(--line-soft); background: var(--surface); flex-wrap: wrap;
}
.ml-scope-seg { display: inline-flex; border: 1px solid var(--line); border-radius: 20px; overflow: hidden; }
.ml-scope-seg button { border: 0; background: transparent; font: inherit; font-size: 11.5px; padding: 2px 10px; color: var(--ink-2); cursor: pointer; }
.ml-scope-seg button.on { background: var(--selected); color: var(--ink); font-weight: 600; }
.ml-scope-found { margin-inline-start: auto; white-space: nowrap; }
.ml-scope-found.about::before { content: '~'; }
.ml-scope-indexing { width: 100%; color: var(--ink-3); font-size: 11px; }

.ml-status-indexing { display: inline-flex; align-items: center; gap: 6px; }
.ml-status-indexing .rw-spinner { width: 11px; height: 11px; border-width: 2px; }

.ml-help-lead { margin: 0 0 10px; color: var(--ink-2); font-size: 12.5px; }
.ml-help-table { display: flex; flex-direction: column; border: 1px solid var(--line); border-radius: var(--r-2); overflow: hidden; max-height: 420px; overflow-y: auto; }
.ml-help-row {
  display: grid; grid-template-columns: minmax(250px, 46%) 1fr; gap: 12px; align-items: baseline; text-align: start;
  padding: 6px 12px; border: 0; border-bottom: 1px solid var(--line-soft); background: transparent; font: inherit; font-size: 12.5px; color: var(--ink); cursor: pointer;
}
.ml-help-row:last-child { border-bottom: 0; }
.ml-help-row:hover { background: var(--hover); }
.ml-help-row code { font-family: var(--mono); font-variant-ligatures: none; font-feature-settings: "liga" 0, "calt" 0; font-size: 11.5px; color: var(--accent-text); text-align: left; white-space: pre; }
.ml-help-row span { color: var(--ink-2); }

.ml-hold-section { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .04em; color: var(--ink-3); margin: 4px 0 6px; }
.ml-hold-section:not(:first-child) { margin-top: 14px; }
.ml-hold-choices { display: flex; flex-direction: column; gap: 6px; }
.ml-hold-choice, .ml-hold-check { display: flex; align-items: flex-start; gap: 8px; font-size: 12.5px; cursor: pointer; line-height: 1.4; }
.ml-hold-choice input, .ml-hold-check input { margin: 2px 0 0; flex: none; }
.ml-hold-check { margin: 4px 0; }
.ml-hold-folders { margin-top: 8px; max-height: 180px; }
.ml-hold-row { display: flex; align-items: center; gap: 10px; font-size: 12.5px; margin: 8px 0 6px 22px; }
.ml-hold-row .rw-select, .ml-hold-row select { flex: 1; }
.ml-hold-dir { display: flex; gap: 6px; align-items: center; }
.ml-hold-dir .rw-input { flex: 1; min-width: 0; }
.ml-hold-note { margin: 10px 0 0; font-size: 11.5px; line-height: 1.45; }
.ml-hold-error { color: var(--bad); font-size: 12px; margin: 8px 0 0; }
.ml-hold-progress { padding: 8px 2px 4px; }
.ml-hold-progress p { margin: 0 0 10px; font-size: 13px; font-variant-numeric: tabular-nums; }
.ml-hold-done { display: flex; gap: 10px; align-items: flex-start; }
.ml-hold-done svg { color: var(--good); flex: none; margin-top: 2px; }
.ml-hold-sentence { margin: 0; font-size: 13px; line-height: 1.5; overflow-wrap: anywhere; }
.ml-hold-files { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 12px; }
.ml-hold-files span { font-family: var(--mono); font-size: 11px; background: var(--sunken); border-radius: 4px; padding: 2px 7px; }
`;

let installed = false;
export function installArchiveStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-mail-archive-css';
  style.textContent = ARCHIVE_CSS;
  document.head.appendChild(style);
}
