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
    from: t('From'),
    to: t('To'),
    cc: t('Cc'),
    bcc: t('Bcc'),
    date: t('Date'),
    messageId: t('Message-ID'),
    source: t('Source'),
    folder: t('Folder'),
    attachments: t('Attachments'),
    name: t('Name'),
    type: t('Type'),
    size: t('Size'),
    sha256: t('SHA-256'),
    noSubject: t('(no subject)'),
    noBody: t('(This message has no text.)'),
    summaryTitle: t('Export summary'),
    exported: t('Exported'),
    selection: t('What was exported'),
    messages: t('Messages'),
    dates: t('Dates'),
    sources: t('Sources'),
    files: t('Files'),
    untouched: t('The sources were opened for reading only; nothing in them was changed.'),
    cancelled: t('Cancelled before the end: only the messages listed in the manifest were exported.'),
    footer: t('Exported by Rutba Office'),
  };
}

/** The search language, as the help lists it: what to type, and what it finds. */
export function searchSyntax() {
  return [
    ['from:amina', t('Sent by someone, by name or address')],
    ['to:bo  cc:legal  bcc:', t('Sent to someone, copied, or blind copied')],
    ['subject:contract', t('A word in the subject')],
    ['"indemnity clause"', t('These words, together and in this order, anywhere in the message')],
    ['has:attachment', t('Carries a file')],
    ['filename:pdf', t('A file whose name holds this')],
    ['is:unread  is:read  is:starred', t('Read or not, starred or not')],
    ['after:2019-01-01  before:2020-01-01', t('Sent on or after a day, or before it')],
    ['received:2019-03  received:>=2019-03', t('In a month or a year, or with >, >=, < and <=')],
    ['received:2019-01-01..2019-06-30', t('Between two days')],
    ['older_than:1y  newer_than:30d', t('Older or newer than so many days (d), weeks (w), months (m) or years (y)')],
    ['size>1mb  larger:5m  smaller:100kb', t('Larger or smaller than a size')],
    ['in:projects', t('Filed in a folder')],
    ['account:northwind', t('In one account or archive')],
    ['invoice OR receipt', t('Either')],
    ['-from:newsletter', t('Leave out what follows the minus')],
  ];
}

export function SearchHelp({ onClose, onInsert }) {
  return (
    <Dialog
      title={t('Search syntax')}
      width={620}
      onClose={onClose}
      actions={<Button primary label={t('Close')} onClick={onClose} />}
    >
      <p className="ml-help-lead">{t('Type words to find them anywhere, the body included. Add any of these; click one to add it to the search.')}</p>
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
  if (kind === 'pst') return t('Outlook data file (.pst)');
  if (kind === 'ost') return t('Outlook offline file (.ost)');
  if (kind === 'olm') return t('Outlook for Mac archive (.olm)');
  if (kind === 'mbox') return t('Mbox archive');
  return kind || '';
}

/** How far an archive's indexing has got, in words: "40,000 of 100,000". */
export const indexedCount = (index) => t('{done} of {total}', { done: (index?.done || 0).toLocaleString(), total: (index?.total || 0).toLocaleString() });

/**
 * The strip at the top of an archive's message list: what it is, that it is
 * read only (or how far its indexing has got), and Export, Stop or Resume
 * indexing, and Close.
 */
export function ArchiveBanner({ archive, onClose, onExport, onStopIndexing, onResumeIndexing }) {
  const indexing = archive.index?.state === 'indexing';
  const paused = archive.index?.state === 'paused';
  const note = archive.missing
    ? t('The file is no longer where it was.')
    : indexing
      ? t('Read only, opened where it lies. Indexing for search: {count} messages, searchable as they are indexed.', { count: indexedCount(archive.index) })
      : paused
        ? t('Read only, opened where it lies. Indexing stopped at {count} messages; what is indexed can be searched.', { count: indexedCount(archive.index) })
        : t('Read only, opened where it lies. Nothing in it is changed.');
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
          {t('Export…')}
        </button>
        {indexing && onStopIndexing ? (
          <button type="button" className="ml-archive-stop" onClick={onStopIndexing}>
            {t('Stop indexing')}
          </button>
        ) : null}
        {paused && onResumeIndexing ? (
          <button type="button" className="ml-archive-resume" onClick={onResumeIndexing}>
            {t('Resume indexing')}
          </button>
        ) : null}
        <button type="button" onClick={onClose}>
          {t('Close archive')}
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
    const [picked] = await shell.dialog.open({ title: t('Export to'), directory: true, defaultPath: dir || undefined });
    if (picked) setDir(picked);
  };

  const run = async () => {
    setState('running');
    setError(null);
    setProgress({ done: 0, total: count });
    shell.store.set({ key: 'mail.exportDir', value: dir }).catch(() => {});
    const selection =
      what === 'results'
        ? t('The search {query}', { query })
        : what === 'folder'
          ? t('The folder {folder} of {source}', { folder: folderOf?.folder || '', source: sourceName || '' })
          : tn(chosen.size, '{count} folder of {source}', '{count} folders of {source}', { source: sourceName || '' });
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
    if (r.mbox) parts.push(t('an mbox'));
    if (r.combinedPages) parts.push(tn(r.combinedPages, 'a PDF of {count} page', 'a PDF of {count} pages'));
    if (r.pdfs) parts.push(tn(r.pdfs, '{count} PDF, one a message', '{count} PDFs, one a message'));
    if (r.attachments) parts.push(tn(r.attachments, '{count} attachment', '{count} attachments'));
    parts.push(t('a manifest with every hash, and a summary'));
    return [
      r.cancelled
        ? t('Stopped after {done} of {total}. What was exported is kept, and the manifest says it stopped.', { done: r.exported.toLocaleString(), total: r.requested.toLocaleString() })
        : tn(r.exported, 'Exported {count} message from {sources} to {dir}:', 'Exported {count} messages from {sources} to {dir}:', {
            sources: tn(r.sources, '{count} mailbox or archive', '{count} mailboxes and archives'),
            dir: r.dir,
          }),
      `${parts.join(', ')}.`,
      t('Nothing in the mailboxes or archives was changed.'),
    ].join(' ');
  };

  if (state === 'done' && result) {
    return (
      <Dialog
        title={t('Export for legal hold')}
        width={560}
        onClose={onClose}
        actions={
          <>
            <Button label={t('Show the folder')} icon="folderOpen" onClick={() => shell.shell.openPath({ path: result.dir })} />
            <Button primary label={t('Close')} onClick={onClose} />
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
      title={t('Export for legal hold')}
      width={580}
      className="ml-hold"
      onClose={running ? undefined : onClose}
      actions={
        running ? (
          <Button label={t('Cancel')} onClick={() => shell.mail.cancelExport({ job: job.current })} />
        ) : (
          <>
            <Button label={t('Cancel')} onClick={onClose} />
            <Button primary label={tn(count, 'Export {count} message', 'Export {count} messages')} disabled={!count || !dir || (!mbox && pdf === 'none')} onClick={run} />
          </>
        )
      }
    >
      {running ? (
        <div className="ml-hold-progress">
          <p>{t('Exported {done} of {total}', { done: progress.done.toLocaleString(), total: (progress.total || count).toLocaleString() })}</p>
          <Progress value={progress.done} max={progress.total || count || 1} />
        </div>
      ) : (
        <>
          <div className="ml-hold-section">{t('What to export')}</div>
          <div className="ml-hold-choices" role="radiogroup">
            {query && resultCount ? (
              <label className="ml-hold-choice">
                <input type="radio" name="hold-what" checked={what === 'results'} onChange={() => setWhat('results')} />
                <span>{tn(resultCount, 'The {count} result of the search {query}', 'The {count} results of the search {query}', { query })}</span>
              </label>
            ) : null}
            {folderOf?.folder ? (
              <label className="ml-hold-choice">
                <input type="radio" name="hold-what" checked={what === 'folder'} onChange={() => setWhat('folder')} />
                <span>{t('This folder: {folder}', { folder: folderOf.folder })}</span>
              </label>
            ) : null}
            {folders?.length > 1 ? (
              <label className="ml-hold-choice">
                <input type="radio" name="hold-what" checked={what === 'folders'} onChange={() => setWhat('folders')} />
                <span>{t('Folders of {source}', { source: sourceName || '' })}</span>
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

          <div className="ml-hold-section">{t('Write')}</div>
          <label className="ml-hold-check">
            <input type="checkbox" className="ml-check" checked={mbox} onChange={(e) => setMbox(e.target.checked)} />
            <span>{t('An mbox file of the messages, as mail programs read them')}</span>
          </label>
          <div className="ml-hold-row">
            <span>{t('PDF')}</span>
            <Select value={pdf} onChange={(e) => setPdf(e.target.value)} aria-label={t('PDF')}>
              <option value="combined">{t('One PDF of them all, after a summary page')}</option>
              <option value="each">{t('One PDF for each message')}</option>
              <option value="both">{t('Both')}</option>
              <option value="none">{t('No PDF')}</option>
            </Select>
          </div>
          <label className="ml-hold-check">
            <input type="checkbox" className="ml-check" checked={attachments} onChange={(e) => setAttachments(e.target.checked)} />
            <span>{t("Write each message's attachments beside it")}</span>
          </label>

          <Field label={t('Into a new folder in')}>
            <div className="ml-hold-dir">
              <Input value={dir} onChange={(e) => setDir(e.target.value)} placeholder={t('Choose a folder')} dir="ltr" />
              <Button label={t('Browse…')} icon="folderOpen" onClick={browse} />
            </div>
          </Field>
          <p className="rw-hint ml-hold-note">
            {t('A manifest (CSV and JSON) with each message\'s date, sender, recipients, subject, Message-ID, where it came from and the SHA-256 of what was written, and a summary page, are always written. Nothing in the mailboxes or archives is changed.')}
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
