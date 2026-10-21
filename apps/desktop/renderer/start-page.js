// An app's entry page (the owner, 2026-10-21: "the entry page of
// applications should show recent files opened by the app beside current
// new and starter templates", and "user should be able to remove the
// already added ones from recent"). Shown when an app is opened as an app,
// from its card on the launcher or its taskbar entry, rather than for a new
// document: New and the app's starter templates on one side, the files
// recently opened in this app on the other, each to open again or take off
// the list (its x, the Delete key, or Remove from the list on its menu).

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Button, Icon, Input, formatWhen, useMenu, t } from '@rutba/office-ui';
import { APPS } from '@rutba/office-formats/registry';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { pickOpen, openInApp } from './shell.js';
import { TemplateGallery } from './template-gallery.js';

/** How many starter templates sit beside Blank before More templates. */
const STARTERS = 7;

/** Whether a recent entry is this app's: kept with its app, or, for an older entry kept without one, its file's kind. */
const isFor = (r, app) => (r.app ? r.app === app : appFor(kindFromExtension(r.path)) === app);
const folderOf = (p) => String(p).replace(/[\\/][^\\/]*$/, '');

/**
 * Whether this window opened as its app (the launcher's card, the taskbar
 * entry) with nothing asked of it: no file, no template, nothing recovered
 * or handed over. Then it shows its entry page rather than a new document.
 */
export function wantsStart() {
  const p = new URLSearchParams(location.search);
  return p.get('start') === '1' && !['file', 'template', 'templateFile', 'recover', 'session'].some((k) => p.get(k));
}

/** Reload this window making a new document from `template` ('blank' for an empty one). */
export function startFrom(template) {
  const params = new URLSearchParams(location.search);
  params.delete('start');
  params.set('template', template);
  location.search = params.toString();
}

/**
 * The files recently opened in `app`, to open again (`onOpen(path)`; by
 * default in a window of their own, this window then closed) or take off the
 * list. `open` adds an Open button beside the search; `title` heads it.
 */
export function RecentFiles({ shell, app, onOpen = null, open = true, title = null, embedded = false }) {
  const meta = APPS[app] || { short: app };
  const [recent, setRecent] = useState(null);
  const [query, setQuery] = useState('');
  const menu = useMenu();
  const keep = (list) => setRecent(list.filter((r) => isFor(r, app)));
  useEffect(() => { shell.app.recent().then(keep).catch(() => setRecent([])); }, [shell, app]);
  const remove = (p) => shell.app.removeRecent({ path: p }).then(keep).catch(() => {});
  const openRecent = onOpen || ((p) => { openInApp(shell, p, app); setTimeout(() => Promise.resolve(shell.win.close({})).catch(() => {}), 300); });
  const openOther = async () => {
    const file = await pickOpen(shell, app);
    if (file) openRecent(file);
  };
  const shown = useMemo(() => {
    const q = query.trim().toLowerCase();
    return (recent || []).filter((r) => !q || r.name.toLowerCase().includes(q) || r.path.toLowerCase().includes(q));
  }, [recent, query]);
  return (
    <section className={`sp-recent${embedded ? ' sp-embed' : ''}`} aria-label={t('Recent')} data-recent={app}>
      <style>{START_PAGE_CSS}</style>
      <header>
        <h2>{title || t('Recent')}</h2>
        <span className="grow" />
        {recent?.length > 4 ? <Input className="rw-input sp-search" aria-label={t('Search recent files')} placeholder={t('Search recent files')} value={query} onChange={(e) => setQuery(e.target.value)} /> : null}
        {open ? <Button icon="open" label={t('Open…')} className="sp-open" onClick={openOther} /> : null}
      </header>
      {recent === null ? null : shown.length ? (
        <ul className="sp-list" role="list">
          {shown.map((r) => (
            <li key={r.path} className="sp-row" data-path={r.path}>
              <button type="button" className="sp-row-main" onClick={() => openRecent(r.path)}
                onKeyDown={(e) => { if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); remove(r.path); } }}
                onContextMenu={(e) => { e.preventDefault(); menu.open(e, [{ label: t('Open'), icon: 'open', run: () => openRecent(r.path) }, { label: t('Remove from the list'), icon: 'close', run: () => remove(r.path) }]); }}>
                <span className="sp-row-icon" data-app={app}><Icon name={app} size={16} /></span>
                <span className="sp-row-words"><strong>{r.name}</strong><small>{folderOf(r.path)}</small></span>
                <span className="sp-row-when">{formatWhen(r.at)}</span>
              </button>
              <button type="button" className="sp-remove" title={t('Remove from the list')} aria-label={t('Remove {name} from the list', { name: r.name })} onClick={() => remove(r.path)}>
                <Icon name="close" size={14} />
              </button>
            </li>
          ))}
        </ul>
      ) : (
        <p className="sp-none">{query ? t('Nothing matches') : t('Files you open in {app} appear here.', { app: meta.short })}</p>
      )}
      {menu.node}
    </section>
  );
}

export function StartPage({ shell, app, onBlank = () => startFrom('blank') }) {
  const [gallery, setGallery] = useState([]);
  const [more, setMore] = useState(false);
  useEffect(() => {
    let live = true;
    Promise.resolve(shell.doc.gallery?.({ app })).then((g) => { if (live) setGallery((g?.templates || []).filter((x) => !x.needs)); }).catch(() => {});
    return () => { live = false; };
  }, [shell, app]);

  return (
    <div className="sp" data-start={app}>
      <section className="sp-new" aria-label={t('New')}>
        <h2>{t('New')}</h2>
        <div className="sp-cards">
          <button type="button" className="sp-card sp-blank" data-template="blank" onClick={onBlank}>
            <span className={`sp-thumb ${app}`}><Icon name="plus" size={26} /></span>
            <span className="sp-name">{t('Blank')}</span>
          </button>
          {gallery.slice(0, STARTERS).map((tpl) => (
            <button key={tpl.id} type="button" className="sp-card" data-template={tpl.id} data-tip={`${t(tpl.name)}: ${t(tpl.description)}`} onClick={() => startFrom(tpl.id)}>
              <span className={`sp-thumb ${app}`}>{tpl.thumbnail ? <img src={tpl.thumbnail} alt="" draggable={false} /> : <Icon name={app} size={22} />}</span>
              <span className="sp-name">{t(tpl.name)}</span>
            </button>
          ))}
          <button type="button" className="sp-card sp-more" onClick={() => setMore(true)}>
            <span className="sp-thumb sp-more-thumb"><Icon name="grid" size={22} /></span>
            <span className="sp-name">{t('More templates')}</span>
          </button>
        </div>
      </section>

      <RecentFiles shell={shell} app={app} />
      {more ? <TemplateGallery shell={shell} app={app} onClose={() => setMore(false)} /> : null}
    </div>
  );
}

export const START_PAGE_CSS = `
.sp { flex: 1; min-height: 0; overflow: auto; display: grid; grid-template-columns: minmax(0, 1.25fr) minmax(320px, 1fr); gap: 28px; padding: 28px 32px; box-sizing: border-box; align-content: start; }
@media (max-width: 1000px) { .sp { grid-template-columns: 1fr; } }
.sp h2 { margin: 0 0 12px; font-size: 12px; font-weight: 600; letter-spacing: .04em; text-transform: uppercase; color: var(--ink-3); }
.sp-cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(132px, 1fr)); gap: 14px; }
.sp-card { display: flex; flex-direction: column; gap: 6px; padding: 0; border: 0; background: none; text-align: start; cursor: pointer; font: inherit; color: var(--ink); }
.sp-thumb { display: grid; place-items: center; aspect-ratio: 4 / 3; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); overflow: hidden; color: var(--ink-3); transition: border-color .12s, box-shadow .12s, transform .12s; }
.sp-thumb img { width: 100%; height: 100%; object-fit: cover; }
.sp-blank .sp-thumb { color: var(--accent); }
.sp-card:hover .sp-thumb, .sp-card:focus-visible .sp-thumb { border-color: var(--accent-line, var(--accent)); box-shadow: var(--shadow-2); transform: translateY(-1px); }
.sp-card:focus-visible { outline: none; }
.sp-card:focus-visible .sp-thumb { outline: 2px solid var(--accent); outline-offset: 2px; }
.sp-name { font-size: 12.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sp-recent header { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
.sp-recent header h2 { margin: 0; }
.sp-search { width: 200px; }
.sp-list { list-style: none; margin: 0; padding: 0; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); }
.sp-row { display: flex; align-items: center; border-bottom: 1px solid var(--line-soft); }
.sp-row:last-child { border-bottom: 0; }
.sp-row-main { flex: 1; min-width: 0; display: flex; align-items: center; gap: 10px; padding: 9px 12px; border: 0; background: none; text-align: start; font: inherit; color: var(--ink); cursor: pointer; }
.sp-row-main:hover, .sp-row-main:focus-visible { background: var(--hover); outline: none; }
.sp-row-icon { flex: none; display: grid; place-items: center; width: 28px; height: 28px; border-radius: var(--r-1, 6px); background: var(--sunken); color: var(--accent); }
.sp-row-words { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.sp-row-words strong { font-weight: 500; font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sp-row-words small { font-size: 11.5px; color: var(--ink-3); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; direction: ltr; }
.sp-row-when { flex: none; font-size: 12px; color: var(--ink-3); }
.sp-remove { flex: none; display: grid; place-items: center; width: 32px; height: 32px; margin-inline-end: 6px; border: 0; border-radius: var(--r-1, 6px); background: none; color: var(--ink-3); cursor: pointer; opacity: 0; }
.sp-row:hover .sp-remove, .sp-remove:focus-visible { opacity: 1; }
.sp-remove:hover { background: var(--hover); color: var(--ink); }
.sp-embed { width: min(640px, 100%); margin: 20px auto 0; text-align: start; }
.sp-none { margin: 0; padding: 16px; border: 1px dashed var(--line-soft); border-radius: var(--r-2); font-size: 13px; color: var(--ink-3); }
`;
