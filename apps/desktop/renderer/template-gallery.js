// File > New: the template gallery, as Word, Excel and PowerPoint open with one.
//
// One gallery for the launcher (every app's templates, by app and by
// category, with a search) and for each app's own File > New (that app's
// templates, a blank one first). The templates are the suite's own, made by
// its engines and pictured by it (packages/templates); a template of one's
// own, kept in Documents\Custom Office Templates, is listed under Personal.
// Choosing one opens a new, untitled copy in a window of its own; the
// template itself is never opened.

import React, { useEffect, useMemo, useState } from 'react';
import { Icon, Search, Dialog, Empty, t, msg } from '@rutba/office-ui';
import { templatesFolder, useUpdate } from './shell.js';
import { useChosenApps } from './apps-choice.js';

/**
 * The gallery's words, listed so the message catalogue carries them: each
 * template's name and description and each category, as packages/templates
 * writes them in English (tests/templates.test.js holds the two lists
 * together).
 */
export const GALLERY_WORDS = [
  msg('template-gallery.business'), msg('template-gallery.finance'), msg('template-gallery.planning'), msg('template-gallery.education'), msg('template-gallery.career'), msg('template-gallery.household'),
  msg('common.letter'), msg('template-gallery.business-letter-details'),
  msg('template-gallery.cv'), msg('template-gallery.one-page-cv'),
  msg('common.report'), msg('template-gallery.report-cover-page'),
  msg('common.invoice'), msg('template-gallery.invoice-print-send'),
  msg('template-gallery.meeting-minutes'), msg('template-gallery.minutes-meeting-came'),
  msg('template-gallery.newsletter'), msg('template-gallery.newsletter-masthead-two'),
  msg('template-gallery.monthly-budget'), msg('template-gallery.income-spending-planned'),
  msg('template-gallery.invoice-adds-itself'),
  msg('template-gallery.timesheet'), msg('template-gallery.week-hours-start'),
  msg('template-gallery.loan-calculator'), msg('template-gallery.monthly-payment-total'),
  msg('template-gallery.to-do-list'), msg('template-gallery.table-tasks-owners'),
  msg('template-gallery.accounts-ledger'), msg('template-gallery.simple-cash-book'),
  msg('common.pitch-deck'), msg('template-gallery.story-business-eight'),
  msg('template-gallery.project-status'), msg('template-gallery.project-stands-progress'),
  msg('template-gallery.photo-album'), msg('template-gallery.pictures-captions-one'),
  msg('template-gallery.training-deck'), msg('template-gallery.course-session'),
  msg('template-gallery.cover-letter'), msg('template-gallery.letter-go-job'),
  msg('template-gallery.memo'), msg('template-gallery.short-note-team'),
  msg('template-gallery.meeting-agenda'), msg('template-gallery.meeting-items-times'),
  msg('template-gallery.project-proposal'), msg('template-gallery.proposal-summary-need'),
  msg('template-gallery.expense-report'), msg('template-gallery.expenses-claim-dates'),
  msg('template-gallery.stock-list'), msg('template-gallery.items-counts-reorder'),
  msg('template-gallery.weekly-planner'), msg('template-gallery.week-hour-dates'),
  msg('template-gallery.grade-book'), msg('template-gallery.marks-class-averages'),
  msg('template-gallery.quarterly-review'), msg('template-gallery.results-quarter-numbers'),
  msg('template-gallery.lesson'), msg('template-gallery.lesson-class-goals'),
  msg('template-gallery.portfolio'), msg('template-gallery.selected-work-dark'),
];

/** The site's page of templates, opened in the browser by More templates. */
export const MORE_TEMPLATES_URL = 'https://office.rutba.io/templates';

const APP_ORDER = ['word', 'sheets', 'slides'];
const APP_NAMES = { word: msg('common.documents'), sheets: msg('common.worksheets'), slides: msg('common.presentations') };
const BLANK = { word: msg('common.blank-document'), sheets: msg('common.blank-worksheet'), slides: msg('common.blank-presentation') };
const EXT_APP = { dotx: 'word', xltx: 'sheets', potx: 'slides' };

/** Templates of one's own, from the folder Office keeps them in: [{ path, name, app }]. */
export async function personalTemplates(shell) {
  try {
    const dir = await templatesFolder(shell);
    // No folder yet is no templates yet: asked to list it, the main process would log an error.
    if (!(await shell.fs.exists({ path: dir }))) return [];
    const list = await shell.fs.list({ path: dir, filter: ['dotx', 'xltx', 'potx'] });
    return (list || [])
      .filter((f) => !f.dir)
      .map((f) => ({ path: f.path, name: String(f.name).replace(/\.[^.]+$/, ''), app: EXT_APP[String(f.name).split('.').pop().toLowerCase()] }))
      .filter((f) => f.app);
  } catch {
    return [];
  }
}

/** A blank page, sheet or slide, drawn rather than pictured. */
function BlankPicture({ app }) {
  return (
    <span className={`tg-blank ${app}`} aria-hidden="true">
      {app === 'sheets' ? <span className="tg-blank-grid" /> : null}
    </span>
  );
}

/**
 * The gallery. `app` limits it to one app's templates (File > New in that
 * app), a blank one first; without it, every app's (the launcher).
 * `onClose` takes it away; choosing a template opens its copy and closes it.
 */
export function TemplateGallery({ shell, app = null, onClose }) {
  const [data, setData] = useState(null);
  const [personal, setPersonal] = useState([]);
  const [query, setQuery] = useState('');
  // What the rail has chosen: 'all', an app, a category or 'personal'.
  const [view, setView] = useState('all');
  // From the launcher, only the apps this person chose (Choose apps).
  const chosen = useChosenApps(shell);
  // A newer template set the update check found (template-sets.js): ready for the next start, or waiting on a newer Office.
  const sets = useUpdate(shell)?.templates || null;
  const kept = (key) => Boolean(app) || chosen.apps.includes(key);

  useEffect(() => {
    let live = true;
    shell.doc.gallery({ app }).then((d) => live && setData(d)).catch(() => live && setData({ categories: [], templates: [] }));
    personalTemplates(shell).then((list) => live && setPersonal(app ? list.filter((p) => p.app === app) : list));
    return () => { live = false; };
  }, [shell, app]);

  const open = (query) => {
    shell.win.create(query);
    onClose?.();
  };

  // A search covers every template, the gallery's and one's own, whatever
  // the rail has chosen, as Word's does; with nothing typed, the rail's choice.
  const searching = Boolean(query.trim());
  const shown = useMemo(() => {
    if (!data) return [];
    const q = query.trim().toLocaleLowerCase();
    return data.templates.filter((tpl) => {
      if (!kept(tpl.app)) return false;
      if (q) return [t(tpl.name), t(tpl.description), tpl.name, tpl.description].some((s) => s.toLocaleLowerCase().includes(q));
      if (view === 'personal') return false;
      if (APP_ORDER.includes(view) && tpl.app !== view) return false;
      if (view !== 'all' && !APP_ORDER.includes(view) && !tpl.categories.includes(view)) return false;
      return true;
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, view, query, chosen.apps]);

  const ownShown = useMemo(() => {
    const q = query.trim().toLocaleLowerCase();
    if (q) return personal.filter((p) => kept(p.app) && p.name.toLocaleLowerCase().includes(q));
    return view === 'personal' ? personal.filter((p) => kept(p.app)) : [];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [personal, view, query, chosen.apps]);

  const rail = [
    { id: 'all', label: t('template-gallery.all-templates'), icon: 'grid' },
    ...(app ? [] : APP_ORDER.filter(kept).map((a) => ({ id: a, label: t(APP_NAMES[a]), icon: a }))),
    ...(data?.categories || []).filter((c) => data.templates.some((tpl) => tpl.categories.includes(c.id))).map((c) => ({ id: c.id, label: t(c.name), heading: c === data.categories[0] ? t('common.categories') : null })),
    { id: 'personal', label: t('template-gallery.personal'), icon: 'star', heading: t('common.yours'), count: personal.length },
  ];

  const title = app ? t(APP_NAMES[app]) : t('template-gallery.new-template');
  const blankFirst = app && view === 'all' && !query.trim();
  const nothing = data && !shown.length && !ownShown.length && !blankFirst;

  return (
    <Dialog title={null} onClose={onClose} className="tg-dialog">
      <style>{CSS}</style>
      <div className="tg" data-app={app || 'home'}>
        <div className="tg-head">
          <h2>{app ? t('common.new') : t('template-gallery.templates')}</h2>
          <span className="tg-sub">{title}</span>
          <Search value={query} onChange={setQuery} placeholder={t('template-gallery.search-templates')} style={{ width: 260 }} autoFocus />
          <button type="button" className="tg-close" data-tip={t('template-gallery.close-the-gallery')} aria-label={t('template-gallery.close-the-gallery')} onClick={onClose}>
            <Icon name="close" size={16} />
          </button>
        </div>
        {sets?.phase === 'ready' ? <div className="tg-set" data-set-phase="ready">{t('template-gallery.newer-templates-set', { set: sets.set })}</div> : null}
        {sets?.phase === 'needs-office' ? <div className="tg-set" data-set-phase="needs-office">{t('template-gallery.newer-templates-set-need', { set: sets.set, version: sets.needs })}</div> : null}
        <div className="tg-main">
          <nav className="tg-rail" aria-label={t('template-gallery.template-categories')}>
            {rail.map((item) => (
              <React.Fragment key={item.id}>
                {item.heading ? <div className="tg-rail-head">{item.heading}</div> : null}
                <button type="button" className={`tg-rail-item${view === item.id && !searching ? ' on' : ''}`} data-view={item.id} onClick={() => { setView(item.id); setQuery(''); }}>
                  {item.icon ? <Icon name={item.icon} size={15} /> : <i className="tg-dot" />}
                  <span>{item.label}</span>
                  {item.count ? <small>{item.count}</small> : null}
                </button>
              </React.Fragment>
            ))}
            {/* The site's Templates page, in the browser: the suite itself fetches nothing for it. */}
            <button type="button" className="tg-rail-item tg-more" data-tip={t('template-gallery.every-template-download')} onClick={() => shell.shell.openExternal({ url: MORE_TEMPLATES_URL }).catch(() => {})}>
              <Icon name="open" size={15} />
              <span>{t('template-gallery.templates-web')}</span>
            </button>
          </nav>
          <div className="tg-grid-wrap" data-searching={searching ? 'all' : undefined}>
            {!data ? <div className="tg-wait">{t('template-gallery.loading-the')}</div> : null}
            {searching && data && !nothing ? <div className="tg-searched" data-role="searched-all">{t('template-gallery.found-among-every')}</div> : null}
            {nothing ? (
              <Empty
                icon={view === 'personal' && !searching ? 'star' : 'find'}
                title={view === 'personal' && !searching ? t('template-gallery.no-templates-own') : t('template-gallery.no-templates-match')}
              >
                {view === 'personal' && !searching
                  ? t('template-gallery.save-document-template')
                  : t('template-gallery.try-other-words')}
              </Empty>
            ) : (
              <div className="tg-grid">
                {blankFirst ? (
                  <button type="button" className="tg-card" data-template="blank" onClick={() => open({ app, query: { template: 'blank' } })}>
                    <span className="tg-thumb"><BlankPicture app={app} /></span>
                    <span className="tg-name"><span className="tpl-glyph" data-app={app}><Icon name={app} size={12} /></span>{t(BLANK[app])}</span>
                  </button>
                ) : null}
                {shown.map((tpl) => (
                  <div key={tpl.id} className="tg-card-wrap">
                  <button
                    type="button"
                    className={`tg-card${tpl.needs ? ' needs' : ''}`}
                    data-template={tpl.id}
                    data-needs={tpl.needs || undefined}
                    data-tip={tpl.needs ? t('template-gallery.name-needs-rutba', { name: t(tpl.name), version: tpl.needs }) : `${t(tpl.name)}: ${t(tpl.description)}`}
                    onClick={() => (tpl.needs ? shell.update.check({ manual: true }).catch(() => {}) : open({ app: tpl.app, query: { template: tpl.id } }))}
                  >
                    <span className={`tg-thumb ${tpl.app}`}>
                      {tpl.thumbnail ? <img src={tpl.thumbnail} alt="" draggable={false} /> : <BlankPicture app={tpl.app} />}
                    </span>
                    <span className="tg-name">
                      <span className="tpl-glyph" data-app={tpl.app}><Icon name={tpl.app} size={12} /></span>
                      {t(tpl.name)}
                    </span>
                    {tpl.needs ? <span className="tg-needs">{t('common.needs-rutba-office', { version: tpl.needs })}</span> : null}
                  </button>
                  {tpl.page ? (
                    <button type="button" className="tg-page-link" data-page={tpl.page} title={t('common.about-name-online', { name: t(tpl.name) })} aria-label={t('common.about-name-online', { name: t(tpl.name) })}
                      onClick={() => shell.shell.openExternal({ url: tpl.page }).catch(() => {})}>
                      <Icon name="globe" size={14} />
                    </button>
                  ) : null}
                  </div>
                ))}
                {ownShown.map((p) => (
                  <button
                    key={p.path}
                    type="button"
                    className="tg-card personal"
                    data-template={p.path}
                    data-tip={t('common.template-new-document', { name: p.name })}
                    onClick={() => open({ app: p.app, query: { templateFile: p.path } })}
                  >
                    <span className={`tg-thumb own ${p.app}`}>
                      <span className="tg-own-glyph" data-app={p.app}><Icon name={p.app} size={30} /></span>
                    </span>
                    <span className="tg-name">
                      <span className="tpl-glyph" data-app={p.app}><Icon name={p.app} size={12} /></span>
                      {p.name}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </Dialog>
  );
}

const CSS = `
.rw-dialog.tg-dialog { width: min(1060px, 94vw); max-width: min(1060px, 94vw); height: min(720px, 88vh); }
.rw-dialog.tg-dialog .rw-dialog-body { padding: 0; flex: 1; min-height: 0; display: flex; overflow: hidden; }
.tg { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.tg-head {
  display: flex; align-items: center; gap: 12px; padding: 16px 18px 14px 22px;
  border-bottom: 1px solid var(--line-soft);
}
.tg-head h2 { margin: 0; font-family: var(--font-display); font-size: 18px; font-weight: 650; letter-spacing: -0.01em; }
.tg-card-wrap { position: relative; display: grid; }
.tg-page-link { position: absolute; top: 8px; inset-inline-end: 8px; display: grid; place-items: center; width: 26px; height: 26px; border: 1px solid var(--line-soft); border-radius: 50%; background: var(--surface); color: var(--ink-3); cursor: pointer; opacity: 0; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12); }
.tg-card-wrap:hover .tg-page-link, .tg-page-link:focus-visible { opacity: 1; }
.tg-page-link:hover { color: var(--accent-text, var(--accent)); }
.tg-set { padding: 6px 20px; font-size: 12.5px; color: var(--ink-2); background: var(--sunken); border-bottom: 1px solid var(--line-soft); }
.tg-searched { color: var(--ink-3); font-size: 12px; margin: 0 0 10px; }
.tg-sub { color: var(--ink-3); font-size: 12.5px; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.tg-close {
  border: 0; background: none; color: var(--ink-3); width: 30px; height: 30px; border-radius: var(--r-2);
  display: grid; place-items: center; cursor: pointer; flex: none;
}
.tg-close:hover { background: var(--hover); color: var(--ink); }
.tg-main { flex: 1; min-height: 0; display: flex; }
.tg-rail {
  width: 196px; flex: none; padding: 12px 10px; overflow: auto; display: flex; flex-direction: column; gap: 1px;
  border-inline-end: 1px solid var(--line-soft); background: var(--surface-2);
}
.tg-rail-head {
  margin: 12px 10px 4px; font-size: 10.5px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--ink-3);
}
.tg-rail-item {
  display: flex; align-items: center; gap: 9px; padding: 7px 10px; border: 0; border-radius: var(--r-2);
  background: none; color: var(--ink-2); font: inherit; font-size: 13px; text-align: start; cursor: pointer;
}
.tg-rail-item span { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.tg-rail-item small { color: var(--ink-3); font-size: 11px; }
.tg-rail-item:hover { background: var(--hover); color: var(--ink); }
.tg-rail-item.on { background: var(--selected); color: var(--accent-text); font-weight: 600; }
.tg-more { margin-top: auto; color: var(--accent-text); }
.tg-dot { width: 15px; height: 15px; display: grid; place-items: center; flex: none; }
.tg-dot::after { content: ''; width: 6px; height: 6px; border-radius: 50%; background: currentColor; opacity: 0.55; }
.tg-grid-wrap { flex: 1; min-width: 0; overflow: auto; padding: 18px 20px 22px; overscroll-behavior: contain; }
.tg-wait { color: var(--ink-3); font-size: 13px; padding: 20px 4px; }
.tg-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(176px, 1fr)); gap: 18px 16px; align-content: start; }
.tg-card {
  display: flex; flex-direction: column; gap: 9px; padding: 0; border: 0; background: none; color: var(--ink);
  font: inherit; text-align: start; cursor: pointer; min-width: 0;
}
.tg-thumb {
  aspect-ratio: 4 / 3; border-radius: var(--r-3); border: 1px solid var(--line); background: var(--sunken);
  display: grid; place-items: center; overflow: hidden; position: relative;
  transition: border-color var(--fast), box-shadow var(--fast), transform var(--fast);
}
.tg-card:hover .tg-thumb, .tg-card:focus-visible .tg-thumb { border-color: var(--accent-line); box-shadow: var(--shadow-2); transform: translateY(-1px); }
.tg-card:focus-visible { outline: none; }
.tg-card.needs .tg-thumb { opacity: .45; filter: grayscale(.6); }
.tg-needs { display: block; font-size: 11px; color: var(--accent-text); margin-top: 2px; }
.tg-card:focus-visible .tg-thumb { outline: 2px solid var(--accent); outline-offset: 2px; }
.tg-thumb img { max-width: calc(100% - 22px); max-height: calc(100% - 22px); object-fit: contain; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.18), 0 4px 14px rgba(0, 0, 0, 0.08); background: #fff; }
.tg-thumb.word img { max-height: calc(100% - 16px); }
.tg-name { display: flex; align-items: center; gap: 7px; font-size: 13px; font-weight: 550; min-width: 0; padding: 0 2px; }
.tg-name .tpl-glyph { width: 18px; height: 18px; border-radius: 5px; }
.tg .tpl-glyph { display: grid; place-items: center; color: #fff; flex: none; }
.tg .tpl-glyph[data-app='word'] { background: #2b5fd9; }
.tg .tpl-glyph[data-app='sheets'] { background: #0f9d58; }
.tg .tpl-glyph[data-app='slides'] { background: #d9534f; }
.tg-blank { display: block; background: #fff; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.18), 0 4px 14px rgba(0, 0, 0, 0.08); }
.tg-blank.word { height: calc(100% - 16px); aspect-ratio: 210 / 297; }
.tg-blank.slides { width: calc(100% - 22px); aspect-ratio: 16 / 9; }
.tg-blank.sheets { width: calc(100% - 22px); height: calc(100% - 22px); position: relative; }
.tg-blank-grid {
  position: absolute; inset: 0;
  background-image: linear-gradient(#e3e6ea 1px, transparent 1px), linear-gradient(90deg, #e3e6ea 1px, transparent 1px);
  background-size: 100% 18px, 54px 100%;
}
.tg-thumb.own { background: color-mix(in srgb, var(--accent) 7%, var(--sunken)); }
.tg-own-glyph { width: 64px; height: 64px; border-radius: 16px; display: grid; place-items: center; color: #fff; }
.tg-own-glyph[data-app='word'] { background: #2b5fd9; }
.tg-own-glyph[data-app='sheets'] { background: #0f9d58; }
.tg-own-glyph[data-app='slides'] { background: #d9534f; }
`;
