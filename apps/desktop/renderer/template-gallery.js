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
import { templatesFolder } from './shell.js';
import { useChosenApps } from './apps-choice.js';

/**
 * The gallery's words, listed so the message catalogue carries them: each
 * template's name and description and each category, as packages/templates
 * writes them in English (tests/templates.test.js holds the two lists
 * together).
 */
export const GALLERY_WORDS = [
  msg('Business'), msg('Finance'), msg('Planning'), msg('Education'), msg('Career'), msg('Household'),
  msg('Letter'), msg('A business letter: your details, the date, the address, a subject line and a signature block'),
  msg('CV'), msg('A one-page CV: a profile, experience, education and skills'),
  msg('Report'), msg('A report with a cover page, a table of contents and numbered sections'),
  msg('Invoice'), msg('An invoice to print or send: who it is from and to, the items, the total and how to pay'),
  msg('Meeting minutes'), msg('Minutes of a meeting: who came, what was discussed and decided, and who does what next'),
  msg('Newsletter'), msg('A newsletter with a masthead, two columns of stories and a dates panel'),
  msg('Monthly budget'), msg('Income and spending, planned against actual, with charts of where the money goes'),
  msg('An invoice that adds itself up: quantities, prices, tax and the total due'),
  msg('Timesheet'), msg('A week of hours: start, finish and breaks, with the hours and pay worked out'),
  msg('Loan calculator'), msg('The monthly payment, the total interest and a year-by-year schedule for a loan'),
  msg('To-do list'), msg('A table of tasks with owners, due dates, priorities and progress'),
  msg('Accounts ledger'), msg('A simple cash book: money in, money out and the running balance'),
  msg('Pitch deck'), msg('The story of a business in eight slides: problem, product, market, traction, team and the ask'),
  msg('Project status'), msg('Where a project stands: progress, milestones, risks and next steps'),
  msg('Photo album'), msg('Pictures with captions, one large and in pairs, on a dark theme'),
  msg('Training deck'), msg('A course session: objectives, the agenda, the lessons, a quick check and a summary'),
  msg('Cover letter'), msg('A letter to go with a job application: the role, what you bring and why this company'),
  msg('Memo'), msg('A short note to a team: to, from, the subject, the background and what happens now'),
  msg('Meeting agenda'), msg('What a meeting is for, its items with times and who leads each, and what to read first'),
  msg('Project proposal'), msg('A proposal with a summary, the need, the plan, the cost and the next steps'),
  msg('Expense report'), msg('Expenses to claim: dates, categories, VAT and receipts, with the total owed to you'),
  msg('Stock list'), msg('Items, counts and reorder levels, the stock value and what to order again'),
  msg('Weekly planner'), msg('A week by the hour, the dates filled in, with priorities and notes for the week'),
  msg('Grade book'), msg('Marks for a class: averages, grades and who needs support, with a chart'),
  msg('Quarterly review'), msg('The results of a quarter: the numbers, revenue by month, what went well, lessons and goals'),
  msg('Lesson'), msg('A lesson for a class: goals, a starter, the main idea, key words, a task and an exit ticket'),
  msg('Portfolio'), msg('Selected work on a dark theme: about you, three projects with pictures, and what clients say'),
];

const APP_ORDER = ['word', 'sheets', 'slides'];
const APP_NAMES = { word: msg('Documents'), sheets: msg('Worksheets'), slides: msg('Presentations') };
const BLANK = { word: msg('Blank document'), sheets: msg('Blank worksheet'), slides: msg('Blank presentation') };
const EXT_APP = { dotx: 'word', xltx: 'sheets', potx: 'slides' };

/** Templates of one's own, from the folder Office keeps them in: [{ path, name, app }]. */
export async function personalTemplates(shell) {
  try {
    const dir = await templatesFolder(shell);
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

  const shown = useMemo(() => {
    if (!data) return [];
    const q = query.trim().toLocaleLowerCase();
    return data.templates.filter((tpl) => {
      if (view === 'personal' || !kept(tpl.app)) return false;
      if (APP_ORDER.includes(view) && tpl.app !== view) return false;
      if (view !== 'all' && !APP_ORDER.includes(view) && !tpl.categories.includes(view)) return false;
      if (!q) return true;
      return [t(tpl.name), t(tpl.description), tpl.name, tpl.description].some((s) => s.toLocaleLowerCase().includes(q));
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, view, query, chosen.apps]);

  const ownShown = useMemo(() => {
    const q = query.trim().toLocaleLowerCase();
    return view === 'personal' || (q && view === 'all') ? personal.filter((p) => kept(p.app) && (!q || p.name.toLocaleLowerCase().includes(q))) : [];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [personal, view, query, chosen.apps]);

  const rail = [
    { id: 'all', label: t('All templates'), icon: 'grid' },
    ...(app ? [] : APP_ORDER.filter(kept).map((a) => ({ id: a, label: t(APP_NAMES[a]), icon: a }))),
    ...(data?.categories || []).filter((c) => data.templates.some((tpl) => tpl.categories.includes(c.id))).map((c) => ({ id: c.id, label: t(c.name), heading: c === data.categories[0] ? t('Categories') : null })),
    { id: 'personal', label: t('Personal'), icon: 'star', heading: t('Yours'), count: personal.length },
  ];

  const title = app ? t(APP_NAMES[app]) : t('New from a template');
  const blankFirst = app && view === 'all' && !query.trim();
  const nothing = data && !shown.length && !ownShown.length && !blankFirst;

  return (
    <Dialog title={null} onClose={onClose} className="tg-dialog">
      <style>{CSS}</style>
      <div className="tg" data-app={app || 'home'}>
        <div className="tg-head">
          <h2>{app ? t('New') : t('Templates')}</h2>
          <span className="tg-sub">{title}</span>
          <Search value={query} onChange={setQuery} placeholder={t('Search templates')} style={{ width: 260 }} autoFocus />
          <button type="button" className="tg-close" data-tip={t('Close the gallery')} aria-label={t('Close the gallery')} onClick={onClose}>
            <Icon name="close" size={16} />
          </button>
        </div>
        <div className="tg-main">
          <nav className="tg-rail" aria-label={t('Template categories')}>
            {rail.map((item) => (
              <React.Fragment key={item.id}>
                {item.heading ? <div className="tg-rail-head">{item.heading}</div> : null}
                <button type="button" className={`tg-rail-item${view === item.id ? ' on' : ''}`} data-view={item.id} onClick={() => setView(item.id)}>
                  {item.icon ? <Icon name={item.icon} size={15} /> : <i className="tg-dot" />}
                  <span>{item.label}</span>
                  {item.count ? <small>{item.count}</small> : null}
                </button>
              </React.Fragment>
            ))}
          </nav>
          <div className="tg-grid-wrap">
            {!data ? <div className="tg-wait">{t('Loading the templates…')}</div> : null}
            {nothing ? (
              <Empty
                icon={view === 'personal' ? 'star' : 'find'}
                title={view === 'personal' ? t('No templates of your own yet') : t('No templates match')}
              >
                {view === 'personal'
                  ? t('Save a document as a template from the app menu (Save as Template) and it is listed here.')
                  : t('Try other words, or look in All templates.')}
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
                  <button
                    key={tpl.id}
                    type="button"
                    className="tg-card"
                    data-template={tpl.id}
                    data-tip={`${t(tpl.name)}: ${t(tpl.description)}`}
                    onClick={() => open({ app: tpl.app, query: { template: tpl.id } })}
                  >
                    <span className={`tg-thumb ${tpl.app}`}>
                      {tpl.thumbnail ? <img src={tpl.thumbnail} alt="" draggable={false} /> : <BlankPicture app={tpl.app} />}
                    </span>
                    <span className="tg-name">
                      <span className="tpl-glyph" data-app={tpl.app}><Icon name={tpl.app} size={12} /></span>
                      {t(tpl.name)}
                    </span>
                  </button>
                ))}
                {ownShown.map((p) => (
                  <button
                    key={p.path}
                    type="button"
                    className="tg-card personal"
                    data-template={p.path}
                    data-tip={t('Your template: a new document made from {name}', { name: p.name })}
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
