// The launcher.
//
// The first thing anybody sees, and the only window that is not about one
// document. It has three jobs: start something new, reopen something recent,
// and make it obvious what the seven apps are — because the whole point of a
// suite is that you did not have to go and find seven separate downloads.

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Icon, Search, Empty, Button, Chip, Spacer, Dialog, useMenu, useToast, formatBytes, formatWhen, basename, t, tn } from '@rutba/office-ui';
import { APPS, APP_GROUPS, NEW_DOCUMENTS, SITE } from '@rutba/office-formats/registry';
import { appFor, kindFromExtension, KINDS } from '@rutba/office-formats/sniff';
import { AppFrame, useAppMenu, pickOpen, openInApp, useFileDrop, templatesFolder } from '../shell.js';
import { WhatsNew } from '../whatsnew.js';
import { TemplateGallery } from '../template-gallery.js';
import { useChosenApps, AppsChooser, setWords } from '../apps-choice.js';
import { PacksNotice, PacksDialog } from '../packs-dialog.js';
import { useProfiles, profileName } from '../profiles.js';

/** The app a recovered copy opens in: a document's by its kind, a window's own (Video's or Image's project) by its name. */
const appOfRecovered = (r) => (r.kind === 'sheet' ? 'sheets' : r.kind === 'deck' ? 'slides' : r.kind === 'video' ? 'video' : r.kind === 'image' ? 'image' : r.kind === 'pdf' ? 'pdf' : 'word');

// A launcher card a build carries only when it was made with its switch.
const ExtraTile = process.env.RUTBA_LOCKER === '1' ? React.lazy(() => import('./locker-tile.js')) : null;
/** One of Start something's chips: an app's mark and what it starts. */
function StartChip({ app, icon, label, tip, className = '', onClick, template }) {
  return (
    <button type="button" className={`home-template${className ? ` ${className}` : ''}`} data-tip={tip} data-template={template} onClick={onClick}>
      <span className={`tpl-glyph${app ? '' : ' more'}`} data-app={app}>
        <Icon name={icon} size={16} />
      </span>
      <span>{label}</span>
    </button>
  );
}

function AppCard({ app, onOpen, onNew }) {
  return (
    <div className="home-card" data-app={app.key}>
      {/* The first line of what it does, and the whole of it on hover. */}
      <button type="button" className="home-card-main" data-tip={`${app.short}: ${app.tagline}`} onClick={onOpen}>
        <span className="home-card-glyph">
          <Icon name={app.icon} size={22} />
        </span>
        <span className="home-card-text">
          <strong>{app.short}</strong>
          <span>{app.tagline}</span>
        </span>
      </button>
      {onNew ? (
        <button type="button" className="home-card-new" onClick={onNew} title={t('New {kind}', { kind: app.noun || app.short.toLowerCase() })}>
          <Icon name="plus" size={14} />
        </button>
      ) : null}
    </div>
  );
}

export default function Home({ app, shell }) {
  const [recent, setRecent] = useState([]);
  const [query, setQuery] = useState('');
  const [version, setVersion] = useState(null);
  const [showAbout, setShowAbout] = useState(() => new URLSearchParams(location.search).has('about'));
  // The release notes, asked for from any window's app menu.
  const [showNotes, setShowNotes] = useState(() => new URLSearchParams(location.search).has('whatsnew'));
  // A right-click on a recent file.
  const rowMenu = useMenu();
  // The path currently turned into a text box, while it is being renamed.
  const [renaming, setRenaming] = useState(null);
  const toast = useToast();
  const [update, setUpdate] = useState(null);
  // What a crash took. The copies are written while a document is open and
  // deleted when it is saved or closed, so anything still here was never
  // either — and the person who lost it is looking at this window.
  const [recovered, setRecovered] = useState([]);

  // Templates of one's own, from the folder Office keeps them in (shell.js).
  const [personal, setPersonal] = useState([]);
  // The template gallery, over the launcher (?templates opens with it).
  const [gallery, setGallery] = useState(() => new URLSearchParams(location.search).has('templates'));
  // The apps this person chose, and the window to choose them in (?apps opens with it).
  const chosen = useChosenApps(shell);
  const [choosing, setChoosing] = useState(() => new URLSearchParams(location.search).has('apps'));
  const has = (key) => chosen.apps.includes(key);
  // The two sets with what is on in each; a set with nothing on leaves the other the whole width.
  const sets = APP_GROUPS.map((g) => ({ id: g.id, apps: g.apps.filter(has) })).filter((g) => g.apps.length);
  // An app another profile's menu sent the person to, turned off in this one (?off=, or told to the Home already open).
  const [offApp, setOffApp] = useState(() => new URLSearchParams(location.search).get('off'));
  useEffect(() => shell.on('app:command', ({ command, args } = {}) => { if (command === 'home.appOff' && args?.app) setOffApp(args.app); }), [shell]);
  // Add or remove apps (packs-dialog.js; ?packs opens with it).
  const [packing, setPacking] = useState(() => new URLSearchParams(location.search).has('packs'));
  // The other profiles, for a recent file's Open in (profiles.js).
  const { list: profiles, me } = useProfiles(shell);
  const others = (profiles?.profiles || []).filter((p) => !p.current);

  const refresh = useCallback(() => {
    shell.app.recent().then(setRecent).catch(() => setRecent([]));
    templatesFolder(shell)
      .then((dir) => shell.fs.list({ path: dir, filter: ['dotx', 'xltx', 'potx'] }))
      .then((list) => setPersonal((list || []).filter((f) => !f.dir).map((f) => ({ path: f.path, name: String(f.name).replace(/\.[^.]+$/, ''), app: { dotx: 'word', xltx: 'sheets', potx: 'slides' }[String(f.name).split('.').pop().toLowerCase()] })).filter((f) => f.app)))
      .catch(() => setPersonal([]));
  }, [shell]);

  useEffect(() => {
    refresh();
    shell.app.version().then(setVersion).catch(() => {});
    shell.update.state().then(setUpdate).catch(() => {});
    shell.doc.recoverable().then(setRecovered).catch(() => setRecovered([]));
  }, [refresh, shell]);

  useEffect(() => shell.on('update:state', setUpdate), [shell]);

  // Back to the window: the recent files and the templates of one's own read
  // again, so a template saved from Documents a moment ago is in the row.
  useEffect(() => {
    const again = () => { if (document.visibilityState !== 'hidden') refresh(); };
    window.addEventListener('focus', again);
    document.addEventListener('visibilitychange', again);
    return () => {
      window.removeEventListener('focus', again);
      document.removeEventListener('visibilitychange', again);
    };
  }, [refresh]);

  // Opened as the app: those with an entry page show it (New, the starters, their recent files).
  const openApp = useCallback((key) => shell.win.create({ app: key, query: { start: '1' } }), [shell]);

  const newDocument = useCallback(
    async (template) => {
      const spec = NEW_DOCUMENTS.find((n) => n.template === template.template && n.app === template.app) || template;
      shell.win.create({ app: spec.app, query: { template: spec.template } });
    },
    [shell]
  );

  const openFile = useCallback(async () => {
    const file = await pickOpen(shell, null);
    if (file) openInApp(shell, file);
    setTimeout(refresh, 400);
  }, [shell, refresh]);

  useFileDrop(
    useCallback(
      (files) => {
        for (const f of files) openInApp(shell, f);
        setTimeout(refresh, 400);
      },
      [shell, refresh]
    )
  );

  const menu = useAppMenu({ shell, appKey: 'home', onOpen: openFile });

  const removeRecent = useCallback(
    (p) => shell.app.removeRecent({ path: p }).then(setRecent).catch(() => {}),
    [shell]
  );

  const startRename = useCallback((r) => setRenaming(r.path), []);

  // Applies a rename typed into the row's own text box. The backend decides
  // what a refusal says — the name is taken, the file has gone, it is open
  // somewhere — this only shows it and leaves the row as it was.
  const applyRename = useCallback(
    async (p, name) => {
      setRenaming(null);
      if (!name || !name.trim()) return;
      try {
        setRecent(await shell.app.renameRecent({ path: p, name: name.trim() }));
      } catch (err) {
        toast(err?.message || t('That name did not work.'), { tone: 'bad' });
      }
    },
    [shell, toast]
  );

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return recent;
    return recent.filter((r) => r.name.toLowerCase().includes(q) || (r.path || '').toLowerCase().includes(q));
  }, [recent, query]);

  // Office's starts: the new documents, the gallery of every template, and the templates of one's own.
  const officeStarts = () => {
    const docs = NEW_DOCUMENTS.filter((tpl) => has(tpl.app));
    return [
      ...docs.map((tpl) => <StartChip key={`${tpl.app}-${tpl.template}`} app={tpl.app} icon={APPS[tpl.app].icon} label={tpl.label} onClick={() => newDocument(tpl)} />),
      ...(docs.length ? [<StartChip key="more" className="home-more" icon="grid" label={t('More templates')} tip={t('More templates: letters, reports, budgets, invoices, decks and more, and your own')} onClick={() => setGallery(true)} />] : []),
      // Each makes a new document, the template left as it is.
      ...personal.filter((tpl) => has(tpl.app)).map((tpl) => <StartChip key={tpl.path} className="personal" app={tpl.app} icon={APPS[tpl.app].icon} label={tpl.name} tip={t('Your template: a new document made from {name}', { name: tpl.name })} template={tpl.path} onClick={() => shell.win.create({ app: tpl.app, query: { templateFile: tpl.path } })} />),
    ];
  };

  // Studio's starts, each with the app that does it. A chip of an app turned off is not offered.
  const studioStarts = () => [
    has('recorder') ? <StartChip key="record" className="home-record" app="recorder" icon="record" label={t('Record a screen')} tip={t('Record a screen: a screen, a window or your camera, with sound, in the Recorder')} onClick={() => shell.win.create({ app: 'recorder' })} /> : null,
    has('broadcast') ? <StartChip key="live" className="home-golive" app="broadcast" icon="broadcast" label={t('Go live')} tip={t('Go live: to the places you stream to, from Live Broadcast')} onClick={() => shell.win.create({ app: 'broadcast' })} /> : null,
    has('video') ? <StartChip key="intro" className="home-intro" app="video" icon="scenes" label={t('Introduction video')} tip={t('Introduction video: a title, you on camera, your screen and an end card, recorded scene by scene in your brand')} onClick={() => shell.win.create({ app: 'video', query: { intro: 1 } })} /> : null,
    has('video') ? <StartChip key="video" className="home-newvideo" app="video" icon="video" label={t('New video')} tip={t('New video: an empty timeline in Video')} onClick={() => shell.win.create({ app: 'video' })} /> : null,
    has('recorder') ? <StartChip key="picture" className="home-picture" app="recorder" icon="picture" label={t('Take a picture')} tip={t('Take a picture: the Recorder opens and takes one of what it shows')} onClick={() => shell.win.create({ app: 'recorder', query: { picture: 1 } })} /> : null,
  ].filter(Boolean);

  return (
    // words-ok: the suite's own name
    <AppFrame app={app} shell={shell} title="Rutba Office" menu={menu} status={<HomeStatus version={version} recent={recent.length} shell={shell} />}>
      <style>{CSS}</style>
      {gallery ? <TemplateGallery shell={shell} onClose={() => setGallery(false)} /> : null}
      {choosing ? <AppsChooser chosen={chosen.apps} addons={chosen.addons} onClose={() => setChoosing(false)} onChoose={async (keys, addons) => { setChoosing(false); await chosen.choose(keys, addons); toast(t('{n} apps on', { n: keys.length }), { tone: 'good', ms: 3000 }); }} /> : null}
      {packing ? <PacksDialog shell={shell} onClose={() => setPacking(false)} /> : null}
      <div className="home">
        <header className="home-hero">
          <div>
            {/* words-ok: the suite's own name */}
            <h1>Rutba Office</h1>
            <p>
              {t('Documents, mail and calls in Office; pictures, video and going live in Studio. Free, open source, and working with the network switched off.')}
            </p>
          </div>
          <div className="home-hero-actions">
            <Button icon="open" label={t('Open a file')} primary onClick={openFile} />
            {has('mail') ? (
              <Button
                icon="import"
                label={t('Import mail')}
                onClick={() => shell.win.create({ app: 'mail', query: { import: 1 } })}
              />
            ) : null}
          </div>
        </header>

        <Announcement shell={shell} />
        <PacksNotice onAdd={() => setPacking(true)} />

        {offApp && APPS[offApp] && me && chosen.loaded && !has(offApp) ? (
          <div className="home-choose home-app-off" role="status">
            <Icon name="info" size={20} />
            <span className="home-choose-text">
              <strong>{t('{app} is off in {name}', { app: APPS[offApp].short, name: profileName(me) })}</strong>
              <span>{t('Each profile keeps its own choice of apps. Turn it on here and it opens from Home.')}</span>
            </span>
            <Button primary label={t('Choose apps…')} className="home-off-choose" onClick={() => setChoosing(true)} />
            <Button ghost icon="close" title={t('Close')} onClick={() => setOffApp(null)} />
          </div>
        ) : null}

        {/* Once, until a choice is made: all the apps are on, and they can be fewer. */}
        {chosen.loaded && !chosen.made ? (
          <div className="home-choose" role="note">
            <Icon name="grid" size={20} />
            <span className="home-choose-text">
              <strong>{t('Choose your apps')}</strong>
              <span>{t('The Office apps are on. Add Talk, Data Files and the Studio apps when you want them, and turn off what you do not use, at any time.')}</span>
            </span>
            <Button label={t('Keep as it is')} className="home-keep-all" onClick={() => chosen.choose(chosen.apps, chosen.addons)} />
            <Button primary label={t('Choose apps…')} className="home-choose-btn" onClick={() => setChoosing(true)} />
          </div>
        ) : null}

        {recovered.length ? (
          <section className="home-section">
            <h2>{t('Recovered')}</h2>
            <div className="home-recovered">
              {recovered.map((r) => (
                <div key={r.file} className="home-recovered-row">
                  <span className="glyph" data-app={appOfRecovered(r)}>
                    <Icon name={appOfRecovered(r)} size={16} />
                  </span>
                  <span className="home-recovered-text">
                    <strong>{r.name}</strong>
                    <small>
                      {r.from
                        ? t('unsaved work from {when} — {file}', { when: new Date(r.at).toLocaleString(), file: r.from })
                        : t('unsaved work from {when} — never saved anywhere', { when: new Date(r.at).toLocaleString() })}
                    </small>
                  </span>
                  <Spacer />
                  <Button
                    primary
                    icon="open"
                    label={t('Recover')}
                    onClick={async () => {
                      // The window opens on the recovered copy, dirty, so the
                      // person decides what it replaces.
                      await shell.win.create({ app: appOfRecovered(r), query: { recover: r.file } });
                      setRecovered((list) => list.filter((e) => e.file !== r.file));
                    }}
                  />
                  <Button
                    icon="trash"
                    label={t('Discard')}
                    title={t('Throw this copy away')}
                    onClick={async () => {
                      await shell.doc.discardRecovery({ file: r.file });
                      setRecovered((list) => list.filter((e) => e.file !== r.file));
                    }}
                  />
                </div>
              ))}
            </div>
          </section>
        ) : null}

        <section className="home-section">
          <div className="home-section-head">
            <h2>{t('Apps')}</h2>
            <Spacer />
            <Button ghost icon="grid" label={t('Choose apps…')} className="home-apps-btn" title={t('Choose which apps are shown and open their kinds of file')} onClick={() => setChoosing(true)} />
            <Button ghost icon="download" label={t('Add or remove apps…')} className="home-packs-btn" title={t('Put apps on this computer or take them off')} onClick={() => setPacking(true)} />
          </div>
          <div className={`home-sets${sets.length === 1 ? ' one' : ''}`}>
            {sets.map((set) => {
              const words = setWords()[set.id];
              const starts = set.id === 'office' ? officeStarts() : studioStarts();
              return (
                <section key={set.id} className="home-set" data-set={set.id} aria-label={words.title}>
                  <div className="home-set-head">
                    <h3>{words.title}</h3>
                    <p>{words.line}</p>
                  </div>
                  <div className="home-grid">
                    {set.apps.map((key) => (
                      <AppCard
                        key={key}
                        app={APPS[key]}
                        onOpen={() => openApp(key)}
                        onNew={NEW_DOCUMENTS.some((n) => n.app === key)
                          ? () => newDocument({ app: key, template: 'blank' })
                          // A new PDF is made of other files: the window opens asking for them.
                          : key === 'pdf' ? () => shell.win.create({ app: 'pdf', query: { combine: 1 } }) : null}
                      />
                    ))}
                    {process.env.RUTBA_LOCKER === '1' && set.id === 'office' ? <React.Suspense fallback={null}><ExtraTile Card={AppCard} onOpen={() => openApp('locker')} /></React.Suspense> : null}
                  </div>
                  {starts.length ? (
                    <div className="home-set-starts">
                      <div className="home-set-label">{t('Start something')}</div>
                      <div className="home-templates">{starts}</div>
                    </div>
                  ) : null}
                </section>
              );
            })}
          </div>
        </section>

        <section className="home-section grow">
          <div className="home-section-head">
            <h2>{t('Recent')}</h2>
            <Spacer />
            <Search
              value={query}
              onChange={setQuery}
              placeholder={t('Search recent files')}
              style={{ width: 240 }}
              onKeyDown={(e) => {
                // Enter opens the first match: type a few letters, press Enter.
                if (e.key === 'Enter' && filtered[0]) openInApp(shell, filtered[0].path);
              }}
            />
          </div>
          {filtered.length ? (
            // A grid of rows, each a file: Enter opens it, F2 renames it, Delete takes it off the list.
            <div className="home-recent" role="grid" aria-label={t('Recent')}>
              {filtered.map((r) => {
                const kind = kindFromExtension(r.path);
                const which = r.app || appFor(kind) || 'home';
                const editing = renaming === r.path;
                const shownName = r.name || basename(r.path);
                return (
                  <div
                    key={r.path}
                    role="row"
                    tabIndex={0}
                    aria-label={`Open ${shownName}`}
                    className="home-recent-row"
                    onClick={() => { if (!editing) openInApp(shell, r.path); }}
                    onKeyDown={(e) => {
                      if (editing) return;
                      if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        openInApp(shell, r.path);
                      } else if (e.key === 'F2') {
                        e.preventDefault();
                        startRename(r);
                      } else if (e.key === 'Delete' || e.key === 'Backspace') {
                        e.preventDefault();
                        removeRecent(r.path);
                      }
                    }}
                    onContextMenu={(e) =>
                      rowMenu.open(e, [
                        { label: t('Open'), icon: 'open', run: () => openInApp(shell, r.path) },
                        // In another profile's windows, started beside this one if it is not open.
                        ...others.slice(0, 6).map((p) => ({ label: t('Open in {name}', { name: profileName(p) }), icon: 'contacts', run: () => shell.profiles.open({ id: p.id, file: r.path }).catch((err) => toast(err.message, { tone: 'bad' })) })),
                        { label: t('Show in folder'), icon: 'folder', run: () => shell.shell.showInFolder({ path: r.path }) },
                        '-',
                        { label: t('Rename…'), icon: 'textbox', run: () => startRename(r) },
                        { label: t('Remove from the list'), icon: 'close', run: () => removeRecent(r.path) },
                        { label: t('Clear the list'), icon: 'trash', run: () => shell.app.clearRecent().then(() => setRecent([])).catch(() => {}) },
                      ])
                    }
                    title={r.path}
                  >
                    <span className="rr-glyph" data-app={which} role="gridcell" aria-hidden="true">
                      <Icon name={APPS[which]?.icon || 'file'} size={15} />
                    </span>
                    <span className="rr-namecell" role="gridcell">
                    {editing ? (
                      <input
                        type="text"
                        aria-label={t('Rename…')}
                        className="rw-input home-recent-rename"
                        autoFocus
                        defaultValue={shownName}
                        onClick={(e) => e.stopPropagation()}
                        onFocus={(e) => {
                          const dot = e.target.value.lastIndexOf('.');
                          e.target.setSelectionRange(0, dot > 0 ? dot : e.target.value.length);
                        }}
                        onKeyDown={(e) => {
                          e.stopPropagation();
                          if (e.key === 'Enter') applyRename(r.path, e.target.value);
                          else if (e.key === 'Escape') setRenaming(null);
                        }}
                        onBlur={() => setRenaming(null)}
                      />
                    ) : (
                      <span className="rr-name">{shownName}</span>
                    )}
                    </span>
                    <span className="rr-kind" role="gridcell">{KINDS[kind]?.label || ''}</span>
                    <span className="rr-when" role="gridcell">{formatWhen(r.at)}</span>
                    <span role="gridcell" className="rr-removecell">
                    <button
                      type="button"
                      className="home-recent-remove"
                      title={t('Remove from the list')}
                      aria-label={t('Remove from the list')}
                      tabIndex={-1}
                      onClick={(e) => {
                        e.stopPropagation();
                        removeRecent(r.path);
                      }}
                    >
                      <Icon name="close" size={11} />
                    </button>
                    </span>
                  </div>
                );
              })}
            </div>
          ) : (
            <Empty icon="clock" title={query ? t('Nothing matches') : t('No recent files yet')}>
              {query
                ? t('Try a different search.')
                : t('Files you open will be listed here. Drop one onto this window to begin.')}
            </Empty>
          )}
        </section>
      </div>

      {showAbout ? (
        <About
          version={version}
          shell={shell}
          update={update}
          onCheck={() => shell.update.check({ manual: true }).then(setUpdate)}
          onInstall={() => shell.update.install()}
          onToggleAuto={(on) => shell.update.setAutomatic({ on }).then(setUpdate)}
          onClose={() => setShowAbout(false)}
        />
      ) : null}
      {showNotes ? <WhatsNew shell={shell} version={version?.version} onClose={() => setShowNotes(false)} /> : null}
      {rowMenu.node}
    </AppFrame>
  );
}

/**
 * The notice board.
 *
 * Nothing is drawn unless there is something to say that has not been said
 * before — no placeholder, no "no announcements", no space reserved. A launcher
 * that shows an empty box every morning trains people not to look at it.
 *
 * Every failure is silence. Offline, blocked, a proxy returning a login page,
 * the endpoint not built yet: all the same, and none of them are the person's
 * problem. There is no loading state for the same reason — a strip that flashes
 * "checking…" on every launch is worse than one that simply appears when it has
 * news.
 */
function Announcement({ shell }) {
  const [notice, setNotice] = useState(null);

  useEffect(() => {
    let live = true;
    shell.announce
      .check()
      .then((r) => live && setNotice(r?.announcement || null))
      .catch(() => {});
    // A second window opening later can find news this one has not seen.
    const off = shell.on('announce:new', (fresh) => setNotice(fresh));
    return () => {
      live = false;
      off?.();
    };
  }, [shell]);

  if (!notice) return null;

  const close = () => {
    setNotice(null);
    shell.announce.dismiss({ id: notice.id }).catch(() => {});
  };

  return (
    <div className={`home-notice ${notice.kind}`} role="status">
      <Icon name={notice.kind === 'security' ? 'shield' : notice.kind === 'release' ? 'download' : 'info'} size={14} />
      <div className="home-notice-text">
        <strong>{notice.title}</strong>
        {notice.body ? <span>{notice.body}</span> : null}
      </div>
      {notice.link ? (
        <button
          type="button"
          className="home-notice-link"
          onClick={() => {
            shell.shell.openExternal({ url: notice.link }).catch(() => {});
            close();
          }}
        >
          {notice.linkLabel}
        </button>
      ) : null}
      <button type="button" className="home-notice-close" onClick={close} title={t('Dismiss')} aria-label={t('Dismiss')}>
        <Icon name="close" size={12} />
      </button>
    </div>
  );
}

function HomeStatus({ version, recent, shell }) {
  const open = (url) => shell.shell.openExternal({ url });
  return (
    <>
      <span>Rutba Office{version ? ` ${version.version}` : ''}</span>

      {/* The footer is where somebody looks for who made this and how to ask. */}
      <a className="home-link" href={SITE.home} onClick={(e) => { e.preventDefault(); open(SITE.home); }}>
        office.rutba.io
      </a>
      <a className="home-link" href={SITE.contact} onClick={(e) => { e.preventDefault(); open(SITE.contact); }}>
        {t('Contact us')}
      </a>
      <Spacer />
      <Chip title={t('Files you have opened')}>{tn(recent, '{count} recent', '{count} recent')}</Chip>
      <Chip title={t('Documents, mail and media all work with no network connection')}>{t('Works offline')}</Chip>
    </>
  );
}

/** What the update service is doing, in words rather than a state name. */
function updateSentence(update) {
  if (!update) return t('Checking…');
  switch (update.state) {
    case 'unpackaged':
      return t('Updates apply to an installed copy; this one is running from source.');
    case 'off':
      return t('Automatic updates are off, so nothing is contacted. You can check whenever you like.');
    case 'checking':
      return t('Looking for a newer release…');
    case 'available':
      return t('Version {version} is available and downloading; you will be asked to restart when it is ready.', { version: update.available });
    case 'downloading':
      return t('Downloading version {version} — {percent}%.', { version: update.available, percent: Math.round(update.percent || 0) });
    case 'ready':
      return t('Version {version} is ready — restart to update, or it installs when you quit.', { version: update.available });
    case 'current':
      return t('This is the latest release.');
    case 'error':
      // A sentence by what went wrong; the updater's own message, headers and
      // stack included, was shown here as it came (2026-10-12).
      if (update.errorKind === 'publishing') return t('A new release is still being published. This copy is fine as it is; check again in a few minutes.');
      if (update.errorKind === 'offline') return t('The check could not reach GitHub, where the releases are kept. Check your connection and try again.');
      if (update.errorKind === 'refused') return t('GitHub turned the check away for now. Try again in an hour.');
      if (update.errorKind === 'unsigned') return t('An update was downloaded but its signature did not check out, so it was not installed. This copy is fine as it is.');
      return t('The last check did not complete: {error}', { error: update.error });
    default:
      return t('No check has run yet. This build contacts nothing on its own.');
  }
}

/**
 * The announcement setting, stated rather than buried.
 *
 * A suite whose mail client names the companies watching you cannot have a
 * quiet channel of its own. So the second — and last — outbound request is
 * described here in the words that are actually true, including the part that
 * is uncomfortable to write down: this is how we know anybody is using it.
 */
function AnnouncementSetting({ shell }) {
  const [status, setStatus] = useState(null);

  useEffect(() => {
    let live = true;
    shell.announce.status().then((s) => live && setStatus(s)).catch(() => {});
    return () => {
      live = false;
    };
  }, [shell]);

  if (!status) return null;

  return (
    <div className="about-update">
      <div className="about-update-line">
        <Icon name="globe" size={15} />
        <span>
          {status.enabled
            ? status.checkedAt
              ? t('Last asked {when}.', { when: formatWhen(new Date(status.checkedAt).toISOString()) })
              : t('Not asked yet.')
            : t('Announcements are off, so this is never contacted.')}
        </span>
      </div>
      <label className="about-auto">
        <input
          type="checkbox"
          checked={status.enabled}
          onChange={async (e) => setStatus({ ...status, ...(await shell.announce.setEnabled({ on: e.target.checked })) })}
        />
        <span>
          {t('Show announcements from office.rutba.io — once a day at most. It sends the version and the operating system, and no identifier of any kind, so what is counted at the other end is "a copy opened somewhere", never you. It is also the only way we know the suite is being used at all.')}
        </span>
      </label>
    </div>
  );
}

function About({ version, shell, update, onCheck, onInstall, onToggleAuto, onClose }) {
  const open = (url) => shell.shell.openExternal({ url });
  return (
    <Dialog
      title={t('About Rutba Office')}
      width={460}
      className="about-dialog"
      onClose={onClose}
      actions={<>
        <Button label={t('Source code')} onClick={() => open(SITE.source)} />
        {update?.state === 'ready' ? (
          <Button label={t('Restart and install')} onClick={onInstall} />
        ) : (
          <Button label={t('Check for updates')} onClick={onCheck} />
        )}
        <Button label={t('Close')} primary onClick={onClose} />
      </>}
    >
        <div>
          <p style={{ marginTop: 0 }}>
            {t('A free office suite: mail, documents, worksheets, presentations, pictures, images and video, on an engine we own.')}
          </p>
          <dl className="about-list">
            <dt>{t('Version')}</dt>
            <dd>{version?.version || '—'}</dd>
            <dt>{t('Platform')}</dt>
            <dd>
              {version?.platform} {version?.arch}
            </dd>
            <dt>{t('Runtime')}</dt>
            <dd>
              Electron {version?.electron} · Chromium {version?.chrome} · Node {version?.node}
            </dd>
            <dt>{t('Website')}</dt>
            <dd>
              <a className="home-link inline" href={SITE.home} onClick={(e) => { e.preventDefault(); open(SITE.home); }}>
                office.rutba.io
              </a>
            </dd>
            <dt>{t('Contact')}</dt>
            <dd>
              <a className="home-link inline" href={SITE.contact} onClick={(e) => { e.preventDefault(); open(SITE.contact); }}>
                office.rutba.io/contact
              </a>
            </dd>
            <dt>{t('Licence')}</dt>
            <dd>{t('GNU AGPL v3.0, or a commercial licence')}</dd>
          </dl>

          <div className="about-update">
            <div className="about-update-line">
              <Icon name={update?.state === 'ready' ? 'download' : update?.state === 'error' ? 'info' : 'refresh'} size={15} />
              <span>{updateSentence(update)}</span>
            </div>
            <label className="about-auto">
              <input type="checkbox" checked={update?.automatic !== false} onChange={(e) => onToggleAuto(e.target.checked)} />
              <span>
                {t('Check for updates automatically — one request to GitHub for the release list, and nothing about you or your files.')}
              </span>
            </label>
          </div>

          <AnnouncementSetting shell={shell} />

          {/*
            The whole list, so nobody has to take our word for it. If a network
            monitor ever shows a fourth thing, that is a bug and we want to hear
            about it.
          */}
          <p className="rw-hint">
            {t('Those two requests, and the mail and calendar servers you set up yourself, are everything Rutba Office ever contacts. Your documents, your mail and what you do with them never leave this computer.')}
          </p>

          <p className="rw-hint">
            {t('Copyright © 2026 Tech Style Ltd. The source is published, and you are free to study, modify and share it under the terms of the AGPL.')}
          </p>
        </div>
    </Dialog>
  );
}

const CSS = `
/* the notice board: one line, easy to ignore, easy to dismiss */
.home-notice {
  display: flex; align-items: center; gap: 10px; margin: 0 0 18px;
  padding: 9px 12px; border-radius: var(--r-2); font-size: 12.5px;
  background: var(--sunken); border: 1px solid var(--line); color: var(--ink-2);
}
.home-notice.release { border-color: var(--accent-line); background: var(--selected); }
.home-notice.security { border-color: color-mix(in srgb, var(--bad) 40%, transparent); background: color-mix(in srgb, var(--bad) 8%, transparent); }
.home-notice-text { flex: 1; min-width: 0; display: flex; gap: 8px; align-items: baseline; }
.home-notice-text strong { color: var(--ink); }
.home-notice-text span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.home-notice-link {
  border: 0; background: none; color: var(--accent-text); font: inherit; font-size: 12.5px;
  font-weight: 600; cursor: pointer; padding: 2px 4px; border-radius: var(--r-1); flex: none;
}
.home-notice-link:hover { background: var(--hover); }
.home-notice-close {
  border: 0; background: none; color: var(--ink-3); cursor: pointer; padding: 2px;
  display: grid; place-items: center; flex: none; border-radius: var(--r-1);
}
.home-notice-close:hover { background: var(--hover); color: var(--ink); }

.home {
  flex: 1; overflow: auto; padding: 24px 30px 28px; display: flex; flex-direction: column; gap: 22px;
  /* A wash of the accent behind the welcome, gone by the first row of tiles. */
  background:
    radial-gradient(900px 360px at 12% -20%, color-mix(in srgb, var(--accent) 12%, transparent), transparent 70%),
    var(--window);
}
.home-hero { display: flex; align-items: flex-start; gap: 24px; flex-wrap: wrap; }
.home-hero h1 {
  margin: 0 0 6px; font-family: var(--font-display); font-size: 27px; font-weight: 650;
  letter-spacing: -0.025em;
}
.home-hero p { margin: 0; max-width: 62ch; color: var(--ink-2); font-size: 13.5px; }
.home-hero-actions { display: flex; gap: 8px; margin-left: auto; }
.home-hero-actions .rw-btn { padding: 7px 14px; border: 1px solid var(--line); }
.home-hero-actions .rw-btn.primary { border-color: transparent; }

.home-section { display: flex; flex-direction: column; gap: 11px; }
/* The recent list takes what the tiles leave; at the launcher's own size
   that is four rows without the page scrolling, more in a taller window. */
.home-section.grow { flex: 1; min-height: 196px; }
.home-section-head { display: flex; align-items: center; gap: 12px; }
.home-section-head h2 { margin-bottom: 0; }
.home-choose { display: flex; align-items: center; gap: 12px; padding: 12px 16px; margin: 0 0 18px; border-radius: var(--r-3); border: 1px solid var(--accent); background: var(--wash-strong); color: var(--ink); }
.home-choose > svg { color: var(--accent-text); flex: none; }
.home-choose-text { flex: 1; min-width: 0; display: grid; gap: 2px; }
.home-choose-text span { font-size: 12.5px; color: var(--ink-2); }
.home-section h2 {
  margin: 0; font-size: 11px; font-weight: 700; letter-spacing: 0.08em;
  text-transform: uppercase; color: var(--ink-3);
}

/* Office and Studio side by side, each its own column of tiles at the same
   width (nine tiles three across, five two across) and its own starts; below
   about 1100 px they stack, Office first, and a set alone takes the width. */
.home-sets { display: grid; grid-template-columns: minmax(0, 9fr) minmax(0, 6fr); gap: 14px; align-items: stretch; }
.home-sets.one { grid-template-columns: minmax(0, 1fr); }
.home-set {
  display: flex; flex-direction: column; gap: 12px; min-width: 0; padding: 14px 16px 16px;
  border: 1px solid var(--line); border-radius: var(--r-3);
  background: color-mix(in srgb, var(--surface) 55%, transparent);
}
.home-set-head h3 { margin: 0; font-family: var(--font-display); font-size: 17px; font-weight: 650; letter-spacing: -0.015em; }
.home-set-head p { margin: 1px 0 0; font-size: 12px; color: var(--ink-3); }
.home-set-starts { display: flex; flex-direction: column; gap: 8px; }
.home-set-label { font-size: 11px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--ink-3); }
.home-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
.home-set[data-set='studio'] .home-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
.home-sets.one .home-set .home-grid { grid-template-columns: repeat(auto-fill, minmax(206px, 1fr)); }
@media (max-width: 1100px) {
  .home-sets { grid-template-columns: minmax(0, 1fr); }
  .home-set .home-grid, .home-set[data-set='studio'] .home-grid { grid-template-columns: repeat(auto-fill, minmax(206px, 1fr)); }
}
.home-card {
  position: relative; display: flex; background: var(--surface);
  border: 1px solid var(--line); border-radius: var(--r-3); overflow: hidden;
  transition: border-color var(--fast), box-shadow var(--fast), transform var(--fast);
}
.home-card:hover { border-color: var(--accent-line); box-shadow: var(--shadow-2); transform: translateY(-1px); }
.home-card-main {
  flex: 1; display: flex; align-items: center; gap: 11px; padding: 10px 12px;
  border: 0; background: transparent; color: var(--ink); text-align: left; min-width: 0;
}
/* The tile wears the app's own colour — the theme names one per app, so a
   new app never draws a blank tile again (Calendar and Contacts did). */
.home-card-glyph {
  width: 36px; height: 36px; border-radius: 10px; display: grid; place-items: center; flex: none;
  color: #fff;
  background: linear-gradient(145deg, color-mix(in srgb, var(--accent) 82%, white), var(--accent) 55%, color-mix(in srgb, var(--accent) 78%, black));
  box-shadow: 0 2px 6px color-mix(in srgb, var(--accent) 35%, transparent), inset 0 1px 0 rgba(255, 255, 255, 0.25);
  transition: transform var(--fast), box-shadow var(--fast);
}
.home-card:hover .home-card-glyph { transform: scale(1.04); box-shadow: 0 4px 12px color-mix(in srgb, var(--accent) 45%, transparent), inset 0 1px 0 rgba(255, 255, 255, 0.25); }
.home-card-text { display: flex; flex-direction: column; gap: 1px; min-width: 0; }
.home-card-text strong { font-size: 13.5px; font-weight: 600; }
.home-card-text span { font-size: 11.5px; color: var(--ink-3); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
/* New, for the apps that have one: over the tile's end, shown only when the tile is worth looking at. */
.home-card-new {
  position: absolute; inset-block: 0; inset-inline-end: 0; width: 32px; border: 0;
  border-inline-start: 1px solid var(--line-soft); background: var(--surface);
  color: var(--ink-3); display: grid; place-items: center; opacity: 0;
  transition: background var(--fast), color var(--fast), opacity var(--fast);
}
.home-card:hover .home-card-new, .home-card:focus-within .home-card-new { opacity: 1; }
.home-card-new:hover { background: var(--hover); color: var(--accent-text); }

.home-templates { display: flex; flex-wrap: wrap; gap: 7px; }
.home-template {
  display: inline-flex; align-items: center; gap: 8px; padding: 6px 12px 6px 7px;
  border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: 22px; font: inherit; font-size: 12.5px;
  transition: border-color var(--fast), background var(--fast);
}
.home-template:hover { border-color: var(--accent-line); background: var(--selected); }
.tpl-glyph { width: 22px; height: 22px; border-radius: 7px; display: grid; place-items: center; color: #fff; flex: none; }
.tpl-glyph[data-app='word'] { background: #2b5fd9; }
.tpl-glyph[data-app='sheets'] { background: #0f9d58; }
.tpl-glyph[data-app='slides'] { background: #d9534f; }
.tpl-glyph[data-app='video'] { background: #c2408f; }
.tpl-glyph[data-app='recorder'] { background: #e03b4c; }
.tpl-glyph[data-app='broadcast'] { background: #ea580c; }
.tpl-glyph[data-app='pdf'] { background: #b3261e; }
.tpl-glyph.more { background: var(--ink-3); }
.home-more { font-weight: 600; }

/* The list takes what is left of the window and scrolls inside it. It used
   to be clipped instead: a flex item that hides its overflow may shrink
   below its content, so the last row in reach was cut in half and there
   was nothing to scroll. */
.home-recent {
  border: 1px solid var(--line); border-radius: var(--r-3); background: var(--surface);
  overflow: auto; min-height: 0; overscroll-behavior: contain;
}
/* Work a crash took, offered back. Marked, because it is not an ordinary row
   in a list: it is the one thing on this window somebody might be looking
   for. */
.home-recovered {
  border: 1px solid var(--accent); border-radius: var(--radius); overflow: hidden;
  background: color-mix(in srgb, var(--accent) 6%, transparent);
}
.home-recovered-row { display: flex; align-items: center; gap: 12px; padding: 10px 13px; border-bottom: 1px solid var(--line-soft); }
.home-recovered-row:last-child { border-bottom: 0; }
.home-recovered-row .glyph { width: 26px; height: 26px; border-radius: 8px; display: grid; place-items: center; color: #fff; flex: none; background: var(--accent); }
.home-recovered-row .glyph[data-app='word'] { background: #2b5fd9; }
.home-recovered-row .glyph[data-app='sheets'] { background: #0f9d58; }
.home-recovered-row .glyph[data-app='slides'] { background: #d9534f; }
.home-recovered-row .glyph[data-app='video'] { background: #c2408f; }
.home-recovered-row .glyph[data-app='pdf'] { background: #b3261e; }
.home-recovered-row .glyph[data-app='image'] { background: #e08b2b; }
.home-recovered-text { display: flex; flex-direction: column; min-width: 0; }
.home-recovered-text small { color: var(--ink-3); font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

.home-recent-row {
  display: flex; align-items: center; gap: 12px; width: 100%; padding: 6px 13px;
  border: 0; border-bottom: 1px solid var(--line-soft); background: transparent; color: var(--ink);
  font: inherit; text-align: left; transition: background var(--fast); cursor: pointer; outline: none;
}
.home-recent-row:last-child { border-bottom: 0; }
.home-recent-row:hover, .home-recent-row:focus-visible { background: var(--hover); }
.rr-glyph { width: 24px; height: 24px; border-radius: 7px; display: grid; place-items: center; color: #fff; flex: none; }
.rr-glyph[data-app] { background: var(--accent); }
.rr-glyph[data-app='home'] { background: var(--n-50); }
.rr-namecell { flex: 1; min-width: 0; display: flex; }
.rr-removecell { flex: none; display: flex; }
.rr-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
/* The row's name becomes this in place, rather than a dialog elsewhere. */
.home-recent-rename { flex: 1; min-width: 0; padding: 2px 6px; font-size: inherit; }
/* Hidden until the row is worth looking at — hovered or focused — so the
   list does not read as a row of little grey crosses at rest. */
.home-recent-remove {
  flex: none; width: 20px; height: 20px; border-radius: 6px; border: 0; background: transparent;
  color: var(--ink-3); display: grid; place-items: center; opacity: 0; transition: opacity var(--fast), background var(--fast), color var(--fast);
}
.home-recent-row:hover .home-recent-remove, .home-recent-row:focus-within .home-recent-remove { opacity: 1; }
.home-recent-remove:hover { background: var(--hover-strong, var(--hover)); color: var(--bad); }
.rr-kind { color: var(--ink-3); font-size: 11.5px; flex: none; }
.rr-when { color: var(--ink-3); font-size: 11.5px; width: 84px; text-align: right; flex: none; }

.about-list { display: grid; grid-template-columns: auto 1fr; gap: 5px 16px; margin: 14px 0; font-size: 12.5px; }
.about-list dt { color: var(--ink-3); }
.about-list dd { margin: 0; }

.home-link {
  color: var(--accent-text); text-decoration: none; font-size: 11.5px;
  border-bottom: 1px solid transparent; transition: border-color var(--fast);
}
.home-link:hover { border-bottom-color: var(--accent); }
.home-link.inline { font-size: inherit; }

.about-update {
  margin: 14px 0 4px; padding: 11px 13px; border-radius: var(--r-2);
  background: var(--sunken); display: flex; flex-direction: column; gap: 9px;
}
.about-update-line { display: flex; align-items: center; gap: 9px; font-size: 12.5px; }
.about-auto { display: flex; align-items: flex-start; gap: 8px; font-size: 11.5px; color: var(--ink-2); }
.about-auto input { margin-top: 2px; accent-color: var(--accent); }
`;
