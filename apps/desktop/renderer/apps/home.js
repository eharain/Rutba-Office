// The launcher.
//
// The first thing anybody sees, and the only window that is not about one
// document. It has three jobs: start something new, reopen something recent,
// and make it obvious what the seven apps are — because the whole point of a
// suite is that you did not have to go and find seven separate downloads.

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Icon, Search, Empty, Button, Chip, Spacer, formatBytes, formatWhen, basename } from '@rutba/office-ui';
import { APPS, NEW_DOCUMENTS, SITE } from '@rutba/office-formats/registry';
import { appFor, kindFromExtension, KINDS } from '@rutba/office-formats/sniff';
import { AppFrame, useAppMenu, pickOpen, openInApp, useFileDrop } from '../shell.js';

const ORDER = ['mail', 'word', 'sheets', 'slides', 'pictures', 'image', 'video'];

function AppCard({ app, onOpen, onNew }) {
  return (
    <div className="home-card" data-app={app.key}>
      <button type="button" className="home-card-main" onClick={onOpen}>
        <span className="home-card-glyph">
          <Icon name={app.icon} size={22} />
        </span>
        <span className="home-card-text">
          <strong>{app.short}</strong>
          <span>{app.tagline}</span>
        </span>
      </button>
      {onNew ? (
        <button type="button" className="home-card-new" onClick={onNew} title={`New ${app.short.toLowerCase()}`}>
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
  const [update, setUpdate] = useState(null);

  const refresh = useCallback(() => {
    shell.app.recent().then(setRecent).catch(() => setRecent([]));
  }, [shell]);

  useEffect(() => {
    refresh();
    shell.app.version().then(setVersion).catch(() => {});
    shell.update.state().then(setUpdate).catch(() => {});
  }, [refresh, shell]);

  useEffect(() => shell.on('update:state', setUpdate), [shell]);

  const openApp = useCallback((key) => shell.win.create({ app: key }), [shell]);

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

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return recent;
    return recent.filter((r) => r.name.toLowerCase().includes(q) || (r.path || '').toLowerCase().includes(q));
  }, [recent, query]);

  return (
    <AppFrame app={app} shell={shell} title="Rutba Office" menu={menu} status={<HomeStatus version={version} recent={recent.length} shell={shell} update={update} />}>
      <style>{CSS}</style>
      <div className="home">
        <header className="home-hero">
          <div>
            <h1>Rutba Office</h1>
            <p>
              Mail, documents, worksheets, presentations and media — free, open source, and working with the
              network switched off.
            </p>
          </div>
          <div className="home-hero-actions">
            <Button icon="open" label="Open a file" primary onClick={openFile} />
            <Button
              icon="import"
              label="Import mail"
              onClick={() => shell.win.create({ app: 'mail', query: { import: 1 } })}
            />
          </div>
        </header>

        <section className="home-section">
          <h2>Apps</h2>
          <div className="home-grid">
            {ORDER.map((key) => (
              <AppCard
                key={key}
                app={APPS[key]}
                onOpen={() => openApp(key)}
                onNew={NEW_DOCUMENTS.some((n) => n.app === key) ? () => newDocument({ app: key, template: 'blank' }) : null}
              />
            ))}
          </div>
        </section>

        <section className="home-section">
          <h2>Start something</h2>
          <div className="home-templates">
            {NEW_DOCUMENTS.map((t) => (
              <button key={`${t.app}-${t.template}`} type="button" className="home-template" onClick={() => newDocument(t)}>
                <span className="tpl-glyph" data-app={t.app}>
                  <Icon name={APPS[t.app].icon} size={16} />
                </span>
                <span>{t.label}</span>
              </button>
            ))}
          </div>
        </section>

        <section className="home-section grow">
          <div className="home-section-head">
            <h2>Recent</h2>
            <Spacer />
            <Search value={query} onChange={setQuery} placeholder="Search recent files" style={{ width: 240 }} />
          </div>
          {filtered.length ? (
            <div className="home-recent">
              {filtered.map((r) => {
                const kind = kindFromExtension(r.path);
                const which = r.app || appFor(kind) || 'home';
                return (
                  <button
                    key={r.path}
                    type="button"
                    className="home-recent-row"
                    onClick={() => openInApp(shell, r.path)}
                    onContextMenu={(e) => {
                      e.preventDefault();
                      shell.shell.showInFolder({ path: r.path });
                    }}
                    title={r.path}
                  >
                    <span className="rr-glyph" data-app={which}>
                      <Icon name={APPS[which]?.icon || 'file'} size={15} />
                    </span>
                    <span className="rr-name">{r.name || basename(r.path)}</span>
                    <span className="rr-kind">{KINDS[kind]?.label || ''}</span>
                    <span className="rr-when">{formatWhen(r.at)}</span>
                  </button>
                );
              })}
            </div>
          ) : (
            <Empty icon="clock" title={query ? 'Nothing matches' : 'No recent files yet'}>
              {query
                ? 'Try a different search.'
                : 'Files you open will be listed here. Drop one onto this window to begin.'}
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
    </AppFrame>
  );
}

function HomeStatus({ version, recent, shell, update }) {
  const open = (url) => shell.shell.openExternal({ url });
  return (
    <>
      <span>Rutba Office{version ? ` ${version.version}` : ''}</span>
      {/* The footer is where somebody looks for who made this and how to ask. */}
      <a className="home-link" href={SITE.home} onClick={(e) => { e.preventDefault(); open(SITE.home); }}>
        office.rutba.io
      </a>
      <a className="home-link" href={SITE.contact} onClick={(e) => { e.preventDefault(); open(SITE.contact); }}>
        Contact us
      </a>
      <Spacer />
      {update?.state === 'ready' ? (
        <Chip title={`Version ${update.available} is downloaded and installs when you quit`}>Update ready</Chip>
      ) : update?.state === 'downloading' ? (
        <Chip title="Downloading in the background">Updating {Math.round(update.percent || 0)}%</Chip>
      ) : null}
      <Chip title="Files you have opened">{recent} recent</Chip>
      <Chip title="Documents, mail and media all work with no network connection">Works offline</Chip>
    </>
  );
}

/** What the update service is doing, in words rather than a state name. */
function updateSentence(update) {
  if (!update) return 'Checking…';
  switch (update.state) {
    case 'unpackaged':
      return 'Updates apply to an installed copy; this one is running from source.';
    case 'off':
      return 'Automatic updates are off. Nothing is contacted.';
    case 'checking':
      return 'Looking for a newer release…';
    case 'available':
      return `Version ${update.available} is available and downloading.`;
    case 'downloading':
      return `Downloading version ${update.available} — ${Math.round(update.percent || 0)}%.`;
    case 'ready':
      return `Version ${update.available} is ready, and installs when you quit.`;
    case 'current':
      return 'This is the latest release.';
    case 'error':
      return `The last check did not complete: ${update.error}`;
    default:
      return 'No check has run yet.';
  }
}

function About({ version, shell, update, onCheck, onInstall, onToggleAuto, onClose }) {
  const open = (url) => shell.shell.openExternal({ url });
  return (
    <div className="rw-scrim" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="rw-dialog" style={{ width: 460 }}>
        <div className="rw-dialog-head">About Rutba Office</div>
        <div className="rw-dialog-body">
          <p style={{ marginTop: 0 }}>
            A free office suite: mail, documents, worksheets, presentations, pictures, images and video, on an
            engine we own.
          </p>
          <dl className="about-list">
            <dt>Version</dt>
            <dd>{version?.version || '—'}</dd>
            <dt>Platform</dt>
            <dd>
              {version?.platform} {version?.arch}
            </dd>
            <dt>Runtime</dt>
            <dd>
              Electron {version?.electron} · Chromium {version?.chrome} · Node {version?.node}
            </dd>
            <dt>Website</dt>
            <dd>
              <a className="home-link inline" href={SITE.home} onClick={(e) => { e.preventDefault(); open(SITE.home); }}>
                office.rutba.io
              </a>
            </dd>
            <dt>Contact</dt>
            <dd>
              <a className="home-link inline" href={SITE.contact} onClick={(e) => { e.preventDefault(); open(SITE.contact); }}>
                office.rutba.io/contact
              </a>
            </dd>
            <dt>Licence</dt>
            <dd>GNU AGPL v3.0, or a commercial licence</dd>
          </dl>

          <div className="about-update">
            <div className="about-update-line">
              <Icon name={update?.state === 'ready' ? 'download' : update?.state === 'error' ? 'info' : 'refresh'} size={15} />
              <span>{updateSentence(update)}</span>
            </div>
            <label className="about-auto">
              <input type="checkbox" checked={update?.automatic !== false} onChange={(e) => onToggleAuto(e.target.checked)} />
              <span>
                Check for updates automatically — one request to GitHub for the release list, and nothing about
                you or your files.
              </span>
            </label>
          </div>

          <p className="rw-hint">
            Copyright © 2026 Tech Style Ltd. The source is published, and you are free to study, modify and
            share it under the terms of the AGPL.
          </p>
        </div>
        <div className="rw-dialog-foot">
          <Button label="Source code" onClick={() => open(SITE.source)} />
          {update?.state === 'ready' ? (
            <Button label="Restart and install" onClick={onInstall} />
          ) : (
            <Button label="Check for updates" onClick={onCheck} />
          )}
          <Button label="Close" primary onClick={onClose} />
        </div>
      </div>
    </div>
  );
}

const CSS = `
.home { flex: 1; overflow: auto; padding: 26px 30px 34px; display: flex; flex-direction: column; gap: 26px; }
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
.home-section.grow { flex: 1; min-height: 220px; }
.home-section-head { display: flex; align-items: center; gap: 12px; }
.home-section h2 {
  margin: 0; font-size: 11px; font-weight: 700; letter-spacing: 0.08em;
  text-transform: uppercase; color: var(--ink-3);
}

.home-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(216px, 1fr)); gap: 10px; }
.home-card {
  position: relative; display: flex; background: var(--surface);
  border: 1px solid var(--line); border-radius: var(--r-3); overflow: hidden;
  transition: border-color var(--fast), box-shadow var(--fast), transform var(--fast);
}
.home-card:hover { border-color: var(--accent-line); box-shadow: var(--shadow-2); transform: translateY(-1px); }
.home-card-main {
  flex: 1; display: flex; align-items: center; gap: 12px; padding: 13px 14px;
  border: 0; background: transparent; color: var(--ink); text-align: left; min-width: 0;
}
.home-card-glyph {
  width: 40px; height: 40px; border-radius: 11px; display: grid; place-items: center; flex: none;
  color: #fff; box-shadow: inset 0 -1px 0 rgba(0,0,0,0.18);
}
.home-card[data-app='mail'] .home-card-glyph { background: #3b7de0; }
.home-card[data-app='word'] .home-card-glyph { background: #2b5fd9; }
.home-card[data-app='sheets'] .home-card-glyph { background: #0f9d58; }
.home-card[data-app='slides'] .home-card-glyph { background: #d9534f; }
.home-card[data-app='pictures'] .home-card-glyph { background: #7b5cd6; }
.home-card[data-app='image'] .home-card-glyph { background: #e08b2b; }
.home-card[data-app='video'] .home-card-glyph { background: #c2408f; }
.home-card-text { display: flex; flex-direction: column; gap: 1px; min-width: 0; }
.home-card-text strong { font-size: 13.5px; font-weight: 600; }
.home-card-text span { font-size: 11.5px; color: var(--ink-3); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.home-card-new {
  width: 34px; border: 0; border-left: 1px solid var(--line-soft); background: transparent;
  color: var(--ink-3); display: grid; place-items: center; transition: background var(--fast), color var(--fast);
}
.home-card-new:hover { background: var(--hover); color: var(--accent); }

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

.home-recent {
  border: 1px solid var(--line); border-radius: var(--r-3); overflow: hidden; background: var(--surface);
}
.home-recent-row {
  display: flex; align-items: center; gap: 12px; width: 100%; padding: 8px 13px;
  border: 0; border-bottom: 1px solid var(--line-soft); background: transparent; color: var(--ink);
  font: inherit; text-align: left; transition: background var(--fast);
}
.home-recent-row:last-child { border-bottom: 0; }
.home-recent-row:hover { background: var(--hover); }
.rr-glyph { width: 26px; height: 26px; border-radius: 8px; display: grid; place-items: center; color: #fff; flex: none; }
.rr-glyph[data-app='mail'] { background: #3b7de0; }
.rr-glyph[data-app='word'] { background: #2b5fd9; }
.rr-glyph[data-app='sheets'] { background: #0f9d58; }
.rr-glyph[data-app='slides'] { background: #d9534f; }
.rr-glyph[data-app='pictures'] { background: #7b5cd6; }
.rr-glyph[data-app='image'] { background: #e08b2b; }
.rr-glyph[data-app='video'] { background: #c2408f; }
.rr-glyph[data-app='home'] { background: var(--n-50); }
.rr-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rr-kind { color: var(--ink-3); font-size: 11.5px; flex: none; }
.rr-when { color: var(--ink-3); font-size: 11.5px; width: 84px; text-align: right; flex: none; }

.about-list { display: grid; grid-template-columns: auto 1fr; gap: 5px 16px; margin: 14px 0; font-size: 12.5px; }
.about-list dt { color: var(--ink-3); }
.about-list dd { margin: 0; }

.home-link {
  color: var(--accent); text-decoration: none; font-size: 11.5px;
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
