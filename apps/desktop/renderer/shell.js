// What every app window has in common.
//
// The frame, the app menu behind the mark in the corner, the open/save
// plumbing, and the two or three behaviours that must be identical everywhere:
// a document window asks before closing on unsaved work, the theme toggle is in
// the same place, and files dropped on a window open in the right app.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Window, TitleBar, Body, StatusBar, Menu, Icon, useToast, useTheme, Button, Progress } from '@rutba/office-ui';
import { APPS, openFilters, saveFilters, NEW_DOCUMENTS } from '@rutba/office-formats/registry';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { WhatsNew } from './whatsnew.js';

/** The menu behind the app mark: new, open, recent, and the way out. */
export function useAppMenu({ shell, appKey, onNew, onOpen, extra = [] }) {
  const [menu, setMenu] = useState(null);
  const [recent, setRecent] = useState([]);
  const { mode, setTheme } = useTheme();

  useEffect(() => {
    shell.app.recent().then(setRecent).catch(() => {});
  }, [shell]);

  const open = useCallback(
    (e) => {
      const rect = e.currentTarget.getBoundingClientRect();
      const items = [];
      if (onNew) items.push({ label: 'New', icon: 'new', key: 'Ctrl+N', run: onNew });
      if (onOpen) items.push({ label: 'Open…', icon: 'open', key: 'Ctrl+O', run: onOpen });
      if (extra.length) items.push('-', ...extra);

      if (recent.length) {
        items.push('-');
        for (const r of recent.slice(0, 8)) {
          items.push({
            label: r.name,
            icon: 'file',
            run: () => shell.win.create({ app: r.app || appFor(kindFromExtension(r.path)) || 'home', file: r.path }),
          });
        }
      }

      items.push('-');
      items.push({
        label: mode === 'dark' ? 'Light theme' : mode === 'light' ? 'Follow system' : 'Dark theme',
        icon: mode === 'dark' ? 'sun' : 'moon',
        run: () => setTheme(mode === 'dark' ? 'light' : mode === 'light' ? 'system' : 'dark'),
      });
      items.push({ label: 'All apps', icon: 'grid', run: () => shell.win.create({ app: 'home' }) });
      items.push('-');
      items.push({ label: 'What’s new', icon: 'star', run: () => shell.win.create({ app: 'home', query: { whatsnew: 1 } }) });
      items.push({ label: 'About Rutba Office', icon: 'info', run: () => shell.win.create({ app: 'home', query: { about: 1 } }) });

      setMenu({ x: rect.left, y: rect.bottom + 4, items });
    },
    [recent, onNew, onOpen, extra, mode, setTheme, shell, appKey]
  );

  return { open, node: menu ? <Menu {...menu} onClose={() => setMenu(null)} /> : null };
}

/**
 * Do not let unsaved work close.
 *
 * The window tells the backend whether it is holding changes; the backend
 * cancels the close and asks with the platform's own dialog. "Save" comes back
 * here as a command, and the window closes once it has actually saved — so a
 * save that is itself cancelled leaves the window open with the work in it.
 *
 * Without this, closing a window with the X threw the document away in
 * silence, which is the single worst thing an editor can do.
 */
export function useDirtyGuard({ shell, dirty, name, onSave }) {
  const save = useRef(onSave);
  save.current = onSave;

  useEffect(() => {
    shell.win.setDirty({ dirty: Boolean(dirty), name: name || '' }).catch(() => {});
  }, [shell, dirty, name]);

  useEffect(() => {
    if (!shell?.on) return undefined;
    return shell.on('app:command', async ({ command }) => {
      if (command !== 'file.saveAndClose') return;
      const saved = await save.current?.();
      if (saved !== false) shell.win.close({ force: true });
    });
  }, [shell]);
}

/** Ask before losing work. Used by every editor's close path. */
export async function confirmDiscard(shell, name) {
  const { response } = await shell.dialog.message({
    type: 'warning',
    message: `Save changes to ${name}?`,
    detail: 'Your changes will be lost if you don’t save them.',
    buttons: ['Save', "Don't save", 'Cancel'],
    defaultId: 0,
    cancelId: 2,
  });
  return response === 0 ? 'save' : response === 1 ? 'discard' : 'cancel';
}

export async function pickOpen(shell, appKey, { multiple = false } = {}) {
  const paths = await shell.dialog.open({
    title: 'Open',
    filters: openFilters(appKey),
    multiple,
  });
  return multiple ? paths : paths[0] || null;
}

export async function pickSave(shell, appKey, defaultPath) {
  return shell.dialog.save({ title: 'Save as', filters: saveFilters(appKey), defaultPath });
}

/**
 * Files dropped onto a window. Chromium would happily navigate to the file and
 * replace the app with it, so the default is refused everywhere and the drop is
 * routed to the app that owns the format.
 */
export function useFileDrop(handler) {
  useEffect(() => {
    const over = (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
    };
    const drop = (e) => {
      e.preventDefault();
      const files = [...(e.dataTransfer?.files || [])]
        .map((f) => window.rutbaOffice?.pathOf?.(f) || f.path)
        .filter(Boolean);
      if (files.length) handler(files);
    };
    window.addEventListener('dragover', over);
    window.addEventListener('drop', drop);
    return () => {
      window.removeEventListener('dragover', over);
      window.removeEventListener('drop', drop);
    };
  }, [handler]);
}

/** Open a path in whichever app owns it, in a window of its own. */
export function openInApp(shell, filePath, fallback = 'home') {
  const which = appFor(kindFromExtension(filePath)) || fallback;
  shell.app.addRecent({ path: filePath, app: which }).catch(() => {});
  return shell.win.create({ app: which, file: filePath });
}

/** What the update service says, kept current in this window. */
export function useUpdate(shell) {
  const [update, setUpdate] = useState(null);
  useEffect(() => {
    if (!shell.update?.state) return;
    shell.update.state().then(setUpdate).catch(() => {});
  }, [shell]);
  useEffect(() => shell.on('update:state', setUpdate), [shell]);
  return update;
}

/** The frame. Apps supply the ribbon, the body and the status bar. */
export function AppFrame({ app, shell, title, subtitle, dirty, ribbon, status, menu, right, children }) {
  const [platform, setPlatform] = useState('win32');
  const update = useUpdate(shell);
  // The release this window said "Not now" to, and the one its status
  // chip asked to see again.
  const [away, setAway] = useState(null);
  const [insist, setInsist] = useState(null);

  useEffect(() => {
    shell.win.state().then((s) => setPlatform(s.platform)).catch(() => {});
  }, [shell]);

  useEffect(() => {
    shell.win.setTitle({ title: `${title ? `${title} — ` : ''}${app?.name || 'Rutba Office'} (beta)` }).catch(() => {});

    shell.win.setDocumentEdited({ edited: Boolean(dirty) }).catch(() => {});
  }, [title, dirty, app, shell]);

  return (
    <Window app={app?.key || 'home'}>
      <TitleBar
        app={app}
        title={title || app?.name}
        subtitle={subtitle}
        dirty={dirty}
        platform={platform}
        shell={shell}
        right={right}
        onMenu={menu?.open}
      />
      {ribbon}
      <Body>{children}</Body>
      <StatusBar>
        {status}
        <UpdateChip update={update} onShow={() => setInsist(update?.available || null)} />
      </StatusBar>
      <style>{UPDATE_CSS}</style>
      <UpdatePrompt
        shell={shell}
        update={update}
        away={away}
        insist={insist}
        onAway={(v) => {
          setAway(v);
          setInsist(null);
        }}
      />
      {menu?.node}
    </Window>
  );
}

/**
 * The word that a release is waiting — in every window, not only the
 * launcher, because a person who works in Rutba Word all day never sees the
 * launcher's chip.
 *
 * The service does the finding and the downloading; this only says so. A
 * release found is announced at once with the download's progress; once it
 * is downloaded the prompt offers "Restart and update", which quits,
 * installs and comes back. "Not now" puts that version away: in this
 * window at once, and through the service for a day everywhere, after
 * which it asks again — the install on quit goes on either way. A newer
 * release is a new question, whatever was put away before.
 *
 * And once a release has installed, the first windows of the new version
 * say so — which version this is now and which it was — with the release's
 * notes behind "What's new". OK marks the arrival seen through the
 * service, so every window drops the card at once.
 */
export function UpdatePrompt({ shell, update, away, insist, onAway }) {
  const [notes, setNotes] = useState(false);
  if (!update) return null;
  const seen = () => shell.update.seen?.().catch(() => {});
  if (!['available', 'downloading', 'ready'].includes(update.state)) {
    const arrived = update.arrived && !update.arrived.seen ? update.arrived : null;
    if (!arrived) return null;
    return (
      <>
        <div className="rw-update" role="status" aria-live="polite" data-state="arrived">
          <div className="rw-update-mark"><Icon name="star" size={18} /></div>
          <div className="rw-update-body">
            <div className="rw-update-title">Rutba Office is now {arrived.to}</div>
            <div className="rw-update-text">Updated from {arrived.from}. What changed is a click away.</div>
            <div className="rw-update-actions">
              <Button primary className="rw-update-whatsnew" label="What’s new" onClick={() => setNotes(true)} />
              <Button ghost className="rw-update-ok" label="OK" onClick={seen} />
            </div>
          </div>
        </div>
        {notes ? (
          <WhatsNew
            shell={shell}
            version={arrived.to}
            onClose={() => {
              setNotes(false);
              seen();
            }}
          />
        ) : null}
      </>
    );
  }
  // Put away — by "Not now" here, or by the service for a day — unless the
  // status chip asked to see it again.
  if ((update.snoozed || away === update.available) && insist !== update.available) return null;
  const ready = update.state === 'ready';
  const percent = Math.round(update.percent || 0);
  const when = update.releasedAt ? new Date(update.releasedAt).toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' }) : null;
  return (
    <>
      <div className="rw-update" role="status" aria-live="polite" data-state={update.state}>
      <div className="rw-update-mark"><Icon name="download" size={18} /></div>
      <div className="rw-update-body">
        <div className="rw-update-title">
          {ready ? `Rutba Office ${update.available} is ready to install` : `Rutba Office ${update.available} is available`}
        </div>
        <div className="rw-update-text">
          {ready
            ? 'Restart to update now. If you carry on working, it installs when you quit.'
            : `Downloading in the background — ${percent}%. You will be asked to restart when it is ready.`}
          {when ? ` Released ${when}.` : ''}
        </div>
        {!ready ? <Progress value={percent} max={100} /> : null}
        <div className="rw-update-actions">
          {ready ? <Button primary className="rw-update-restart" label="Restart and update" onClick={() => shell.update.install().catch(() => {})} /> : null}
          <Button ghost className="rw-update-later" label="Not now" onClick={() => { onAway(update.available); shell.update.snooze?.({ version: update.available }).catch(() => {}); }} />
        </div>
      </div>
      </div>
    </>
  );
}

/**
 * The word in the status bar: a download counting up, or a release ready.
 * The prompt can be put away; the chip stays, and brings it back.
 */
function UpdateChip({ update, onShow }) {
  if (!update) return null;
  if (update.state === 'downloading') {
    return (
      <span className="chip rw-update-chip" data-state="downloading" data-tip={`Downloading Rutba Office ${update.available} in the background`}>
        <Icon name="download" size={11} /> Updating {Math.round(update.percent || 0)}%
      </span>
    );
  }
  if (update.state === 'ready') {
    return (
      <button
        type="button"
        className="chip rw-update-chip"
        data-state="ready"
        data-tip={`Rutba Office ${update.available} is downloaded — restart to update, or it installs when you quit`}
        onClick={onShow}
      >
        <Icon name="download" size={11} /> Update ready
      </button>
    );
  }
  return null;
}

const UPDATE_CSS = `
.rw-update {
  position: fixed; right: 18px; bottom: 46px; width: 380px; max-width: calc(100vw - 36px);
  display: flex; gap: 12px; padding: 14px 16px; z-index: 60; font-size: 13px;
  background: var(--surface); color: var(--ink); border: 1px solid var(--line);
  border-radius: var(--r-3); box-shadow: var(--shadow-3);
}
.rw-update-mark {
  flex: none; width: 34px; height: 34px; border-radius: 10px; display: grid; place-items: center;
  background: var(--accent-soft); color: var(--accent);
}
.rw-update-body { flex: 1; min-width: 0; }
.rw-update-title { font-weight: 600; font-size: 13.5px; }
.rw-update-text { margin-top: 3px; color: var(--ink-2); line-height: 1.4; }
.rw-update .rw-progress { margin-top: 8px; }
.rw-update-actions { display: flex; gap: 8px; margin-top: 10px; align-items: center; }
.rw-status .rw-update-chip { color: var(--accent); background: var(--accent-soft); }
.rw-status button.rw-update-chip { border: 0; font: inherit; cursor: pointer; }
.rw-status button.rw-update-chip:hover { filter: brightness(0.94); }
`;

export { APPS, NEW_DOCUMENTS };
