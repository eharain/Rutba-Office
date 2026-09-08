// What every app window has in common.
//
// The frame, the app menu behind the mark in the corner, the open/save
// plumbing, and the two or three behaviours that must be identical everywhere:
// a document window asks before closing on unsaved work, the theme toggle is in
// the same place, and files dropped on a window open in the right app.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Window, TitleBar, Body, StatusBar, Menu, Icon, useToast, useTheme, Button } from '@rutba/office-ui';
import { APPS, openFilters, saveFilters, NEW_DOCUMENTS } from '@rutba/office-formats/registry';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';

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

/** The frame. Apps supply the ribbon, the body and the status bar. */
export function AppFrame({ app, shell, title, subtitle, dirty, ribbon, status, menu, right, children }) {
  const [platform, setPlatform] = useState('win32');

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
      <StatusBar>{status}</StatusBar>
      {menu?.node}
    </Window>
  );
}

export { APPS, NEW_DOCUMENTS };
