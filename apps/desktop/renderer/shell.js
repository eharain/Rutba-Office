// What every app window has in common.
//
// The frame, the app menu behind the mark in the corner, the open/save
// plumbing, and the two or three behaviours that must be identical everywhere:
// a document window asks before closing on unsaved work, the theme toggle is in
// the same place, and files dropped on a window open in the right app.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Window, TitleBar, Body, StatusBar, Menu, Dialog, Icon, useToast, useTheme, Button, Progress, t, msg, language, windowCommands, prettyKey } from '@rutba/office-ui';
import { APPS, APP_GROUPS, openFilters, saveFilters, NEW_DOCUMENTS } from '@rutba/office-formats/registry';
import { appFor, kindFromExtension } from '@rutba/office-formats/sniff';
import { pathOf, installedPacks } from '@rutba/office-shell/client';
import { WhatsNew } from './whatsnew.js';
import { filtersInLanguage } from './registry-words.js';
import { TemplateGallery } from './template-gallery.js';
import { AiBadge, AiConfirm, AI_TOOLS_CSS } from './ai-tools.js';

// The AI Tools window is AI Tools' own, in its pack: loaded when it is opened.
const AiToolsDialog = React.lazy(() => import('./ai-tools-dialog.js').then((m) => ({ default: m.AiToolsDialog })));
import { useChosenApps, setWords } from './apps-choice.js';
import { ProfileBadge, ManageProfiles, KeystoreNotice, PROFILES_CSS, useProfiles, profileName, openManageProfiles } from './profiles.js';

/**
 * The languages the windows can be shown in, each named in its own words
 * (none, the first, follows the system). A language is offered once its
 * catalogue is whole.
 */
const LANGUAGES = [[null, null], ['en', 'English'], ['ur', 'اردو'], ['ar', 'العربية'], ['hi', 'हिन्दी']];

/**
 * The digits a window in one of these languages may write its numbers in,
 * Western first (the default, none stored), then the language's own.
 */
const DIGITS = {
  ar: [[null, msg('Western digits (123)')], ['arab', msg('Arabic-Indic digits (١٢٣)')]],
  ur: [[null, msg('Western digits (123)')], ['arabext', msg('Eastern Arabic-Indic digits (۱۲۳)')]],
  hi: [[null, msg('Western digits (123)')], ['deva', msg('Devanagari digits (१२३)')]],
};

/** A window's address with its language set, or taken off for the system's. */
function withLanguage(href, tag, key = 'lang') {
  const url = new URL(href);
  if (tag) url.searchParams.set(key, tag);
  else url.searchParams.delete(key);
  return url.toString();
}

/**
 * Help ▸ Keyboard Shortcuts: every command this window has a key for, as the
 * ribbon and the keys run it, by name; read from the window's own commands,
 * so it is never behind what the keys do.
 */
function ShortcutsSheet({ onClose }) {
  const [rows] = useState(() => {
    const seen = new Set();
    const out = [];
    for (const c of Object.values(windowCommands())) {
      if (!c?.key || !c.label) continue;
      const label = String(c.label).replace(/…$/, '');
      const key = prettyKey(c.key);
      if (seen.has(`${label}\u0000${key}`)) continue;
      seen.add(`${label}\u0000${key}`);
      out.push({ label, key });
    }
    return out.sort((a, b) => a.label.localeCompare(b.label));
  });
  return (
    <Dialog title={t('Keyboard Shortcuts')} className="rw-shortcuts" width={480} onClose={onClose} actions={<Button label={t('Close')} onClick={onClose} />}>
      {rows.length ? (
        <table className="rw-shortcuts-list">
          <tbody>{rows.map((r) => <tr key={`${r.label}:${r.key}`}><td>{r.label}</td><td><kbd>{r.key}</kbd></td></tr>)}</tbody>
        </table>
      ) : <p>{t('This window has no keyboard shortcuts of its own.')}</p>}
    </Dialog>
  );
}

/** The site's name for each app, for its how-to films (office.rutba.io/how-to). */
const SITE_APP = { sheets: 'worksheets', slides: 'presentation', base: 'data-files' };

/** How many recent files each list shows. */
const RECENT_MAX = 15;

/**
 * The menu behind the app mark, in the same order in every app (the owner,
 * 2026-11-02, D-2026-11-02-09: "menu in 2.22"); a row an app does not have
 * is left out, never greyed:
 *   New, New from Template…, the app's other new things, Open…, Open Recent ▸
 *   Save, Save As…, Export ▸, Print…          the app's slots
 *   the app's own few                           `own`
 *   Window ▸, Settings ▸, Help ▸
 *   Close Window
 * An app fills the slots it has (`save` and `print` are what each runs;
 * `saveAs` too, or the list of its ways to save a copy; `exports` and `own`
 * lists of items; `newItems` its other new things). A slot the app leaves
 * empty is taken from its commands when it has one that can run now
 * (file.save, file.saveAs, file.export, file.pdf in the document apps,
 * file.print), so the menu says what the ribbon and the keys do. Open Recent has this app's files first, then the other apps' in
 * a submenu of their own.
 */
export function useAppMenu({ shell, appKey, onNew, onOpen, newItems = [], save, saveAs, exports = [], print, own = [] }) {
  const [menu, setMenu] = useState(null);
  const [shortcuts, setShortcuts] = useState(false);
  const { mode, setTheme } = useTheme();
  const toast = useToast();
  // The language setting as stored: a tag, or none for the system's.
  const [chosenLanguage, setChosenLanguage] = useState(undefined);
  const [chosenDigits, setChosenDigits] = useState(null);
  // File > New's template gallery, in the three document apps.
  const [gallery, setGallery] = useState(false);
  const galleryApp = ['word', 'sheets', 'slides'].includes(appKey) ? appKey : null;
  // The add-ons added in Choose apps: AI Tools is in this menu once it is.
  const { addons: chosenAddons, apps: chosenAppKeys } = useChosenApps(shell);

  useEffect(() => {
    shell.store.get({ key: 'language', fallback: null }).then((v) => setChosenLanguage(v || null)).catch(() => setChosenLanguage(null));
    shell.store.get({ key: 'digits', fallback: null }).then((v) => setChosenDigits(v || null)).catch(() => setChosenDigits(null));
  }, [shell]);

  /**
   * A window keeps the language it opened in, as Office's do: the choice is
   * kept for the windows opened after it, and the launcher, which holds no
   * document, opens again in it at once.
   */
  const chooseLanguage = useCallback(async (tag) => {
    await shell.store.set({ key: 'language', value: tag || null });
    setChosenLanguage(tag || null);
    if (appKey === 'home') {
      window.location.replace(withLanguage(window.location.href, tag));
      return;
    }
    toast(t('The windows you open from now on are in the language chosen. Restart Rutba Office to see every window in it.'), { ms: 5000 });
  }, [shell, appKey, toast]);

  /** The digits, kept the same way: the launcher at once, other windows as they open. */
  const chooseDigits = useCallback(async (system) => {
    await shell.store.set({ key: 'digits', value: system || null });
    setChosenDigits(system || null);
    if (appKey === 'home') {
      window.location.replace(withLanguage(window.location.href, system, 'digits'));
      return;
    }
    toast(t('The windows you open from now on write numbers in the digits chosen.'), { ms: 5000 });
  }, [shell, appKey, toast]);

  const open = useCallback(
    async (e) => {
      const rect = e.currentTarget.getBoundingClientRect();
      // Right to left, the menu drops from the mark's far edge and its submenus open to the left.
      const rtl = getComputedStyle(e.currentTarget).direction === 'rtl';
      const keyboard = e.detail === 0;
      // The recent files and this app's windows as they are now, not as they were when the window opened.
      const [recent, windows] = await Promise.all([shell.app.recent().catch(() => []), shell.win.list?.().catch(() => []) ?? []]);
      const appOf = (r) => r.app || appFor(kindFromExtension(r.path)) || 'home';
      const fileRow = (r) => ({ label: r.name, icon: 'file', title: r.path, run: () => shell.win.create({ app: appOf(r), file: r.path }) });
      const groups = [];

      // New and open.
      const first = [];
      if (onNew) first.push({ label: t('New'), icon: 'new', key: prettyKey('Mod+N'), run: onNew });
      if (galleryApp) first.push({ label: t('New from Template…'), icon: 'grid', run: () => setGallery(true) });
      first.push(...newItems);
      if (onOpen) first.push({ label: t('Open…'), icon: 'open', key: prettyKey('Mod+O'), run: onOpen });
      const mine = appKey === 'home' ? recent : recent.filter((r) => appOf(r) === appKey);
      const others = appKey === 'home' ? [] : recent.filter((r) => appOf(r) !== appKey);
      if (recent.length) {
        const list = mine.slice(0, RECENT_MAX).map(fileRow);
        if (others.length) list.push(...(list.length ? ['-'] : []), { label: t('Other apps’ files'), icon: 'grid', items: others.slice(0, RECENT_MAX).map(fileRow) });
        list.push('-', { label: t('Clear the List'), icon: 'trash', run: () => shell.app.clearRecent().catch(() => {}) });
        first.push({ label: t('Open Recent'), icon: 'clock', items: list });
      }
      groups.push(first);

      // The document's own: save, export, print; an empty slot from the window's commands, when one can run now.
      const cmds = windowCommands();
      const can = (id) => {
        const c = cmds[id];
        return c && !c.disabled && (!c.when || c.when()) ? c : null;
      };
      const asItem = (c, label) => ({ label: label || c.label, icon: c.icon, key: prettyKey(c.key), run: () => c.run() });
      const doc = [];
      if (save) doc.push({ label: t('Save'), icon: 'save', key: prettyKey('Mod+S'), run: save });
      else if (can('file.save')) doc.push(asItem(can('file.save'), t('Save')));
      // Save As: one way, or the app's several (a copy, the other form of the file, a template), each a row.
      if (typeof saveAs === 'function') doc.push({ label: t('Save As…'), icon: 'save', key: prettyKey('Mod+Shift+S'), run: saveAs });
      else if (Array.isArray(saveAs)) doc.push(...saveAs.filter(Boolean));
      else if (can('file.saveAs')) doc.push(asItem(can('file.saveAs')));
      const out = exports.length ? exports : [can('file.export'), galleryApp ? can('file.pdf') : null].filter(Boolean).map((c) => asItem(c));
      if (out.length === 1) doc.push(out[0]);
      else if (out.length) doc.push({ label: t('Export'), icon: 'export', items: out });
      if (print) doc.push({ label: t('Print…'), icon: 'print', key: prettyKey('Mod+P'), run: print });
      else if (can('file.print')) doc.push(asItem(can('file.print'), t('Print…')));
      groups.push(doc);

      // The app's own few.
      groups.push(own);

      // Windows, settings and help.
      const otherWindows = windows.filter((w) => !w.current);
      const settings = [
        { label: t('Theme'), icon: 'theme', items: [
          { label: t('Light theme'), checked: mode === 'light', run: () => setTheme('light') },
          { label: t('Dark theme'), checked: mode === 'dark', run: () => setTheme('dark') },
          { label: t('As the system is'), checked: mode === 'system', run: () => setTheme('system') },
        ] },
        // The language, each named in its own words, the window's own ticked.
        { label: t('Language'), icon: 'globe', items: LANGUAGES.map(([tag, name]) => ({ label: tag ? name : t('As the system is'), checked: (tag || null) === (chosenLanguage ?? null), run: () => chooseLanguage(tag) })) },
      ];
      // The digits, in a window whose language has digits of its own.
      const digits = DIGITS[language()];
      if (digits) settings.push({ label: t('Digits'), icon: 'listNumber', items: digits.map(([system, label]) => ({ label: t(label), checked: (system || null) === (chosenDigits ?? null), run: () => chooseDigits(system) })) });
      // Profiles: the people, or the parts of one's work, each in its own windows (profiles.js).
      settings.push('-', { label: t('Profiles…'), icon: 'contacts', run: () => openManageProfiles() });
      // AI Tools is an add-on: in the menu once it is added in Choose apps.
      if (chosenAddons.includes('ai')) settings.push({ label: t('AI Tools…'), icon: 'wand', run: () => window.dispatchEvent(new Event('rutba-ai-tools')) });
      const site = appKey === 'home' ? '' : `#${SITE_APP[appKey] || appKey}`;
      // The two sets, each with the apps on in it (an app turned off leaves the menus).
      const appSets = APP_GROUPS
        .map((g) => ({
          label: setWords()[g.id].title,
          icon: g.id === 'office' ? 'word' : 'video',
          items: g.apps.filter((k) => chosenAppKeys.includes(k)).map((k) => ({ label: APPS[k].short, icon: APPS[k].icon, run: () => shell.win.create({ app: k, query: { start: '1' } }) })),
        }))
        .filter((s) => s.items.length);
      groups.push([
        { label: t('Window'), icon: 'window', items: [
          ...otherWindows.map((w) => ({ label: w.name || t('Untitled'), icon: 'window', run: () => shell.win.focus({ id: w.id }) })),
          ...(otherWindows.length ? ['-'] : []),
          ...appSets,
          { label: t('All apps'), icon: 'grid', run: () => shell.win.create({ app: 'home' }) },
        ] },
        { label: t('Settings'), icon: 'settings', items: settings },
        { label: t('Help'), icon: 'info', items: [
          { label: t('What’s new'), icon: 'star', run: () => shell.win.create({ app: 'home', query: { whatsnew: 1 } }) },
          { label: t('Keyboard Shortcuts'), icon: 'mouse', run: () => setShortcuts(true) },
          { label: t('How-to Films'), icon: 'play', run: () => shell.shell.openExternal({ url: `https://office.rutba.io/how-to${site}` }).catch(() => {}) },
          { label: t('Contact Us'), icon: 'mail', run: () => shell.shell.openExternal({ url: 'https://office.rutba.io/contact' }).catch(() => {}) },
          '-',
          { label: t('About Rutba Office'), icon: 'info', run: () => shell.win.create({ app: 'home', query: { about: 1 } }) },
        ] },
      ]);
      // The launcher is the suite itself: closing it is quitting, which its own window does.
      if (appKey !== 'home') groups.push([{ label: t('Close Window'), icon: 'close', key: prettyKey('Mod+W'), run: () => shell.win.close({}).catch(() => {}) }]);

      const items = [];
      for (const g of groups.filter((x) => x.length)) items.push(...(items.length ? ['-'] : []), ...g);
      setMenu({ x: rtl ? rect.right : rect.left, y: rect.bottom + 4, rtl, items, keyboard });
    },
    [onNew, onOpen, newItems, save, saveAs, exports, print, own, mode, setTheme, shell, appKey, galleryApp, chosenLanguage, chooseLanguage, chosenDigits, chooseDigits, chosenAddons, chosenAppKeys]
  );

  return {
    open,
    openGallery: () => setGallery(true),
    node: (
      <>
        {menu ? <Menu {...menu} onClose={() => setMenu(null)} /> : null}
        {shortcuts ? <ShortcutsSheet onClose={() => setShortcuts(false)} /> : null}
        {gallery && galleryApp ? <TemplateGallery shell={shell} app={galleryApp} onClose={() => setGallery(false)} /> : null}
      </>
    ),
  };
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
    message: t('Save changes to {name}?', { name }),
    detail: t('Your changes will be lost if you don’t save them.'),
    buttons: [t('Save'), t("Don't save"), t('Cancel')],
    defaultId: 0,
    cancelId: 2,
  });
  return response === 0 ? 'save' : response === 1 ? 'discard' : 'cancel';
}

export async function pickOpen(shell, appKey, { multiple = false } = {}) {
  const paths = await shell.dialog.open({
    title: t('Open'),
    // The launcher offers the kinds of file of the apps installed, and no others.
    filters: filtersInLanguage(openFilters(appKey, { apps: installedPacks() })),
    multiple,
  });
  return multiple ? paths : paths[0] || null;
}

/**
 * View → Window → Arrange All, Side by Side, Cascade: this app's windows laid
 * out by the shell, which knows where every one of them is. Says so when
 * there is nothing to arrange beside this one.
 */
export async function arrangeWindows(shell, toast, mode) {
  const { count } = await shell.win.arrange({ mode });
  if (count < 2) {
    toast(mode === 'sideBySide'
      ? t('Open a second window first — Side by Side lays two of this app\'s windows next to each other.')
      : t('This is the only window to arrange.'), { ms: 4000 });
  }
  return count;
}

/**
 * View → Window → Switch Windows (or Unhide, with `hiddenOnly`): this app's
 * windows as a menu, opened where the button was pressed. The point is taken
 * before the list is asked for, since the press is over by the time it comes.
 */
export async function openWindowMenu(e, menu, shell, { hiddenOnly = false, toast = null } = {}) {
  const at = { clientX: e.clientX, clientY: e.clientY, preventDefault() {}, stopPropagation() {} };
  e.preventDefault?.();
  const list = (await shell.win.list()).filter((w) => !hiddenOnly || w.hidden);
  if (!list.length) {
    toast?.(hiddenOnly ? t('No window of this app is hidden.') : t('There are no other windows.'), { ms: 3500 });
    return;
  }
  menu.open(at, list.map((w) => ({
    label: w.name || t('Untitled'),
    icon: w.current ? 'check' : undefined,
    run: () => (hiddenOnly ? shell.win.unhide({ id: w.id }) : shell.win.focus({ id: w.id })),
  })));
}

/**
 * The two forms of Open XML Documents, Worksheets and Presentations save,
 * named as Word, Excel and PowerPoint name them: the usual one and
 * ISO/IEC 29500 Strict.
 */
const OPEN_XML = {
  word: { ext: 'docx', usual: () => t('Word Document (.docx)'), strict: () => t('Strict Open XML Document (.docx)'), asUsual: () => t('Save as Word Document…'), asStrict: () => t('Save as Strict Open XML Document…') },
  sheets: { ext: 'xlsx', usual: () => t('Excel Workbook (.xlsx)'), strict: () => t('Strict Open XML Spreadsheet (.xlsx)'), asUsual: () => t('Save as Excel Workbook…'), asStrict: () => t('Save as Strict Open XML Spreadsheet…') },
  slides: { ext: 'pptx', usual: () => t('PowerPoint Presentation (.pptx)'), strict: () => t('Strict Open XML Presentation (.pptx)'), asUsual: () => t('Save as PowerPoint Presentation…'), asStrict: () => t('Save as Strict Open XML Presentation…') },
};

/**
 * Save As. A Strict file is saved as Strict, and the dialog names the form
 * so: its Open XML filter called what Word, Excel or PowerPoint calls Strict.
 */
export async function pickSave(shell, appKey, defaultPath, conformance = null) {
  const form = OPEN_XML[appKey];
  const filters = filtersInLanguage(saveFilters(appKey)).map((f) => (form && conformance === 'strict' && f.extensions[0] === form.ext ? { ...f, name: form.strict() } : f));
  return shell.dialog.save({ title: t('Save as'), filters, defaultPath });
}

/** File → Save as the other form's words: Strict for a file in the usual form, the usual form for a Strict one. */
export function otherFormLabel(appKey, doc) {
  const form = OPEN_XML[appKey];
  return doc?.conformance === 'strict' ? form.asUsual() : form.asStrict();
}

/** File → Save as the other form: { target, conformance }, or null when the dialog is cancelled. */
export async function pickSaveOtherForm(shell, appKey, doc) {
  const form = OPEN_XML[appKey];
  const strict = doc?.conformance !== 'strict';
  const base = String(doc?.path || doc?.name || '').replace(/\.[^.\\/]+$/, '') || 'Untitled';
  const target = await shell.dialog.save({ title: t('Save as'), filters: [{ name: strict ? form.strict() : form.usual(), extensions: [form.ext] }], defaultPath: `${base}.${form.ext}` });
  return target ? { target, conformance: strict ? 'strict' : 'transitional' } : null;
}

/** Each app's template, by the name Office gives it. */
export const TEMPLATE_KINDS = { word: { ext: 'dotx', label: t('Word Template (.dotx)') }, sheets: { ext: 'xltx', label: t('Excel Template (.xltx)') }, slides: { ext: 'potx', label: t('PowerPoint Template (.potx)') } };

/**
 * Where templates of one's own are kept: Office's own folder for them,
 * Documents\Custom Office Templates, so Word, Excel and PowerPoint list the
 * same ones as File → New here does.
 */
export async function templatesFolder(shell) {
  const { documents } = await shell.app.paths();
  const sep = String(documents).includes('\\') ? '\\' : '/';
  return `${documents}${sep}Custom Office Templates`;
}

/** File → Save as Template: a name in the templates folder, made if it is not there yet. */
export async function pickSaveTemplate(shell, appKey, name) {
  const kind = TEMPLATE_KINDS[appKey];
  const dir = await templatesFolder(shell);
  await shell.fs.mkdirp({ path: dir }).catch(() => {});
  const base = String(name || 'Template').replace(/\.[^.]+$/, '');
  const sep = dir.includes('\\') ? '\\' : '/';
  return shell.dialog.save({ title: t('Save as Template'), filters: [{ name: kind.label, extensions: [kind.ext] }], defaultPath: `${dir}${sep}${base}.${kind.ext}` });
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
      const files = [...(e.dataTransfer?.files || [])].map(pathOf).filter(Boolean);
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
export function AppFrame({ app, shell, title, subtitle, dirty, ribbon, status, menu, right, titleAfter, children }) {
  const [platform, setPlatform] = useState('win32');
  const update = useUpdate(shell);
  // The release this window said "Not now" to, and the one its status
  // chip asked to see again.
  const [away, setAway] = useState(null);
  const [insist, setInsist] = useState(null);
  // AI Tools, from the title bar's badge or the app menu (an 'rutba-ai-tools' event).
  const [aiTools, setAiTools] = useState(() => typeof location !== 'undefined' && new URLSearchParams(location.search).has('aitools'));
  useEffect(() => {
    const open = () => setAiTools(true);
    window.addEventListener('rutba-ai-tools', open);
    return () => window.removeEventListener('rutba-ai-tools', open);
  }, []);
  // Manage Profiles, from the profile menu or the app menu (a 'rutba-profiles' event), or ?profiles.
  const [profilesOpen, setProfilesOpen] = useState(() => (typeof location !== 'undefined' && new URLSearchParams(location.search).has('profiles') ? { mode: null } : null));
  useEffect(() => {
    const open = (e) => setProfilesOpen({ mode: e?.detail?.mode || null });
    window.addEventListener('rutba-profiles', open);
    return () => window.removeEventListener('rutba-profiles', open);
  }, []);
  // The profile the window belongs to, in its title too (the taskbar's), once there are two.
  const { me: profile } = useProfiles(shell);

  useEffect(() => {
    shell.win.state().then((s) => setPlatform(s.platform)).catch(() => {});
  }, [shell]);

  useEffect(() => {
    const of = profile && profile.count > 1 ? ` · ${profileName(profile)}` : '';
    shell.win.setTitle({ title: `${title ? `${title} \u2014 ` : ''}${app?.name || 'Rutba Office'} (beta)${of}` }).catch(() => {});

    shell.win.setDocumentEdited({ edited: Boolean(dirty) }).catch(() => {});
  }, [title, dirty, app, shell, profile]);

  return (
    <Window app={app?.key || 'home'}>
      <TitleBar
        app={app}
        title={title || app?.name}
        subtitle={subtitle}
        dirty={dirty}
        titleAfter={titleAfter}
        platform={platform}
        shell={shell}
        right={<>{right}<AiBadge shell={shell} onOpen={() => setAiTools(true)} /><ProfileBadge shell={shell} appKey={app?.key || 'home'} always={(app?.key || 'home') === 'home'} /></>}
        onMenu={menu?.open}
      />
      <style>{PROFILES_CSS}</style>
      <style>{AI_TOOLS_CSS}</style>
      {aiTools ? <React.Suspense fallback={null}><AiToolsDialog shell={shell} onClose={() => setAiTools(false)} /></React.Suspense> : null}
      <AiConfirm shell={shell} />
      {profilesOpen ? <ManageProfiles shell={shell} mode={profilesOpen.mode} onClose={() => setProfilesOpen(null)} /> : null}
      <KeystoreNotice shell={shell} />
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
 * launcher, because a person who works in Documents all day never sees the
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
            <div className="rw-update-title">{t('Rutba Office is now {version}', { version: arrived.to })}</div>
            <div className="rw-update-text">{t('Updated from {version}. What changed is a click away.', { version: arrived.from })}</div>
            <div className="rw-update-actions">
              <Button primary className="rw-update-whatsnew" label={t('What’s new')} onClick={() => setNotes(true)} />
              <Button ghost className="rw-update-ok" label={t('OK')} onClick={seen} />
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
          {ready ? t('Rutba Office {version} is ready to install', { version: update.available }) : t('Rutba Office {version} is available', { version: update.available })}
        </div>
        <div className="rw-update-text">
          {ready
            ? t('Restart to update now. If you carry on working, it installs when you quit.')
            : t('Downloading in the background — {percent}%. You will be asked to restart when it is ready.', { percent })}
          {when ? ` ${t('Released {date}.', { date: when })}` : ''}
        </div>
        {!ready ? <Progress value={percent} max={100} /> : null}
        <div className="rw-update-actions">
          {ready ? <Button primary className="rw-update-restart" label={t('Restart and update')} onClick={() => shell.update.install().catch(() => {})} /> : null}
          <Button ghost className="rw-update-later" label={t('Not now')} onClick={() => { onAway(update.available); shell.update.snooze?.({ version: update.available }).catch(() => {}); }} />
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
      <span className="chip rw-update-chip" data-state="downloading" data-tip={t('Downloading Rutba Office {version} in the background', { version: update.available })}>
        <Icon name="download" size={11} /> {t('Updating {percent}%', { percent: Math.round(update.percent || 0) })}
      </span>
    );
  }
  if (update.state === 'ready') {
    return (
      <button
        type="button"
        className="chip rw-update-chip"
        data-state="ready"
        data-tip={t('Rutba Office {version} is downloaded — restart to update, or it installs when you quit', { version: update.available })}
        onClick={onShow}
      >
        <Icon name="download" size={11} /> {t('Update ready')}
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
  background: var(--accent-soft); color: var(--accent-text);
}
.rw-update-body { flex: 1; min-width: 0; }
.rw-update-title { font-weight: 600; font-size: 13.5px; }
.rw-update-text { margin-top: 3px; color: var(--ink-2); line-height: 1.4; }
.rw-update .rw-progress { margin-top: 8px; }
.rw-update-actions { display: flex; gap: 8px; margin-top: 10px; align-items: center; }
.rw-status .rw-update-chip { color: var(--accent-text); background: var(--accent-soft); }
.rw-status button.rw-update-chip { border: 0; font: inherit; cursor: pointer; }
.rw-status button.rw-update-chip:hover { filter: brightness(0.94); }
`;

export { APPS, NEW_DOCUMENTS };
