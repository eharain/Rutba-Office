// Choose apps (owner, 2026-10-20: "since the number of applications have
// grown we should let users to choose what they wanted to install and what
// not"). Every app is one tile to turn on or off. An app turned off leaves
// the launcher and the templates and claims no file types, and stays
// installed: turning it on again is one click here, with no download.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Icon, t, msg } from '@rutba/office-ui';
import { APPS, APP_ORDER as ALL_APPS, APP_GROUPS, ADDONS as ALL_ADDONS, APPS_CHOICE_KEY, chosenApps as chosenOf, choiceFor, addonsChosen as addonsOf } from '@rutba/office-formats/registry';
import { isInstalled } from '@rutba/office-shell/client';

// Only the apps on this computer (docs/APP-PACKS.md): one not installed is
// neither offered here nor counted as chosen, whatever the stored choice says.
const APP_ORDER = ALL_APPS.filter(isInstalled);
const ADDONS = ALL_ADDONS.filter(isInstalled);
const chosenApps = (choice) => chosenOf(choice).filter(isInstalled);
const addonsChosen = (choice) => addonsOf(choice).filter(isInstalled);

/** The add-ons, as Choose apps shows them. */
const ADDON_INFO = () => ({ ai: { key: 'ai', short: t('common.ai-tools'), tagline: t('apps-choice.let-ai-apps'), icon: 'wand' } });
/** The sets' names and their one line each: Home, Choose apps and the Window menu say the same. */
export const setWords = () => ({
  office: { title: t('common.office'), line: t('apps-choice.mail-calls-documents') },
  studio: { title: t('common.studio'), line: t('apps-choice.pictures-video-recording') },
  addons: { title: t('apps-choice.add-ons'), line: t('apps-choice.work-inside-apps') },
});

/**
 * The installer's Choose apps page says these too (the rest of its words are
 * the ones above and the apps' names): marked here for the catalogue, and
 * taken from it into the installer by tools/make-installer-words.mjs.
 */
export const INSTALLER_WORDS = [
  msg('apps-choice.apps-want-see'),
  msg('apps-choice.tick-apps-add'),
  msg('apps-choice.choose-one-app'),
  msg('apps-choice.text-recognition'),
  msg('apps-choice.apps-to-install'),
  msg('apps-choice.only-apps-tick'),
  msg('apps-choice.tick-apps-install'),
];

/**
 * The apps this person has chosen, read again whenever the window comes
 * forward (the choice may have changed in another window), and a way to
 * change it. `made` is false until a choice has been made at all.
 */
export function useChosenApps(shell) {
  const [state, setState] = useState({ apps: chosenApps(null), addons: [], made: true, loaded: false });
  const read = useCallback(() => {
    shell.store.get({ key: APPS_CHOICE_KEY, fallback: null })
      .then((choice) => setState({ apps: chosenApps(choice), addons: addonsChosen(choice), made: Boolean(choice), loaded: true }))
      .catch(() => setState((s) => ({ ...s, loaded: true })));
  }, [shell]);
  useEffect(() => {
    read();
    window.addEventListener('focus', read);
    return () => window.removeEventListener('focus', read);
  }, [read]);
  const choose = useCallback(async (keys, addons = []) => {
    // An app not installed is kept on in the stored choice: added later, it shows.
    const choice = choiceFor([...keys, ...ALL_APPS.filter((k) => !isInstalled(k))], [...addons, ...ALL_ADDONS.filter((a) => !isInstalled(a))]);
    await shell.store.set({ key: APPS_CHOICE_KEY, value: choice });
    setState({ apps: chosenApps(choice), addons: addonsChosen(choice), made: true, loaded: true });
    // The kinds of file Windows sends to Office follow the choice.
    shell.defaults?.follow?.({ apps: chosenApps(choice) })?.catch?.(() => {});
  }, [shell]);
  return { ...state, choose };
}

/** The Choose apps window: a tile for each app, turned on or off. */
export function AppsChooser({ chosen, addons = [], onChoose, onClose }) {
  const [on, setOn] = useState(() => new Set([...chosen, ...addons]));
  const flip = (key) => setOn((s) => {
    const next = new Set(s);
    if (next.has(key)) next.delete(key);
    else next.add(key);
    return next;
  });
  const count = APP_ORDER.filter((k) => on.has(k)).length;
  const tile = (item) => {
    const picked = on.has(item.key);
    return (
      <button key={item.key} type="button" role="checkbox" aria-checked={picked} className={`apps-tile${picked ? ' on' : ''}`} data-app={item.key} data-tip={`${item.short}: ${item.tagline}`} onClick={() => flip(item.key)}>
        <span className="apps-glyph"><Icon name={item.icon} size={15} /></span>
        <strong className="apps-name">{item.short}</strong>
        <span className="apps-check" aria-hidden="true" />
      </button>
    );
  };
  const words = setWords();
  const head = (id) => (
    <div className="apps-group-head">
      <h3>{words[id].title}</h3>
      <span>{words[id].line}</span>
    </div>
  );
  return (
    <Dialog
      title={t('apps-choice.choose-apps')}
      width={720}
      className="apps-chooser"
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.all-apps')} className="apps-all" onClick={() => setOn((s) => new Set([...APP_ORDER, ...ADDONS.filter((a) => s.has(a))]))} />
          <span style={{ flex: 1 }} />
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.done')} className="apps-done" disabled={!count} onClick={() => onChoose(APP_ORDER.filter((k) => on.has(k)), ADDONS.filter((a) => on.has(a)))} />
        </>
      }
    >
      <style>{CHOOSER_CSS}</style>
      <p className="apps-lead">{t('apps-choice.turn-apps-add')}</p>
      {APP_GROUPS.filter((g) => g.apps.some(isInstalled)).map((g) => (
        <section key={g.id} className="apps-group" data-group={g.id}>
          {head(g.id)}
          <div className="apps-grid" role="group" aria-label={words[g.id].title}>{g.apps.filter(isInstalled).map((key) => tile(APPS[key]))}</div>
        </section>
      ))}
      {ADDONS.length ? (
        <section className="apps-group" data-group="addons">
          {head('addons')}
          <div className="apps-grid" role="group" aria-label={words.addons.title}>{ADDONS.map((a) => tile(ADDON_INFO()[a]))}</div>
        </section>
      ) : null}
      <p className="apps-count" aria-live="polite">{count ? t('apps-choice.n-total-apps', { n: count, total: APP_ORDER.length }) : t('apps-choice.choose-one-app')}</p>
    </Dialog>
  );
}

const CHOOSER_CSS = `
.apps-lead { margin: 0 0 14px; color: var(--ink-2); font-size: 13px; line-height: 1.5; }
.apps-group + .apps-group { margin-top: 14px; }
.apps-group-head { display: flex; align-items: baseline; gap: 10px; margin: 0 2px 7px; }
.apps-group-head h3 { margin: 0; font-size: 14.5px; font-weight: 650; color: var(--ink); }
.apps-group-head span { font-size: 12px; color: var(--ink-3); }
.apps-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px; }
.apps-tile { display: flex; align-items: center; gap: 10px; min-height: 42px; padding: 7px 10px; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--surface); color: var(--ink-2); font: inherit; text-align: start; cursor: pointer; }
.apps-tile:hover { border-color: var(--accent-line); background: var(--wash); }
.apps-tile.on { color: var(--ink); }
.apps-glyph { width: 26px; height: 26px; border-radius: 7px; display: grid; place-items: center; background: var(--sunken); flex: none; }
.apps-tile.on .apps-glyph { background: var(--accent-fill); color: #fff; }
.apps-name { flex: 1; min-width: 0; font-weight: 500; font-size: 13.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.apps-check { position: relative; width: 34px; height: 20px; border-radius: 10px; background: var(--n-50, var(--line)); flex: none; transition: background var(--fast); }
.apps-check::after { content: ''; position: absolute; top: 3px; inset-inline-start: 3px; width: 14px; height: 14px; border-radius: 50%; background: #fff; transition: inset-inline-start var(--fast); }
.apps-tile.on .apps-check { background: var(--accent-fill); }
.apps-tile.on .apps-check::after { inset-inline-start: 17px; }
.apps-count { margin: 12px 0 0; font-size: 12px; color: var(--ink-3); }
`;
