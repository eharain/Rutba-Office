// Choose apps (owner, 2026-10-20: "since the number of applications have
// grown we should let users to choose what they wanted to install and what
// not"). Every app is one tile to turn on or off. An app turned off leaves
// the launcher and the templates and claims no file types, and stays
// installed: turning it on again is one click here, with no download.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Icon, t } from '@rutba/office-ui';
import { APPS, APP_ORDER, APP_GROUPS, ADDONS, APPS_CHOICE_KEY, chosenApps, choiceFor, addonsChosen } from '@rutba/office-formats/registry';

/** The add-ons, as Choose apps shows them. */
const ADDON_INFO = () => ({ ai: { key: 'ai', short: t('AI Tools'), tagline: t('Let AI apps on this computer make and read documents with Office'), icon: 'wand' } });
const GROUP_TITLES = () => ({ office: t('Office apps'), more: t('More apps'), addons: t('Add-ons') });

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
    const choice = choiceFor(keys, addons);
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
      <button key={item.key} type="button" role="checkbox" aria-checked={picked} className={`apps-tile${picked ? ' on' : ''}`} data-app={item.key} onClick={() => flip(item.key)}>
        <span className="apps-glyph"><Icon name={item.icon} size={20} /></span>
        <span className="apps-text">
          <strong>{item.short}</strong>
          <span>{item.tagline}</span>
        </span>
        <span className="apps-check" aria-hidden="true">{picked ? <Icon name="check" size={14} /> : null}</span>
      </button>
    );
  };
  const titles = GROUP_TITLES();
  return (
    <Dialog
      title={t('Choose apps')}
      width={720}
      className="apps-chooser"
      onClose={onClose}
      actions={
        <>
          <Button label={t('All apps')} className="apps-all" onClick={() => setOn((s) => new Set([...APP_ORDER, ...ADDONS.filter((a) => s.has(a))]))} />
          <span style={{ flex: 1 }} />
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Done')} className="apps-done" disabled={!count} onClick={() => onChoose(APP_ORDER.filter((k) => on.has(k)), ADDONS.filter((a) => on.has(a)))} />
        </>
      }
    >
      <style>{CHOOSER_CSS}</style>
      <p className="apps-lead">{t('Turn on the apps and add-ons you want, now or at any time. An app turned off is not shown and does not open its kinds of file; it stays installed, so it can be turned on again here.')}</p>
      {APP_GROUPS.map((g) => (
        <section key={g.id} className="apps-group" data-group={g.id}>
          <h3>{titles[g.id]}</h3>
          <div className="apps-grid" role="group" aria-label={titles[g.id]}>{g.apps.map((key) => tile(APPS[key]))}</div>
        </section>
      ))}
      <section className="apps-group" data-group="addons">
        <h3>{titles.addons}</h3>
        <div className="apps-grid" role="group" aria-label={titles.addons}>{ADDONS.map((a) => tile(ADDON_INFO()[a]))}</div>
      </section>
      <p className="apps-count" aria-live="polite">{count ? t('{n} of {total} apps on', { n: count, total: APP_ORDER.length }) : t('Choose one app at least.')}</p>
    </Dialog>
  );
}

const CHOOSER_CSS = `
.apps-lead { margin: 0 0 14px; color: var(--ink-2); font-size: 13px; line-height: 1.5; }
.apps-group h3 { margin: 14px 0 8px; font-size: 12px; font-weight: 650; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-3); }
.apps-group:first-of-type h3 { margin-top: 0; }
.apps-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 8px; }
.apps-tile { display: flex; align-items: center; gap: 10px; padding: 10px 12px; border-radius: var(--r-2); border: 1px solid var(--line-soft); background: var(--surface); color: var(--ink-3); font: inherit; text-align: start; cursor: pointer; }
.apps-tile:hover { border-color: var(--line); background: var(--wash); }
.apps-tile.on { color: var(--ink); border-color: var(--accent); background: var(--wash-strong); }
.apps-glyph { width: 34px; height: 34px; border-radius: 9px; display: grid; place-items: center; background: var(--sunken); flex: none; }
.apps-tile.on .apps-glyph { background: var(--accent-fill); color: #fff; }
.apps-text { flex: 1; min-width: 0; display: grid; }
.apps-text strong { font-weight: 600; }
.apps-text span { font-size: 11.5px; color: var(--ink-3); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.apps-check { width: 20px; height: 20px; border-radius: 6px; border: 1.5px solid var(--line); display: grid; place-items: center; flex: none; }
.apps-tile.on .apps-check { border-color: var(--accent); background: var(--accent-fill); color: #fff; }
.apps-count { margin: 12px 0 0; font-size: 12px; color: var(--ink-3); }
`;
