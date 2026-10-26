// Add or remove apps (docs/APP-PACKS.md 5.4), and what the launcher says
// about the apps that are not here.
//
// Choose apps turns an installed app on or off. This puts an app on the
// computer or takes it off: added, its pack comes from the release of this
// very version and is checked before it is put in place; removed, it goes at
// the next start. Either way the change takes effect when Rutba Office
// starts again, and the person's documents, mail and settings stay.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Icon, formatBytes, t } from '@rutba/office-ui';
import { APPS } from '@rutba/office-formats/registry';
import { refusedPacks } from '@rutba/office-shell/client';
import { PACKS } from '@rutba/office-shell/contract';

const iconOf = (id) => APPS[id]?.icon || { recorder: 'record', ocr: 'scanText', ai: 'wand' }[id] || 'grid';
const nameOf = (id) => (APPS[id] ? APPS[id].short : t(PACKS.find((p) => p.id === id)?.name || id));

/**
 * The launcher's note about an app that is not here: the one a file or a
 * link asked for (the window was made for the launcher instead, `?missing=`),
 * and those on the computer that would not load, with why.
 */
export function PacksNotice({ onAdd }) {
  const missing = new URLSearchParams(location.search).get('missing');
  const [shown, setShown] = useState(true);
  const refused = refusedPacks();
  if (!shown || (!missing && !refused.length)) return null;
  return (
    <div className="home-packs-note" role="note">
      <style>{PACKS_CSS}</style>
      <Icon name="info" size={18} />
      <span className="home-packs-text">
        {missing ? <strong>{t('{app} is not installed on this computer.', { app: nameOf(missing) })}</strong> : null}
        {missing ? <span>{t('Add it with Add or remove apps, and it opens what it opens from the next start.')}</span> : null}
        {refused.map((r) => (
          <span key={r.id} data-pack={r.id}>{t(r.reason)}</span>
        ))}
      </span>
      {onAdd ? <Button primary label={t('Add or remove apps…')} className="home-packs-add" onClick={onAdd} /> : null}
      <Button ghost icon="close" title={t('Close')} onClick={() => setShown(false)} />
    </div>
  );
}

const STATE_WORDS = () => ({
  installed: t('Installed'),
  absent: t('Not installed'),
  adding: t('Added: here from the next start'),
  removing: t('Removed: goes at the next start'),
  refused: t('Installed, but would not load'),
});

/** The Add or remove apps window. */
export function PacksDialog({ shell, onClose }) {
  const [list, setList] = useState(null);
  const [busy, setBusy] = useState(null);
  const [errors, setErrors] = useState({});
  const read = useCallback(() => {
    shell.packs.list().then(setList).catch((err) => setList({ error: err?.message || String(err), apps: [] }));
  }, [shell]);
  useEffect(() => {
    read();
    return shell.on('packs:changed', () => read());
  }, [shell, read]);

  const act = async (id, verb) => {
    setBusy(id);
    setErrors((e) => ({ ...e, [id]: null }));
    try {
      setList(await shell.packs[verb]({ id }));
    } catch (err) {
      setErrors((e) => ({ ...e, [id]: err?.message || String(err) }));
    } finally {
      setBusy(null);
    }
  };
  const restart = async () => {
    setBusy('restart');
    try {
      await shell.packs.restart();
    } catch (err) {
      setErrors((e) => ({ ...e, restart: err?.message || String(err) }));
      setBusy(null);
    }
  };

  const words = STATE_WORDS();
  return (
    <Dialog
      title={t('Add or remove apps')}
      width={640}
      className="packs-dialog"
      onClose={onClose}
      actions={
        <>
          {list?.restart ? <Button primary label={list.perMachine ? t('Restart and allow the change…') : t('Restart now')} className="packs-restart" disabled={busy === 'restart'} onClick={restart} /> : null}
          <span style={{ flex: 1 }} />
          <Button label={list?.restart ? t('Later') : t('Close')} onClick={onClose} />
        </>
      }
    >
      <style>{PACKS_CSS}</style>
      {!list ? <p className="packs-lead">{t('Reading what is installed…')}</p> : null}
      {list && !list.managed && !list.error ? (
        <p className="packs-lead">{t('Every app is built into this copy of Rutba Office, so there is nothing to add or remove. The installed copy lets you choose.')}</p>
      ) : null}
      {list?.managed ? (
        <p className="packs-lead">
          {t('An app added or removed here changes at the next start. Removing an app keeps your documents, mail and settings; add it again at any time.')}
          {list.perMachine ? ` ${t('This copy is installed for everyone on this computer, so Windows asks for permission when it restarts.')}` : ''}
        </p>
      ) : null}
      {list?.error ? <p className="packs-error" role="alert">{t(list.error)}</p> : null}
      {errors.restart ? <p className="packs-error" role="alert">{t(errors.restart)}</p> : null}
      {list?.managed ? (
        <div className="packs-list" role="list">
          {list.apps.map((a) => (
            <div key={a.id} className={`packs-row ${a.state}`} role="listitem" data-pack={a.id} data-state={a.state}>
              <span className="packs-glyph"><Icon name={iconOf(a.id)} size={18} /></span>
              <span className="packs-text">
                <strong>{nameOf(a.id)}</strong>
                <span>
                  {words[a.state]}
                  {a.size ? ` · ${formatBytes(a.size)}` : ''}
                </span>
                {a.reason ? <span className="packs-reason">{t(a.reason)}</span> : null}
                {errors[a.id] ? <span className="packs-error" role="alert">{t(errors[a.id])}</span> : null}
              </span>
              {a.state === 'installed' ? <Button label={t('Remove')} className="packs-remove" disabled={Boolean(busy)} onClick={() => act(a.id, 'remove')} /> : null}
              {a.state === 'removing' ? <Button label={t('Keep')} className="packs-keep" disabled={Boolean(busy)} onClick={() => act(a.id, 'keep')} /> : null}
              {a.state === 'adding' ? <Button label={t('Do not add')} className="packs-undo" disabled={Boolean(busy)} onClick={() => act(a.id, 'remove')} /> : null}
              {a.state === 'absent' || a.state === 'refused' ? (
                <Button primary={a.state === 'absent'} label={busy === a.id ? t('Adding…') : a.state === 'refused' ? t('Install again') : t('Add')} className="packs-add" disabled={Boolean(busy)} onClick={() => act(a.id, 'add')} />
              ) : null}
            </div>
          ))}
        </div>
      ) : null}
    </Dialog>
  );
}

export const PACKS_CSS = `
.home-packs-note { display: flex; align-items: center; gap: 12px; padding: 12px 16px; margin: 0 0 18px; border-radius: var(--r-3); border: 1px solid var(--line); background: var(--wash); color: var(--ink); }
.home-packs-note > svg { color: var(--accent-text); flex: none; }
.home-packs-text { flex: 1; min-width: 0; display: grid; gap: 2px; }
.home-packs-text span { font-size: 12.5px; color: var(--ink-2); }
.packs-lead { margin: 0 0 14px; color: var(--ink-2); font-size: 13px; line-height: 1.5; }
.packs-list { display: grid; gap: 6px; }
.packs-row { display: flex; align-items: center; gap: 10px; padding: 8px 12px; border-radius: var(--r-2); border: 1px solid var(--line-soft); background: var(--surface); }
.packs-row.absent { color: var(--ink-3); }
.packs-glyph { width: 32px; height: 32px; border-radius: 9px; display: grid; place-items: center; background: var(--sunken); flex: none; }
.packs-row.installed .packs-glyph, .packs-row.adding .packs-glyph { background: var(--accent-fill); color: #fff; }
.packs-text { flex: 1; min-width: 0; display: grid; }
.packs-text strong { font-weight: 600; }
.packs-text span { font-size: 11.5px; color: var(--ink-3); }
.packs-reason { white-space: normal; }
.packs-error { color: var(--bad-text, var(--ink)); font-size: 12px; white-space: normal; }
`;
