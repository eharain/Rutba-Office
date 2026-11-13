// Voices and languages (C5): what the packages feed offers this version of
// Rutba Office, by kind, and what is on this computer, with Add, Pause,
// Cancel and Remove. It cannot live in a pack: it is how the first pack
// arrives. The main process does the fetching and the checking
// (apps/desktop/main/addons-service.js); this only shows it and asks.
//
// It also says what the update service is waiting for (C11), why an add-on
// is not used, and when a start went without add-ons (C13), each in this
// window, never a dialog of the system's.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Icon, formatBytes, msg, t, tn, language } from '@rutba/office-ui';

/**
 * Removing Rutba Office asks whether what was added here goes too (C9), and
 * says how much room it takes: marked here for the catalogue, and taken from
 * it into the uninstaller by tools/make-installer-words.mjs.
 */
export const UNINSTALLER_WORDS = [
  msg('Remove the voices and languages added to Rutba Office as well?'),
  msg('They take {size} on this computer. Keep them if you will install Rutba Office again.'),
];

/**
 * The installer's Languages page (L5), in the words of Home's offer below:
 * marked here for the catalogue, and taken into the installer by
 * tools/make-installer-words.mjs.
 */
export const INSTALLER_LANGUAGE_WORDS = [
  msg('Rutba Office in the languages this computer uses.'),
  msg('Tick the languages to add. They are added the first time Rutba Office starts while this computer is online. Any language can be added or removed later, in Voices and languages.'),
  msg('{name} (draft)'),
];

/** A language's name with its draft mark when its pack is a draft: by the pack's status alone, never its tag. */
export const draftMarked = (name, status) => (status === 'draft' ? t('{name} (draft)', { name }) : name);

/** What Restart can say when it does not restart (addons-service.js restartNow), for the catalogue. */
export const RESTART_WORDS = [
  msg('Windows did not give permission to change the add-ons of a copy installed for everyone. Nothing was changed; try again, or ask whoever looks after this computer.'),
];

/** What Add from a file can say (addons-service.js addFromFile), for the catalogue. */
export const FROM_FILE_WORDS = [
  msg('This file is not an add-on: open a .zip published with the Rutba Office add-ons.'),
  msg('This file is not an add-on: it should hold one add-on and its folder.'),
];

/** A kind of add-on, as a person reads it. */
export const kindWords = () => ({
  language: t('language'),
  dictionary: t('dictionary'),
  font: t('font'),
  voice: t('voice'),
  model: t('speech model'),
  pronunciation: t('pronunciation'),
  engine: t('speech engine'),
});

/** "French (language), speech engine": what an update waits for. */
export function waitingList(items = []) {
  const k = kindWords();
  return items.map((w) => (w.kind === 'engine' ? k.engine : `${w.name || w.id} (${k[w.kind] || w.kind || t('add-on')})`)).join(', ');
}

/**
 * What stops until each add-on has its version for the update, said once
 * before "Update anyway" (private design 5.2, step 4).
 */
export function whatStops(items = [], version) {
  const lines = [];
  if (items.some((w) => ['engine', 'voice', 'model', 'pronunciation'].includes(w.kind))) {
    lines.push(t('Speech will use the Windows voice until the speech engine for {version} is ready.', { version }));
  }
  for (const w of items.filter((x) => x.kind === 'language')) {
    lines.push(t('{name} will show new labels in English until its {version} version is ready.', { name: w.name || w.id, version }));
  }
  for (const w of items.filter((x) => !['engine', 'voice', 'model', 'pronunciation', 'language'].includes(x.kind))) {
    lines.push(t('{name} is not used until its {version} version is ready.', { name: w.name || w.id, version }));
  }
  return lines;
}

/** The sentence About and this page give for an update that waits, or null. */
export function waitingSentence(update) {
  if (update?.state !== 'waiting') return null;
  if (update.addonsChecked === false) {
    return t('Rutba Office {version} is out, but your voices and languages could not be checked for it. It is asked again at the next check.', { version: update.available });
  }
  return t('Rutba Office {version} is out. It will be offered once these are ready for it: {list}.', { version: update.available, list: waitingList(update.waitingFor) });
}

/** What an update no longer has, named before it is installed. */
export function retiredSentences(update) {
  const k = kindWords();
  return (update?.retired || []).map((r) => t('Rutba Office {version} no longer has {name} ({kind}). It stays on this computer, unused, until you remove it.', { version: update.available, name: r.name || r.id, kind: k[r.kind] || r.kind || t('add-on') }));
}

const STATE_WORDS = () => ({
  loaded: t('Added'),
  added: t('Added: used from the next start'),
  waiting: t('Added, waiting for what it needs'),
  refused: t('Added, but not used'),
  removing: t('Removed: goes at the next start'),
  off: t('Turned off: withdrawn by its makers'),
  deferred: t('Added: not ready in time at this start'),
});
const STATUS_WORDS = () => ({
  add: t('Not added'),
  installed: t('Added'),
  update: t('A newer revision is ready'),
  adding: t('Adding…'),
  paused: t('Paused'),
  staged: t('Added: moved in at Restart'),
  failed: t('Could not be added'),
  unavailable: t('Not ready for this version yet'),
  offline: t('Can be added when this computer is online'),
});

/** A language's name as the window's language writes it, by Intl, or null. */
export function nameInWindow(tag, windowLanguage = language()) {
  try {
    return new Intl.DisplayNames([windowLanguage, 'en'], { type: 'language' }).of(tag) || null;
  } catch {
    return null;
  }
}

/** A language row's title (L6): its name in its own words, and in the window's when that differs. */
export function languageTitle(p, windowLanguage = language()) {
  const own = p.nativeName || p.name || p.id;
  const here = (p.tag && nameInWindow(p.tag, windowLanguage)) || p.name || null;
  return here && here.toLocaleLowerCase() !== String(own).toLocaleLowerCase() ? `${own} · ${here}` : own;
}

/**
 * How complete a language is for this version, as a whole percentage, or null:
 * added, counted against this Office's own words ({ found, of }); offered, its
 * labels made from that version's English over English's own (feed coverage).
 */
export function completeness(p, englishLabels = null) {
  const c = p?.coverage;
  if (!c) return null;
  if (Number.isFinite(c.found) && c.of > 0) return Math.min(100, Math.floor((100 * c.found) / c.of));
  if (Number.isFinite(c.labels) && englishLabels > 0) return Math.min(100, Math.floor((100 * c.labels) / englishLabels));
  return null;
}

/** A language's details line: how complete, draft or reviewed. */
export function languageDetails(p, englishLabels = null) {
  const parts = [];
  const pct = completeness(p, englishLabels);
  if (pct !== null) parts.push(t('{pct}% translated for this version', { pct }));
  const status = p.packStatus || p.status;
  if (status === 'reviewed') parts.push(t('Reviewed'));
  else if (status === 'draft') parts.push(t('Draft, not yet reviewed'));
  return parts.join(' · ');
}

/** Where an offered add-on is shown: languages with their dictionaries, voices, then the speech parts. */
function sections(offered, known = [], { online = true } = {}) {
  // English is part of Office and never listed. A language the list knows with no pack offered is shown, to say it is coming.
  const langs = offered.filter((p) => p.kind === 'language' && p.id !== 'lang-en');
  const have = new Set(langs.map((p) => p.id));
  const coming = known.filter((l) => l.pack && l.pack !== 'lang-en' && !have.has(l.pack)).map((l) => ({ ...l, id: l.pack, kind: 'language', status: online ? 'unavailable' : 'offline', coming: true }));
  const dictsOf = (lang) => offered.filter((p) => p.kind === 'dictionary' && (p.requires || []).some((r) => r === lang.id || (lang.provides || []).includes(r)));
  const placed = new Set([...langs.map((p) => p.id), ...langs.flatMap((l) => dictsOf(l).map((d) => d.id))]);
  return [
    { id: 'languages', title: t('Languages'), rows: [...langs.flatMap((l) => [{ ...l, depth: 0 }, ...dictsOf(l).map((d) => ({ ...d, depth: 1 }))]), ...coming.map((l) => ({ ...l, depth: 0 }))] },
    { id: 'voices', title: t('Voices'), rows: offered.filter((p) => p.kind === 'voice').map((p) => ({ ...p, depth: 0 })) },
    { id: 'other', title: t('Fonts and dictionaries'), rows: offered.filter((p) => ['font', 'dictionary'].includes(p.kind) && !placed.has(p.id)).map((p) => ({ ...p, depth: 0 })) },
    { id: 'speech', title: t('Speech engine and models'), lead: t('These come by themselves with the first voice you add.'), rows: offered.filter((p) => ['engine', 'model', 'pronunciation'].includes(p.kind)).map((p) => ({ ...p, depth: 0 })) },
  ].filter((s) => s.rows.length);
}

/** The Voices and languages window. */
export function VoicesDialog({ shell, onClose }) {
  const [list, setList] = useState(null);
  const [update, setUpdate] = useState(null);
  const [busy, setBusy] = useState(null);
  const [errors, setErrors] = useState({});
  const [confirmAnyway, setConfirmAnyway] = useState(false);
  const [progress, setProgress] = useState({});

  const read = useCallback(() => {
    shell.addons.list().then(setList).catch((err) => setList({ error: err?.message || String(err), installed: [], offered: [] }));
  }, [shell]);
  useEffect(() => {
    // Opening the page reads this version's list of add-ons: the one request it makes, said in the README.
    shell.addons.feed({ refresh: true }).then(setList).catch(() => read());
    shell.update.state().then(setUpdate).catch(() => {});
    const offs = [
      shell.on('addons:changed', () => read()),
      shell.on('addons:progress', (p) => {
        setProgress((all) => ({ ...all, [p.id]: p }));
        if (p.state !== 'adding') read();
      }),
      shell.on('update:state', setUpdate),
    ];
    return () => offs.forEach((off) => off?.());
  }, [shell, read]);

  const act = async (id, verb) => {
    setBusy(id);
    setErrors((e) => ({ ...e, [id]: null }));
    try {
      setList(await shell.addons[verb]({ id }));
    } catch (err) {
      setErrors((e) => ({ ...e, [id]: err?.message || String(err) }));
    } finally {
      setBusy(null);
    }
  };
  // A voice or language from its file (L5): for a computer with no network.
  const fromFile = async () => {
    const [file] = (await shell.dialog.open({ title: t('Add from a file'), filters: [{ name: t('Add-on files'), extensions: ['zip'] }] }).catch(() => null)) || [];
    if (!file) return;
    setBusy('file');
    setErrors((e) => ({ ...e, file: null }));
    try {
      setList(await shell.addons.addFromFile({ file }));
    } catch (err) {
      setErrors((e) => ({ ...e, file: err?.message || String(err) }));
    } finally {
      setBusy(null);
    }
  };
  const restart = async (verb = 'restart') => {
    setBusy('restart');
    try {
      const r = await shell.addons[verb]();
      // Not restarted: unsaved work, or Windows' permission not given. Nothing was changed.
      if (r && r.restarting === false) {
        setBusy(null);
        if (r.unsaved) {
          const names = (r.unsaved || []).filter(Boolean);
          await shell.dialog.message({
            type: 'info',
            message: t('Save your work before restarting'),
            detail: `${names.length ? `${names.join('\n')}\n\n` : ''}${t('Restarting closes every window, and these have changes not saved yet. Save or close them, then restart again. Nothing has been closed.')}`,
            buttons: [t('OK')],
          });
        } else if (r.reason) setErrors((e) => ({ ...e, restart: r.reason }));
      }
    } catch (err) {
      setErrors((e) => ({ ...e, restart: err?.message || String(err) }));
      setBusy(null);
    }
  };

  const words = STATE_WORDS();
  const statusWords = STATUS_WORDS();
  const kinds = kindWords();
  const installed = new Map((list?.installed || []).map((p) => [p.id, p]));
  const waiting = waitingSentence(update);
  // English's own count of labels for this version: what an offered language's coverage is measured against.
  const englishLabels = (list?.offered || []).find((p) => p.id === 'lang-en')?.coverage?.labels || null;
  const gone = retiredSentences(update);

  const row = (p) => {
    const mine = installed.get(p.id);
    const job = progress[p.id]?.state === 'adding' ? progress[p.id] : (list?.jobs || []).find((j) => j.id === p.id);
    const status = job?.state || p.status;
    const pct = job?.total ? Math.round((100 * (job.done || 0)) / job.total) : null;
    return (
      <div key={p.id} className={`voices-row depth-${p.depth} ${status}`} role="listitem" data-addon={p.id} data-status={status}>
        <span className="voices-text">
          <strong>{p.kind === 'language' ? languageTitle(mine?.tag ? { ...p, ...mine, name: p.name || mine.name } : p) : p.nativeName && p.nativeName !== p.name ? `${p.name} · ${p.nativeName}` : p.name || p.id}</strong>
          {p.kind === 'language' && !p.coming ? <span>{languageDetails(mine?.coverage ? { ...p, coverage: mine.coverage } : p, englishLabels)}</span> : null}
          {p.kind === 'language' && (mine?.review || p.review) ? <span className="voices-reason">{mine?.review || p.review}</span> : null}
          <span>
            {mine ? words[mine.state] || statusWords[status] : statusWords[status]}
            {p.size ? ` · ${formatBytes(p.size)}` : ''}
            {p.licence ? ` · ${p.licence}` : ''}
            {status === 'adding' && pct !== null ? ` · ${pct}%` : ''}
          </span>
          {mine?.reason ? <span className="voices-reason">{t(mine.reason)}</span> : null}
          {p.withdrawn ? <span className="voices-reason">{t('Withdrawn by its makers.')}</span> : null}
          {job?.error ? <span className="voices-error" role="alert">{t(job.error)}</span> : null}
          {errors[p.id] ? <span className="voices-error" role="alert">{t(errors[p.id])}</span> : null}
        </span>
        {status === 'add' || status === 'failed' || status === 'update' ? (
          <Button primary={status === 'add'} label={status === 'update' ? t('Update') : status === 'failed' ? t('Try again') : t('Add')} disabled={Boolean(busy) || p.withdrawn} onClick={() => act(p.id, 'add')} />
        ) : null}
        {status === 'adding' ? <Button label={t('Pause')} onClick={() => act(p.id, 'pause')} /> : null}
        {status === 'paused' ? <Button primary label={t('Resume')} onClick={() => act(p.id, 'add')} /> : null}
        {status === 'adding' || status === 'paused' || status === 'staged' ? <Button ghost label={t('Cancel')} onClick={() => act(p.id, 'cancel')} /> : null}
        {mine && mine.state !== 'removing' && p.id !== 'lang-en' && !['adding', 'paused'].includes(status) ? (
          <Button label={t('Remove')} disabled={Boolean(busy) || mine.inUse} title={mine.inUse ? t('A window is open in this language. Close it, or open the windows in another language, before you remove it.') : undefined} onClick={() => act(p.id, 'remove')} />
        ) : null}
        {mine?.state === 'removing' ? <Button label={t('Keep')} disabled={Boolean(busy)} onClick={() => act(p.id, 'cancel')} /> : null}
      </div>
    );
  };

  // What is here and no longer offered (retired, or offline with no list): still listed, with Remove.
  const offeredIds = new Set((list?.offered || []).map((p) => p.id));
  const unlisted = (list?.installed || []).filter((p) => !offeredIds.has(p.id) && p.id !== 'lang-en');

  return (
    <Dialog
      title={t('Voices and languages')}
      width={680}
      className="voices-dialog"
      onClose={onClose}
      actions={
        <>
          {list?.skipped === 'failed-starts' ? <Button primary label={t('Try again')} disabled={busy === 'restart'} onClick={() => restart('tryAgain')} /> : null}
          {list?.restart ? <Button primary label={list.perMachine ? t('Restart and allow the change…') : t('Restart now')} disabled={busy === 'restart'} onClick={() => restart()} /> : null}
          {list?.dir ? <Button label={t('Add from a file…')} disabled={Boolean(busy)} onClick={fromFile} /> : null}
          <span style={{ flex: 1 }} />
          <Button label={list?.restart ? t('Later') : t('Close')} onClick={onClose} />
        </>
      }
    >
      <style>{VOICES_CSS}</style>
      <p className="voices-lead">
        {t('Speech is made on this computer. What it reads never leaves it.')}
        {' '}
        {t('An added voice or language is used from the next start. Removing one keeps your documents and settings.')}
        {list?.perMachine ? ` ${t('This copy is installed for everyone on this computer, so Windows asks for permission when it restarts.')}` : ''}
      </p>
      {list?.said ? <p className="voices-note" role="note"><Icon name="info" size={15} /> {t(list.said)}</p> : null}
      {waiting ? (
        <div className="voices-note" role="note">
          <Icon name="download" size={15} />
          <span>{waiting}</span>
          {update?.state === 'waiting' ? <Button label={t('Update anyway…')} onClick={() => setConfirmAnyway(true)} /> : null}
        </div>
      ) : null}
      {gone.map((line) => <p key={line} className="voices-note" role="note">{line}</p>)}
      {list?.engine?.problem ? <p className="voices-note" role="note"><Icon name="info" size={15} /> {t(list.engine.problem)} {t('Read Aloud uses the Windows voice.')}</p> : null}
      {list?.error ? <p className="voices-error" role="alert">{t(list.error)}</p> : null}
      {errors.restart ? <p className="voices-error" role="alert">{t(errors.restart)}</p> : null}
      {errors.file ? <p className="voices-error" role="alert">{t(errors.file)}</p> : null}
      {list && !list.dir ? <p className="voices-lead">{t('This copy of Rutba Office has no add-ons folder: voices and languages are added to an installed copy.')}</p> : null}
      {list && list.dir && !list.feed ? <p className="voices-lead">{t('The list of voices and languages could not be read. What is added below still works; try again when this computer is online.')}</p> : null}
      {list?.feed?.from === 'bundled' ? <p className="voices-lead">{t('Offline: this is the list that came with this version.')}</p> : null}
      {list?.dir && list.feed?.from !== 'network' ? <p className="voices-lead">{t('With no network, a voice or language can be added from its file: a .zip published with the Rutba Office add-ons, brought from another computer. Choose Add from a file.')}</p> : null}

      {sections(list?.offered || [], list?.languages || [], { online: Boolean(list?.feed) }).map((s) => (
        <section key={s.id} className="voices-section" aria-label={s.title}>
          <h3>{s.title}</h3>
          {s.lead ? <p className="voices-lead small">{s.lead}</p> : null}
          <div className="voices-list" role="list">{s.rows.map(row)}</div>
        </section>
      ))}
      {unlisted.length ? (
        <section className="voices-section" aria-label={t('Added here')}>
          <h3>{t('Added here, not on the list for this version')}</h3>
          <div className="voices-list" role="list">
            {unlisted.map((p) => row({ ...p, depth: 0, status: 'installed', name: `${p.name || p.id} (${kinds[p.kind] || p.kind})` }))}
          </div>
        </section>
      ) : null}

      {confirmAnyway ? (
        <Dialog
          title={t('Update to {version} anyway?', { version: update?.available })}
          width={460}
          onClose={() => setConfirmAnyway(false)}
          actions={<>
            <span style={{ flex: 1 }} />
            <Button label={t('Cancel')} onClick={() => setConfirmAnyway(false)} />
            <Button primary label={t('Update anyway')} onClick={() => { setConfirmAnyway(false); shell.update.anyway({ version: update?.available }).then(setUpdate).catch(() => {}); }} />
          </>}
        >
          {whatStops(update?.waitingFor, update?.available).map((line) => <p key={line} className="voices-lead">{line}</p>)}
          <p className="voices-lead">{t('Nothing is removed. Each comes back by itself once its version for {version} is ready.', { version: update?.available })}</p>
        </Dialog>
      ) : null}
    </Dialog>
  );
}

/**
 * Home's offer of the computer's own languages (L5), once, at a first start
 * with no installer's choice; and, while the computer is offline, what waits
 * to be added. The main process decides (first-languages.js); this asks.
 */
export function LanguagesNotice({ shell, onOpen }) {
  const [state, setState] = useState(null);
  const [busy, setBusy] = useState(false);
  const read = useCallback(() => {
    shell.addons.offer?.().then(setState).catch(() => setState(null));
  }, [shell]);
  useEffect(() => {
    read();
    const off = shell.on('addons:changed', () => read());
    return () => off?.();
  }, [shell, read]);
  if (!state) return null;
  const drafts = new Set(state.drafts || []);
  const names = (ids) => ids.map((id) => draftMarked(nameInWindow(id.replace(/^lang-/, '')) || id, drafts.has(id) ? 'draft' : null)).join(', ');
  const call = async (verb, args) => {
    setBusy(true);
    try {
      setState(await shell.addons[verb](args));
    } catch {
      read();
    } finally {
      setBusy(false);
    }
  };
  if (state.moved) {
    const name = nameInWindow(state.moved.tag) || state.moved.tag;
    return (
      <div className="home-choose home-languages" role="note">
        <Icon name="globe" size={20} />
        <span className="home-choose-text">
          <span>{state.moved.added
            ? t('{language} now comes as a language pack, and was added with this update.', { language: name })
            : t('{language} is no longer built into Rutba Office. Add its language pack in Voices and languages.', { language: name })}</span>
        </span>
        {state.moved.added ? null : <Button primary label={t('Voices and languages…')} onClick={onOpen} />}
        <Button ghost icon="close" title={t('Close')} disabled={busy} onClick={() => call('dismissOffer', { moved: true })} />
      </div>
    );
  }
  if (state.offer?.length) {
    return (
      <div className="home-choose home-languages" role="note">
        <Icon name="globe" size={20} />
        <span className="home-choose-text">
          <strong>{t('Rutba Office in your languages')}</strong>
          <span>{tn(state.offer.length, 'This computer uses {languages}. Add it, and choose it from the language menu after Rutba Office restarts.', 'This computer uses {languages}. Add them, and choose one from the language menu after Rutba Office restarts.', { languages: names(state.offer) })}</span>
          {!state.online ? <span>{tn(state.offer.length, 'This computer is offline: the language is added once it is online. It can also be added from a file, in Voices and languages.', 'This computer is offline: the languages are added once it is online. One can also be added from a file, in Voices and languages.')}</span> : null}
        </span>
        <Button label={t('Not now')} disabled={busy} onClick={() => call('dismissOffer')} />
        <Button primary label={state.online ? t('Add') : t('Add when online')} disabled={busy} onClick={() => call('takeOffer')} />
      </div>
    );
  }
  if (state.waiting?.length && !state.online) {
    return (
      <div className="home-choose home-languages" role="note">
        <Icon name="globe" size={20} />
        <span className="home-choose-text">
          <span>{tn(state.waiting.length, '{languages} will be added when this computer is online. A language can also be added from a file, in Voices and languages.', '{languages} will be added when this computer is online. A language can also be added from a file, in Voices and languages.', { languages: names(state.waiting) })}</span>
        </span>
        <Button label={t('Voices and languages…')} onClick={onOpen} />
      </div>
    );
  }
  return null;
}

export const VOICES_CSS = `
.voices-lead { margin: 0 0 12px; color: var(--ink-2); font-size: 13px; line-height: 1.5; }
.voices-lead.small { font-size: 12px; margin: -4px 0 8px; }
.voices-note { display: flex; align-items: center; gap: 8px; margin: 0 0 12px; padding: 8px 12px; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--wash); font-size: 12.5px; }
.voices-note > span { flex: 1; }
.voices-section { margin: 0 0 14px; }
.voices-section h3 { margin: 0 0 6px; font-size: 12px; font-weight: 600; color: var(--ink-3); text-transform: none; }
.voices-list { display: grid; gap: 6px; }
.voices-row { display: flex; align-items: center; gap: 8px; padding: 8px 12px; border-radius: var(--r-2); border: 1px solid var(--line-soft); background: var(--surface); }
.voices-row.depth-1 { margin-inline-start: 24px; }
.voices-text { flex: 1; min-width: 0; display: grid; }
.voices-text strong { font-weight: 600; }
.voices-text span { font-size: 11.5px; color: var(--ink-3); }
.voices-reason { white-space: normal; }
.voices-error { color: var(--bad-text, var(--ink)); font-size: 12px; white-space: normal; }
`;
