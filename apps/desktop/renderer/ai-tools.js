// AI Tools: the person's switch for letting AI apps on this computer make and
// read documents with Rutba Office (owner, 2026-10-16, D-2026-10-16-06: "it
// should be off unless some one opens it for single session or want to keep it
// always open"; 2026-10-17, D-2026-10-17-04: "protect it from remote
// connections"). Off, for this session, or always; who may connect, this
// computer only or other machines too with a key (2026-10-21, D-2026-10-21-06:
// "the user should have choice to open remote connections or local access
// only"); what the tools may reach; how to connect an AI app; and what the
// tools have done.
//
// What they may reach (D-2026-10-26-04, -06): each mail account, archive,
// calendar, address book, database and folder has its own step, Off, Read only,
// Read and prepare, or Read, prepare and act. The one setting is shown here
// and in its own app's settings (AiStepPicker, AiThingsList), and a change in
// either is heard by both ('ai:changed'). What the tools ask to send or write
// elsewhere waits for the person's press in every window (AiConfirm).
//
// The badge, the confirmations and the step pickers other apps show are the
// core's; the AI Tools window itself is AI Tools' own (ai-tools-dialog.js),
// loaded when it is opened.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Icon, Select, useToast, formatWhen, t, tn } from '@rutba/office-ui';

export const MODES = () => [
  ['off', t('Off'), t('No AI app can use them.')],
  ['session', t('On for this session'), t('Until Rutba Office closes.')],
  ['always', t('Always on'), t('Until you turn them off here.')],
];

export const REACH = () => [
  ['local', t('This computer only'), t('Nothing listens on the network.')],
  ['remote', t('Other machines too'), t('Over the network, with a key.')],
];

/** The four steps, in the words every app's settings use. */
export const STEP_NAMES = () => ({ off: t('Off'), read: t('Read only'), prepare: t('Read and prepare'), act: t('Read, prepare and act') });
export const STEP_NOTES = () => ({
  off: t('The AI tools cannot see it.'),
  read: t('The AI tools may read it, and change nothing.'),
  prepare: t('The AI tools may read it and prepare new things kept on this computer: drafts, events, contacts, files and records. Nothing is sent.'),
  act: t('The AI tools may also ask to send, or to write to a server. Each one is shown to you in Rutba Office first and done only when you press its button.'),
});

/** The groups the AI Tools window lists, in order. */
export const GROUPS = () => [
  [['mail', 'archive'], t('Mail accounts and archives'), 'mail'],
  [['calendar'], t('Calendars'), 'calendar'],
  [['book'], t('Address books'), 'contacts'],
  [['database'], t('Databases'), 'grid'],
  [['folder'], t('Folders'), 'open'],
];

/** A thing's name as its row shows it. */
export const thingName = (x) => (x.kind === 'book' && !x.name ? t('On this computer') : x.name);

/** The step chosen for one thing, from its steps; a change goes to the one setting. */
export function StepSelect({ thing, onSet, label = t('AI tools') }) {
  const names = STEP_NAMES();
  return (
    <Select
      className="ai-step-select"
      aria-label={t('{what}: {name}', { what: label, name: thingName(thing) })}
      data-ai-kind={thing.kind}
      data-ai-id={thing.id}
      value={thing.step}
      title={STEP_NOTES()[thing.step]}
      onChange={(e) => onSet(e.target.value)}
    >
      {thing.steps.map((k) => <option key={k} value={k}>{names[k]}</option>)}
    </Select>
  );
}

/**
 * One thing's AI tools setting in its own app's settings (a mail account in
 * Mail's Accounts, a database in Data Files): the same setting the AI Tools
 * window lists. Not shown when the AI tools are not among the person's apps.
 */
export function AiStepPicker({ shell, kind, id, name = '', className = '' }) {
  const toast = useToast();
  const [s, setS] = useState(null);
  const read = useCallback(() => {
    if (!shell?.ai?.stepOf || !id) return;
    shell.ai.stepOf({ kind, id }).then(setS).catch(() => setS(null));
  }, [shell, kind, id]);
  useEffect(() => { read(); }, [read]);
  useEffect(() => (shell?.on ? shell.on('ai:changed', (c) => { if (c?.kind === kind) read(); }) : undefined), [shell, kind, read]);
  if (!s?.addon) return null;
  const set = async (step) => {
    try {
      setS(await shell.ai.setStep({ kind, id, step }));
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    }
  };
  return (
    <label className={`ai-step ${className}`} data-ai-step={s.step}>
      <span className="ai-step-label">{t('AI tools')}</span>
      <StepSelect thing={{ kind, id, name, step: s.step, steps: s.steps }} onSet={set} />
    </label>
  );
}

/**
 * Every thing of a kind with its AI tools setting, for an app's own settings
 * (Calendar's and Contacts' Accounts): the same settings the AI Tools window lists.
 */
export function AiThingsList({ shell, kind, title = t('AI tools') }) {
  const toast = useToast();
  const [list, setList] = useState(null);
  const read = useCallback(() => {
    if (!shell?.ai?.things) return;
    shell.ai.things().then((r) => setList(r?.addon ? (r.things || []).filter((x) => x.kind === kind) : null)).catch(() => setList(null));
  }, [shell, kind]);
  useEffect(() => { read(); }, [read]);
  useEffect(() => (shell?.on ? shell.on('ai:changed', read) : undefined), [shell, read]);
  useEffect(() => (shell?.on ? shell.on('dav:changed', read) : undefined), [shell, read]);
  if (!list?.length) return null;
  const set = async (x, step) => {
    try {
      await shell.ai.setStep({ kind: x.kind, id: x.id, step });
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    }
    read();
  };
  return (
    <div className="ai-things-own" data-ai-kind={kind}>
      <div className="ai-things-own-h">{title}</div>
      <p className="ai-note">{t('What the AI tools may do with each. Everything is Off until you choose. The same settings are in AI Tools.')}</p>
      <ul className="ai-things">
        {list.map((x) => (
          <li key={x.id} className="ai-thing" data-ai-id={x.id}>
            <span className="ai-thing-name">{thingName(x)}{x.server ? <small>{` ${t('kept in step with a server')}`}</small> : null}</span>
            <StepSelect thing={x} onSet={(step) => set(x, step)} />
          </li>
        ))}
      </ul>
    </div>
  );
}

/** A request's title and its button, by what it would do. */
const ASKS = () => ({
  mail: [t('An AI app wants to send this message'), t('Send')],
  invitation: [t('An AI app wants to send this invitation'), t('Send Invitation')],
  answer: [t('An AI app wants to answer this invitation'), t('Send Answer')],
  event: [t('An AI app wants to add this event to a calendar kept on a server'), t('Add Event')],
  contact: [t('An AI app wants to add this contact to an address book kept on a server'), t('Add Contact')],
  database: [t('An AI app wants to write this to a database engine'), t('Write It')],
});

/** The words for each line a request shows. */
const FIELD_NAMES = () => ({
  from: t('From'), to: t('To'), cc: t('Cc'), bcc: t('Bcc'), subject: t('Subject'), attachments: t('Attachments'),
  calendar: t('Calendar'), summary: t('Event'), start: t('Starts'), end: t('Ends'), location: t('Location'), answer: t('Answer'),
  book: t('Address book'), name: t('Name'), emails: t('Email'), phones: t('Phone'),
  database: t('Database'), table: t('Table'), record: t('Record'), values: t('Values'),
});
const ANSWERS = () => ({ accept: t('Accept'), tentative: t('Tentative'), decline: t('Decline') });

/** A line's value as the person reads it: a time in their own clock. */
const fieldValue = (key, value) => {
  if ((key === 'start' || key === 'end') && value) return formatWhen(value);
  if (key === 'answer') return ANSWERS()[value] || value;
  return String(value);
};

/**
 * What the AI tools asked to do that leaves this computer or tells another
 * person (D-2026-10-26-05), shown in every window until it is answered in
 * one: the action's own button, Keep as draft where there is a draft to
 * keep, and Cancel. Nothing is sent until the press; Cancel does nothing.
 */
export function AiConfirm({ shell }) {
  const toast = useToast();
  const [pending, setPending] = useState([]);
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    if (!shell?.ai?.pending || !shell?.on) return undefined;
    shell.ai.pending().then((p) => setPending(p || [])).catch(() => {});
    return shell.on('ai:pending', (p) => setPending(Array.isArray(p) ? p : []));
  }, [shell]);
  const [r] = pending;
  if (!r) return null;
  const [title, action] = ASKS()[r.kind] || [t('An AI app wants to do this'), t('Do It')];
  const names = FIELD_NAMES();
  const press = async (choice) => {
    setBusy(true);
    try {
      const done = await shell.ai.answer({ id: r.id, choice });
      if (done?.error) toast(done.error, { tone: 'bad', ms: 7000 });
      else if (choice === 'act') toast(t('Done.'), { tone: 'good' });
      else if (choice === 'draft') toast(t('Kept as a draft. Nothing was sent.'), { tone: 'good', ms: 5000 });
      else toast(t('Cancelled. Nothing was done.'), { ms: 4000 });
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    } finally {
      setBusy(false);
    }
  };
  return (
    <Dialog
      title={title}
      width={560}
      className="ai-confirm"
      onClose={() => press('cancel')}
      actions={(
        <>
          {pending.length > 1 ? <span className="ai-confirm-more">{tn(pending.length - 1, '{count} more waiting', '{count} more waiting')}</span> : null}
          <span style={{ flex: 1 }} />
          <Button label={t('Cancel')} className="ai-confirm-cancel" disabled={busy} onClick={() => press('cancel')} />
          {r.draft ? <Button label={t('Keep as draft')} className="ai-confirm-draft" disabled={busy} onClick={() => press('draft')} /> : null}
          <Button primary label={action} className="ai-confirm-act" disabled={busy} onClick={() => press('act')} />
        </>
      )}
    >
      <p className="ai-note" data-ai-kind={r.kind}>
        {r.draft
          ? t('Nothing is sent until you press {action}. Keep as draft keeps it on this computer instead; Cancel does nothing.', { action })
          : t('Nothing is done until you press {action}. Cancel does nothing.', { action })}
      </p>
      <dl className="ai-confirm-fields">
        {Object.entries(r.fields || {}).filter(([, v]) => v !== null && v !== undefined && v !== '').map(([k, v]) => (
          <React.Fragment key={k}>
            <dt>{names[k] || k}</dt>
            <dd className={`selectable ai-confirm-${k}`}>{fieldValue(k, v)}</dd>
          </React.Fragment>
        ))}
      </dl>
      {r.body ? <pre className="ai-confirm-body selectable">{r.body}</pre> : null}
    </Dialog>
  );
}

export const AI_TOOLS_CSS = `
.ai-step { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-3); }
.ai-step-select { font-size: 12px; }
.ai-step-key { display: grid; grid-template-columns: max-content 1fr; gap: 2px 10px; margin: 0 0 10px; font-size: 12px; }
.ai-step-key dt { font-weight: 600; }
.ai-step-key dd { margin: 0; color: var(--ink-3); }
.ai-group { margin: 0 0 8px; }
.ai-group-h { display: flex; align-items: center; gap: 6px; margin: 8px 0 4px; font-size: 12.5px; font-weight: 600; }
.ai-things { list-style: none; margin: 0 0 6px; padding: 0; display: grid; gap: 4px; }
.ai-thing { display: flex; align-items: center; gap: 8px; }
.ai-thing-name { flex: 1; min-width: 0; overflow-wrap: anywhere; font-size: 12.5px; }
.ai-thing-name small { color: var(--ink-3); }
.ai-things-own { border-top: 1px solid var(--line); padding-top: 10px; }
.ai-things-own-h { font-weight: 600; margin-bottom: 4px; }
.ai-confirm-fields { display: grid; grid-template-columns: max-content 1fr; gap: 4px 12px; margin: 0 0 8px; font-size: 13px; }
.ai-confirm-fields dt { color: var(--ink-3); }
.ai-confirm-fields dd { margin: 0; overflow-wrap: anywhere; white-space: pre-wrap; }
.ai-confirm-body { margin: 0; padding: 8px 10px; border-radius: var(--r-2); background: var(--sunken); font: 12.5px/1.5 inherit; white-space: pre-wrap; overflow-wrap: anywhere; max-height: 240px; overflow: auto; }
.ai-confirm-more { color: var(--ink-3); font-size: 12px; }
.ai-tools-dialog .ai-lead { margin-top: 0; }
.ai-connections { list-style: none; margin: 0 0 8px; padding: 0; display: grid; gap: 4px; }
.ai-connections li { display: grid; grid-template-columns: 10px minmax(120px, 1.2fr) auto auto 1.4fr; align-items: center; gap: 8px; font-size: 12.5px; }
.ai-conn-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--good, #2fbf8a); }
.ai-conn-who { font-weight: 600; overflow-wrap: anywhere; }
.ai-conn-where, .ai-conn-when, .ai-conn-last { color: var(--ink-3); }
.ai-badge { position: relative; }
.ai-badge.live { color: var(--accent-text); }
.ai-badge-count { position: absolute; top: 1px; right: 0; min-width: 14px; height: 14px; padding: 0 3px; border-radius: 7px; background: var(--good, #2fbf8a); color: #fff; font-size: 10px; line-height: 14px; text-align: center; box-sizing: border-box; }
.ai-local { display: flex; align-items: center; gap: 6px; margin: 0 0 12px; padding: 8px 10px; border-radius: var(--r-2); background: var(--sunken); color: var(--ink-2); font-size: 12.5px; }
.ai-local.open { background: color-mix(in srgb, var(--accent) 10%, var(--sunken)); }
.ai-remote { display: grid; gap: 8px; margin: 0 0 8px; padding: 10px; border: 1px solid var(--line); border-radius: var(--r-2); }
.ai-warn { display: flex; align-items: flex-start; gap: 6px; margin: 0; color: var(--ink-2); font-size: 12.5px; line-height: 1.45; }
.ai-remote-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; font-size: 12.5px; }
.ai-remote-label { width: 64px; color: var(--ink-3); }
.ai-urls { display: grid; gap: 2px; font: 12px/1.4 ui-monospace, Consolas, monospace; overflow-wrap: anywhere; }
.ai-key { flex: 1; min-width: 120px; padding: 4px 8px; border-radius: var(--r-1, 4px); background: var(--sunken); font: 12px ui-monospace, Consolas, monospace; overflow-wrap: anywhere; }
.ai-port { width: 96px; }
.ai-listening { margin: 0; }
.ai-listening.bad { color: var(--bad, #c0392b); }
.ai-log-where { color: var(--ink-3); font-weight: 400; }
.ai-modes { border: 0; margin: 0 0 12px; padding: 0; display: grid; gap: 6px; }
.ai-mode { display: grid; grid-template-columns: auto auto 1fr; align-items: center; gap: 8px; padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--r-2); cursor: pointer; }
.ai-mode.on { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
.ai-mode-label { font-weight: 600; }
.ai-mode-note { color: var(--ink-3); font-size: 12px; }
.ai-h { font-size: 12px; text-transform: uppercase; letter-spacing: .04em; color: var(--ink-3); margin: 16px 0 6px; }
.ai-folders, .ai-log { list-style: none; margin: 0 0 8px; padding: 0; display: grid; gap: 4px; }
.ai-folders li { display: flex; align-items: center; gap: 8px; }
.ai-folder-path { flex: 1; min-width: 0; overflow-wrap: anywhere; font-size: 12.5px; }
.ai-config { margin: 0 0 6px; padding: 8px 10px; border-radius: var(--r-2); background: var(--sunken); font: 11.5px/1.45 ui-monospace, Consolas, monospace; white-space: pre-wrap; overflow-wrap: anywhere; max-height: 160px; overflow: auto; }
.ai-note, .ai-none { color: var(--ink-3); font-size: 12.5px; margin: 0 0 6px; }
.ai-log li { display: grid; grid-template-columns: 120px 150px 1fr; gap: 8px; font-size: 12px; }
.ai-log li.bad .ai-log-tool { color: var(--bad, #c0392b); }
.ai-log-what { overflow-wrap: anywhere; color: var(--ink-2); }
`;

/**
 * The badge every window's title bar shows while the AI tools are on: lit
 * when an AI app is connected, with how many; a click opens AI Tools, which
 * says who is connected, from where, since when and what it last did (owner,
 * 2026-10-17: "the ui should show indicate that and clicking should give a
 * detail who is connected and from"). Nothing is shown while they are off.
 */
export function AiBadge({ shell, onOpen }) {
  const [state, setState] = useState(null);
  useEffect(() => {
    if (!shell?.ai?.state) return undefined;
    let alive = true;
    const read = () => shell.ai.state().then((st) => { if (alive) setState(st); }).catch(() => {});
    read();
    const timer = setInterval(read, 5000);
    return () => { alive = false; clearInterval(timer); };
  }, [shell]);
  if (!state || state.mode === 'off') return null;
  const n = state.connections?.length || 0;
  const tip = n
    ? t('AI tools on: {apps} connected. Click for the details.', { apps: state.connections.map((c) => c.client).join(', ') })
    : (state.mode === 'session' ? t('AI tools on for this session; no AI app connected. Click to change.') : t('AI tools always on; no AI app connected. Click to change.'));
  return (
    <button type="button" className={`rw-btn ghost ai-badge${n ? ' live' : ''}`} data-tip={tip} aria-label={tip} data-connections={n} onClick={onOpen}>
      <Icon name="wand" size={13} />
      {n ? <span className="ai-badge-count" aria-hidden="true">{n}</span> : null}
    </button>
  );
}
