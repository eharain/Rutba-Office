// Notifications (D-2026-11-22-23): what Rutba Office tells the person of
// while they are doing something else, and whether it stays by the clock to
// do so. Opened from Settings in any app's menu and from the icon's own
// menu; drawn by the launcher (?notifications opens with it).
//
// The settings are the main process's (watch.js, the `notify` calls): this
// window shows them and sends each change as it is made.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Select, Icon, t } from '@rutba/office-ui';

/** The kinds a notification comes in, in the order they are listed. */
const KIND_ROWS = ['call', 'missed', 'chat', 'mail', 'reminder', 'meeting'];

const kindWords = () => ({
  call: [t('notify-dialog.incoming-calls'), t('notify-dialog.call-talk-while')],
  missed: [t('talk.missed-calls'), t('notify-dialog.call-talk-nobody')],
  chat: [t('mail.messages'), t('notify-dialog.message-talk')],
  mail: [t('common.new-email'), t('notify-dialog.new-mail-arrives')],
  reminder: [t('notify-dialog.event-reminders'), t('notify-dialog.event-calendar-reminder')],
  meeting: [t('notify-dialog.meetings-about-start'), t('notify-dialog.meeting-few-minutes')],
});

/**
 * Why none of this copy's notifications is shown, in words: this copy hands
 * Windows none, Windows' own switches are off, or Windows said it did not
 * show the last one (watch.js, `block`).
 */
const blockWords = (block) => ({
  copy: t('notify-dialog.only-installed-copy-rutba'),
  policy: t('notify-dialog.notifications-turned-off-computer'),
  all: t('notify-dialog.notifications-turned-off'),
  app: t('notify-dialog.notifications-rutba'),
  refused: t('notify-dialog.windows-did-not'),
})[block] || '';
/** The blocks the person can lift in Windows' own settings. */
const LIFTED_IN_WINDOWS = ['all', 'app', 'refused'];

/** Every quarter of an hour of the day, as "22:00". */
const TIMES = Array.from({ length: 96 }, (_, i) => `${String(Math.floor(i / 4)).padStart(2, '0')}:${String((i % 4) * 15).padStart(2, '0')}`);

/** A time of day as the person's own clock writes it. */
const shownTime = (hhmm) => {
  const d = new Date(2000, 0, 1, Number(hhmm.slice(0, 2)), Number(hhmm.slice(3)));
  return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
};

/** The window. A kind is listed only when this copy watches for it (its app is installed and on). */
export function NotifyDialog({ shell, onClose }) {
  const [state, setState] = useState(null);
  const [error, setError] = useState('');

  useEffect(() => {
    let live = true;
    shell.notify?.state().then((s) => { if (live) setState(s); }).catch((e) => { if (live) setError(e?.message || String(e)); });
    const off = shell.on?.('notify:state', (s) => { if (live) setState(s); });
    // Back from Windows' own settings: what was turned on there is read again.
    const again = () => { shell.notify?.state().then((s) => { if (live) setState(s); }).catch(() => {}); };
    window.addEventListener('focus', again);
    return () => { live = false; off?.(); window.removeEventListener('focus', again); };
  }, [shell]);

  const change = useCallback((settings) => {
    shell.notify.set({ settings }).then(setState).catch((e) => setError(e?.message || String(e)));
  }, [shell]);

  const s = state?.settings;
  const words = kindWords();
  const rows = KIND_ROWS.filter((kind) => (state?.watched || []).includes(kind));
  const blocked = blockWords(state?.block);

  return (
    <Dialog
      title={t('common.notifications')}
      width={560}
      className="notify-dialog"
      onClose={onClose}
      actions={
        <>
          {state?.supported ? <Button label={t('notify-dialog.show-one-now')} title={t('notify-dialog.shows-notification-see')} onClick={() => shell.notify.test().catch(() => {})} /> : null}
          <span style={{ flex: 1 }} />
          <Button primary label={t('common.close')} onClick={onClose} />
        </>
      }
    >
      <style>{NOTIFY_CSS}</style>
      {error ? <p className="notify-error" role="alert">{error}</p> : null}
      {!s ? null : (
        <>
          {blocked ? (
            <p className="notify-note notify-block" role="note" data-block={state.block}>
              <Icon name="info" size={15} />
              <span>{blocked}</span>
              {state.systemPage && LIFTED_IN_WINDOWS.includes(state.block) ? <Button className="notify-windows" label={t('notify-dialog.open-settings-windows')} onClick={() => shell.notify.windowsSettings().catch(() => {})} /> : null}
            </p>
          ) : state.supported ? null : <p className="notify-note" role="note"><Icon name="info" size={15} /> <span>{t('notify-dialog.computer-does-not')}</span></p>}
          {state.paused ? (
            <p className="notify-note" role="note">
              <Icon name="info" size={15} />
              <span>{t('common.notifications-paused', { time: new Date(state.pausedUntil).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) })}</span>
              <Button label={t('common.turn-notifications-back')} onClick={() => shell.notify.resume().then(setState).catch(() => {})} />
            </p>
          ) : null}

          <section className="notify-section" aria-label={t('notify-dialog.tell-me-of')}>
            <h3>{t('notify-dialog.tell-me-of')}</h3>
            {rows.length ? null : <p className="notify-lead">{t('notify-dialog.talk-mail-calendar')}</p>}
            {rows.map((kind) => (
              <label key={kind} className="notify-row">
                <input type="checkbox" className={`notify-kind-${kind}`} checked={s.kinds[kind] !== false} onChange={(e) => change({ kinds: { [kind]: e.target.checked } })} />
                <span className="notify-text"><strong>{words[kind][0]}</strong><span>{words[kind][1]}</span></span>
              </label>
            ))}
          </section>

          <section className="notify-section" aria-label={t('notify-dialog.how')}>
            <h3>{t('notify-dialog.how')}</h3>
            <label className="notify-row">
              <input type="checkbox" className="notify-sound" checked={s.sound} onChange={(e) => change({ sound: e.target.checked })} />
              <span className="notify-text"><strong>{t('notify-dialog.play-a-sound')}</strong><span>{t('notify-dialog.sound-windows-plays')}</span></span>
            </label>
            <label className="notify-row">
              <input type="checkbox" className="notify-show-text" checked={s.text} onChange={(e) => change({ text: e.target.checked })} />
              <span className="notify-text"><strong>{t('notify-dialog.show-says')}</strong><span>{t('notify-dialog.words-message-subject')}</span></span>
            </label>
            <label className="notify-row">
              <input type="checkbox" className="notify-quiet" checked={s.quiet.on} onChange={(e) => change({ quiet: { on: e.target.checked } })} />
              <span className="notify-text"><strong>{t('notify-dialog.quiet-hours')}</strong><span>{t('notify-dialog.no-notifications-between')}</span></span>
            </label>
            <div className="notify-hours" aria-disabled={!s.quiet.on}>
              <label>
                <span>{t('common.from')}</span>
                <Select className="notify-quiet-from" value={s.quiet.from} disabled={!s.quiet.on} onChange={(e) => change({ quiet: { from: e.target.value } })}>
                  {TIMES.map((v) => <option key={v} value={v}>{shownTime(v)}</option>)}
                </Select>
              </label>
              <label>
                <span>{t('calendar.until')}</span>
                <Select className="notify-quiet-to" value={s.quiet.to} disabled={!s.quiet.on} onChange={(e) => change({ quiet: { to: e.target.value } })}>
                  {TIMES.map((v) => <option key={v} value={v}>{shownTime(v)}</option>)}
                </Select>
              </label>
            </div>
          </section>

          <section className="notify-section" aria-label={t('notify-dialog.keeping-watch')}>
            <h3>{t('notify-dialog.keeping-watch')}</h3>
            <label className="notify-row">
              <input type="checkbox" className="notify-keep-running" checked={s.keepRunning} onChange={(e) => change({ keepRunning: e.target.checked })} />
              <span className="notify-text"><strong>{t('notify-dialog.keep-rutba-office')}</strong><span>{t('notify-dialog.stays-clock-while')}</span></span>
            </label>
            <label className="notify-row">
              <input type="checkbox" className="notify-startup" checked={state.startup.on} disabled={!state.startup.can || !s.keepRunning} onChange={(e) => change({ startup: e.target.checked })} />
              <span className="notify-text">
                <strong>{t('notify-dialog.start-rutba-office')}</strong>
                <span>{state.startup.can ? t('notify-dialog.starts-clock-no') : t('notify-dialog.only-installed-copy')}</span>
              </span>
            </label>
            <p className="notify-lead">{t('notify-dialog.rutba-office-watches')}</p>
          </section>
        </>
      )}
    </Dialog>
  );
}

export const NOTIFY_CSS = `
.notify-section { margin: 0 0 14px; }
.notify-section h3 { margin: 0 0 6px; font-size: 12px; font-weight: 600; color: var(--ink-3); text-transform: none; }
.notify-row { display: flex; align-items: flex-start; gap: 10px; padding: 7px 12px; border-radius: var(--r-2); border: 1px solid var(--line-soft); background: var(--surface); margin: 0 0 6px; }
.notify-row input { margin-top: 3px; }
.notify-text { flex: 1; min-width: 0; display: grid; }
.notify-text strong { font-weight: 600; }
.notify-text span { font-size: 11.5px; color: var(--ink-3); white-space: normal; }
.notify-hours { display: flex; gap: 16px; margin: 0 0 6px; padding-inline-start: 34px; }
.notify-hours label { display: flex; align-items: center; gap: 8px; font-size: 12.5px; }
.notify-hours[aria-disabled="true"] { opacity: 0.55; }
.notify-lead { margin: 8px 0 0; color: var(--ink-2); font-size: 12px; line-height: 1.5; }
.notify-note { display: flex; align-items: center; gap: 8px; margin: 0 0 12px; padding: 8px 12px; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--wash); font-size: 12.5px; }
.notify-note > span { flex: 1; }
.notify-error { color: var(--bad-text, var(--ink)); font-size: 12px; white-space: normal; }
`;
