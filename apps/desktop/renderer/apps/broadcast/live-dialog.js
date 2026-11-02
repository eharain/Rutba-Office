// Go Live's destinations: where the program is streamed to.
//
// Each destination is a platform (YouTube, Facebook, Twitch, or any RTMP
// server), the server address (filled in for the platforms, and changeable),
// and the stream key the platform gives, pasted here once. The key goes to
// the main process and is kept encrypted by the operating system's keystore;
// this window only ever sees its last four characters again. A destination
// can be turned off without being forgotten.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Input, Icon, t } from '@rutba/office-ui';
import { Toggle } from '../../recording/panel.js';

/** The shape a server address takes, shown in its empty field: not words to translate. */
const SERVER_SHAPE = 'rtmps://…';

const PLATFORM_NAMES = () => ({ youtube: 'YouTube', facebook: 'Facebook', twitch: 'Twitch', custom: t('RTMP server') });
const KEY_HELP = () => ({
  youtube: t('In YouTube Studio: Create, Go live, then Stream. Copy the stream key.'),
  facebook: t('In Facebook Live Producer: choose Streaming software, then copy the stream key.'),
  twitch: t('In the Twitch Creator Dashboard: Settings, then Stream. Copy the primary stream key.'),
  custom: t('The server address and the stream key your streaming service gives you.'),
});

/** One destination, as a row that opens to change it. */
function Destination({ d, platforms, onSave, onRemove, busy }) {
  const [open, setOpen] = useState(!d.hasKey);
  const [url, setUrl] = useState(d.url);
  const [name, setName] = useState(d.name);
  const [key, setKey] = useState('');
  const names = PLATFORM_NAMES();
  const save = async () => {
    const ok = await onSave({ id: d.id, platform: d.platform, name, url, on: d.on, key: key || undefined });
    if (ok) { setKey(''); setOpen(false); }
  };
  return (
    <li className={`lv-dest${d.on ? ' on' : ''}`} data-platform={d.platform} data-id={d.id}>
      <div className="lv-row">
        <Toggle on={d.on} label={d.name || names[d.platform]} disabled={busy} onChange={(v) => onSave({ id: d.id, platform: d.platform, name: d.name, url: d.url, on: v })} className="lv-on" />
        <span className={`lv-key${d.hasKey ? ' kept' : ''}`}>{d.hasKey ? t('Key kept {hint}', { hint: d.keyHint }) : t('No stream key yet')}</span>
        <Button ghost icon={open ? 'chevronUp' : 'settings'} title={open ? t('Close') : t('Change')} onClick={() => setOpen((x) => !x)} />
        <Button ghost icon="trash" title={t('Remove this destination')} disabled={busy} onClick={() => onRemove(d.id)} className="lv-remove" />
      </div>
      {open ? (
        <div className="lv-edit">
          <p className="lv-help">{KEY_HELP()[d.platform] || ''}</p>
          <label className="lv-field">
            <span>{t('Name')}</span>
            <Input value={name} onChange={(e) => setName(e.target.value)} aria-label={t('Name')} />
          </label>
          <label className="lv-field">
            <span>{t('Server')}</span>
            <Input className="lv-url" value={url} placeholder={SERVER_SHAPE} onChange={(e) => setUrl(e.target.value)} aria-label={t('Server')} />
          </label>
          <label className="lv-field">
            <span>{d.hasKey ? t('New stream key') : t('Stream key')}</span>
            <Input className="lv-keyfield" type="password" autoComplete="off" value={key} placeholder={d.hasKey ? t('Leave empty to keep the one kept') : ''} onChange={(e) => setKey(e.target.value)} aria-label={t('Stream key')} />
          </label>
          <div className="lv-actions">
            <Button primary label={t('Save')} className="lv-save" disabled={busy || !url.trim() || (!d.hasKey && !key.trim())} onClick={save} />
          </div>
        </div>
      ) : null}
    </li>
  );
}

/**
 * @param {{ shell, toast, onClose, recordCopy, onRecordCopy, kbps, busy }} props
 */
export function LiveDialog({ shell, toast, onClose, recordCopy, onRecordCopy, kbps, busy = false }) {
  const [list, setList] = useState(null);
  const [platforms, setPlatforms] = useState(null);
  const names = PLATFORM_NAMES();

  useEffect(() => {
    // A row added while the list was loading is kept when it arrives.
    const merge = (found) => setList((l) => [...found, ...((l || []).filter((d) => d.draft))]);
    shell.live.destinations().then(merge).catch(() => merge([]));
    shell.live.platforms().then(setPlatforms).catch(() => setPlatforms(null));
  }, [shell]);

  const save = useCallback(async (d) => {
    try {
      setList(await shell.live.saveDestination(d));
      return true;
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 7000 });
      return false;
    }
  }, [shell, toast]);
  const remove = useCallback(async (id) => {
    setList(await shell.live.removeDestination({ id }).catch(() => list));
  }, [shell, list]);
  const add = useCallback(async (platform) => {
    const url = platforms?.[platform]?.url || '';
    if (platform === 'custom') {
      // A custom server has no address to fill in: it is added with a placeholder that is not saved until it is right.
      setList((l) => [...(l || []), { id: '', platform, name: names.custom, url: '', on: true, hasKey: false, keyHint: '', draft: true }]);
      return;
    }
    await save({ platform, name: names[platform], url, on: true });
  }, [platforms, save, names]);

  const on = (list || []).filter((d) => d.on && d.hasKey).length;
  const upload = on ? ((kbps + 128) * on * 1.15) / 1000 : 0;

  return (
    <Dialog
      title={t('Go Live: where to stream')}
      className="lv-dialog"
      onClose={onClose}
      actions={<Button primary label={t('Done')} onClick={onClose} className="lv-done" />}
    >
      <p className="lv-lead">{t('The program is streamed to every destination that is on, at once. Stream keys are kept encrypted on this computer and are sent only to the server they belong to.')}</p>
      {list == null ? null : list.length ? (
        <ul className="lv-list">
          {list.map((d) => (d.draft ? (
            <Destination key="draft" d={d} platforms={platforms} busy={busy} onSave={async (x) => { const ok = await save({ ...x, id: undefined }); return ok; }} onRemove={() => setList((l) => l.filter((y) => !y.draft))} />
          ) : (
            <Destination key={d.id} d={d} platforms={platforms} busy={busy} onSave={save} onRemove={remove} />
          )))}
        </ul>
      ) : <p className="lv-none">{t('No destination yet. Add the platform you stream to.')}</p>}
      <div className="lv-add">
        <span>{t('Add')}</span>
        {['youtube', 'facebook', 'twitch', 'custom'].map((p) => (
          <Button key={p} icon={p === 'custom' ? 'globe' : 'plus'} label={names[p]} className={`lv-add-${p}`} disabled={busy} onClick={() => add(p)} />
        ))}
      </div>
      <div className="lv-foot">
        <Toggle on={recordCopy} icon="record" label={t('Record a copy on this computer while live')} disabled={busy} onChange={onRecordCopy} className="lv-copy" />
        <p className="lv-need">
          <Icon name="info" size={14} />
          {on
            ? t('Streaming to {count} at {kbps} kbit/s needs an upload of about {mbps} Mbit/s.', { count: on, kbps, mbps: upload.toFixed(1) })
            : t('Turn on a destination with a stream key to go live.')}
        </p>
      </div>
    </Dialog>
  );
}

export const LIVE_CSS = `
.lv-dialog { width: min(640px, 92vw); }
.lv-lead { margin: 0 0 12px; color: var(--ink-2); font-size: 13px; line-height: 1.5; }
.lv-list { list-style: none; margin: 0; padding: 0; display: grid; gap: 8px; }
.lv-dest { border: 1px solid var(--line-soft); border-radius: var(--r-2); padding: 6px 8px; background: var(--surface); }
.lv-dest.on { border-color: color-mix(in srgb, var(--accent) 45%, var(--line-soft)); }
.lv-row { display: flex; align-items: center; gap: 8px; }
.lv-row .rec-toggle { flex: 1; min-width: 0; }
.lv-key { font-size: 12px; color: var(--ink-3); white-space: nowrap; font-variant-numeric: tabular-nums; }
.lv-key.kept { color: var(--good, #2bb673); }
.lv-edit { display: grid; gap: 8px; padding: 8px 4px 4px 34px; }
.lv-help { margin: 0; font-size: 12px; color: var(--ink-3); line-height: 1.45; }
.lv-field { display: grid; grid-template-columns: 120px 1fr; align-items: center; gap: 8px; font-size: 12.5px; color: var(--ink-2); }
.lv-actions { display: flex; justify-content: flex-end; }
.lv-none { color: var(--ink-3); font-size: 13px; }
.lv-add { display: flex; align-items: center; flex-wrap: wrap; gap: 6px; margin-top: 14px; }
.lv-add > span { font-size: 12px; color: var(--ink-3); margin-inline-end: 4px; }
.lv-foot { margin-top: 14px; padding-top: 12px; border-top: 1px solid var(--line-soft); display: grid; gap: 8px; }
.lv-need { display: flex; align-items: center; gap: 6px; margin: 0; font-size: 12px; color: var(--ink-3); }
.st-live-badge { position: absolute; top: 12px; right: 12px; display: inline-flex; align-items: center; gap: 6px; background: #e5484d; color: #fff; font-size: 11.5px; font-weight: 800; letter-spacing: 0.06em; padding: 4px 10px; border-radius: 6px; }
.st-live-badge i { width: 7px; height: 7px; border-radius: 50%; background: #fff; animation: rec-blink 1.2s infinite; }
.st-dests { display: flex; flex-wrap: wrap; gap: 6px; justify-content: center; }
.st-dest { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; border-radius: 999px; border: 1px solid var(--line-soft); background: var(--surface); font-size: 12px; color: var(--ink-2); }
.st-dest i { width: 7px; height: 7px; border-radius: 50%; background: var(--n-40, #a7abb3); }
.st-dest[data-state="live"] i { background: #e5484d; }
.st-dest[data-state="connecting"] i, .st-dest[data-state="reconnecting"] i { background: #f5a524; }
.st-dest[data-state="error"] { color: var(--bad, #c0392b); }
`;
