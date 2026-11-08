// My streams, and adding one (owner, 6 Nov: "the live stream key entry screen is
// very scary and long it should list the the configured streams and allow
// adding additional").
//
// A list with nothing to fill in: each stream's logo, name, platform and shape,
// and a menu with Change and Remove. Add a stream is its own short screen. The
// key goes to the main process and is kept encrypted by the operating system's
// keystore; this window only ever sees its last four characters again.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Dialog, Button, Input, Icon, Menu, t } from '@rutba/office-ui';

/** The shape a server address takes, shown in its empty field: not words to translate. */
const SERVER_SHAPE = 'rtmps://…';

/** The platforms, in the order they are offered. */
const PLATFORM_ORDER = ['youtube', 'facebook', 'twitch', 'tiktok', 'instagram', 'linkedin', 'x', 'kick', 'rumble', 'custom'];

/** A platform's name as shown: the service's own, or words for the one that is not a service. */
export const platformName = (id, platforms) => (id === 'custom' ? t('RTMP server') : platforms?.[id]?.name || id);

/**
 * Where a platform's key is found, in its own words: one line, and a link to
 * the page when the platform's help names one. Written from each platform's
 * help page; TikTok's and Rumble's could not be read, so they say only that
 * the address and key come from the platform.
 */
const keyHelp = (platform) => ({
  youtube: { line: t('In YouTube Studio, choose Create, Go live, then Stream, and copy the stream key. The first time, YouTube can take up to 24 hours to allow it.'), link: t('Open YouTube Studio'), url: 'https://studio.youtube.com/' },
  facebook: { line: t('On Facebook, choose Go live, then Streaming software, and copy the stream key.'), link: t('Open Facebook Live'), url: 'https://www.facebook.com/live/create' },
  twitch: { line: t('In the Twitch Creator Dashboard, open Settings, then Stream, and copy the stream key.'), link: t('Open the Creator Dashboard'), url: 'https://dashboard.twitch.tv/settings/stream' },
  tiktok: { line: t('The server address and the stream key come from TikTok. Copy both here.') },
  instagram: { line: t('On instagram.com, choose Add post, then Live, and copy the address and the stream key. Instagram gives a new key each time, so add it again before you go live.'), link: t('Open Instagram'), url: 'https://www.instagram.com/' },
  linkedin: { line: t('In LinkedIn Live Studio, open Manage streams, choose Prepare to go live, then Get URL, and copy both. Only members who meet LinkedIn\'s access criteria can go live.'), link: t('Open Live Studio'), url: 'https://www.linkedin.com/video/golive/manage/' },
  x: { line: t('In X Live Studio, create a source, then copy its address and stream key.'), link: t('Open Live Studio'), url: 'https://studio.x.com/live' },
  kick: { line: t('In the Kick creator dashboard, open Channel, then Stream URL and Key, and copy both.'), link: t('Open the Kick dashboard'), url: 'https://dashboard.kick.com/stream' },
  rumble: { line: t('The server address and the stream key come from Rumble. Copy both here.') },
  custom: { line: t('The server address and the stream key your streaming service gives you.') },
}[platform] || { line: '' });

export function Logo({ platform, platforms, size = 30, muted = false }) {
  const p = platforms?.[platform];
  return <span className="lv-logo" aria-hidden="true" style={{ width: size, height: size, background: muted ? undefined : p?.colour, fontSize: Math.round(size * 0.4) }}>{p?.mark || '?'}</span>;
}

/** Add a stream, or Change one: choose the platform, one line on where the key is, the key, and the rest out of the way. */
function StreamForm({ shell, toast, platforms, editing, focusKey, onDone, onCancel, onSaved }) {
  const fresh = !editing;
  const [platform, setPlatform] = useState(editing?.platform || 'youtube');
  const known = platforms?.[platform];
  const fixed = Boolean(known?.url);
  const [name, setName] = useState(editing?.name || '');
  const [key, setKey] = useState('');
  const [shape, setShape] = useState(editing?.shape || (known?.tall ? 'tall' : 'wide'));
  const [url, setUrl] = useState(editing?.url && editing.url !== known?.url ? editing.url : '');
  const [more, setMore] = useState(false);
  const [busy, setBusy] = useState(false);
  const keyLine = useRef(null);
  const label = platformName(platform, platforms);

  useEffect(() => { if (focusKey) keyLine.current?.querySelector('input')?.focus(); }, [focusKey]);

  const pick = (p) => {
    setPlatform(p);
    if (fresh) setShape(platforms?.[p]?.tall ? 'tall' : 'wide');
    setUrl('');
    setMore(false);
  };

  const paste = async () => {
    const text = await shell.clipboard.readText().catch(() => '');
    if (text) setKey(text.trim());
  };

  const help = keyHelp(platform);
  const needsAddress = !fixed && !editing?.url && !url.trim();

  const save = async () => {
    setBusy(true);
    try {
      const list = await shell.live.saveDestination({
        id: editing?.id,
        platform,
        name: name.trim() || undefined,
        url: url.trim() || (fixed ? known.url : undefined),
        shape,
        key: key.trim() || undefined,
      });
      onSaved?.(editing?.id || null, list);
      onDone();
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 7000 });
      setBusy(false);
    }
  };

  return (
    <Dialog
      title={fresh ? t('Add a stream') : t('Change {name}', { name: editing.name })}
      className="lv-dialog lv-form"
      onClose={onCancel}
      actions={(
        <>
          <Button label={t('Cancel')} className="lv-cancel" onClick={onCancel} />
          <Button primary label={fresh ? t('Add') : t('Save')} className="lv-save" disabled={busy || needsAddress} onClick={save} />
        </>
      )}
    >
      {fresh ? (
        <div className="lv-grid" role="radiogroup" aria-label={t('Where to')}>
          {PLATFORM_ORDER.map((p) => (
            <button key={p} type="button" role="radio" aria-checked={p === platform} className={`lv-pt${p === platform ? ' on' : ''}`} data-platform={p} onClick={() => pick(p)}>
              <Logo platform={p} platforms={platforms} size={24} />{platformName(p, platforms)}
            </button>
          ))}
        </div>
      ) : null}
      <p className="lv-help">
        {help.line}
        {help.url ? <>{' '}<a href={help.url} onClick={(e) => { e.preventDefault(); shell.shell.openExternal({ url: help.url }).catch(() => {}); }}>{help.link}</a></> : null}
      </p>
      {!fixed ? (
        <label className="lv-field">
          <span>{t('Server address')}</span>
          <Input className="lv-url" value={url || (editing?.url || '')} placeholder={SERVER_SHAPE} onChange={(e) => setUrl(e.target.value)} aria-label={t('Server address')} />
        </label>
      ) : null}
      <label className="lv-field">
        <span>{editing?.hasKey ? t('New stream key') : t('Stream key')}</span>
        <span className="lv-keyline" ref={keyLine}>
          <Input className="lv-keyfield" type="password" autoComplete="off" value={key} placeholder={editing?.hasKey ? t('Leave empty to keep the one kept') : ''} onChange={(e) => setKey(e.target.value)} aria-label={t('Stream key')} />
          <Button label={t('Paste')} className="lv-paste" onClick={paste} />
        </span>
      </label>
      {fresh ? (
        <label className="lv-field">
          <span>{t('Name (optional)')}</span>
          <Input className="lv-name" value={name} placeholder={t('My {platform}', { platform: label })} onChange={(e) => setName(e.target.value)} aria-label={t('Name (optional)')} />
        </label>
      ) : (
        <label className="lv-field">
          <span>{t('Name')}</span>
          <Input className="lv-name" value={name} onChange={(e) => setName(e.target.value)} aria-label={t('Name')} />
        </label>
      )}
      <div className="lv-field">
        <span>{t('Shape')}</span>
        <span className="lv-seg" role="radiogroup" aria-label={t('Shape')}>
          <button type="button" role="radio" aria-checked={shape === 'wide'} className={shape === 'wide' ? 'on' : ''} data-shape="wide" onClick={() => setShape('wide')}>{t('Wide')}</button>
          <button type="button" role="radio" aria-checked={shape === 'tall'} className={shape === 'tall' ? 'on' : ''} data-shape="tall" onClick={() => setShape('tall')}>{t('Tall')}</button>
        </span>
        <small className="lv-small">{t('For wide and tall on one platform, add it twice, with a key for each.')}</small>
      </div>
      {fixed ? (
        more ? (
          <label className="lv-field">
            <span>{t('Server address')}</span>
            <Input className="lv-url" value={url} placeholder={known?.url || SERVER_SHAPE} onChange={(e) => setUrl(e.target.value)} aria-label={t('Server address')} />
          </label>
        ) : <button type="button" className="lv-more" onClick={() => setMore(true)}>{t('Show more: server address')}</button>
      ) : null}
      <small className="lv-small">{t('Kept encrypted on this computer, and sent only to {platform}.', { platform: label })}</small>
    </Dialog>
  );
}

/**
 * My streams: the list, and the screens it opens. Props are those Go Live's old
 * box had, so the studio window opens it as it did.
 * @param {{ shell, toast, onClose, onChanged?: (list) => void, onKeyChanged?: (id) => void, start?: { add: boolean, id?: string } }} props
 */
export function LiveDialog({ shell, toast, onClose, onChanged, onKeyChanged, start = null }) {
  const [list, setList] = useState(null);
  const [platforms, setPlatforms] = useState(null);
  const [screen, setScreen] = useState(start?.add ? { add: true } : null);
  const [menu, setMenu] = useState(null);
  const started = useRef(false);

  const reload = useCallback(async () => {
    const found = await shell.live.destinations().catch(() => []);
    setList(found);
    onChanged?.(found);
    return found;
  }, [shell, onChanged]);

  useEffect(() => {
    reload().then((found) => {
      if (start?.id && !started.current) {
        started.current = true;
        const d = found.find((x) => x.id === start.id);
        if (d) setScreen({ editing: d, focusKey: true });
      }
    });
    shell.live.platforms().then(setPlatforms).catch(() => setPlatforms(null));
  }, [shell, reload, start]);

  const remove = async (id) => {
    const next = await shell.live.removeDestination({ id }).catch(() => list);
    setList(next);
    onChanged?.(next);
  };

  if (screen) {
    return (
      <StreamForm shell={shell} toast={toast} platforms={platforms} editing={screen.editing || null} focusKey={screen.focusKey}
        onSaved={(id, next) => { setList(next); onChanged?.(next); if (id) onKeyChanged?.(id); }}
        onDone={() => setScreen(null)} onCancel={() => setScreen(null)} />
    );
  }

  return (
    <Dialog title={t('My streams')} className="lv-dialog lv-mine" onClose={onClose} actions={<Button primary label={t('Done')} onClick={onClose} className="lv-done" />}>
      <p className="lv-lead">{t('Set up once here. The icons on the Live panel choose which ones go out.')}</p>
      {list == null ? null : list.length ? (
        <ul className="lv-list">
          {list.map((d) => (
            <li key={d.id} className="lv-dest" data-id={d.id} data-platform={d.platform} data-shape={d.shape}>
              <Logo platform={d.platform} platforms={platforms} />
              <span className="lv-who">
                <b>{d.name}</b>
                <small className={d.hasKey ? '' : 'wn'}>{d.hasKey
                  ? `${platformName(d.platform, platforms)} · ${d.shape === 'tall' ? t('Tall') : t('Wide')}`
                  : `${d.shape === 'tall' ? t('Tall') : t('Wide')} · ${t('Needs a key')}`}</small>
              </span>
              {d.hasKey ? null : <Button label={t('Add key')} className="lv-addkey" onClick={() => setScreen({ editing: d, focusKey: true })} />}
              <button type="button" className="lv-more-btn" aria-label={t('More for {name}', { name: d.name })} title={t('More for {name}', { name: d.name })}
                onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); setMenu({ x: r.left, y: r.bottom + 2, d }); }}>
                <Icon name="more" size={16} />
              </button>
            </li>
          ))}
        </ul>
      ) : <p className="lv-none">{t('Nothing set up yet. Add the first place you stream to.')}</p>}
      <div className="lv-add">
        <Button icon="plus" label={t('Add a stream')} className="lv-add-stream" onClick={() => setScreen({ add: true })} />
      </div>
      {menu ? (
        <Menu x={menu.x} y={menu.y} onClose={() => setMenu(null)} items={[
          { label: t('Change'), icon: 'settings', run: () => setScreen({ editing: menu.d }) },
          { label: t('Remove'), icon: 'trash', run: () => remove(menu.d.id) },
        ]} />
      ) : null}
    </Dialog>
  );
}

export const LIVE_CSS = `
.lv-dialog { width: min(452px, 92vw); }
.lv-lead { margin: 0 0 10px; color: var(--ink-2); font-size: 13px; line-height: 1.5; }
.lv-list { list-style: none; margin: 0; padding: 0; display: grid; gap: 6px; }
.lv-dest { display: flex; align-items: center; gap: 11px; min-height: 52px; padding: 6px 8px 6px 10px; border: 1px solid var(--line-soft); border-radius: 9px; background: var(--surface); }
.lv-logo { flex: none; display: grid; place-items: center; border-radius: 8px; background: var(--n-50, #5e6571); color: #fff; font-weight: 700; }
.lv-who { flex: 1; min-width: 0; display: grid; }
.lv-who b { font-size: 13.5px; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.lv-who small { font-size: 12px; color: var(--ink-2); }
.lv-who small.wn { color: #8a4b05; font-weight: 600; }
.lv-more-btn { width: 32px; height: 32px; display: grid; place-items: center; border: 0; border-radius: 6px; background: transparent; color: var(--ink-2); cursor: pointer; }
.lv-more-btn:hover { background: var(--hover, rgba(0,0,0,.06)); }
.lv-none { color: var(--ink-3); font-size: 13px; margin: 0; }
.lv-add { display: flex; justify-content: flex-start; margin-top: 10px; }
.lv-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 6px; }
.lv-pt { display: flex; align-items: center; gap: 8px; height: 44px; padding: 0 9px; border: 1px solid var(--line-soft); border-radius: 8px; background: var(--surface); color: var(--ink); font: inherit; font-size: 13px; cursor: pointer; }
.lv-pt:hover { border-color: var(--line); }
.lv-pt.on { border-color: var(--accent); background: color-mix(in srgb, var(--accent) 12%, var(--surface)); font-weight: 600; }
.lv-help { margin: 12px 0 10px; font-size: 13px; line-height: 1.5; color: var(--ink-2); }
.lv-help a { color: var(--accent); }
.lv-url, .lv-keyfield, .lv-name { height: 34px; padding: 0 10px; border: 1px solid var(--line); border-radius: 7px; box-sizing: border-box; width: 100%; font: inherit; background: var(--surface); color: var(--ink); }
.lv-keyfield { flex: 1; min-width: 0; width: auto; }
.lv-paste, .lv-addkey, .lv-add-stream { border: 1px solid var(--line) !important; }
.lv-form .lv-field { display: flex; flex-direction: column; gap: 4px; margin-bottom: 10px; font-size: 12.5px; color: var(--ink-2); }
.lv-keyline { display: flex; gap: 6px; }
.lv-keyline .rw-input { flex: 1; min-width: 0; }
.lv-seg { display: inline-flex; align-self: flex-start; padding: 2px; gap: 2px; background: var(--n-10, #eceef1); border: 1px solid var(--line-soft); border-radius: 8px; }
.lv-seg button { height: 28px; padding: 0 14px; border: 0; border-radius: 6px; background: transparent; color: var(--ink-2); font: inherit; font-size: 13px; cursor: pointer; }
.lv-seg button.on { background: var(--surface); color: var(--ink); font-weight: 600; box-shadow: 0 1px 2px rgba(15,20,30,.08); }
.lv-small { font-size: 12px; color: var(--ink-3); }
.lv-more { align-self: flex-start; display: block; border: 0; padding: 0; height: 26px; background: transparent; color: var(--accent); font: inherit; font-size: 13px; cursor: pointer; margin-bottom: 4px; }
.st-live-badge { position: absolute; top: 12px; right: 12px; display: inline-flex; align-items: center; gap: 6px; background: #e5484d; color: #fff; font-size: 11.5px; font-weight: 800; letter-spacing: 0.06em; padding: 4px 10px; border-radius: 6px; }
.st-live-badge i { width: 7px; height: 7px; border-radius: 50%; background: #fff; animation: rec-blink 1.2s infinite; }
`;
