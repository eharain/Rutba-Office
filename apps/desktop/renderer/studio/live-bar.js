// The bar at the top of the Live panel (owner, 8 Nov: "the icons on top should
// let user enable the streaming destinations ... the live go or no go should
// be simple"). An icon for each stream, grouped under Wide and Tall, its state
// in a word beneath it; click to turn it on or off at any time; click Wide or
// Tall for the whole group. One button holds the final say: Go live, which
// sends to whatever is on, and End broadcast, which stops everything whatever
// the icons say. Every icon off is offline with the broadcast still open.
//
// Stands alone. The rules are @rutba/stream/bar; the encoders are LiveSession.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Icon, t } from '@rutba/office-ui';
import { barModel, reconcileOn, toggleGroup, toggleStream, uploadMbps } from '@rutba/stream/bar';
import { LiveDialog, Logo, platformName } from '../apps/broadcast/live-dialog.js';
import { LiveSession } from '../apps/broadcast/live-session.js';
import { makeTallProgram } from './tall-program.js';

const clock = (ms) => {
  const s = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = String(s % 60).padStart(2, '0');
  return h ? `${h}:${String(m).padStart(2, '0')}:${ss}` : `${m}:${ss}`;
};

const mbps = (kbps) => (kbps / 1000).toFixed(1);

const ShapeIcon = ({ shape }) => (shape === 'wide'
  ? <svg width="20" height="12" viewBox="0 0 20 12" aria-hidden="true"><rect x="1" y="1" width="18" height="10" rx="1.5" fill="none" stroke="currentColor" strokeWidth="1.6" /></svg>
  : <svg width="12" height="20" viewBox="0 0 12 20" aria-hidden="true"><rect x="1" y="1" width="10" height="18" rx="1.5" fill="none" stroke="currentColor" strokeWidth="1.6" /></svg>);

function Tile({ tile, platforms, onToggle, onAddKey }) {
  const d = tile.stream;
  const word = { off: t('common.off'), on: t('common.on'), live: t('common.live'), retrying: t('studio.live-bar.retrying'), nokey: t('common.add-key'), refused: t('common.add-key') }[tile.state];
  const name = `${d.name}, ${platformName(d.platform, platforms)}`;
  const label = tile.state === 'nokey' ? t('studio.live-bar.name-needs-key', { name }) : `${name}, ${word}`;
  const tone = tile.state === 'live' ? ' ok' : tile.state === 'retrying' || tile.state === 'nokey' || tile.state === 'refused' ? ' wn' : tile.state === 'on' ? ' b' : '';
  const noKey = tile.state === 'nokey';
  const tip = tile.state === 'refused' ? `${name}: ${t('studio.live-bar.could-not-connect')}` : label;
  return (
    <div className={`lv-tile${tile.on ? ' on' : ''}`} data-id={d.id} data-state={tile.state} data-on={tile.on ? '1' : '0'}>
      <button type="button" className="lv-tile-go" aria-pressed={noKey ? undefined : tile.on} aria-label={label} title={tip} onClick={() => (noKey ? onAddKey(d.id) : onToggle(d.id))}>
        <Logo platform={d.platform} platforms={platforms} size={26} muted={!tile.on} />
        {tile.state === 'refused' ? null : <span className={`lv-word${tone}`}>{word}</span>}
      </button>
      {tile.state === 'refused' ? <button type="button" className="lv-word wn lv-tile-key" title={tip} onClick={() => onAddKey(d.id)}>{word}</button> : null}
    </div>
  );
}

function Group({ group, label, platforms, onToggleGroup, onToggle, onAddKey }) {
  return (
    <div className="lv-grp" role="group" aria-label={label}>
      <button type="button" className={`lv-shape${group.on ? ' on' : ''}`} data-shape={group.shape} aria-pressed={group.on} aria-label={label} title={label} onClick={() => onToggleGroup(group.shape)}>
        <ShapeIcon shape={group.shape} />{group.shape === 'wide' ? t('common.wide') : t('common.tall')}
      </button>
      {group.tiles.map((x) => <Tile key={x.id} tile={x} platforms={platforms} onToggle={onToggle} onAddKey={onAddKey} />)}
    </div>
  );
}

/**
 * The bar. Stands on its own: give it the studio's parts and it does the rest.
 *
 * @param {object} p
 * @param {object} p.shell
 * @param {(text: string, o?: object) => void} p.toast
 * @param {{ current: object }} p.program  ref to the wide On air Program
 * @param {object} p.media  StudioMedia, for the sound
 * @param {number} p.fps
 * @param {number} p.kbps  each picture's bitrate
 * @param {boolean} [p.copy]  whether a copy is being kept, for its words
 * @param {boolean} [p.blocked]  Go live held back (a countdown under way)
 * @param {string[]} [p.on]  which streams are on, when a setup holds that; with `p.onChangeOn`. Left out, they are kept with the streams.
 * @param {(ids: string[]) => void} [p.onChangeOn]
 * @param {(phase: 'off'|'starting'|'live'|'ending', since: number) => void} [p.onPhase]  told as the broadcast opens and ends
 * @param {number} [p.size]  the studio's output size, 720 or 1080: what the tall picture is made at
 * @param {(() => object|null)} [p.makeTall]  a tall picture's Program, in place of the one made here
 * @param {string} [p.rateText]  said when the picture's bitrate has been lowered
 * @param {number|null} [p.viewers]  how many are watching, when a platform says
 * @param {{ current: object }} [p.control]  filled with { end, setList }: the window ends the broadcast when it closes, and tells the bar of a change made in My streams from the settings
 */
export function LiveBar({ shell, toast, program, media, fps, kbps, copy = false, blocked = false, on: onProp, onChangeOn, onPhase, size = 1080, makeTall = null, rateText = '', viewers = null, control = null }) {
  const [streams, setStreams] = useState([]);
  const [platforms, setPlatforms] = useState(null);
  const [onKept, setOnKept] = useState([]);
  const [status, setStatus] = useState({ destinations: [], rates: {} });
  const [phase, setPhaseState] = useState('off');
  const [since, setSince] = useState(0);
  const [dialog, setDialog] = useState(null); // null, or { start } for My streams
  const [asking, setAsking] = useState(false);
  const [, setTick] = useState(0);
  const session = useRef(null);
  const memory = useRef({});
  const loaded = useRef(false);

  const tallFor = useCallback(() => (makeTall ? makeTall() : makeTallProgram({ wide: program.current, media, size })), [makeTall, program, media, size]);
  const on = onProp === undefined ? onKept : onProp;
  const open = phase !== 'off';
  const model = useMemo(() => barModel({ streams, on, status: status.destinations, open }), [streams, on, status, open]);

  const setPhase = useCallback((p, at = 0) => { setPhaseState(p); onPhase?.(p, at); }, [onPhase]);

  // My streams and what can be chosen, kept as they change.
  const setList = useCallback((list) => {
    setStreams(list);
    if (!loaded.current) {
      loaded.current = true;
      // The on flags the streams already carry are the first choices.
      setOnKept(reconcileOn(list.filter((d) => d.on).map((d) => d.id), list));
    }
  }, []);
  useEffect(() => {
    let alive = true;
    shell.live.destinations().then((list) => { if (alive) setList(list); }).catch(() => {});
    shell.live.platforms().then((p) => { if (alive) setPlatforms(p); }).catch(() => {});
    return () => { alive = false; };
  }, [shell, setList]);

  useEffect(() => shell.on('live:status', (payload = {}) => {
    setStatus({ destinations: payload.destinations || [], rates: payload.rates || {} });
    session.current?.onStatus(payload);
  }), [shell]);

  // Gone while live: the broadcast ends with the window.
  useEffect(() => () => { session.current?.end().catch(() => {}); session.current = null; }, []);

  useEffect(() => {
    if (!open) return undefined;
    const timer = setInterval(() => setTick((n) => n + 1), 1000);
    return () => clearInterval(timer);
  }, [open]);

  /** Change which streams are on, and if the broadcast is open, send or stop them now. */
  const choose = useCallback(async (ids) => {
    if (onProp === undefined) {
      setOnKept(ids);
      shell.live.setOn({ ids }).catch(() => {});
    } else {
      onChangeOn?.(ids);
    }
    if (session.current) {
      try {
        await session.current.setOn(ids, streams);
      } catch (err) {
        toast(err?.message || String(err), { tone: 'bad', ms: 7000 });
      }
    }
  }, [onProp, onChangeOn, shell, streams, toast]);

  const onToggle = useCallback((id) => choose(toggleStream(on, streams, id)), [choose, on, streams]);
  const onToggleGroup = useCallback((shape) => {
    const next = toggleGroup(on, streams, shape, memory.current);
    memory.current = next.memory;
    choose(next.on);
  }, [choose, on, streams]);

  const goLive = useCallback(async () => {
    if (phase !== 'off' || !model.canGo || !program.current) return;
    setPhase('starting');
    const s = new LiveSession({ shell, media, program: program.current, fps, kbps, makeTall: tallFor });
    try {
      const r = await s.open({ ids: model.onIds, streams });
      if (r?.skipped?.length) toast(t('studio.live-bar.not-streamed-names', { names: r.skipped.map((x) => x.name).join(', ') }), { ms: 6000 });
      session.current = s;
      const at = Date.now();
      setSince(at);
      setPhase('live', at);
    } catch (err) {
      await s.end().catch(() => {});
      setPhase('off');
      toast(err?.message || String(err), { tone: 'bad', ms: 8000 });
    }
  }, [phase, model, program, shell, media, fps, kbps, streams, tallFor, setPhase, toast]);

  const endBroadcast = useCallback(async () => {
    setAsking(false);
    const s = session.current;
    if (!s) return;
    setPhase('ending');
    session.current = null;
    await s.end().catch(() => {});
    setPhase('off');
    setStatus({ destinations: [], rates: {} });
    toast(t('studio.live-bar.stream-ended'), { ms: 4000 });
  }, [setPhase, toast]);

  useEffect(() => {
    if (!control) return undefined;
    control.current = { end: endBroadcast, setList };
    return () => { control.current = null; };
  }, [control, endBroadcast, setList]);

  const rate = (shape) => status.rates?.[shape]?.kbps || kbps;
  const upload = uploadMbps(model.tiles, { wide: kbps, tall: kbps });
  const copyWord = copy ? ` · ${t('studio.live-bar.keeping-a-copy')}` : '';
  let note;
  if (!open) note = model.canGo ? t('studio.live-bar.goes-count-about', { count: model.count, mbps: upload.toFixed(1) }) : t('studio.live-bar.turn-stream-go');
  else if (model.air === 'offline') note = `${t('studio.live-bar.nothing-going-out')}${copyWord}`;
  else {
    const parts = [];
    if (model.wide.on) parts.push(t('studio.live-bar.wide-mbps-mb', { mbps: mbps(rate('wide')) }));
    if (model.tall.on) parts.push(t('studio.live-bar.tall-mbps-mb', { mbps: mbps(rate('tall')) }));
    note = `${parts.join(' · ')}${copyWord}`;
  }
  const elapsed = open && since ? clock(Date.now() - since) : '';
  const airWord = { ready: t('common.ready'), live: `${t('studio.live')} ${elapsed}`, offline: `${t('studio.live-bar.offline')} ${elapsed}`, connecting: t('common.connecting') }[model.air];

  return (
    <div className="lv-bar" role="group" aria-label={t('studio.live-bar.where-it-goes')} data-air={model.air}>
      <span className={`lv-air ${model.air}`}>{model.air === 'live' ? <i /> : null}{airWord}</span>
      <div className="lv-bar-streams">
        <Group group={model.wide} label={t('studio.live-bar.wide-streams')} platforms={platforms} onToggleGroup={onToggleGroup} onToggle={onToggle} onAddKey={(id) => setDialog({ start: { id } })} />
        <span className="lv-sep" aria-hidden="true" />
        <Group group={model.tall} label={t('studio.live-bar.tall-streams')} platforms={platforms} onToggleGroup={onToggleGroup} onToggle={onToggle} onAddKey={(id) => setDialog({ start: { id } })} />
      </div>
      <button type="button" className="lv-cfg" onClick={() => setDialog({ start: null })} title={t('common.my-streams')}>
        <Icon name="plus" size={18} />{t('common.my-streams')}
      </button>
      <span className="lv-note">{note}</span>
      {rateText && model.air !== 'offline' ? <span className="st-live-rate">{rateText}</span> : null}
      {open && viewers != null ? <span className="st-viewers" data-tip={t('studio.live-bar.watching-now')}><Icon name="eye" size={13} />{viewers}</span> : null}
      {phase === 'off' || phase === 'starting' ? (
        <button type="button" className={`lv-go${model.canGo ? '' : ' no'}`} disabled={phase === 'starting' || !model.canGo || blocked} title={model.canGo ? t('common.go-live') : t('studio.live-bar.turn-stream-go')} onClick={goLive}>
          <i />{phase === 'starting' ? t('studio.live-bar.starting') : t('common.go-live')}
        </button>
      ) : (
        <button type="button" className="lv-end" disabled={phase === 'ending'} title={t('studio.live-bar.end')} onClick={() => setAsking(true)}>{phase === 'ending' ? t('studio.live-bar.ending') : t('studio.live-bar.end-broadcast')}</button>
      )}
      {dialog ? (
        <LiveDialog shell={shell} toast={toast} start={dialog.start} onChanged={setList} onClose={() => setDialog(null)}
          onKeyChanged={(id) => { if (session.current && on.includes(id)) session.current.restart(id).catch(() => {}); }} />
      ) : null}
      {asking ? (
        <Dialog title={t('studio.live-bar.end-the-broadcast')} className="lv-ask" onClose={() => setAsking(false)} actions={(
          <>
            <Button label={t('studio.live-bar.keep-going')} className="lv-keep" onClick={() => setAsking(false)} />
            <Button primary label={t('studio.live-bar.end')} className="lv-end-now" onClick={endBroadcast} />
          </>
        )}>
          <p className="lv-lead">{t('studio.live-bar.every-stream-stops')}</p>
        </Dialog>
      ) : null}
    </div>
  );
}

export const LIVE_BAR_CSS = `
.st-panel:has(.lv-bar[data-air="offline"]) .st-live-badge { display: none; }
.lv-bar { flex: none; display: flex; align-items: center; gap: 12px; min-height: 72px; padding: 0 16px; background: var(--surface); }
.lv-air { flex: none; display: inline-flex; align-items: center; gap: 7px; height: 28px; padding: 0 11px; border-radius: 14px; font-size: 13px; font-weight: 700; letter-spacing: 0.04em; font-variant-numeric: tabular-nums; white-space: nowrap; }
.lv-air i { width: 9px; height: 9px; border-radius: 50%; background: #fff; display: block; }
.lv-air.ready { background: var(--n-10, #eceef1); color: var(--ink-2); }
.lv-air.connecting { background: var(--n-10, #eceef1); color: #8a4b05; }
.lv-air.live { background: #c42b31; color: #fff; }
.lv-air.offline { background: #454b55; color: #fff; }
.lv-bar-streams { flex: 0 1 auto; min-width: 0; display: flex; align-items: center; gap: 12px; overflow-x: auto; padding: 2px 0; scrollbar-width: thin; }
.lv-grp { flex: none; display: flex; align-items: center; gap: 4px; }
.lv-sep { flex: none; width: 1px; height: 44px; background: var(--line-soft); margin: 0 4px; }
.lv-shape { flex: none; width: 48px; height: 56px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 5px; border: 0; border-radius: 8px; background: transparent; color: var(--ink-3); font: inherit; font-size: 11.5px; cursor: pointer; padding: 0; }
.lv-shape:hover { background: var(--hover, rgba(0,0,0,.05)); }
.lv-shape.on { color: var(--ink); font-weight: 700; }
.lv-tile { flex: none; width: 60px; height: 56px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 3px; border: 1px solid var(--line-soft); border-radius: 9px; background: var(--surface); box-sizing: border-box; }
.lv-tile.on { border: 2px solid var(--accent); background: color-mix(in srgb, var(--accent) 12%, var(--surface)); }
.lv-tile-go { width: 100%; flex: 1; min-height: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 3px; border: 0; background: transparent; padding: 0; cursor: pointer; font: inherit; color: inherit; }
.lv-word { font-size: 11px; color: var(--ink-2); line-height: 1.2; white-space: nowrap; }
.lv-word.ok { color: #0b7a44; font-weight: 600; }
.lv-word.wn { color: #8a4b05; font-weight: 600; }
.lv-word.b { color: var(--accent); font-weight: 600; }
.lv-tile-key { border: 0; background: transparent; padding: 0 0 3px; cursor: pointer; font: inherit; font-size: 11px; font-weight: 600; color: #8a4b05; text-decoration: underline; }
.lv-tile .lv-logo { border-radius: 7px; }
.lv-cfg { flex: none; width: 72px; height: 56px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 4px; border: 0; border-radius: 8px; background: transparent; color: var(--accent); font: inherit; font-size: 11.5px; cursor: pointer; padding: 0; }
.lv-cfg:hover { background: var(--hover, rgba(0,0,0,.05)); }
.lv-note { flex: 1; min-width: 0; font-size: 12.5px; color: var(--ink-2); text-align: end; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.lv-go { flex: none; display: inline-flex; align-items: center; justify-content: center; gap: 9px; height: 44px; min-width: 118px; padding: 0 20px; border-radius: 22px; border: 0; background: #d93a40; color: #fff; font: inherit; font-size: 15px; font-weight: 700; white-space: nowrap; cursor: pointer; }
.lv-go i { width: 11px; height: 11px; border-radius: 50%; background: #fff; display: block; }
.lv-go.no { background: var(--n-20, #e1e4e9); color: var(--ink-3); cursor: default; }
.lv-go.no i { background: var(--n-40, #8d949f); }
.lv-end { flex: none; display: inline-flex; align-items: center; height: 44px; padding: 0 16px; border-radius: 8px; border: 1px solid #d93a40; background: var(--surface); color: #b42329; font: inherit; font-size: 14px; font-weight: 700; white-space: nowrap; cursor: pointer; }
.lv-ask .lv-lead { margin: 0; }
`;
