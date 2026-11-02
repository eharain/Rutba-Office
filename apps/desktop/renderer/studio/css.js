// The Recorder's studio (owner, 2026-11-02): the first line across the top,
// the share in the middle with the selected source's bar over it and the
// scenes under it, and one list of every source on the right. The recorder
// dialog's own pieces (rec-*) are reused for the countdown and the meters.

export const STUDIO_CSS = `
.st-top { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; padding: 8px 12px; border-bottom: 1px solid var(--line-soft); background: var(--chrome, var(--surface)); }
.st-acts { display: inline-flex; gap: 6px; }
.st-act { display: inline-flex; align-items: center; gap: 7px; height: 32px; padding: 0 12px; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; font-size: 13px; font-weight: 600; cursor: pointer; white-space: nowrap; font-variant-numeric: tabular-nums; }
.st-act:hover:not(:disabled) { background: var(--hover); border-color: var(--accent-line); }
.st-act:disabled { opacity: 0.5; cursor: default; }
.st-act svg { flex: none; }
.st-act.rec { background: #d93a40; border-color: #d93a40; color: #fff; }
.st-act.rec:hover:not(:disabled) { background: #c2323a; border-color: #c2323a; }
.st-act.on { background: #d93a40; border-color: #d93a40; color: #fff; box-shadow: 0 0 0 3px rgba(229,72,77,0.25); }
.st-act.on.held { background: #c98a12; border-color: #c98a12; box-shadow: 0 0 0 3px rgba(245,165,36,0.25); }
.st-act .st-dot { width: 9px; height: 9px; border-radius: 50%; background: currentColor; flex: none; }
.st-act.on .st-dot { background: #fff; animation: rec-blink 1.2s infinite; }
.st-act.on.held .st-dot { animation: none; }
.st-act-pause { width: 32px; padding: 0; justify-content: center; }
.st-viewers { display: inline-flex; align-items: center; gap: 4px; padding-inline-start: 6px; margin-inline-start: 2px; border-inline-start: 1px solid rgba(255,255,255,0.45); font-weight: 600; }
.st-switch { display: inline-flex; padding: 2px; gap: 2px; background: var(--sunken); border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.st-switch button { display: inline-flex; align-items: center; gap: 6px; height: 26px; padding: 0 10px; border: 0; border-radius: calc(var(--r-2) - 2px); background: transparent; color: var(--ink-2); font: inherit; font-size: 12.5px; cursor: pointer; white-space: nowrap; }
.st-switch button:hover:not(.on) { background: var(--hover); color: var(--ink); }
.st-switch button.on { background: var(--surface); color: var(--ink); box-shadow: var(--shadow-1); font-weight: 600; }
.st-switch button.on svg { color: var(--accent-text); }
.st-switch-label { font-size: 11px; color: var(--ink-3); margin-inline-end: 2px; }
.st-top-gap { flex: 1; }
.st-gear { width: 32px; height: 32px; display: grid; place-items: center; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--surface); color: var(--ink-2); cursor: pointer; }
.st-gear:hover { background: var(--hover); color: var(--ink); }

.st-stage { height: 100%; display: flex; flex-direction: column; align-items: stretch; gap: 10px; padding: 12px 16px 14px; background: var(--window); container-type: size; min-width: 0; }
.st-bar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; min-height: 34px; }
.st-bar-name { display: inline-flex; align-items: center; gap: 6px; font-weight: 600; font-size: 13px; color: var(--ink); max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-bar-name > span { min-width: 0; overflow: hidden; text-overflow: ellipsis; }
.st-bar-hint { color: var(--ink-3); font-size: 12.5px; }
.st-bar .rw-input { width: 180px; }
.st-bar .st-level { width: 120px; }
.st-tool { display: inline-flex; align-items: center; gap: 6px; height: 28px; padding: 0 10px; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; font-size: 12.5px; cursor: pointer; white-space: nowrap; }
.st-tool:hover:not(:disabled) { background: var(--hover); border-color: var(--accent-line); }
.st-tool.on { background: var(--wash-strong); border-color: var(--accent); color: var(--accent-text); font-weight: 600; }
.st-tool:disabled { opacity: 0.5; cursor: default; }
.st-tool.icon { width: 28px; padding: 0; justify-content: center; }
.st-sep { width: 1px; height: 20px; background: var(--line-soft); margin: 0 2px; }

.st-share { flex: 1; min-height: 0; display: flex; align-items: center; justify-content: center; }
.st-frame { position: relative; width: min(100%, calc((100cqh - 140px) * 16 / 9)); aspect-ratio: 16 / 9; border-radius: 10px; overflow: hidden; background: #000; outline: none;
  box-shadow: 0 0 0 1px rgba(255,255,255,0.06), 0 12px 32px rgba(0,0,0,0.3); }
.st-frame:focus-visible { box-shadow: 0 0 0 2px var(--accent), 0 12px 32px rgba(0,0,0,0.3); }
.st-frame.live { box-shadow: 0 0 0 2px #e5484d, 0 12px 32px rgba(229,72,77,0.25); }
.st-frame.live.held { box-shadow: 0 0 0 2px #f5a524, 0 12px 32px rgba(0,0,0,0.3); }
.st-canvas { display: block; width: 100%; height: 100%; }
.st-boxes { position: absolute; inset: 0; }
.st-box { position: absolute; cursor: move; border-radius: 2px; }
.st-box:hover { box-shadow: inset 0 0 0 1px rgba(255,255,255,0.55); }
.st-box.circle { border-radius: 50%; }
.st-box.on { box-shadow: inset 0 0 0 2px var(--accent, #4c8dff), 0 0 0 1px rgba(0,0,0,0.4); }
.st-box.clipping { cursor: crosshair; box-shadow: inset 0 0 0 2px #ffd54a; }
.st-handle { position: absolute; width: 12px; height: 12px; margin: -6px; background: #fff; border: 2px solid var(--accent, #4c8dff); border-radius: 3px; }
.st-handle[data-c="nw"] { left: 0; top: 0; cursor: nwse-resize; }
.st-handle[data-c="ne"] { left: 100%; top: 0; cursor: nesw-resize; }
.st-handle[data-c="sw"] { left: 0; top: 100%; cursor: nesw-resize; }
.st-handle[data-c="se"] { left: 100%; top: 100%; cursor: nwse-resize; }
.st-badge { top: 12px; left: 12px; }
.st-live-badge { position: absolute; top: 12px; right: 12px; display: inline-flex; align-items: center; gap: 6px; background: #d93a40; color: #fff; font-size: 11.5px; font-weight: 700; padding: 4px 10px; border-radius: 999px; font-variant-numeric: tabular-nums; }
.st-live-badge i { width: 8px; height: 8px; border-radius: 50%; background: #fff; animation: rec-blink 1.2s infinite; }
.st-empty { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; gap: 10px; color: rgba(255,255,255,0.78); font-size: 13px; text-align: center; padding: 24px; pointer-events: none; }

.st-scenes { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.st-scenes-label { font-size: 11px; color: var(--ink-3); margin-inline-end: 4px; }
.st-scene { display: inline-flex; align-items: center; gap: 6px; height: 28px; padding: 0 10px; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; font-size: 12.5px; cursor: pointer; white-space: nowrap; }
.st-scene:hover { background: var(--hover); }
.st-scene.on { border-color: var(--accent); background: var(--wash-strong); color: var(--accent-text); font-weight: 600; }
.st-scene kbd { font: inherit; font-size: 11px; color: var(--ink-3); }

.st-under { display: flex; flex-direction: column; gap: 8px; }
/* Live Broadcast's preview: what goes out next beside what is on air, with Cut and Fade between. */
.st-share.duo { gap: 14px; align-items: flex-start; }
.st-pane { display: flex; flex-direction: column; gap: 6px; flex: 1 1 0; min-width: 0; align-items: center; }
.st-pane.off { display: none; }
.st-share.duo .st-pane:not(.on-air) { order: -1; }
.st-pane-label { align-self: flex-start; display: inline-flex; align-items: center; gap: 6px; font-size: 11px; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: var(--ink-3); }
.st-pane.on-air .st-pane-label { color: #d93a40; }
.st-pane.on-air .st-pane-label i { width: 7px; height: 7px; border-radius: 50%; background: #d93a40; }
.st-share.duo .st-frame { width: min(100%, calc((100cqh - 200px) * 16 / 9)); }
.st-share.duo .st-pane.on-air .st-frame { box-shadow: 0 0 0 2px #d93a40, 0 12px 32px rgba(0,0,0,0.3); }
.st-take { display: flex; gap: 8px; justify-content: center; }
.st-take .st-act { min-width: 96px; justify-content: center; }
/* The shots: the one on air marked red, the one picked next marked as picked; a paired microphone and the rotation on each. */
.st-scene.air { border-color: #d93a40; box-shadow: inset 0 0 0 1px #d93a40; }
.st-scene .st-scene-tags { display: inline-flex; gap: 2px; color: var(--ink-3); }
.st-scene.on .st-scene-tags { color: inherit; }
.st-switching { display: inline-flex; align-items: center; gap: 6px; margin-inline-start: auto; flex-wrap: wrap; }
.st-switching .st-tool { height: 28px; }
.st-switching select { height: 28px; }

.st-dests { display: flex; flex-wrap: wrap; gap: 6px; }
.st-flash { position: absolute; inset: 0; background: #fff; pointer-events: none; animation: st-flash 420ms ease-out forwards; }
@keyframes st-flash { from { opacity: 0.85; } to { opacity: 0; } }
.st-saved, .st-shot { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; padding: 6px 10px; border-radius: var(--r-2); background: var(--surface); border: 1px solid var(--line-soft); box-shadow: var(--shadow-1); color: var(--ink); max-width: 100%; }
.st-saved > svg { color: var(--good, #2bb673); }
.st-saved-name, .st-shot-name { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 280px; }
.st-saved-meta, .st-shot-meta { color: var(--ink-3); font-size: 12px; font-variant-numeric: tabular-nums; }
.st-shot img { width: 64px; height: 36px; object-fit: contain; background: var(--sunken); border-radius: 4px; }

.st-list { display: flex; flex-direction: column; gap: 6px; padding: 6px 10px 12px; }
.st-list h3 { margin: 8px 2px 2px; font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-3); font-weight: 650; display: flex; align-items: center; gap: 6px; }
.st-list h3 span { margin-inline-start: auto; text-transform: none; letter-spacing: 0; font-weight: 400; }
.st-drop { min-height: 8px; border-radius: var(--r-2); }
.st-drop.over { outline: 2px dashed var(--accent); outline-offset: -2px; min-height: 36px; }
.st-card { position: relative; border-radius: var(--r-2); border: 1px solid var(--line-soft); background: var(--surface); cursor: pointer; outline: none; }
.st-card:hover { border-color: var(--line); }
.st-card:focus-visible { box-shadow: 0 0 0 2px var(--accent-soft); }
.st-card.sel { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.st-card.dragging { opacity: 0.4; }
.st-card.over-top::before, .st-card.over-bottom::after { content: ''; position: absolute; left: 0; right: 0; height: 3px; background: var(--accent); border-radius: 2px; }
.st-card.over-top::before { top: -5px; }
.st-card.over-bottom::after { bottom: -5px; }
.st-pv { position: relative; aspect-ratio: 16 / 9; background: #14171d; border-radius: calc(var(--r-2) - 1px); overflow: hidden; display: grid; place-items: center; color: rgba(255,255,255,0.55); }
.st-pv video, .st-pv img { width: 100%; height: 100%; object-fit: contain; display: block; }
.st-pv video.mirror { transform: scaleX(-1); }
.st-pv.hidden video, .st-pv.hidden img { opacity: 0.55; }
.st-pv-title { padding: 6px 10px; width: 86%; background: rgba(12,14,20,0.82); border-inline-start: 4px solid #e03b4c; color: #fff; font-size: 12px; font-weight: 700; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
.st-pv-name { position: absolute; inset-inline-start: 6px; bottom: 6px; max-width: calc(100% - 12px); display: inline-flex; align-items: center; gap: 5px; padding: 2px 7px; border-radius: 5px; background: rgba(10,12,16,0.78); color: #fff; font-size: 11.5px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
.st-pv-name svg { flex: none; }
.st-pv-state { position: absolute; inset-inline-start: 6px; top: 6px; padding: 1px 7px; border-radius: 5px; background: rgba(10,12,16,0.78); color: #fff; font-size: 11px; }
.st-pv-btns { position: absolute; inset-inline-end: 5px; top: 5px; display: flex; gap: 4px; }
.st-pv-btn { width: 26px; height: 26px; display: grid; place-items: center; border-radius: 6px; border: 0; background: rgba(10,12,16,0.72); color: #fff; cursor: pointer; padding: 0; }
.st-pv-btn:hover { background: rgba(10,12,16,0.92); }
.st-pv-btn.on { background: var(--accent-fill, #2f6fec); }
.st-pv-btn.stop { background: rgba(217,58,64,0.9); }
.st-pv-grip { position: absolute; inset-inline-start: 0; top: 0; bottom: 0; width: 14px; cursor: grab; }
.st-snd { display: grid; grid-template-columns: 24px 1fr auto; align-items: center; gap: 8px; padding: 8px 8px 8px 10px; }
.st-snd-name { font-size: 12.5px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; color: var(--ink); }
.st-snd .rec-meter { grid-column: 2; }
.st-snd-btns { grid-row: 1 / span 2; grid-column: 3; display: flex; gap: 4px; }
.st-snd-btns .st-pv-btn { background: var(--sunken); color: var(--ink-2); }
.st-snd-btns .st-pv-btn.on { background: var(--accent-fill, #2f6fec); color: #fff; }
.st-snd-btns .st-pv-btn.stop { background: rgba(217,58,64,0.9); color: #fff; }
.st-snd > svg { grid-row: 1 / span 2; color: var(--ink-2); }
.st-snd.off .st-snd-name { color: var(--ink-3); }
.st-err { color: var(--bad, #d93a40); font-size: 11.5px; padding: 0 10px 6px; }
.st-add { margin-top: 6px; }

.st-settings { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 16px; }
.st-settings .wide { grid-column: 1 / -1; }
.st-settings-row { display: flex; align-items: center; gap: 8px; }
.st-settings-row span { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink-2); font-size: 12.5px; }
`;
