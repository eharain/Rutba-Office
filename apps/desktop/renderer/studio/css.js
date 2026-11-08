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

.st-panel { height: 100%; display: flex; flex-direction: column; min-width: 0; min-height: 0; background: var(--window); }
.st-stage { flex: 1; min-height: 0; display: flex; flex-direction: column; align-items: stretch; gap: 10px; padding: 12px 16px; background: var(--window); min-width: 0; }
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

.st-share { flex: 1; min-height: 0; min-width: 0; display: flex; align-items: center; justify-content: center; container-type: size; }
.st-canvas-slot { position: absolute; inset: 0; }
.st-frame { position: relative; width: min(100cqw, calc(100cqh * 16 / 9)); aspect-ratio: 16 / 9; border-radius: 10px; overflow: hidden; background: #000; outline: none;
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

.st-under-frame { flex: none; display: flex; flex-direction: column; gap: 8px; min-width: 0; }
.st-scenes { display: flex; align-items: stretch; gap: 10px; overflow-x: auto; padding: 2px 2px 6px; }
.st-scenes-head { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; min-height: 30px; }
.st-scenes-hint { color: var(--ink-3); font-size: 12px; }
.st-dock-gap { flex: 1; }
.st-scenes-label { font-size: 11px; color: var(--ink-3); margin-inline-end: 4px; }
.st-scene { flex: none; width: 172px; display: flex; flex-direction: column; gap: 0; padding: 0; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; font-size: 12.5px; cursor: pointer; overflow: hidden; text-align: start; }
.st-scenes.small .st-scene { width: 128px; }
.st-scene:hover { border-color: var(--accent-line); }
.st-scene:focus-visible { outline: 2px solid var(--ink-3); outline-offset: 2px; }
.st-scene.on { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
.st-scene.next { border-color: #2bb673; box-shadow: 0 0 0 1px #2bb673; }
.st-scene.lost { border-color: #f5a524; }
.st-thumb { position: relative; display: grid; place-items: center; aspect-ratio: 16 / 9; background: #14171d; color: rgba(255,255,255,0.7); }
.st-thumb-canvas { display: block; width: 100%; height: 100%; }
.st-thumb-note { position: absolute; inset: auto 6px 6px 6px; text-align: center; color: #fff; font-size: 11.5px; background: rgba(10,12,16,0.72); border-radius: 4px; padding: 2px 4px; }
.st-tag { position: absolute; inset-block-start: 6px; inset-inline-start: 6px; padding: 1px 6px; border-radius: 4px; font-size: 10.5px; font-weight: 700; color: #fff; letter-spacing: 0.03em; }
.st-tag.air { background: #d93a40; }
.st-tag.next { background: #2bb673; }
.st-lost-mark { position: absolute; inset-block-start: 6px; inset-inline-end: 6px; width: 20px; height: 20px; display: grid; place-items: center; border-radius: 50%; background: #f5a524; color: #201500; }
.st-scene-foot { display: flex; align-items: center; gap: 6px; padding: 5px 8px; min-height: 28px; min-width: 0; }
.st-scene-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-scene kbd { font: inherit; font-size: 11px; font-weight: 600; color: var(--ink-3); background: var(--sunken); border-radius: 4px; padding: 0 5px; }
.st-scene-save { align-items: center; justify-content: center; gap: 6px; flex-direction: row; color: var(--ink-2); border-style: dashed; min-height: 80px; }
.st-scene-save:disabled { opacity: 0.5; cursor: default; }
.st-pick { display: flex; align-items: stretch; gap: 8px; overflow-x: auto; padding: 6px; background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.st-pick-label { align-self: center; color: var(--ink-2); font-size: 12.5px; white-space: nowrap; padding-inline: 6px; }
.st-pick-item { flex: none; width: 120px; display: flex; flex-direction: column; gap: 4px; padding: 4px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: transparent; color: var(--ink); font: inherit; font-size: 11.5px; cursor: pointer; }
.st-pick-item:hover, .st-pick-item.on { border-color: var(--accent); }
.st-pick-item img { width: 100%; aspect-ratio: 16 / 9; object-fit: contain; background: #14171d; border-radius: 3px; }
.st-pick-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-pick-close { align-self: flex-start; width: 26px; height: 26px; display: grid; place-items: center; border: 0; border-radius: 6px; background: transparent; color: var(--ink-2); cursor: pointer; position: sticky; inset-inline-end: 0; }
.st-pick-close:hover { background: var(--hover); }

/* The top line: the setup, the two panels in the middle, the settings. */
.st-setup-btn b { font-weight: 650; }
.st-setup-word { color: var(--ink-3); font-weight: 400; }
.st-setup-btn.locked { display: inline-flex; align-items: center; gap: 5px; height: 32px; padding: 0 10px; color: var(--ink-2); font-size: 13px; }
.st-undo { display: inline-flex; gap: 4px; }
.st-hist { width: 32px; height: 32px; display: grid; place-items: center; border-radius: var(--r-2); border: 1px solid transparent; background: transparent; color: var(--ink-2); cursor: pointer; }
.st-hist:hover:not(:disabled) { background: var(--hover); color: var(--ink); }
.st-hist:disabled { opacity: 0.35; cursor: default; }
.st-seg { display: inline-flex; padding: 2px; gap: 2px; background: var(--sunken); border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.st-seg button { height: 28px; padding: 0 22px; border: 0; border-radius: calc(var(--r-2) - 2px); background: transparent; color: var(--ink-2); font: inherit; font-size: 13px; font-weight: 600; cursor: pointer; white-space: nowrap; }
.st-seg button:hover:not(.on):not(:disabled) { background: var(--hover); color: var(--ink); }
.st-seg button:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.st-seg button.on { background: var(--surface); color: var(--ink); box-shadow: var(--shadow-1); }
.st-seg.small button { height: 24px; padding: 0 10px; font-size: 12px; font-weight: 500; }

/* The bar at the foot of Record and Live. */
.st-dock { flex: none; display: flex; align-items: center; gap: 8px; min-height: 68px; padding: 10px 16px; border-top: 1px solid var(--line-soft); background: var(--chrome, var(--surface)); flex-wrap: wrap; }
.st-bb { display: inline-flex; align-items: center; gap: 8px; height: 40px; padding: 0 12px; border-radius: var(--r-2); border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; font-size: 13px; cursor: pointer; white-space: nowrap; }
.st-bb:hover:not(:disabled) { background: var(--hover); border-color: var(--accent-line); }
.st-bb:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.st-bb:disabled { opacity: 0.5; cursor: default; }
.st-bb.icon { width: 40px; padding: 0; justify-content: center; }
.st-bb.off { color: var(--bad, #d93a40); }
.st-bb.on svg { color: var(--accent-text); }
.st-bb-name { max-width: 150px; overflow: hidden; text-overflow: ellipsis; }
.st-bb-split { display: inline-flex; }
.st-bb-split .st-bb:first-child { border-start-end-radius: 0; border-end-end-radius: 0; }
.st-bb-split .st-bb:last-child { width: 26px; border-start-start-radius: 0; border-end-start-radius: 0; margin-inline-start: -1px; }
.st-meter { width: 74px; }
.st-rec { display: inline-flex; align-items: center; justify-content: center; gap: 10px; height: 48px; min-width: 176px; padding: 0 26px; border-radius: 24px; border: 0; background: #d93a40; color: #fff; font: inherit; font-size: 15px; font-weight: 700; cursor: pointer; white-space: nowrap; font-variant-numeric: tabular-nums; }
.st-rec:hover:not(:disabled) { background: #c2323a; }
.st-rec:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.st-rec:disabled { opacity: 0.5; cursor: default; }
.st-rec .st-dot { width: 12px; height: 12px; border-radius: 50%; background: #fff; }
.st-rec .st-square { width: 12px; height: 12px; border-radius: 2px; background: #fff; }
.st-rec.on { background: #2a2d33; }
.st-rec.on.held { background: #c98a12; }
.st-round { width: 48px; height: 48px; display: grid; place-items: center; border-radius: 50%; border: 1px solid var(--line); background: var(--surface); color: var(--ink); cursor: pointer; }
.st-round:hover:not(:disabled) { background: var(--hover); }
.st-round:disabled { opacity: 0.5; cursor: default; }
.st-round.on { background: #c98a12; border-color: #c98a12; color: #fff; }
.st-check { display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink); cursor: pointer; white-space: nowrap; }

/* The Live panel's bar. */
.st-livebar { flex: none; border-bottom: 1px solid var(--line-soft); background: var(--chrome, var(--surface)); }
.st-live-rate { color: var(--ink-3); font-size: 12px; font-variant-numeric: tabular-nums; white-space: nowrap; }
.st-flip-slot:empty { display: none; }
.st-take { flex: none; align-self: center; display: flex; flex-direction: column; align-items: center; gap: 6px; }
.st-take-btn { min-width: 96px; justify-content: center; height: 44px; font-size: 14px; }
.st-act.primary { background: var(--accent-fill, #2f6fec); border-color: var(--accent-fill, #2f6fec); color: #fff; }
.st-act.primary:hover:not(:disabled) { background: var(--accent-fill, #2f6fec); filter: brightness(1.08); }
.st-take kbd { font: inherit; font-size: 11px; color: var(--ink-3); }
.st-pane-frame { flex: 1; min-height: 0; width: 100%; display: flex; align-items: center; justify-content: center; container-type: size; }
.st-share:not(.duo) .st-pane.on-air { flex: 1; height: 100%; }
.st-share:not(.duo) .st-pane-label { display: none; }

/* Set up's three columns. */
.st-setup-grid { flex: 1; min-height: 0; display: grid; grid-template-columns: 268px minmax(0, 1fr) 272px; }
.st-setup-grid > .st-stage { min-height: 0; }
.st-col { min-height: 0; display: flex; flex-direction: column; background: var(--chrome, var(--surface)); }
.st-sources { border-inline-end: 1px solid var(--line-soft); }
.st-props { border-inline-start: 1px solid var(--line-soft); }
.st-col-head { flex: none; display: flex; align-items: center; justify-content: space-between; gap: 8px; min-height: 44px; padding: 6px 12px; font-size: 11px; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; color: var(--ink-3); }
.st-col-name { display: inline-flex; align-items: center; gap: 5px; min-width: 0; max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; text-transform: none; letter-spacing: 0; font-size: 12.5px; font-weight: 600; color: var(--ink); }
.st-col-body { flex: 1; min-height: 0; overflow: auto; padding: 2px 12px 14px; display: flex; flex-direction: column; gap: 12px; }
.st-col-note { margin: 0; color: var(--ink-3); font-size: 12px; line-height: 1.45; }
.st-col-section { display: flex; flex-direction: column; gap: 8px; padding-bottom: 12px; border-bottom: 1px solid var(--line-soft); }
.st-col-section:last-child { border-bottom: 0; }
.st-col-section h3 { margin: 0; display: flex; align-items: center; justify-content: space-between; font-size: 12.5px; font-weight: 650; color: var(--ink); }
.st-col-section kbd { font: inherit; font-size: 11px; color: var(--ink-3); background: var(--sunken); border-radius: 4px; padding: 0 5px; }
.st-col-row { display: flex; align-items: center; gap: 8px; }
.st-col-row > span { flex: 1; min-width: 0; font-size: 12.5px; color: var(--ink-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-col-actions { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.st-sources .st-list { flex: 1; min-height: 0; overflow: auto; padding: 0 8px 12px; gap: 3px; }
.st-src { display: flex; align-items: center; gap: 8px; min-height: 44px; padding: 5px 6px 5px 8px; border-color: transparent; background: transparent; }
.st-src:hover { background: var(--hover); border-color: transparent; }
.st-src.sel { background: var(--wash-strong); border-color: var(--accent-line); box-shadow: none; }
.st-src.miss { background: rgba(245,165,36,0.12); border-color: rgba(245,165,36,0.55); }
.st-src.off .st-src-pic { opacity: 0.5; }
.st-src-pic { flex: none; width: 44px; height: 26px; display: grid; place-items: center; border-radius: 4px; overflow: hidden; background: #2a2e36; color: rgba(255,255,255,0.7); }
.st-src-pic video, .st-src-pic img { width: 100%; height: 100%; object-fit: cover; display: block; }
.st-src-text { flex: 1; min-width: 0; display: flex; flex-direction: column; line-height: 1.25; }
.st-src-name { font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink); }
.st-src-text small { font-size: 11px; color: var(--ink-3); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-src-text small.st-src-err { color: var(--bad, #d93a40); }
.st-src.unset .st-src-name { white-space: normal; }
.st-src-miss { font-size: 11px; font-weight: 700; color: #a56a00; }
.st-src-choose { flex: none; align-self: flex-start; margin-top: 3px; height: 24px; padding: 0 8px; border-radius: 5px; border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; font-size: 12px; cursor: pointer; }
.st-src-choose:hover { border-color: var(--accent-line); }
.st-src .st-pv-btn { width: 24px; height: 24px; background: transparent; color: var(--ink-3); }
.st-src .st-pv-btn:hover { background: var(--sunken); color: var(--ink); }
.st-src .st-pv-btn.on { background: transparent; color: var(--accent-text); }
.st-src .st-pv-btn.stop { background: transparent; color: var(--ink-3); }
.st-src .st-run, .st-src .st-more { opacity: 0; }
.st-src:hover .st-run, .st-src:hover .st-more, .st-src:focus-within .st-run, .st-src:focus-within .st-more, .st-src.sel .st-run, .st-src.sel .st-more { opacity: 1; }
.st-frame-head { flex: none; display: flex; align-items: baseline; gap: 10px; min-height: 22px; font-size: 13px; }
.st-frame-size, .st-frame-hint { color: var(--ink-3); font-size: 12px; font-variant-numeric: tabular-nums; }
.st-places { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 6px; }
.st-place { position: relative; height: 40px; padding: 0; border: 1px solid var(--line); border-radius: 6px; background: var(--surface); cursor: pointer; overflow: hidden; }
.st-place:hover { border-color: var(--accent); }
.st-place i { position: absolute; background: var(--ink-3); border-radius: 2px; }
.st-fourfields { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.st-fourfields input { width: 100%; }
.st-props .st-bar { display: flex; flex-direction: column; align-items: stretch; gap: 12px; min-height: 0; }
.st-props .st-bar .rw-input { width: 100%; }
.st-props .st-level { width: 100%; }
.st-shapes { width: 100%; }
.st-shapes button { flex: 1; padding: 0 4px; }
.st-auto-seg { width: 100%; }
.st-auto-seg button { flex: 1; padding: 0 4px; font-size: 11.5px; }
.st-add .rw-btn { height: 28px; }

/* Set up: what stays on air, and the bands while running. */
.st-band { flex: none; padding: 6px 16px; font-size: 12.5px; font-weight: 600; color: #fff; }
.st-band.live { background: #d93a40; }
.st-band.rec { background: #2a2d33; }
.st-mirror { position: absolute; inset-block-start: 10px; inset-inline-end: 10px; width: 200px; border-radius: 6px; overflow: hidden; box-shadow: 0 0 0 2px #d93a40, 0 8px 20px rgba(0,0,0,0.4); background: #000; pointer-events: none; }
.st-mirror span { position: absolute; inset-block-start: 4px; inset-inline-start: 4px; padding: 1px 6px; background: #d93a40; color: #fff; border-radius: 4px; font-size: 10.5px; font-weight: 700; }
.st-mirror-canvas { display: block; width: 100%; aspect-ratio: 16 / 9; }
.st-guide { position: absolute; background: #ff5da2; pointer-events: none; }
.st-guide.x { inset-block: 0; width: 1px; }
.st-guide.y { inset-inline: 0; height: 1px; }

.st-under { display: flex; flex-direction: column; gap: 8px; }
/* Live Broadcast's preview: what goes out next beside what is on air, with Cut and Fade between. */
.st-share.duo { gap: 14px; align-items: stretch; container-type: normal; }
.st-pane { display: flex; flex-direction: column; gap: 6px; flex: 3 1 0; min-width: 0; min-height: 0; align-items: center; }
.st-next-pane { flex: 2 1 0; }
.st-pane.off { display: none; }
.st-pane-label { align-self: flex-start; display: inline-flex; align-items: center; gap: 6px; font-size: 11px; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: var(--ink-3); }
.st-pane.on-air .st-pane-label { color: #d93a40; }
.st-pane.on-air .st-pane-label i { width: 7px; height: 7px; border-radius: 50%; background: #d93a40; }
.st-share.duo .st-pane.on-air .st-frame { box-shadow: 0 0 0 2px #d93a40, 0 12px 32px rgba(0,0,0,0.3); }
/* The shots: the one on air marked red, the one picked next marked as picked; a paired microphone and the rotation on each. */
.st-scene.air { border-color: #d93a40; box-shadow: inset 0 0 0 1px #d93a40; }
.st-scene .st-scene-tags { display: inline-flex; gap: 2px; color: var(--ink-3); }
.st-scene.on .st-scene-tags { color: inherit; }
.st-switching { display: inline-flex; align-items: center; gap: 6px; margin-inline-start: auto; flex-wrap: wrap; }
.st-switching .st-tool { height: 28px; }
.st-switching select { height: 28px; }

.st-flash { position: absolute; inset: 0; background: #fff; pointer-events: none; animation: st-flash 420ms ease-out forwards; }
@keyframes st-flash { from { opacity: 0.85; } to { opacity: 0; } }
.st-saved { display: flex; flex-direction: column; gap: 10px; padding: 14px; border-radius: var(--r-2); background: var(--surface); border: 1px solid var(--line-soft); box-shadow: var(--shadow-1); color: var(--ink); }
.st-saved-top { display: flex; gap: 14px; align-items: center; min-width: 0; }
.st-saved-pic { flex: none; width: 128px; aspect-ratio: 16 / 9; display: grid; place-items: center; background: #14171d; color: rgba(255,255,255,0.7); border-radius: 6px; overflow: hidden; }
.st-saved-pic img { width: 100%; height: 100%; object-fit: cover; display: block; }
.st-saved-text { min-width: 0; }
.st-saved-kicker { color: var(--good, #2bb673); font-size: 12px; font-weight: 700; }
.st-saved-name { font-size: 15px; }
.st-saved-btns { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.st-shot { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; padding: 6px 10px; border-radius: var(--r-2); background: var(--surface); border: 1px solid var(--line-soft); box-shadow: var(--shadow-1); color: var(--ink); max-width: 100%; }
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

.st-settings { display: flex; flex-direction: column; gap: 18px; }
.st-into { display: inline-flex; align-items: center; gap: 6px; margin-inline-start: 10px; padding: 3px 10px; border-radius: 999px; background: var(--wash-strong); color: var(--accent-text); font-size: 12px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 320px; }
.st-set-section { display: flex; flex-direction: column; gap: 10px; }
.st-set-section + .st-into { display: inline-flex; align-items: center; gap: 6px; margin-inline-start: 10px; padding: 3px 10px; border-radius: 999px; background: var(--wash-strong); color: var(--accent-text); font-size: 12px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 320px; }
.st-set-section { padding-top: 16px; border-top: 1px solid var(--line-soft); }
.st-set-head { display: flex; align-items: baseline; gap: 10px; flex-wrap: wrap; }
.st-set-head h2 { margin: 0; font-size: 13.5px; font-weight: 650; color: var(--ink); }
.st-set-head span { font-size: 12px; color: var(--ink-3); }
.st-set-row { display: grid; grid-template-columns: 170px minmax(0, 1fr); align-items: center; gap: 12px; min-height: 32px; }
.st-set-label { font-size: 13px; color: var(--ink-2); }
.st-set-value { min-width: 0; }
.st-set-value .rw-select { width: 100%; }
.st-set-key { font: inherit; font-size: 12px; font-weight: 600; color: var(--ink-2); background: var(--sunken); border-radius: 4px; padding: 2px 7px; }
.st-settings-row { display: flex; align-items: center; gap: 8px; }
.st-open-title { margin: 0 0 2px; font-size: 18px; }
.st-open-sub { margin: 0 0 12px; color: var(--ink-2); }
.st-open-rows { display: flex; flex-direction: column; border: 1px solid var(--line-soft); border-radius: var(--r-2); max-height: 300px; overflow: auto; }
.st-open-row { display: flex; align-items: center; gap: 10px; padding: 8px 12px; border-bottom: 1px solid var(--line-soft); }
.st-open-row:last-child { border-bottom: 0; }
.st-open-name { flex: 1; min-width: 0; font-weight: 600; }
.st-open-name small { display: block; font-weight: 400; color: var(--ink-3); font-size: 12px; }
.st-open-how { color: var(--ink-3); font-size: 12px; }
.st-open-ok { color: var(--good, #2bb673); display: inline-flex; }
.st-open-no { color: #c98a12; display: inline-flex; }
.st-settings-row span { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink-2); font-size: 12.5px; }
`;
