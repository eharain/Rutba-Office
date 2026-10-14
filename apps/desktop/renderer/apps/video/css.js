// The Video window's look. The monitor and the timeline are dark in either
// theme, as an editor's are: pictures are judged against black, and a
// clip's colours read on a dark track. The ribbon and the properties pane
// follow the suite's theme.

export const CSS = `
.vd-editor { flex: 1; min-height: 0; display: flex; flex-direction: column; background: #121418; }
.vd-top { flex: 1; min-height: 0; display: flex; }
.vd-monitor { flex: 1; min-width: 0; display: flex; flex-direction: column; background: #0b0c0f; }
.vd-stage { flex: 1; min-height: 0; display: grid; place-items: center; position: relative; overflow: hidden;
  background: radial-gradient(ellipse at center, #16181d 0%, #0b0c0f 70%); }
.vd-stage .rw-empty { color: rgba(255,255,255,0.78); }
.vd-stage .rw-empty h2, .vd-stage .rw-empty .rw-empty-title { color: #fff; }
.vd-frame { position: relative; box-shadow: 0 0 0 1px rgba(255,255,255,0.08), 0 12px 40px rgba(0,0,0,0.55); background: #000; }
.vd-canvas { width: 100%; height: 100%; display: block; cursor: default; }
.vd-outline { position: absolute; border: 1.5px solid #ffd54a; box-shadow: 0 0 0 1px rgba(0,0,0,0.6); pointer-events: none; border-radius: 2px; }
.vd-export-canvas { position: fixed; left: -10000px; top: 0; width: 16px; height: 9px; pointer-events: none; }

.vd-transport { display: flex; align-items: center; gap: 4px; padding: 6px 12px; background: #15171c; border-top: 1px solid #23262d; color: #d9dce3; }
.vd-transport .rw-btn { color: #cfd3db; }
.vd-transport .rw-btn:hover:not(:disabled) { background: rgba(255,255,255,0.08); }
.vd-transport .vd-play { width: 34px; height: 34px; border-radius: 50%; background: var(--accent); color: #fff; margin: 0 4px; }
.vd-transport .vd-play:hover:not(:disabled) { background: color-mix(in srgb, var(--accent) 85%, #fff); }
.vd-time { font-size: 12px; font-variant-numeric: tabular-nums; min-width: 58px; text-align: center; color: #f1f3f7; font-weight: 600; }
.vd-time.total { color: #8b919c; font-weight: 500; }
.vd-scrub { flex: 1; accent-color: var(--accent); min-width: 80px; }

.vd-split { height: 6px; flex: none; cursor: row-resize; background: #0e1013; border-top: 1px solid #23262d; border-bottom: 1px solid #23262d; }
.vd-split:hover { background: #1c1f26; }
/* Its grip, so it reads as something to take hold of. */
.vd-split { display: grid; place-items: center; outline: none; }
.vd-split::after { content: ''; width: 36px; height: 2px; border-radius: 1px; background: #3a3f49; }
.vd-split:hover::after, .vd-split:focus-visible::after { background: var(--accent); }
.vd-split:focus-visible { background: #1c1f26; }
/* Time runs left to right in every language, as it does in every editor; the words in it follow the window. */
.vd-timeline { flex: none; display: flex; min-height: 0; background: #17191e; direction: ltr; }
.vd-transport { direction: ltr; }
.vd-tl-scroll { flex: 1; overflow: auto; position: relative; min-height: 0; }
.vd-tl-scroll::-webkit-scrollbar { height: 11px; width: 11px; }
.vd-tl-scroll::-webkit-scrollbar-thumb { background: #3a3f49; border-radius: 6px; border: 2px solid #17191e; }
.vd-tl-scroll::-webkit-scrollbar-track { background: #17191e; }
.vd-tl-inner { position: relative; }

.vd-ruler { position: sticky; top: 0; z-index: 6; display: flex; height: 28px; background: #1d2026; border-bottom: 1px solid #2c3038; }
.vd-ruler-corner { position: sticky; left: 0; z-index: 7; width: 120px; flex: none; background: #1d2026; border-right: 1px solid #2c3038; display: flex; align-items: center; justify-content: center; }
.vd-ruler-time { font-size: 12px; font-weight: 700; color: #ffd54a; font-variant-numeric: tabular-nums; }
.vd-ruler-track { position: relative; height: 100%; cursor: ew-resize; }
.vd-tick { position: absolute; top: 0; bottom: 0; border-left: 1px solid #4a505b; }
.vd-tick span { position: absolute; left: 4px; top: 3px; font-size: 10px; color: #9aa1ad; white-space: nowrap; font-variant-numeric: tabular-nums; }
.vd-tick.minor { top: 19px; border-left-color: #353a43; }
.vd-range { position: absolute; top: 0; bottom: 0; background: rgba(255, 213, 74, 0.16); border-left: 2px solid #ffd54a; border-right: 2px solid #ffd54a; }
.vd-marker { position: absolute; bottom: 1px; width: 10px; height: 12px; margin-left: -5px; padding: 0; border: 0; background: #4fc3f7; clip-path: polygon(0 0, 100% 0, 100% 62%, 50% 100%, 0 62%); cursor: pointer; }
.vd-playhead-cap { position: absolute; bottom: 0; width: 13px; height: 13px; margin-left: -6px; background: #ff4d5e; clip-path: polygon(0 0, 100% 0, 100% 50%, 50% 100%, 0 50%); pointer-events: none; }
.vd-playhead { position: absolute; top: 14px; width: 2px; margin-left: -1px; background: #ff4d5e; pointer-events: none; z-index: 5; box-shadow: 0 0 6px rgba(255,77,94,0.5); }
.vd-snapline { position: absolute; top: 0; width: 1px; background: #ffd54a; pointer-events: none; z-index: 5; }

.vd-row { display: flex; border-bottom: 1px solid #23262d; }
.vd-row.first-sound { border-top: 2px solid #2f343d; }
.vd-row-head { position: sticky; left: 0; z-index: 4; flex: none; display: flex; align-items: center; gap: 2px; padding: 0 6px 0 10px; background: #1d2026; border-right: 1px solid #2c3038; color: #c8ccd4; }
.vd-row.pic .vd-row-head { box-shadow: inset 3px 0 0 #5b8def; }
.vd-row.snd .vd-row-head { box-shadow: inset 3px 0 0 #2fbf8a; }
.vd-row-name { font-size: 11.5px; font-weight: 700; letter-spacing: 0.02em; flex: 1; }
/* The grips: the track names' column wider at the ruler's corner, each kind of track taller at the foot of its name. */
.vd-head-grip { position: absolute; top: 0; bottom: 0; right: -3px; width: 6px; cursor: col-resize; z-index: 8; }
.vd-lane-grip { position: absolute; left: 0; right: 0; bottom: -3px; height: 6px; cursor: row-resize; z-index: 5; }
.vd-head-grip:hover, .vd-lane-grip:hover { background: color-mix(in srgb, var(--accent) 60%, transparent); }
.vd-row-btn { width: 24px; height: 22px; border: 0; border-radius: 5px; background: transparent; color: #8b919c; display: grid; place-items: center; position: relative; cursor: pointer; }
.vd-row-btn:hover { background: rgba(255,255,255,0.08); color: #e6e9ef; }
.vd-row-btn.on { color: #ffd54a; }
.vd-slash { position: absolute; width: 16px; height: 1.5px; background: #ffd54a; transform: rotate(-45deg); }
.vd-row-track { position: relative; background: #1a1c21; }
.vd-row.pic .vd-row-track { background: repeating-linear-gradient(90deg, #1a1c21 0, #1a1c21 59px, #1c1e24 59px, #1c1e24 60px); }
.vd-row.empty .vd-row-track { background: #16181c; }
.vd-row-track.dim .vd-clip { opacity: 0.4; }

.vd-clip { position: absolute; top: 3px; border-radius: 6px; overflow: hidden; cursor: grab; box-sizing: border-box;
  border: 1px solid rgba(255,255,255,0.14); box-shadow: 0 1px 2px rgba(0,0,0,0.4); user-select: none; touch-action: none; }
.vd-clip:active { cursor: grabbing; }
.vd-clip.k-video { background: #2c3a5a; }
.vd-clip.k-image { background: #3d3360; }
.vd-clip.k-sound { background: #1a4a3c; border-color: rgba(120, 230, 190, 0.25); }
.vd-clip.k-text { background: linear-gradient(180deg, #7b4bd1, #5f37a8); }
.vd-clip.k-shape { background: linear-gradient(180deg, #b7792a, #8c5b1d); }
.vd-zoom-row { display: flex; align-items: center; gap: 6px; margin: 4px 0; }
.vd-zoom-row .rw-select { flex: 1; min-width: 0; padding: 3px 6px; }
.vd-zoom-len { font-size: 11px; color: var(--ink-2); white-space: nowrap; }
.vd-zoom-on, .vd-zoom-off { margin-top: 6px; }
.vd-export-frame { display: flex; align-items: center; gap: 10px; margin: 8px 0 6px; font-size: 12px; color: var(--ink-2); line-height: 1.45; }
.vd-export-shape { flex: none; display: block; border: 1.5px solid currentColor; border-radius: 3px; opacity: 0.8; }
.vd-export-shape.s-landscape { width: 32px; height: 18px; }
.vd-export-shape.s-square { width: 22px; height: 22px; }
.vd-export-shape.s-vertical { width: 15px; height: 26px; }
.vd-export-warn { margin-top: 6px; font-size: 12px; color: var(--bad, #c62828); }
.vd-clip.k-caption { background: linear-gradient(180deg, #2f6f8f, #22566f); }
.vd-clip.k-pointer { background: repeating-linear-gradient(135deg, #5a4a1f 0 6px, #4a3d19 6px 12px); }
.vd-clip.selected { border: 2px solid #ffd54a; box-shadow: 0 0 0 1px rgba(0,0,0,0.6), 0 0 12px rgba(255,213,74,0.35); z-index: 2; }
.vd-clip.missing { background: repeating-linear-gradient(45deg, #4a1f25 0 8px, #3a171c 8px 16px); }
.vd-clip-art { position: absolute; left: 0; top: 0; pointer-events: none; }
.vd-clip.k-video .vd-clip-art, .vd-clip.k-image .vd-clip-art { opacity: 0.92; }
.vd-clip-label { position: absolute; left: 0; right: 0; bottom: 0; display: flex; align-items: center; gap: 5px; padding: 2px 7px 3px;
  font-size: 11px; color: #fff; pointer-events: none; background: linear-gradient(180deg, rgba(0,0,0,0) 0%, rgba(0,0,0,0.72) 100%); text-shadow: 0 1px 2px rgba(0,0,0,0.6); }
.vd-clip.k-sound .vd-clip-label, .vd-clip.k-text .vd-clip-label, .vd-clip.k-shape .vd-clip-label { top: 0; bottom: auto; background: none; }
.vd-clip.k-text .vd-clip-label, .vd-clip.k-shape .vd-clip-label { bottom: 0; top: 0; }
.vd-clip.k-pointer .vd-clip-label { top: 0; bottom: auto; background: none; }
.vd-clip.k-sound .vd-clip-label, .vd-clip.k-pointer .vd-clip-label { padding-right: 24px; }
.vd-clip-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; min-width: 0; font-weight: 600; }
.vd-clip-len { color: rgba(255,255,255,0.75); font-variant-numeric: tabular-nums; font-size: 10px; }
.vd-link { position: absolute; top: 3px; right: 5px; width: 15px; height: 15px; border-radius: 4px; background: rgba(0,0,0,0.5); color: #fff; display: grid; place-items: center; pointer-events: none; }
.vd-tr { position: absolute; top: 0; bottom: 0; left: 0; display: grid; place-items: center; color: #fff; pointer-events: none;
  background: linear-gradient(90deg, rgba(255,255,255,0.0), rgba(255,255,255,0.28)); border-right: 1px dashed rgba(255,255,255,0.7); }
.vd-tr svg { filter: drop-shadow(0 1px 2px rgba(0,0,0,0.6)); }
.vd-trim { position: absolute; top: 0; bottom: 0; width: 8px; cursor: ew-resize; z-index: 3; }
.vd-trim.start { left: 0; }
.vd-trim.end { right: 0; }
.vd-clip:hover .vd-trim, .vd-clip.selected .vd-trim { background: rgba(255, 213, 74, 0.55); }
.vd-clip.selected .vd-trim.start { border-radius: 5px 0 0 5px; }
.vd-clip.selected .vd-trim.end { border-radius: 0 5px 5px 0; }

.vd-inspector { display: flex; flex-direction: column; gap: 4px; font-size: 12.5px; padding: 2px 14px 16px; }
.vd-insp-head { display: flex; align-items: center; gap: 10px; padding: 4px 2px 10px; border-bottom: 1px solid var(--line-soft); }
.vd-insp-kind { width: 30px; height: 30px; border-radius: 8px; display: grid; place-items: center; color: #fff; flex: none; background: #4a5568; }
.vd-insp-kind.k-video { background: #3b5ba8; }
.vd-insp-kind.k-image { background: #6b55b0; }
.vd-insp-kind.k-sound { background: #23896a; }
.vd-insp-kind.k-text { background: #7b4bd1; }
.vd-insp-kind.k-shape { background: #b7792a; }
.vd-insp-kind.k-project { background: var(--accent); }
.vd-insp-title { display: flex; flex-direction: column; min-width: 0; }
.vd-insp-title strong { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vd-insp-title small { color: var(--ink-3); }
.vd-insp-section { padding: 8px 0 10px; border-bottom: 1px solid var(--line-soft); }
.vd-insp-section h3 { margin: 0 0 6px; font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--ink-3); font-weight: 700; }
.vd-prop { display: flex; align-items: center; gap: 8px; min-height: 30px; }
.vd-prop-label { width: 92px; flex: none; color: var(--ink-2); }
.vd-prop-value { flex: 1; min-width: 0; display: flex; align-items: center; gap: 4px; }
.vd-prop-value .rw-select, .vd-prop-value select { flex: 1; min-width: 0; }
.vd-prop-note { color: var(--ink-3); font-size: 12px; }
.vd-num { width: 76px; }
.vd-slider { display: flex; align-items: center; gap: 6px; flex: 1; min-width: 0; }
.vd-slider input { flex: 1; width: 0; accent-color: var(--accent); min-width: 0; }
.vd-slider-value { width: 38px; text-align: end; font-variant-numeric: tabular-nums; color: var(--ink-2); font-size: 11.5px; }
.vd-text { width: 100%; resize: vertical; margin-bottom: 6px; font: inherit; box-sizing: border-box; }
.vd-swatches { display: flex; gap: 4px; flex-wrap: wrap; }
.vd-swatch { width: 20px; height: 20px; border-radius: 50%; border: 1px solid var(--line-strong); cursor: pointer; padding: 0; }
.vd-swatch.on { box-shadow: 0 0 0 2px var(--surface), 0 0 0 4px var(--accent); }
.vd-warn { color: #c0392b; font-size: 12px; margin-top: 4px; }
.vd-marker-row { display: flex; align-items: center; gap: 4px; margin-bottom: 4px; }
.vd-marker-row .rw-input { flex: 1; min-width: 0; }
.vd-marker-time { border: 0; background: var(--sunken); border-radius: 5px; padding: 4px 6px; font-variant-numeric: tabular-nums; font-size: 11.5px; color: var(--ink-2); cursor: pointer; }
.vd-source { display: flex; align-items: center; gap: 6px; padding: 3px 0; color: var(--ink-2); }
.vd-source.missing { color: #c0392b; }
.vd-source-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vd-source-len { color: var(--ink-3); font-variant-numeric: tabular-nums; font-size: 11.5px; }
.flip-x svg { transform: scaleX(-1); }
.turn-90 svg { transform: rotate(90deg); }

.vd-title-btn { min-width: 64px; }
.vd-title-tile { width: 44px; height: 26px; border-radius: 4px; background: linear-gradient(135deg, #3b4a6b, #1c2333); position: relative; overflow: hidden; display: block; margin-bottom: 2px; box-shadow: inset 0 0 0 1px rgba(0,0,0,0.25); }
.vd-title-tile .band, .vd-title-tile .words { position: absolute; display: block; }
.vd-title-tile.p-title .words { left: 9px; right: 9px; top: 10px; height: 5px; background: #fff; border-radius: 2px; }
.vd-title-tile.p-lowerThird .band { left: 3px; bottom: 3px; width: 26px; height: 8px; background: var(--accent); border-radius: 2px; }
.vd-title-tile.p-lowerThird .words { left: 5px; bottom: 5px; width: 16px; height: 3px; background: #fff; border-radius: 1px; }
.vd-title-tile.p-endCard { background: #101218; }
.vd-title-tile.p-endCard .words { left: 11px; right: 11px; top: 9px; height: 7px; background: #fff; border-radius: 2px; }
.vd-title-tile.p-callout .words { right: 4px; top: 3px; width: 18px; height: 7px; background: var(--accent); border-radius: 4px; }
.vd-ribbon-field { display: flex; flex-direction: column; gap: 4px; font-size: 11.5px; color: var(--ink-2); padding: 2px 6px; }

.vd-export-format { font-weight: 600; }

/* A picture in picture's corner, picked from four small frames. */
.vd-corners { display: inline-grid; grid-template-columns: repeat(2, 28px); gap: 4px; direction: ltr; }
.vd-corner { width: 28px; height: 18px; padding: 0; border: 1px solid var(--line-strong); border-radius: 4px; background: var(--surface); position: relative; cursor: pointer; }
.vd-corner:hover { border-color: var(--accent-line); background: var(--hover); }
.vd-corner i { position: absolute; width: 9px; height: 6px; border-radius: 1.5px; background: var(--ink-3); }
.vd-corner.c-tl i { top: 2px; left: 2px; }
.vd-corner.c-tr i { top: 2px; right: 2px; }
.vd-corner.c-bl i { bottom: 2px; left: 2px; }
.vd-corner.c-br i { bottom: 2px; right: 2px; }
.vd-corner.on { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.vd-corner.on i { background: var(--accent); }
.vd-inset-full { margin-top: 6px; }
.vd-start-actions { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; margin-top: 14px; }
.vd-export-line { margin-bottom: 10px; font-variant-numeric: tabular-nums; }
`;
