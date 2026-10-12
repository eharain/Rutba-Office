// The recorder's look: the suite's dialog, wider, with the frame dark as an
// editor's monitor is in either theme, and the settings in the suite's own
// colours on the right. The frame is laid out left to right in every
// language (a screen is not mirrored); the settings follow the window.

export const RECORDER_CSS = `
.rw-dialog.rec-dialog { width: min(1120px, 96vw); max-width: min(1120px, 96vw); max-height: 92vh; }
.rec-dialog .rw-dialog-body { padding: 0 20px 4px; overflow: hidden; display: flex; min-height: 0; }
.rec-dialog .rw-dialog-foot { align-items: center; }
.rec-body { flex: 1; min-height: 0; display: grid; grid-template-columns: minmax(0, 1fr) 300px; gap: 18px; }

.rec-stage { min-width: 0; min-height: 0; display: flex; flex-direction: column; gap: 8px; padding-top: 4px; container-type: inline-size; }
.rec-stage.with-prompter .rec-frame { width: min(100cqw, calc((92vh - 400px) * 16 / 9)); }
.rec-stage.with-prompter .rec-prompter { margin-bottom: 0; }
.rec-frame { position: relative; width: min(100cqw, calc((92vh - 250px) * 16 / 9)); aspect-ratio: 16 / 9; flex: none; align-self: center; outline: none;
  background: radial-gradient(ellipse at center, #1a1d23 0%, #0b0c0f 75%); border-radius: 12px; overflow: hidden;
  box-shadow: 0 0 0 1px rgba(255,255,255,0.06), 0 10px 30px rgba(0,0,0,0.28); user-select: none; touch-action: none; }
.rec-frame.drawing { cursor: crosshair; }
.rec-frame.live { box-shadow: 0 0 0 2px #e5484d, 0 10px 30px rgba(229,72,77,0.25); }
.rec-frame.live.held { box-shadow: 0 0 0 2px #f5a524, 0 10px 30px rgba(0,0,0,0.28); }
.rec-screen { position: absolute; inset: 0; width: 100%; height: 100%; object-fit: contain; pointer-events: none; }
.rec-empty { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 10px; color: rgba(255,255,255,0.72); font-size: 13px; text-align: center; padding: 24px; }
.rec-empty .rw-spinner { border-color: rgba(255,255,255,0.18); border-top-color: #fff; }
.rec-region { position: absolute; border: 1.5px dashed #ffd54a; box-shadow: 0 0 0 9999px rgba(0,0,0,0.5); pointer-events: none; border-radius: 2px; }
.rec-region span { position: absolute; top: -1px; left: -1px; transform: translateY(-100%); background: #ffd54a; color: #1a1300; font-size: 11px; font-weight: 700; padding: 1px 7px; border-radius: 4px 4px 0 0; }
.rec-hint { position: absolute; left: 50%; bottom: 14px; transform: translateX(-50%); background: rgba(10,12,16,0.78); color: #fff; font-size: 12px; padding: 6px 12px; border-radius: 999px; pointer-events: none; white-space: nowrap; }
.rec-inset { position: absolute; overflow: hidden; background: #22262e; cursor: grab; box-shadow: 0 0 0 2px rgba(255,255,255,0.9), 0 6px 20px rgba(0,0,0,0.45); transition: left 160ms var(--ease), top 160ms var(--ease); }
.rec-inset.dragging { transition: none; cursor: grabbing; box-shadow: 0 0 0 2px #ffd54a, 0 10px 28px rgba(0,0,0,0.55); }
.rec-inset.shape-circle { border-radius: 50%; }
.rec-inset.shape-rounded { border-radius: 14%; }
.rec-inset.shape-rect { border-radius: 2px; }
.rec-frame.live .rec-inset, .rec-frame:not(.drawing) .rec-inset { pointer-events: auto; }
.rec-cam { width: 100%; height: 100%; object-fit: cover; display: block; pointer-events: none; }
.rec-cam-off { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 4px; color: rgba(255,255,255,0.75); text-align: center; padding: 8px; }
.rec-cam-off small { font-size: 10.5px; line-height: 1.25; }
.rec-badge { position: absolute; top: 12px; left: 12px; display: inline-flex; align-items: center; gap: 6px; background: rgba(10,12,16,0.72); color: #fff; font-size: 11.5px; font-weight: 700; letter-spacing: 0.03em; padding: 4px 10px 4px 8px; border-radius: 999px; }
.rec-badge i { width: 8px; height: 8px; border-radius: 50%; background: #e5484d; animation: rec-blink 1.2s infinite; }
.rec-frame.held .rec-badge i { background: #f5a524; animation: none; }
.rec-count { position: absolute; inset: 0; display: grid; place-items: center; background: rgba(8,10,14,0.45); pointer-events: none; }
.rec-count span { width: 120px; height: 120px; border-radius: 50%; display: grid; place-items: center; font-size: 60px; font-weight: 700; color: #fff; background: rgba(229,72,77,0.92); box-shadow: 0 0 0 8px rgba(229,72,77,0.25); font-variant-numeric: tabular-nums; animation: rec-pop 900ms var(--ease); }
@keyframes rec-blink { 50% { opacity: 0.3; } }
@keyframes rec-pop { 0% { transform: scale(1.25); opacity: 0.4; } 30% { transform: scale(1); opacity: 1; } }
.rec-under { display: flex; align-items: center; gap: 8px; min-height: 26px; font-size: 12px; color: var(--ink-3); }
.rec-under > span { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

/* The settings scroll beside the frame: they take the frame's height, never set it. */
.rec-side { contain: size; min-height: 380px; overflow: auto; display: flex; flex-direction: column; gap: 4px; padding-inline-end: 4px; margin-inline-end: -4px; }
.rec-side[aria-disabled="true"] { opacity: 0.62; }
.rec-sec { display: flex; flex-direction: column; gap: 8px; padding: 10px 0 12px; border-bottom: 1px solid var(--line-soft); }
.rec-sec:last-child { border-bottom: 0; }
.rec-sec h3 { margin: 0; font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-3); font-weight: 650; }
.rec-sub { display: flex; flex-direction: column; gap: 8px; padding-inline-start: 26px; }
.rec-sub .rw-select { width: 100%; }

.rec-seg { display: flex; padding: 2px; gap: 2px; background: var(--sunken); border-radius: var(--r-2); border: 1px solid var(--line-soft); }
.rec-seg button { flex: 1; min-width: 0; display: inline-flex; align-items: center; justify-content: center; gap: 6px; padding: 5px 6px; border: 0; border-radius: calc(var(--r-2) - 2px); background: transparent; color: var(--ink-2); font: inherit; font-size: 12px; cursor: pointer; white-space: nowrap; }
.rec-seg button:hover:not(:disabled):not(.on) { background: var(--hover); color: var(--ink); }
.rec-seg button.on { background: var(--surface); color: var(--ink); box-shadow: var(--shadow-1); font-weight: 600; }
.rec-seg button.on svg { color: var(--accent); }
.rec-seg button:disabled { cursor: default; }
.rec-seg button span { overflow: hidden; text-overflow: ellipsis; }

.rec-sources { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; max-height: 210px; overflow: auto; padding: 1px; }
.rec-source { display: flex; flex-direction: column; gap: 4px; padding: 5px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); color: var(--ink); font: inherit; font-size: 11.5px; cursor: pointer; text-align: start; min-width: 0; }
.rec-source:hover:not(:disabled) { border-color: var(--accent-line); background: var(--hover); }
.rec-source.on { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.rec-source:disabled { cursor: default; }
.rec-source img { width: 100%; height: 64px; object-fit: contain; background: var(--sunken); border-radius: var(--r-1); }
.rec-source span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rec-wait { grid-column: 1 / -1; display: flex; align-items: center; gap: 8px; color: var(--ink-3); font-size: 12px; padding: 8px 2px; }
.rec-refresh { align-self: flex-start; }

.rec-toggle { display: flex; align-items: center; gap: 10px; cursor: pointer; font-size: 13px; color: var(--ink); min-height: 28px; }
.rec-toggle svg { color: var(--ink-2); flex: none; }
.rec-toggle-label { flex: 1; min-width: 0; }
.rec-toggle input { position: absolute; opacity: 0; width: 1px; height: 1px; pointer-events: none; }
.rec-switch { flex: none; width: 32px; height: 18px; border-radius: 999px; background: var(--n-30, #c5c8ce); position: relative; transition: background var(--fast); }
.rec-switch::after { content: ''; position: absolute; top: 2px; inset-inline-start: 2px; width: 14px; height: 14px; border-radius: 50%; background: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.3); transition: inset-inline-start var(--fast); }
.rec-toggle input:checked + .rec-switch { background: var(--accent); }
.rec-toggle input:checked + .rec-switch::after { inset-inline-start: 16px; }
.rec-toggle input:focus-visible + .rec-switch { box-shadow: 0 0 0 3px var(--accent-soft); }
.rec-toggle.off { cursor: default; }

.rec-meter { display: block; position: relative; height: 6px; border-radius: 999px; background: var(--sunken); overflow: hidden; border: 1px solid var(--line-soft); }
.rec-meter i { position: absolute; inset-block: 0; inset-inline-start: 0; width: 0; background: var(--good, #2bb673); border-radius: 999px; transition: width 90ms linear; }
.rec-meter i[data-level="loud"] { background: #e5484d; }
.rec-meter i[data-level="quiet"] { background: var(--n-40, #a7abb3); }
.rec-foot-info .rec-meter { width: 90px; }
.rec-err { color: var(--bad); font-size: 11.5px; }

.rec-field { display: grid; grid-template-columns: 96px minmax(0, 1fr); align-items: center; gap: 8px; font-size: 12.5px; color: var(--ink-2); }
.rec-field .rw-select { width: 100%; }
.rec-rate { color: var(--ink-3); font-size: 11.5px; }

.rec-foot-info { flex: 1; min-width: 0; display: flex; align-items: center; gap: 10px; font-size: 12.5px; color: var(--ink-2); }
.rec-disk { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rec-disk.low { color: var(--bad); font-weight: 600; }
.rec-dot { width: 10px; height: 10px; border-radius: 50%; background: #e5484d; flex: none; animation: rec-blink 1.2s infinite; }
.rec-dot.held { background: #f5a524; animation: none; }
.rec-time { font-variant-numeric: tabular-nums; font-size: 15px; min-width: 48px; color: var(--ink); }
.rec-chip { font-size: 11px; font-weight: 700; padding: 2px 8px; border-radius: 999px; background: color-mix(in srgb, #f5a524 22%, transparent); color: var(--ink); }
.rec-keys { color: var(--ink-3); font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rec-dialog .rw-btn.rec-go, .rec-dialog .rw-btn.rec-stop { background: #d93036; border-color: #d93036; color: #fff; }
.rec-dialog .rw-btn.rec-go:hover:not(:disabled), .rec-dialog .rw-btn.rec-stop:hover:not(:disabled) { background: #c42a30; }
.rec-dialog .rw-btn.rec-go:disabled { opacity: 0.5; }

@media (max-width: 860px) {
  .rec-body { grid-template-columns: 1fr; overflow: auto; }
  .rec-side { contain: none; }
}
`;
