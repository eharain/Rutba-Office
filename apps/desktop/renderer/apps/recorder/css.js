// The Recorder window: scenes on the left, the program in the middle, the
// sources on the right. The recorder dialog's own pieces (rec-*) are reused.

export const STUDIO_CSS = `
.st-scenes { display: grid; gap: 6px; padding: 10px; }
.st-scene { display: flex; align-items: center; gap: 10px; padding: 9px 10px; border-radius: var(--r-2); border: 1px solid var(--line-soft); background: var(--surface); color: var(--ink); font: inherit; text-align: start; cursor: pointer; position: relative; }
.st-scene:hover { border-color: var(--line); background: var(--wash); }
.st-scene.on { border-color: var(--accent); background: var(--wash-strong); box-shadow: inset 3px 0 0 var(--accent); }
.st-scene-mark { width: 30px; height: 30px; border-radius: 8px; display: grid; place-items: center; background: var(--sunken); color: var(--accent-text); flex: none; }
.st-scene.on .st-scene-mark { background: var(--accent-fill); color: #fff; }
.st-live-dot { position: absolute; inset-inline-end: 10px; width: 8px; height: 8px; border-radius: 50%; background: #e5484d; animation: rec-blink 1.2s infinite; }
.st-hint { margin: 4px 12px 12px; font-size: 12px; color: var(--ink-3); line-height: 1.45; }

.st-stage { height: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 12px; padding: 18px; background: var(--window); container-type: size; }
.st-frame { position: relative; width: min(100%, calc((100cqh - 80px) * 16 / 9)); aspect-ratio: 16 / 9; border-radius: 12px; overflow: hidden; background: #000;
  box-shadow: 0 0 0 1px rgba(255,255,255,0.06), 0 12px 32px rgba(0,0,0,0.3); }
.st-frame.live { box-shadow: 0 0 0 2px #e5484d, 0 12px 32px rgba(229,72,77,0.25); }
.st-frame.live.held { box-shadow: 0 0 0 2px #f5a524, 0 12px 32px rgba(0,0,0,0.3); }
.st-canvas { display: block; width: 100%; height: 100%; }
.st-frame.drawing { cursor: crosshair; }
.st-region-bar { display: flex; align-items: center; gap: 8px; font-size: 12.5px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.st-badge { top: 12px; left: 12px; }
.st-wait { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; gap: 10px; color: rgba(255,255,255,0.78); font-size: 13px; text-align: center; padding: 24px; }
.st-wait .rw-spinner { border-color: rgba(255,255,255,0.18); border-top-color: #fff; }
.st-saved { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; padding: 8px 12px; border-radius: var(--r-2); background: var(--surface); border: 1px solid var(--line-soft); box-shadow: var(--shadow-1); color: var(--ink); max-width: 100%; }
.st-saved > svg { color: var(--good, #2bb673); }
.st-saved-name { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 320px; }
.st-saved-meta { color: var(--ink-3); font-size: 12px; font-variant-numeric: tabular-nums; }

.st-side { padding: 4px 12px 16px; display: grid; gap: 4px; }
.st-side[aria-disabled="true"] .st-kinds, .st-side[aria-disabled="true"] .st-sources { opacity: 0.6; }
.st-sources { max-height: 220px; }
.st-device { color: var(--ink-3); font-size: 12px; }
.st-folder { display: block; color: var(--ink-3); font-size: 12px; margin-top: 6px; overflow-wrap: anywhere; }
.st-logo { display: flex; align-items: center; gap: 8px; margin-top: 8px; font-size: 12px; color: var(--ink-2); }
.st-logo img { width: 40px; height: 28px; object-fit: contain; background: var(--sunken); border-radius: 4px; }
.st-logo span { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-corners button, .st-logo .rec-seg button { min-width: 30px; justify-content: center; }
.st-brb-field { margin-top: 8px; }
`;
