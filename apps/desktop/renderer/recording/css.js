// The studio's shared recording marks: the countdown, the red badge, the switch, the teleprompter.

export const RECORDER_CSS = `

.rec-region { position: absolute; border: 1.5px dashed #ffd54a; box-shadow: 0 0 0 9999px rgba(0,0,0,0.5); pointer-events: none; border-radius: 2px; }
.rec-region span { position: absolute; top: -1px; left: -1px; transform: translateY(-100%); background: #ffd54a; color: #1a1300; font-size: 11px; font-weight: 700; padding: 1px 7px; border-radius: 4px 4px 0 0; }
.rec-hint { position: absolute; left: 50%; bottom: 14px; transform: translateX(-50%); background: rgba(10,12,16,0.78); color: #fff; font-size: 12px; padding: 6px 12px; border-radius: 999px; pointer-events: none; white-space: nowrap; }
.rec-badge { position: absolute; top: 12px; left: 12px; display: inline-flex; align-items: center; gap: 6px; background: rgba(10,12,16,0.72); color: #fff; font-size: 11.5px; font-weight: 700; letter-spacing: 0.03em; padding: 4px 10px 4px 8px; border-radius: 999px; }
.rec-badge i { width: 8px; height: 8px; border-radius: 50%; background: #e5484d; animation: rec-blink 1.2s infinite; }
.rec-count { position: absolute; inset: 0; display: grid; place-items: center; background: rgba(8,10,14,0.45); pointer-events: none; }
.rec-count span { width: 120px; height: 120px; border-radius: 50%; display: grid; place-items: center; font-size: 60px; font-weight: 700; color: #fff; background: rgba(229,72,77,0.92); box-shadow: 0 0 0 8px rgba(229,72,77,0.25); font-variant-numeric: tabular-nums; animation: rec-pop 900ms var(--ease); }
@keyframes rec-blink { 50% { opacity: 0.3; } }
@keyframes rec-pop { 0% { transform: scale(1.25); opacity: 0.4; } 30% { transform: scale(1); opacity: 1; } }

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

/* A voice alone: the microphone and its level as it has gone. */
`;

/** The teleprompter, an introduction's scene script above the frame (RecorderPanel's `prompter`). */
export const PROMPTER_CSS = `
/* The teleprompter, above the frame while a scene records. */
.rec-prompter { flex: none; border-radius: 10px; background: #0d0f14; color: #fff; overflow: hidden; display: flex; flex-direction: column; margin-bottom: 8px; border: 1px solid rgba(255,255,255,0.08); }
.rec-prompter-head { display: flex; align-items: center; gap: 8px; padding: 6px 10px; font-size: 12px; color: rgba(255,255,255,0.75); border-bottom: 1px solid rgba(255,255,255,0.08); }
.rec-prompter-head b { color: #fff; font-weight: 600; flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rec-prompter-speed { flex: none; background: rgba(255,255,255,0.08); border-color: rgba(255,255,255,0.12); }
.rec-prompter-speed button { flex: none; min-width: 54px; padding: 2px 10px; font-size: 11.5px; color: rgba(255,255,255,0.72); }
.rec-prompter-speed button:hover:not(:disabled):not(.on) { background: rgba(255,255,255,0.1); color: #fff; }
.rec-prompter-speed button.on { background: rgba(255,255,255,0.2); color: #fff; box-shadow: none; }
.rec-prompter-view { position: relative; height: 92px; overflow: hidden; }
.rec-prompter-view::before, .rec-prompter-view::after { content: ''; position: absolute; left: 0; right: 0; height: 22px; z-index: 1; pointer-events: none; }
.rec-prompter-view::before { top: 0; background: linear-gradient(#0d0f14, rgba(13,15,20,0)); }
.rec-prompter-view::after { bottom: 0; background: linear-gradient(rgba(13,15,20,0), #0d0f14); }
.rec-prompter-text { padding: 18px 18px 60px; font-size: 21px; line-height: 1.45; font-weight: 600; white-space: pre-wrap; transition: transform 260ms linear; }
.rec-prompter-empty { padding: 14px 18px; font-size: 13px; color: rgba(255,255,255,0.6); }
`;
