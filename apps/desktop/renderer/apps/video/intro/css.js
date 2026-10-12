// The introduction's panel, layout drawings, brand kit and teleprompter.
export const INTRO_CSS = `
.vi-scenes { display: flex; flex-direction: column; height: 100%; min-height: 0; min-width: 0; overflow: hidden; }
.vi-list { list-style: none; margin: 0; padding: 8px; display: flex; flex-direction: column; gap: 6px; overflow: auto; flex: 1; min-height: 0; }
.vi-scene { border: 1px solid var(--line); border-radius: 8px; background: var(--surface); }
.vi-scene.open { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent) inset; }
.vi-scene-head { display: flex; align-items: center; gap: 8px; width: 100%; padding: 7px 8px; background: none; border: 0; font: inherit; color: var(--ink); text-align: start; cursor: pointer; border-radius: 8px; }
.vi-scene-head:hover { background: var(--hover); }
.vi-num { flex: none; width: 18px; height: 18px; border-radius: 50%; background: var(--sunken); color: var(--ink-2); font-size: 11px; font-weight: 700; display: grid; place-items: center; }
.vi-scene.open .vi-num { background: var(--accent); color: var(--ink-on-accent); }
.vi-scene-name { flex: 1; min-width: 0; display: flex; flex-direction: column; line-height: 1.25; }
.vi-scene-name b { font-size: 12.5px; font-weight: 600; }
.vi-scene-name small { font-size: 11px; color: var(--ink-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vi-scene.waiting .vi-scene-name small { color: var(--bad); }
.vi-len { font-size: 11px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.vi-scene-body { padding: 2px 10px 10px; display: flex; flex-direction: column; gap: 6px; }
.vi-label { font-size: 11px; font-weight: 600; color: var(--ink-2); text-transform: none; margin-top: 4px; }
.vi-layouts { display: grid; grid-template-columns: repeat(5, 1fr); gap: 4px; }
.vi-layout { height: 34px; border: 1px solid var(--line); border-radius: 6px; background: var(--surface-2); display: grid; place-items: center; cursor: pointer; padding: 0; }
.vi-layout:hover { border-color: var(--line-strong); }
.vi-layout.on { border-color: var(--accent); background: var(--selected); }
.vi-script { resize: vertical; min-height: 64px; font: inherit; font-size: 12.5px; line-height: 1.45; }
.vi-note { font-size: 11px; color: var(--ink-2); }
.vi-record { align-self: flex-start; margin-top: 2px; }
.vi-takes { display: flex; flex-direction: column; gap: 2px; margin-top: 2px; }
.vi-take { display: flex; align-items: center; gap: 6px; font-size: 12px; padding: 3px 4px; border-radius: 5px; min-height: 26px; }
.vi-take.on { background: var(--selected); }
.vi-take-len { color: var(--ink-2); font-variant-numeric: tabular-nums; font-size: 11px; }
.vi-take-used { margin-inline-start: auto; font-size: 11px; color: var(--ink-2); }
.vi-take .vi-use { margin-inline-start: auto; padding: 2px 8px; min-height: 0; }
.vi-title, .vi-subtitle { font-size: 12.5px; }
.vi-row { display: flex; align-items: center; gap: 6px; }
.vi-row .vi-label { margin: 0; }
.vi-seconds { width: 64px; }
.vi-unit { font-size: 11.5px; color: var(--ink-2); }
.vi-scene-tools { display: flex; align-items: center; gap: 2px; margin-top: 4px; border-top: 1px solid var(--line-soft); padding-top: 6px; }
.vi-grow { flex: 1; }
.vi-foot { display: flex; flex-wrap: wrap; gap: 6px; padding: 8px; border-top: 1px solid var(--line); flex: none; }
.vi-foot .rw-btn { flex: 1 1 auto; justify-content: center; }

/* A layout, drawn small: the frame, the screen and the camera where they sit. */
.vi-glyph { position: relative; flex: none; display: block; width: 34px; height: 20px; border-radius: 3px; background: #2a2e36; overflow: hidden; }
.vi-glyph i { position: absolute; display: block; }
.vi-glyph.g-ribbon { width: 46px; height: 27px; border-radius: 4px; }
.vi-glyph.g-ribbon.l-screenCamera .cam { width: 11px; height: 11px; }
.vi-glyph.g-ribbon.l-sideBySide .cam { width: 17px; height: 17px; top: 5px; }
.vi-glyph.g-ribbon.l-sideBySide .scr { width: 23px; height: 14px; top: 7px; }
.vi-glyph .scr { background: #6f8fd6; }
.vi-glyph .cam { background: #f2b25c; }
.vi-glyph .ln1, .vi-glyph .ln2 { background: #fff; border-radius: 1px; height: 2px; }
.vi-glyph.l-title .scr, .vi-glyph.l-title .cam, .vi-glyph.l-end .scr, .vi-glyph.l-end .cam { display: none; }
.vi-glyph.l-title .ln1, .vi-glyph.l-end .ln1 { left: 8px; right: 8px; top: 7px; height: 3px; }
.vi-glyph.l-title .ln2, .vi-glyph.l-end .ln2 { left: 11px; right: 11px; top: 12px; background: #ffc107; }
.vi-glyph.l-end .ln1 { left: 6px; right: 6px; }
.vi-glyph.l-talking .scr { display: none; }
.vi-glyph.l-talking .cam { inset: 0; border-radius: 0; }
.vi-glyph.l-talking .ln1 { left: 3px; width: 13px; bottom: 5px; }
.vi-glyph.l-talking .ln2 { left: 3px; width: 9px; bottom: 2px; background: #ffc107; }
.vi-glyph.l-screenCamera .scr { inset: 0; }
.vi-glyph.l-screenCamera .cam { right: 2px; bottom: 2px; width: 8px; height: 8px; border-radius: 50%; }
.vi-glyph.l-screenCamera .ln1, .vi-glyph.l-screenCamera .ln2, .vi-glyph.l-sideBySide .ln1, .vi-glyph.l-sideBySide .ln2 { display: none; }
.vi-glyph.l-sideBySide .cam { left: 2px; top: 4px; width: 12px; height: 12px; border-radius: 2px; }
.vi-glyph.l-sideBySide .scr { right: 2px; top: 5px; width: 17px; height: 10px; }

/* The brand kit. */
.vi-brand { display: grid; grid-template-columns: 1fr 230px; gap: 18px; }
.vi-brand-form { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
.vi-logo-row { display: flex; align-items: center; gap: 6px; }
.vi-logo-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12.5px; color: var(--ink-2); }
.vi-colours { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; }
.vi-colour { display: flex; align-items: center; gap: 7px; padding: 5px 6px; border: 1px solid var(--line); border-radius: 7px; cursor: pointer; min-width: 0; }
.vi-colour input { width: 26px; height: 26px; padding: 0; border: 0; background: none; cursor: pointer; flex: none; }
.vi-colour-text { display: flex; flex-direction: column; min-width: 0; line-height: 1.2; }
.vi-colour-text b { font-size: 12px; font-weight: 600; }
.vi-colour-text small { font-size: 10.5px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.vi-check { display: flex; align-items: center; gap: 7px; font-size: 12.5px; margin-top: 6px; }
.vi-brand-look { display: flex; flex-direction: column; gap: 8px; }
.vi-brand-card { aspect-ratio: 16 / 9; border-radius: 8px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 3px; padding: 10px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.2); overflow: hidden; }
.vi-brand-card b { font-size: 15px; line-height: 1.2; max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vi-brand-card span { font-size: 11px; max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vi-brand-logo { max-width: 36%; max-height: 34%; object-fit: contain; margin-bottom: 3px; }
.vi-brand-scene { position: relative; aspect-ratio: 16 / 9; border-radius: 8px; overflow: hidden; background: linear-gradient(160deg, #5d6b82, #2e3644); }
.vi-brand-face { position: absolute; left: 50%; top: 22%; width: 34%; aspect-ratio: 1; transform: translateX(-50%); border-radius: 50% 50% 46% 46%; background: radial-gradient(circle at 50% 38%, #d9b49a 0 32%, transparent 33%), radial-gradient(ellipse at 50% 100%, #3c4a62 0 58%, transparent 59%); }
.vi-brand-third { position: absolute; left: 6%; bottom: 10%; max-width: 70%; display: flex; flex-direction: column; padding: 4px 8px; border-radius: 3px; line-height: 1.2; }
.vi-brand-third.rtl { left: auto; right: 6%; text-align: right; }
.vi-brand-third b { font-size: 11px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vi-brand-third small { font-size: 9px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.vi-brand-mark-img { position: absolute; right: 4%; top: 6%; width: 12%; object-fit: contain; }

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
