// Talk's look: plain, roomy, nothing that is not needed.

export const TALK_CSS = `
.tk-wait { height: 100%; display: grid; place-items: center; }
.tk-avatar { position: relative; flex: none; display: grid; place-items: center; border-radius: 50%; background: var(--accent-fill); color: #fff; font-weight: 650; }
.tk-dot { position: absolute; right: -1px; bottom: -1px; width: 30%; height: 30%; min-width: 8px; min-height: 8px; border-radius: 50%; border: 2px solid var(--window); background: var(--ink-4, #9aa0a6); }
.tk-dot.online { background: #2bb673; }
.tk-dot.busy { background: #e5484d; }
.tk-dot.away { background: #f5a524; }
.tk-who { flex: 1; min-width: 0; display: grid; text-align: start; }
.tk-who strong { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.tk-who small { color: var(--ink-3); font-size: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.tk-none { color: var(--ink-3); font-size: 13px; line-height: 1.5; margin: 8px 12px; }

.tk-people { display: grid; gap: 4px; padding: 6px 8px; }
.tk-section h3 { margin: 10px 6px 6px; font-size: 11.5px; font-weight: 650; text-transform: uppercase; letter-spacing: 0.04em; color: var(--ink-3); }
.tk-section-head { display: flex; align-items: center; justify-content: space-between; }
.tk-contact, .tk-request { display: flex; align-items: center; gap: 10px; width: 100%; padding: 8px; border-radius: var(--r-2); border: 0; background: transparent; color: var(--ink); font: inherit; cursor: pointer; }
.tk-contact:hover { background: var(--wash); }
.tk-contact.on { background: var(--wash-strong); box-shadow: inset 3px 0 0 var(--accent); }
.tk-request { cursor: default; background: var(--sunken); }
.tk-waiting { display: grid; gap: 2px; margin: 6px 12px; color: var(--ink-3); }
.tk-foot { display: flex; justify-content: space-between; padding: 8px; border-top: 1px solid var(--line-soft); margin-top: 8px; }

.tk-conversation { height: 100%; display: flex; flex-direction: column; }
.tk-head { display: flex; align-items: center; gap: 12px; padding: 12px 18px; border-bottom: 1px solid var(--line-soft); }
.tk-grow { flex: 1; }
.tk-messages { flex: 1; min-height: 0; overflow: auto; padding: 18px; display: flex; flex-direction: column; gap: 6px; }
.tk-bubble { max-width: 70%; align-self: flex-start; padding: 8px 12px; border-radius: 14px; background: var(--sunken); color: var(--ink); line-height: 1.45; display: grid; gap: 2px; }
.tk-bubble.mine { align-self: flex-end; background: var(--accent-fill); color: #fff; }
.tk-bubble time { font-size: 10.5px; opacity: 0.7; justify-self: end; }
.tk-text { white-space: pre-wrap; overflow-wrap: anywhere; }
.tk-compose { display: flex; align-items: flex-end; gap: 8px; padding: 12px 18px; border-top: 1px solid var(--line-soft); }
.tk-input { flex: 1; resize: none; min-height: 38px; max-height: 140px; padding: 9px 12px; border-radius: 19px; border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; }
.tk-input:focus { outline: 2px solid var(--accent); outline-offset: 0; }
.tk-empty { height: 100%; display: grid; place-items: center; }

.tk-callview { height: 100%; display: flex; flex-direction: column; background: #0e1116; color: #fff; }
.tk-stage { position: relative; flex: 1; min-height: 0; display: grid; place-items: center; }
.tk-remote { width: 100%; height: 100%; object-fit: contain; background: #000; }
.tk-remote-still { display: grid; justify-items: center; gap: 12px; font-size: 18px; }
.tk-hidden { position: absolute; width: 1px; height: 1px; opacity: 0; pointer-events: none; }
.tk-self { position: absolute; right: 16px; bottom: 16px; width: 22%; min-width: 160px; aspect-ratio: 16 / 9; object-fit: cover; border-radius: 10px; background: #222; box-shadow: 0 6px 20px rgba(0,0,0,0.5); transform: scaleX(-1); }
.tk-status { position: absolute; left: 16px; top: 14px; font-size: 13px; padding: 4px 10px; border-radius: 12px; background: rgba(0,0,0,0.45); font-variant-numeric: tabular-nums; }
.tk-controls { display: flex; justify-content: center; gap: 18px; padding: 16px; }
.tk-round { width: 52px; height: 52px; border-radius: 50%; border: 0; display: grid; place-items: center; background: rgba(255,255,255,0.14); color: #fff; cursor: pointer; }
.tk-round:hover { background: rgba(255,255,255,0.24); }
.tk-round.off { background: #fff; color: #111; }
.tk-round.on { background: var(--accent-fill); }
.tk-round.end { background: #e5484d; transform: rotate(135deg); }
.tk-code { margin: 0 0 14px; text-align: center; font-size: 12px; opacity: 0.75; font-variant-numeric: tabular-nums; }

.tk-ring { position: fixed; right: 20px; bottom: 44px; z-index: 70; display: flex; align-items: center; gap: 12px; padding: 14px 16px; border-radius: 14px; background: var(--surface); color: var(--ink); box-shadow: 0 12px 32px rgba(0,0,0,0.3); border: 1px solid var(--line-soft); min-width: 360px; }

.tk-signin { max-width: 520px; margin: 40px auto; display: grid; gap: 18px; }
.tk-soon { color: var(--ink-3); font-size: 12.5px; }
.tk-error { color: #c0392b; }
.tk-test { border: 1px solid var(--line-soft); border-radius: var(--r-2); padding: 10px 14px; }
.tk-test summary { cursor: pointer; color: var(--ink-2); }
.tk-field { display: grid; gap: 4px; margin: 10px 0; font-size: 12.5px; color: var(--ink-2); }
.tk-lead { margin: 0 0 10px; color: var(--ink-2); font-size: 13px; line-height: 1.5; }
.tk-relay-list { display: grid; gap: 6px; }
.tk-relay { display: flex; align-items: center; gap: 10px; padding: 8px 10px; border-radius: var(--r-2); border: 1px solid var(--line-soft); cursor: pointer; }
.tk-relay span { display: grid; }
.tk-relay small { color: var(--ink-3); font-size: 12px; }

.tk-asked { display: flex; align-items: center; gap: 4px; }
.tk-asked small { flex: 1; min-width: 0; display: grid; }
.tk-asked strong { font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.tk-told { font-size: 11.5px; }
.tk-invite-actions { display: grid; justify-items: start; padding: 0 8px; }
.tk-how { display: grid; gap: 6px; margin: 12px 0 0; padding: 0; border: 0; }
.tk-how legend { padding: 0; margin-bottom: 4px; font-size: 12.5px; color: var(--ink-2); }
.tk-how .tk-relay:has(input:disabled) { opacity: 0.6; cursor: default; }
.tk-link-make { display: flex; gap: 8px; margin-bottom: 12px; }
.tk-link-make .tk-link-note { flex: 1; }
.tk-link { display: grid; gap: 6px; padding: 10px; margin-top: 8px; border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.tk-link-row { display: flex; gap: 8px; }
.tk-link-url { flex: 1; min-width: 0; font-family: var(--mono, monospace); font-size: 12px; }
.tk-link-info { color: var(--ink-3); font-size: 12px; }
.tk-share { display: flex; flex-wrap: wrap; gap: 4px; }
.tk-qr { width: 168px; height: 168px; justify-self: start; border-radius: var(--r-2); }
.tk-invitation { position: fixed; z-index: 70; top: 56px; left: 50%; transform: translateX(-50%); display: flex; align-items: center; gap: 12px; min-width: 380px; max-width: calc(100% - 32px); padding: 12px 14px; border-radius: 14px; background: var(--surface); color: var(--ink); border: 1px solid var(--line-soft); box-shadow: 0 12px 32px rgba(0,0,0,0.3); }
.tk-own { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; margin: 4px 0 0 30px; }
.tk-own .tk-own-urls { grid-column: 1 / -1; }
.tk-statusbar { color: var(--ink-3); }

.tk-ie-to { display: flex; flex-wrap: wrap; align-items: center; gap: 4px; padding: 4px; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); }
.tk-ie-to .tk-ie-typed { flex: 1; min-width: 160px; border: 0; background: transparent; }
.tk-ie-chip { display: inline-flex; align-items: center; gap: 4px; padding: 2px 4px 2px 10px; border-radius: 12px; background: var(--wash-strong); color: var(--ink); font-size: 12.5px; max-width: 100%; }
.tk-ie-chip button { border: 0; background: transparent; color: var(--ink-3); cursor: pointer; font-size: 14px; line-height: 1; padding: 0 4px; }
.tk-ie-contacts { margin: 6px 0 10px; }
.tk-ie-contacts summary { cursor: pointer; color: var(--ink-2); font-size: 12.5px; }
.tk-ie-contacts .tk-ie-search { width: 100%; margin: 6px 0; }
.tk-ie-list { display: grid; gap: 2px; max-height: 180px; overflow: auto; }
.tk-ie-pick { display: flex; align-items: center; gap: 10px; padding: 6px 8px; border-radius: var(--r-2); cursor: pointer; }
.tk-ie-pick:hover { background: var(--wash); }
.tk-ie-pick .tk-connect-now, .tk-ie-chip .tk-connect-now { flex: none; }
.tk-ie-chip .tk-connect-now { min-height: 0; padding: 0 8px; font-size: 12px; }
.tk-ie-pick[data-reach="contact"], .tk-ie-pick[data-reach="asked"] { opacity: 0.7; cursor: default; }
.tk-foot { flex-wrap: wrap; gap: 4px; }
.tk-addr-list { display: grid; gap: 6px; }
.tk-addr-row { display: flex; align-items: center; gap: 10px; padding: 8px 10px; border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.tk-addr-shown { flex: 1; min-width: 0; display: flex; align-items: center; gap: 10px; cursor: pointer; }
.tk-addr-find { display: flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink-2); white-space: nowrap; }
.tk-addr-add { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-top: 12px; }
.tk-addr-add .tk-addr-new { flex: 1; min-width: 200px; }
.tk-addr-add .tk-addr-code { width: 120px; letter-spacing: 0.2em; }
.tk-addr-later, .tk-addr-sender { margin: 10px 0 0; }
.tk-ie-rutba { color: var(--ink-3); line-height: 1.45; }
.tk-ie-note { width: 100%; resize: vertical; font: inherit; line-height: 1.45; padding: 8px 10px; }
.tk-ie-results { list-style: none; margin: 0; padding: 0; display: grid; gap: 6px; }
.tk-ie-result { display: grid; padding: 6px 8px; border-radius: var(--r-2); background: var(--sunken); }
.tk-ie-result small { color: var(--ink-3); }
`;
