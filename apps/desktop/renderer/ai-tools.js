// AI Tools: the person's switch for letting AI apps on this computer make and
// read documents with Rutba Office (owner, 2026-10-16, D-2026-10-16-06: "it
// should be off unless some one opens it for single session or want to keep it
// always open"; 2026-10-17, D-2026-10-17-04: "protect it from remote
// connections"). Off, for this session, or always; the folders the tools may
// use; how to connect an AI app; and what the tools have done.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Icon, useToast, formatWhen, t } from '@rutba/office-ui';

const MODES = () => [
  ['off', t('Off'), t('No AI app can use them.')],
  ['session', t('On for this session'), t('Until Rutba Office closes.')],
  ['always', t('Always on'), t('Until you turn them off here.')],
];

/** The lines an AI app's settings take to start the tools (MCP), as most of them write it. */
export function connectJson(connect) {
  if (!connect) return '';
  return JSON.stringify({ mcpServers: { 'rutba-office': { command: connect.command, args: connect.args, env: connect.env } } }, null, 2);
}

export function AiToolsDialog({ shell, onClose }) {
  const toast = useToast();
  const [state, setState] = useState(null);
  const load = useCallback(() => shell.ai.state().then(setState).catch((err) => toast(err.message, { tone: 'bad' })), [shell, toast]);
  useEffect(() => { load(); }, [load]);

  const act = async (fn) => {
    try {
      setState(await fn());
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    }
  };
  /** A folder the person chooses for the tools; null when they cancel. */
  const chooseFolder = async () => {
    const [folder] = (await shell.dialog.open({ title: t('Choose a folder the AI tools may use'), directory: true })) || [];
    return folder || null;
  };
  const addFolder = async () => {
    const folder = await chooseFolder();
    if (folder) act(() => shell.ai.addFolder({ path: folder }));
  };
  /**
   * On, off, or for this session. Turned on with no folder yet, the person
   * chooses one first (owner, 2026-10-17: "the user should have choice to
   * setup different one"); none is made for them, and cancelling leaves the
   * tools off.
   */
  const setMode = async (mode) => {
    if (mode !== 'off' && !state?.folders?.length) {
      const folder = await chooseFolder();
      if (!folder) {
        toast(t('The AI tools stay off until you choose a folder for them.'), { ms: 5000 });
        return;
      }
      try {
        await shell.ai.addFolder({ path: folder });
      } catch (err) {
        toast(err.message, { tone: 'bad', ms: 6000 });
        return;
      }
    }
    act(() => shell.ai.setMode({ mode }));
  };
  const copy = async () => {
    await shell.clipboard.writeText({ text: connectJson(state?.connect) }).catch(() => {});
    toast(t('Copied. Paste it into your AI app’s MCP settings.'), { ms: 4000 });
  };

  return (
    <Dialog title={t('AI Tools')} width={620} className="ai-tools-dialog" onClose={onClose} actions={<Button label={t('Close')} primary onClick={onClose} />}>
      <p className="ai-lead">
        {t('Let an AI app on this computer make and read documents, workbooks, presentations and PDFs with Rutba Office, in the folders you choose. It never overwrites or deletes a file.')}
      </p>
      <p className="ai-local">
        <Icon name="lock" size={13} />
        {' '}
        {t('Only AI apps on this computer can use these tools. Nothing listens on the network, so no other machine can connect.')}
      </p>

      <fieldset className="ai-modes" aria-label={t('AI tools')}>
        {MODES().map(([key, label, note]) => (
          <label key={key} className={`ai-mode${state?.mode === key ? ' on' : ''}`}>
            <input type="radio" name="ai-mode" value={key} checked={state?.mode === key} onChange={() => setMode(key)} />
            <span className="ai-mode-label">{label}</span>
            <span className="ai-mode-note">{note}</span>
          </label>
        ))}
      </fieldset>

      <h3 className="ai-h">{t('Connected now')}</h3>
      {state?.connections?.length ? (
        <ul className="ai-connections">
          {state.connections.map((c) => (
            <li key={c.pid}>
              <span className="ai-conn-dot" aria-hidden="true" />
              <span className="ai-conn-who">{c.client}{c.clientVersion ? ` ${c.clientVersion}` : ''}</span>
              <span className="ai-conn-where">{c.where === 'this computer' ? t('This computer') : c.where}</span>
              <span className="ai-conn-when">{t('Since {when}', { when: formatWhen(c.since) })}</span>
              <span className="ai-conn-last">{c.last ? t('Last: {tool}, {when}', { tool: c.lastTool, when: formatWhen(c.last) }) : t('No call yet')}</span>
            </li>
          ))}
        </ul>
      ) : (
        <p className="ai-none">{state?.mode === 'off' ? t('Nothing can connect while the tools are off.') : t('No AI app is connected.')}</p>
      )}

      <h3 className="ai-h">{t('Folders the tools may use')}</h3>
      {state?.folders?.length ? (
        <ul className="ai-folders">
          {state.folders.map((f) => (
            <li key={f}>
              <Icon name="open" size={13} />
              <span className="ai-folder-path selectable">{f}</span>
              <Button label={t('Remove')} onClick={() => act(() => shell.ai.removeFolder({ path: f }))} />
            </li>
          ))}
        </ul>
      ) : (
        <p className="ai-none">{t('None yet. Turning the tools on asks you to choose one, and you can add more or remove any here.')}</p>
      )}
      <Button icon="plus" label={t('Add Folder…')} className="ai-add-folder" onClick={addFolder} />

      <h3 className="ai-h">{t('Connect an AI app')}</h3>
      <p className="ai-note">{t('Add these lines to the AI app’s list of MCP servers, in its settings. The app starts the tools itself when it needs them.')}</p>
      <pre className="ai-config selectable">{connectJson(state?.connect)}</pre>
      <Button icon="copy" label={t('Copy')} className="ai-copy" onClick={copy} />

      <h3 className="ai-h">{t('What the tools have done')}</h3>
      {state?.log?.length ? (
        <ul className="ai-log">
          {state.log.slice(0, 20).map((e, i) => (
            <li key={`${e.at}-${i}`} className={e.ok ? 'ok' : 'bad'}>
              <span className="ai-log-when">{formatWhen(e.at)}</span>
              <span className="ai-log-tool">{e.tool}</span>
              <span className="ai-log-what">{e.ok ? (e.paths || []).join(', ') : e.error}</span>
            </li>
          ))}
        </ul>
      ) : (
        <p className="ai-none">{t('Nothing yet.')}</p>
      )}
    </Dialog>
  );
}

export const AI_TOOLS_CSS = `
.ai-tools-dialog .ai-lead { margin-top: 0; }
.ai-connections { list-style: none; margin: 0 0 8px; padding: 0; display: grid; gap: 4px; }
.ai-connections li { display: grid; grid-template-columns: 10px minmax(120px, 1.2fr) auto auto 1.4fr; align-items: center; gap: 8px; font-size: 12.5px; }
.ai-conn-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--good, #2fbf8a); }
.ai-conn-who { font-weight: 600; overflow-wrap: anywhere; }
.ai-conn-where, .ai-conn-when, .ai-conn-last { color: var(--ink-3); }
.ai-badge { position: relative; }
.ai-badge.live { color: var(--accent-text); }
.ai-badge-count { position: absolute; top: 1px; right: 0; min-width: 14px; height: 14px; padding: 0 3px; border-radius: 7px; background: var(--good, #2fbf8a); color: #fff; font-size: 10px; line-height: 14px; text-align: center; box-sizing: border-box; }
.ai-local { display: flex; align-items: center; gap: 6px; margin: 0 0 12px; padding: 8px 10px; border-radius: var(--r-2); background: var(--sunken); color: var(--ink-2); font-size: 12.5px; }
.ai-modes { border: 0; margin: 0 0 12px; padding: 0; display: grid; gap: 6px; }
.ai-mode { display: grid; grid-template-columns: auto auto 1fr; align-items: center; gap: 8px; padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--r-2); cursor: pointer; }
.ai-mode.on { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
.ai-mode-label { font-weight: 600; }
.ai-mode-note { color: var(--ink-3); font-size: 12px; }
.ai-h { font-size: 12px; text-transform: uppercase; letter-spacing: .04em; color: var(--ink-3); margin: 16px 0 6px; }
.ai-folders, .ai-log { list-style: none; margin: 0 0 8px; padding: 0; display: grid; gap: 4px; }
.ai-folders li { display: flex; align-items: center; gap: 8px; }
.ai-folder-path { flex: 1; min-width: 0; overflow-wrap: anywhere; font-size: 12.5px; }
.ai-config { margin: 0 0 6px; padding: 8px 10px; border-radius: var(--r-2); background: var(--sunken); font: 11.5px/1.45 ui-monospace, Consolas, monospace; white-space: pre-wrap; overflow-wrap: anywhere; max-height: 160px; overflow: auto; }
.ai-note, .ai-none { color: var(--ink-3); font-size: 12.5px; margin: 0 0 6px; }
.ai-log li { display: grid; grid-template-columns: 120px 150px 1fr; gap: 8px; font-size: 12px; }
.ai-log li.bad .ai-log-tool { color: var(--bad, #c0392b); }
.ai-log-what { overflow-wrap: anywhere; color: var(--ink-2); }
`;

/**
 * The badge every window's title bar shows while the AI tools are on: lit
 * when an AI app is connected, with how many; a click opens AI Tools, which
 * says who is connected, from where, since when and what it last did (owner,
 * 2026-10-17: "the ui should show indicate that and clicking should give a
 * detail who is connected and from"). Nothing is shown while they are off.
 */
export function AiBadge({ shell, onOpen }) {
  const [state, setState] = useState(null);
  useEffect(() => {
    if (!shell?.ai?.state) return undefined;
    let alive = true;
    const read = () => shell.ai.state().then((st) => { if (alive) setState(st); }).catch(() => {});
    read();
    const timer = setInterval(read, 5000);
    return () => { alive = false; clearInterval(timer); };
  }, [shell]);
  if (!state || state.mode === 'off') return null;
  const n = state.connections?.length || 0;
  const tip = n
    ? t('AI tools on: {apps} connected from this computer. Click for the details.', { apps: state.connections.map((c) => c.client).join(', ') })
    : (state.mode === 'session' ? t('AI tools on for this session; no AI app connected. Click to change.') : t('AI tools always on; no AI app connected. Click to change.'));
  return (
    <button type="button" className={`rw-btn ghost ai-badge${n ? ' live' : ''}`} data-tip={tip} aria-label={tip} data-connections={n} onClick={onOpen}>
      <Icon name="wand" size={13} />
      {n ? <span className="ai-badge-count" aria-hidden="true">{n}</span> : null}
    </button>
  );
}
