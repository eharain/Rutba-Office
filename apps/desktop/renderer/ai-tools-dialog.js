// The AI Tools window (ai-tools.js says what it is for): AI Tools' own code,
// in its pack, which the shell loads when the window is opened.

import React, { useCallback, useEffect, useState } from 'react';
import { Dialog, Button, Icon, Input, Select, useToast, formatWhen, t } from '@rutba/office-ui';
import { MODES, REACH, STEP_NAMES, STEP_NOTES, GROUPS, thingName, StepSelect } from './ai-tools.js';
/** The lines an AI app's settings take to start the tools (MCP), as most of them write it. */
export function connectJson(connect) {
  if (!connect) return '';
  return JSON.stringify({ mcpServers: { 'rutba-office': { command: connect.command, args: connect.args, env: connect.env } } }, null, 2);
}

/** Where another machine reaches the tools: each of this computer's addresses, at the port. */
export const remoteUrls = (remote) => (remote ? (remote.addresses?.length ? remote.addresses : ['localhost']).map((a) => `http://${a}:${remote.port}/mcp`) : []);

/** The lines an AI app on another machine takes: the address, and the key it shows. */
export function remoteJson(remote) {
  if (!remote) return '';
  return JSON.stringify({ mcpServers: { 'rutba-office': { type: 'http', url: remoteUrls(remote)[0], headers: { Authorization: `Bearer ${remote.key}` } } } }, null, 2);
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
  const copy = async (text = connectJson(state?.connect)) => {
    await shell.clipboard.writeText({ text }).catch(() => {});
    toast(t('Copied. Paste it into your AI app’s MCP settings.'), { ms: 4000 });
  };
  const [keyShown, setKeyShown] = useState(false);
  const [portText, setPortText] = useState('');
  const remote = state?.remote || null;

  return (
    <Dialog title={t('AI Tools')} width={620} className="ai-tools-dialog" onClose={onClose} actions={<Button label={t('Close')} primary onClick={onClose} />}>
      <p className="ai-lead">
        {t('Let an AI app on this computer make and read documents, workbooks, presentations and PDFs with Rutba Office, in the folders you choose, and use the mail, calendars, contacts and databases you allow. It never overwrites or deletes a file.')}
      </p>
      <p className={`ai-local${remote ? ' open' : ''}`}>
        <Icon name={remote ? 'globe' : 'lock'} size={13} />
        {' '}
        {remote
          ? t('AI apps on this computer, and on other machines that have the key below, can use these tools while they are on.')
          : t('Only AI apps on this computer can use these tools. Nothing listens on the network, so no other machine can connect.')}
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

      <h3 className="ai-h">{t('Who can connect')}</h3>
      <fieldset className="ai-modes ai-reach" aria-label={t('Who can connect')}>
        {REACH().map(([key, label, note]) => (
          <label key={key} className={`ai-mode${state?.reach === key ? ' on' : ''}`}>
            <input type="radio" name="ai-reach" value={key} checked={state?.reach === key} onChange={() => act(() => shell.ai.setReach({ reach: key }))} />
            <span className="ai-mode-label">{label}</span>
            <span className="ai-mode-note">{note}</span>
          </label>
        ))}
      </fieldset>
      {remote ? (
        <div className="ai-remote">
          <p className="ai-warn">
            <Icon name="shield" size={13} />
            {' '}
            {t('The connection is not encrypted. Use it on a network you trust, or reach it through a secure tunnel. Anyone with the key can use the tools while they are on.')}
          </p>
          <div className="ai-remote-row">
            <span className="ai-remote-label">{t('Address')}</span>
            <span className="ai-urls selectable">{remoteUrls(remote).map((u) => <span key={u}>{u}</span>)}</span>
          </div>
          <div className="ai-remote-row">
            <span className="ai-remote-label">{t('Key')}</span>
            <code className="ai-key selectable">{keyShown ? remote.key : '\u2022'.repeat(16)}</code>
            <Button label={keyShown ? t('Hide') : t('Show')} className="ai-key-show" onClick={() => setKeyShown(!keyShown)} />
            <Button icon="copy" label={t('Copy')} className="ai-key-copy" onClick={() => copy(remote.key)} />
            <Button label={t('New Key')} className="ai-key-new" title={t('A new key: the old one, and every connection made with it, stop working at once.')} onClick={() => act(() => shell.ai.renewKey())} />
          </div>
          <div className="ai-remote-row">
            <span className="ai-remote-label">{t('Port')}</span>
            <Input className="ai-port" type="number" min={1} max={65535} aria-label={t('Port')} value={portText || String(remote.port)} onChange={(e) => setPortText(e.target.value)} />
            <Button label={t('Use')} className="ai-port-use" disabled={!portText || Number(portText) === remote.port} onClick={() => act(() => shell.ai.setReach({ reach: 'remote', port: Number(portText) })).then(() => setPortText(''))} />
          </div>
          <p className={`ai-note ai-listening${remote.error ? ' bad' : ''}`}>
            {remote.error || (state.mode === 'off' ? t('Nothing listens until the tools are turned on.') : remote.listening ? t('Listening on port {port}.', { port: remote.port }) : '')}
          </p>
        </div>
      ) : null}

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

      <h3 className="ai-h">{t('What the tools may reach')}</h3>
      <p className="ai-note">{t('Each mail account, calendar, address book, database and folder has its own setting, also shown in its own app. Everything is Off until you choose. Nothing is ever deleted or moved, and anything sent waits for your press here first.')}</p>
      <dl className="ai-step-key">
        {Object.entries(STEP_NAMES()).map(([k, name]) => (
          <React.Fragment key={k}><dt>{name}</dt><dd>{STEP_NOTES()[k]}</dd></React.Fragment>
        ))}
      </dl>
      {GROUPS().map(([kinds, title, icon]) => {
        const rows = (state?.things || []).filter((x) => kinds.includes(x.kind));
        const isFolders = kinds[0] === 'folder';
        if (!rows.length && !isFolders) return null;
        return (
          <section key={kinds[0]} className={`ai-group${isFolders ? ' ai-folders' : ''}`} data-ai-group={kinds[0]}>
            <h4 className="ai-group-h"><Icon name={icon} size={13} /> {title}</h4>
            {rows.length ? (
              <ul className="ai-things">
                {rows.map((x) => (
                  <li key={`${x.kind}:${x.id}`} className="ai-thing" data-ai-kind={x.kind} data-ai-id={x.id}>
                    <span className={`ai-thing-name${x.kind === 'folder' ? ' ai-folder-path selectable' : ''}`}>
                      {thingName(x)}
                      {x.kind === 'archive' ? <small>{` ${t('archive, read only')}`}</small> : null}
                      {x.kind === 'mail' && x.local ? <small>{` ${t('brought in from a file')}`}</small> : null}
                      {x.server ? <small>{` ${t('kept in step with a server')}`}</small> : null}
                      {x.kind === 'database' && x.role ? <small>{` ${t('works as {role}', { role: x.role })}`}</small> : null}
                      {x.kind === 'database' && x.missing ? <small>{` ${t('not found')}`}</small> : null}
                    </span>
                    <StepSelect thing={x} onSet={(step) => act(async () => { await shell.ai.setStep({ kind: x.kind, id: x.id, step }); return shell.ai.state(); })} />
                    {x.kind === 'folder' ? <Button label={t('Remove')} onClick={() => act(() => shell.ai.removeFolder({ path: x.id }))} /> : null}
                  </li>
                ))}
              </ul>
            ) : (
              <p className="ai-none">{t('None yet. Turning the tools on asks you to choose one, and you can add more or remove any here.')}</p>
            )}
          </section>
        );
      })}
      <Button icon="plus" label={t('Add Folder…')} className="ai-add-folder" onClick={addFolder} />
      <p className="ai-note">{t('A protected database is reached only as a role an Admin chose for the AI tools in Data Files, never one that may change its design.')}</p>

      <h3 className="ai-h">{t('Connect an AI app')}</h3>
      <p className="ai-note">{t('Add these lines to the AI app’s list of MCP servers, in its settings. The app starts the tools itself when it needs them.')}</p>
      <pre className="ai-config selectable">{connectJson(state?.connect)}</pre>
      <Button icon="copy" label={t('Copy')} className="ai-copy" onClick={() => copy()} />
      <p className="ai-note ai-copy-again">{t('Lines copied from Rutba Office before this version name a file the AI app cannot find. Copy these and put them in its settings in place of the old ones. They open the tools in this profile.')}</p>
      {remote ? (
        <>
          <p className="ai-note">{t('On another machine, an AI app that connects to an MCP server at an address takes these lines, the key with them.')}</p>
          <pre className="ai-config ai-config-remote selectable">{keyShown ? remoteJson(remote) : remoteJson({ ...remote, key: '\u2022'.repeat(16) })}</pre>
          <Button icon="copy" label={t('Copy')} className="ai-copy-remote" onClick={() => copy(remoteJson(remote))} />
        </>
      ) : null}

      <h3 className="ai-h">{t('What the tools have done')}</h3>
      {state?.log?.length ? (
        <ul className="ai-log">
          {state.log.slice(0, 20).map((e, i) => (
            <li key={`${e.at}-${i}`} className={e.ok ? 'ok' : 'bad'}>
              <span className="ai-log-when">{formatWhen(e.at)}</span>
              <span className="ai-log-tool">{e.tool}{e.where ? <small className="ai-log-where">{` ${t('from {where}', { where: e.where })}`}</small> : null}</span>
              <span className="ai-log-what">{e.thing ? `${e.thing}, ${STEP_NAMES()[e.step] || ''}${e.ok ? '' : `: ${e.error}`}` : e.ok ? (e.paths || []).join(', ') : e.error}</span>
            </li>
          ))}
        </ul>
      ) : (
        <p className="ai-none">{t('Nothing yet.')}</p>
      )}
    </Dialog>
  );
}
