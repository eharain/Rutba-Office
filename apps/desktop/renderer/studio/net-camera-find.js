// Add › Cameras on the network: Office asks the local network for cameras
// that speak ONVIF, lists those that answer by name, and reads the stream
// address of the one chosen from the camera itself. The person gives only the
// camera's sign-in (owner, 22 Nov: "camera on (onvif, rtsp etc)").
//
// The asking and the reading are the main process's (main/net-cameras.js).
// This window is handed a list to choose from and never an address to send
// anything to: it names a found camera by a token of that list.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Dialog, Field, Input, t } from '@rutba/office-ui';
import { NET_CAMERA_CSS } from './net-camera-style.js';
import { netCameraTrouble } from './net-camera-words.js';

const refusal = (reason) => {
  if (reason === 'no-keystore') return t('studio.net-camera-dialog.computer-no-keystore');
  if (reason === 'gone') return t('studio.net-camera-find.camera-no-longer');
  if (reason === 'unreachable' || reason === 'timeout' || reason === 'closed') return t('studio.net-camera-find.camera-stopped-answering');
  if (reason === 'protocol') return t('studio.net-camera-find.camera-did-not');
  return netCameraTrouble(reason);
};

/**
 * @param {object} p
 * @param {object} p.shell
 * @param {(camera: { id: string, name: string }) => void} p.onDone  a camera was added, or one added before was chosen
 * @param {() => void} p.onByAddress  the person would rather type the address
 * @param {() => void} p.onClose
 */
export function NetCameraFindDialog({ shell, onDone, onByAddress, onClose }) {
  const [found, setFound] = useState(null);
  const [picked, setPicked] = useState(null);
  const [user, setUser] = useState('');
  const [password, setPassword] = useState('');
  const [name, setName] = useState('');
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const alive = useRef(true);
  const round = useRef(0);

  const look = () => {
    round.current += 1;
    const mine = round.current;
    setFound(null);
    setPicked(null);
    setError(null);
    shell.capture.cameraFind().then((r) => {
      if (alive.current && round.current === mine) setFound(r?.cameras || []);
    }, () => {
      if (alive.current && round.current === mine) setFound([]);
    });
  };
  useEffect(() => {
    alive.current = true;
    look();
    return () => { alive.current = false; };
  }, [shell]);

  const titleOf = (c) => c.name || c.hardware || c.host;
  const pick = (c) => {
    if (c.known) {
      onDone({ id: c.known, name: c.knownName || titleOf(c) });
      return;
    }
    setPicked(c);
    setName(titleOf(c));
    setError(null);
  };
  const add = async () => {
    if (!picked || busy) return;
    setBusy(true);
    setError(null);
    try {
      const r = await shell.capture.cameraAdopt({ key: picked.key, user: user.trim(), name: name.trim(), ...(password ? { password } : {}) });
      if (!alive.current) return;
      if (r?.ok) onDone({ id: r.camera.id, name: r.camera.name });
      else setError(refusal(r?.reason));
    } catch {
      if (alive.current) setError(refusal('unreachable'));
    }
    if (alive.current) setBusy(false);
  };
  const enter = (e) => { if (e.key === 'Enter') add(); };

  return (
    <Dialog title={t('studio.net-camera-find.cameras-network')} className="st-web-dialog st-cam-dialog st-cam-find" onClose={onClose} actions={(
      <>
        <Button label={t('common.cancel')} onClick={onClose} />
        {picked ? <Button primary label={t('common.add')} className="st-cam-ok" disabled={busy} onClick={add} /> : null}
      </>
    )}>
      <style>{NET_CAMERA_CSS + NET_CAMERA_FIND_CSS}</style>
      {found === null ? <p className="st-cam-note st-cam-looking" role="status">{t('studio.net-camera-find.looking-cameras-network')}</p> : null}
      {found && !found.length ? (
        <p className="st-cam-note st-cam-none" role="status">{t('studio.net-camera-find.no-camera-answered')}</p>
      ) : null}
      {found && found.length ? (
        <div className="st-cam-known" role="group" aria-label={t('studio.net-camera-find.cameras-that')}>
          {found.map((c) => (
            <button key={c.key} type="button" className={`st-cam-pick${picked?.key === c.key ? ' is-picked' : ''}`} data-host={c.host} aria-pressed={picked?.key === c.key} onClick={() => pick(c)}>
              <b>{titleOf(c)}</b>
              <small>{[c.name && c.hardware ? c.hardware : '', c.host, c.known ? t('studio.net-camera-find.added-before') : ''].filter(Boolean).join(' · ')}</small>
            </button>
          ))}
        </div>
      ) : null}
      {picked ? (
        <>
          <p className="st-cam-note">{t('studio.net-camera-find.give-user-name')}</p>
          <div className="st-cam-pair">
            <Field label={t('common.user-name')}>
              <Input className="st-web-url st-cam-user" value={user} autoFocus autoComplete="off" onChange={(e) => setUser(e.target.value)} onKeyDown={enter} aria-label={t('common.user-name')} />
            </Field>
            <Field label={t('common.password')}>
              <Input className="st-web-url st-cam-pass" type="password" value={password} autoComplete="new-password" onChange={(e) => setPassword(e.target.value)} onKeyDown={enter} aria-label={t('common.password')} />
            </Field>
          </div>
          <Field label={t('common.name')}>
            <Input className="st-web-url st-cam-name" value={name} maxLength={80} onChange={(e) => setName(e.target.value)} onKeyDown={enter} aria-label={t('common.name')} />
          </Field>
        </>
      ) : null}
      {error ? <p className="st-cam-err" role="alert">{error}</p> : null}
      {found ? (
        <div className="st-cam-more">
          <Button ghost label={t('studio.net-camera-find.look-again')} className="st-cam-again" disabled={busy} onClick={look} />
          <Button ghost label={t('studio.net-camera-find.add-by-address')} className="st-cam-by-address" disabled={busy} onClick={onByAddress} />
        </div>
      ) : null}
      {picked ? <p className="st-cam-note">{t('studio.net-camera-dialog.password-kept-computers')}</p> : null}
    </Dialog>
  );
}

const NET_CAMERA_FIND_CSS = `
.st-cam-find .st-cam-pick.is-picked { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
.st-cam-more { display: flex; gap: 8px; margin: 10px 0 0; }
`;
