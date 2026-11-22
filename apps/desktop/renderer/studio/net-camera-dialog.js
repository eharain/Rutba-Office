// Add › Camera by address, and a network camera's details afterwards (owner,
// 22 Nov: "in the sources we need screens, camera and camera on (onvif, rtsp
// etc)").
//
// The person types the camera's RTSP address, its user name and its password.
// The main process keeps them (main/net-cameras.js): the password in the
// operating system's keystore, never handed back here, so this dialog can
// say only that one is kept and take a new one. A source in a scene names its
// camera by an id; no address and no password is written into a setup.
//
// Cameras added before are offered again, since the same camera is wanted in
// more than one scene.

import React, { useEffect, useState } from 'react';
import { Button, Dialog, Field, Input, t } from '@rutba/office-ui';
import { addItem, updateItem } from '@rutba/studio/stack';

const refusal = (reason) => (reason === 'no-keystore'
  ? t('studio.net-camera-dialog.computer-no-keystore')
  : t('studio.net-camera-dialog.not-camera-address'));

/**
 * @param {object} p
 * @param {object} p.shell
 * @param {string|null} p.editId  a camera's id to show its details; null to add one
 * @param {(camera: { id: string, name: string }) => void} p.onDone  added (or chosen again), or saved
 * @param {() => void} p.onClose
 */
export function NetCameraDialog({ shell, editId = null, onDone, onClose }) {
  const [cameras, setCameras] = useState(null);
  const [url, setUrl] = useState('');
  const [user, setUser] = useState('');
  const [password, setPassword] = useState('');
  const [name, setName] = useState('');
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const known = editId && cameras ? cameras.find((c) => c.id === editId) || null : null;

  useEffect(() => {
    let gone = false;
    shell.capture.cameraList().then((list) => {
      if (gone) return;
      setCameras(list);
      const was = editId ? list.find((c) => c.id === editId) : null;
      if (was) {
        setUrl(was.url);
        setUser(was.user);
        setName(was.name);
      }
    }, () => { if (!gone) setCameras([]); });
    return () => { gone = true; };
  }, [shell, editId]);

  const ok = Boolean(url.trim()) && !busy && cameras !== null;
  const save = async () => {
    if (!ok) return;
    setBusy(true);
    setError(null);
    try {
      const r = await shell.capture.cameraSave({ ...(editId ? { id: editId } : {}), url: url.trim(), user: user.trim(), name: name.trim(), ...(password ? { password } : {}) });
      if (r?.ok) onDone({ id: r.camera.id, name: r.camera.name });
      else setError(refusal(r?.reason));
    } catch {
      setError(refusal('address'));
    }
    setBusy(false);
  };
  const forget = async (c) => {
    try {
      setCameras(await shell.capture.cameraRemove({ id: c.id }));
    } catch {
      // Left as it was.
    }
  };
  const enter = (e) => { if (e.key === 'Enter') save(); };
  const others = !editId && cameras ? cameras : [];

  return (
    <Dialog title={editId ? t('studio.net-camera-dialog.camera-details') : t('studio.net-camera-dialog.add-camera-address')} className="st-web-dialog st-cam-dialog" onClose={onClose} actions={(
      <>
        <Button label={t('common.cancel')} onClick={onClose} />
        <Button primary label={editId ? t('common.save') : t('common.add')} className="st-cam-ok" disabled={!ok} onClick={save} />
      </>
    )}>
      <style>{NET_CAMERA_CSS}</style>
      {others.length ? (
        <div className="st-cam-known" role="group" aria-label={t('studio.net-camera-dialog.cameras-added-before')}>
          <h4>{t('studio.net-camera-dialog.cameras-added-before')}</h4>
          {others.map((c) => (
            <div key={c.id} className="st-cam-row" data-id={c.id}>
              <button type="button" className="st-cam-pick" onClick={() => onDone({ id: c.id, name: c.name })}>
                <b>{c.name}</b>
                <small>{c.url}</small>
              </button>
              <Button ghost label={t('studio.net-camera-dialog.forget')} title={t('studio.net-camera-dialog.forget-camera-password')} className="st-cam-forget" onClick={() => forget(c)} />
            </div>
          ))}
        </div>
      ) : null}
      {editId && cameras && !known ? <p className="st-cam-note">{t('studio.net-camera-dialog.camera-set-another')}</p> : null}
      <Field label={t('common.address')} hint={t('studio.net-camera-dialog.address-starts-rtsp')}>
        {/* words-ok: how a camera address starts */}
        <Input className="st-web-url st-cam-url" value={url} placeholder="rtsp://" autoFocus={!others.length} onChange={(e) => setUrl(e.target.value)} onKeyDown={enter} aria-label={t('common.address')} />
      </Field>
      <div className="st-cam-pair">
        <Field label={t('common.user-name')}>
          <Input className="st-web-url st-cam-user" value={user} autoComplete="off" onChange={(e) => setUser(e.target.value)} onKeyDown={enter} aria-label={t('common.user-name')} />
        </Field>
        <Field label={t('common.password')}>
          <Input className="st-web-url st-cam-pass" type="password" value={password} autoComplete="new-password" placeholder={known?.hasPassword ? t('studio.net-camera-dialog.kept-type-new') : ''} onChange={(e) => setPassword(e.target.value)} onKeyDown={enter} aria-label={t('common.password')} />
        </Field>
      </div>
      <Field label={t('common.name')}>
        <Input className="st-web-url st-cam-name" value={name} maxLength={80} onChange={(e) => setName(e.target.value)} onKeyDown={enter} aria-label={t('common.name')} />
      </Field>
      {error ? <p className="st-cam-err" role="alert">{error}</p> : null}
      <p className="st-cam-note">{t('studio.net-camera-dialog.password-kept-computers')}</p>
    </Dialog>
  );
}

/**
 * The dialog for a studio window: `ask()` to add a camera, `ask(item)` for
 * the details of the one a source shows. Adding puts the source in the scene.
 * @param {{ shell: object, changeStack: Function, setSelected?: (id: string) => void }} p
 * @returns {{ ask: (item?: object|null) => void, dialog: React.ReactNode }}
 */
export function useNetCamera({ shell, changeStack, setSelected = null }) {
  const [asked, setAsked] = useState(null);
  const done = (camera) => {
    const item = asked?.item || null;
    setAsked(null);
    if (item) {
      changeStack((s) => updateItem(s, item.id, { name: camera.name }));
      return;
    }
    let id = null;
    changeStack((s) => {
      const r = addItem(s, { kind: 'netcam', ref: camera.id, name: camera.name });
      id = r.id;
      return r.stack;
    });
    if (id) setSelected?.(id);
  };
  return {
    ask: (item = null) => setAsked({ item: item?.kind === 'netcam' ? item : null }),
    dialog: asked ? <NetCameraDialog shell={shell} editId={asked.item ? asked.item.ref : null} onDone={done} onClose={() => setAsked(null)} /> : null,
  };
}

export const NET_CAMERA_CSS = `
.st-cam-dialog { width: min(520px, 92vw); }
.st-cam-pair { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.st-cam-known { margin: 0 0 12px; display: flex; flex-direction: column; gap: 6px; }
.st-cam-known h4 { margin: 0; font-size: 12px; font-weight: 600; color: var(--ink-2); }
.st-cam-row { display: flex; align-items: center; gap: 8px; }
.st-cam-pick { flex: 1; min-width: 0; display: flex; flex-direction: column; align-items: flex-start; gap: 1px; padding: 7px 10px; border: 1px solid var(--line); border-radius: 7px; background: var(--surface); color: var(--ink); font: inherit; text-align: start; cursor: pointer; }
.st-cam-pick:hover { border-color: var(--accent); }
.st-cam-pick b { font-size: 13px; max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-cam-pick small { font-size: 11.5px; color: var(--ink-3); max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; direction: ltr; }
.st-cam-url { direction: ltr; }
.st-cam-note { margin: 10px 0 0; font-size: 12px; line-height: 1.45; color: var(--ink-3); }
.st-cam-err { margin: 10px 0 0; font-size: 12.5px; color: var(--danger, #b3261e); }
`;
