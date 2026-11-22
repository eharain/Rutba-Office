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
import { NetCameraFindDialog } from './net-camera-find.js';
import { NET_CAMERA_CSS } from './net-camera-style.js';

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
 * The dialogs for a studio window: `ask()` to add a camera by its address,
 * `ask(item)` for the details of the one a source shows, `find()` to choose
 * one of those that answer on the network. Adding puts the source in the scene.
 * @param {{ shell: object, changeStack: Function, setSelected?: (id: string) => void }} p
 * @returns {{ ask: (item?: object|null) => void, find: () => void, dialog: React.ReactNode }}
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
    find: () => setAsked({ find: true }),
    dialog: !asked ? null : asked.find
      ? <NetCameraFindDialog shell={shell} onDone={done} onByAddress={() => setAsked({ item: null })} onClose={() => setAsked(null)} />
      : <NetCameraDialog shell={shell} editId={asked.item ? asked.item.ref : null} onDone={done} onClose={() => setAsked(null)} />,
  };
}
