// The bottom bar of the Record and Live panels (A2, A3): the microphone with
// its level and mute, the computer's sound, the camera, then whatever the panel
// puts at its right.

import React, { useState } from 'react';
import { Icon, Menu, t } from '@rutba/office-ui';
import { Meter } from './parts.js';

const at = (e) => { const r = e.currentTarget.getBoundingClientRect(); return { x: r.left, y: r.top - 4, above: true }; };

/** @param {{ s: object, children?: any }} p  `s` is the studio; children go on the right */
export function BottomBar({ s, children }) {
  const [menu, setMenu] = useState(null);
  const { named, devices, media } = s;
  const mic = named.sounds.find((x) => x.kind === 'mic') || null;
  const system = named.sounds.find((x) => x.kind === 'system') || null;
  const camera = named.items.find((x) => x.kind === 'camera') || null;
  const micLive = Boolean(mic?.live && media.soundState(mic.id) === 'live' && !mic.muted);
  const cameraOn = Boolean(camera) && !named.cameraOff;
  const open = (e, items) => { const p = at(e); setMenu({ x: p.x, y: Math.max(40, p.y - items.length * 32 - 8), items }); };

  const micItems = () => devices.mics.length
    ? devices.mics.map((m) => ({ label: m.name, checked: (mic?.ref || devices.mics[0]?.id) === m.id, run: () => s.useMic(m) }))
    : [{ label: t('studio.no-microphone-found'), disabled: true }];
  const cameraItems = () => [
    ...(devices.cameras.length ? devices.cameras.map((c) => ({ label: c.name, checked: (camera?.ref || devices.cameras[0]?.id) === c.id, run: () => s.useCamera(c) })) : [{ label: t('common.no-camera-found'), disabled: true }]),
  ];

  return (
    <div className="st-dock" role="group" aria-label={t('studio.bottom-bar.sound-and-camera')}>
      {mic ? (
        <>
          <button type="button" className="st-bb st-mic-pick" aria-haspopup="menu" data-tip={t('studio.bottom-bar.choose-the')} onClick={(e) => open(e, micItems())}>
            <Icon name="mic" size={16} /><span className="st-bb-name">{mic.name || t('common.microphone')}</span>
            <Meter media={media} id={mic.id} live={micLive} />
            <Icon name="chevronDown" size={11} />
          </button>
          <button type="button" className={`st-bb icon st-mic-mute${mic.muted ? ' off' : ''}`} aria-pressed={mic.muted} aria-label={mic.muted ? t('studio.bottom-bar.unmute-the') : t('studio.bottom-bar.mute-the')} data-tip={mic.muted ? t('studio.bottom-bar.unmute-the') : t('studio.bottom-bar.mute-the')} onClick={() => s.act.sound(mic.id, { muted: !mic.muted, live: true })}>
            <Icon name={mic.muted ? 'micOff' : 'mic'} size={16} />
          </button>
        </>
      ) : null}
      {system ? (
        <button type="button" className={`st-bb st-system${system.live ? ' on' : ''}`} aria-pressed={system.live} data-tip={t('studio.bottom-bar.record-sound-computer')} onClick={() => s.act.sound(system.id, { live: !system.live })}>
          <Icon name="speaker" size={16} />{system.live ? t('studio.bottom-bar.computer-sound-on') : t('studio.bottom-bar.computer-sound-off')}
        </button>
      ) : null}
      {camera ? (
        <span className="st-bb-split">
          <button type="button" className={`st-bb st-camera-toggle${cameraOn ? ' on' : ''}`} aria-pressed={cameraOn} data-tip={t('studio.bottom-bar.turn-camera-off')} onClick={s.toggleCamera}>
            <Icon name="camera" size={16} /><span className="st-bb-name">{camera.name || t('common.camera')}: {cameraOn ? t('common.on') : t('common.off')}</span>
          </button>
          <button type="button" className="st-bb icon st-camera-pick" aria-haspopup="menu" aria-label={t('studio.bottom-bar.choose-the-camera')} data-tip={t('studio.bottom-bar.choose-the-camera')} onClick={(e) => open(e, cameraItems())}><Icon name="chevronDown" size={11} /></button>
        </span>
      ) : null}
      <span className="st-dock-gap" />
      {children}
      {menu ? <Menu x={menu.x} y={menu.y} items={menu.items} onClose={() => setMenu(null)} /> : null}
    </div>
  );
}
