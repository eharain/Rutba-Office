// The Video ribbon: Home, Insert, Edit, Transitions, Clip and View, each
// button a command from the window's table (its label, its key, whether it
// can run now), so a button, its shortcut and the menu can never disagree.

import React from 'react';
import { Ribbon, Group, Button, Rows, Select, prettyKey, t } from '@rutba/office-ui';
import { TITLE_PRESETS } from '@rutba/studio/project';
import { LAYOUTS } from '@rutba/studio/intro';
import { LayoutGlyph } from './intro/scenes.js';

/** A command as a ribbon button: its label, its tip beginning with the label and naming the key. */
function C({ cmds, id, tall, label, icon, pressed }) {
  const c = cmds[id];
  if (!c) return null;
  const key = c.key || c.shortcut;
  const tip = `${c.label}${key ? ` (${prettyKey(key)})` : ''}${c.tip ? `: ${c.tip}` : ''}`;
  return (
    <Button
      tall={tall}
      icon={icon ?? c.icon}
      label={label ?? c.label}
      title={tip}
      pressed={pressed ?? c.pressed}
      disabled={c.disabled}
      onClick={() => c.run()}
    />
  );
}

/** A title preset drawn as a small frame, so the gallery shows what it adds. */
export function TitleTile({ preset }) {
  return (
    <span className={`vd-title-tile p-${preset}`} aria-hidden="true">
      <span className="band" />
      <span className="words" />
    </span>
  );
}

export const TITLE_NAMES = () => ({ title: t('common.title'), lowerThird: t('video.ribbon.lower-third'), endCard: t('video.ribbon.end-card'), callout: t('video.ribbon.callout') });

export default function VideoRibbon({ tab, setTab, cmds, playing, transitionSeconds, setTransitionSeconds, selection, intro = false }) {
  const tabs = [
    { id: 'home', label: t('common.home') },
    ...(intro ? [{ id: 'intro', label: t('video.ribbon.introduction'), contextual: true }] : []),
    { id: 'insert', label: t('common.insert') },
    { id: 'edit', label: t('common.edit') },
    { id: 'transitions', label: t('common.transitions') },
    { id: 'clip', label: t('common.clip') },
    { id: 'view', label: t('common.view') },
  ];
  const names = TITLE_NAMES();
  return (
    <Ribbon
      tabs={tabs}
      active={tab}
      onTab={setTab}
      quick={
        <>
          <C cmds={cmds} id="file.save" label="" />
          <C cmds={cmds} id="edit.undo" label="" />
          <C cmds={cmds} id="edit.redo" label="" />
          <C cmds={cmds} id="play.toggle" label="" icon={playing ? 'pause' : 'play'} />
        </>
      }
    >
      {tab === 'home' ? (
        <>
          <Group label={t('common.clipboard')}>
            <C cmds={cmds} id="edit.paste" tall />
            <Rows>
              <C cmds={cmds} id="edit.cut" />
              <C cmds={cmds} id="edit.copy" />
              <C cmds={cmds} id="edit.duplicate" />
            </Rows>
          </Group>
          <Group label={t('common.editing')}>
            <C cmds={cmds} id="edit.split" tall />
            <Rows>
              <C cmds={cmds} id="edit.trimStart" />
              <C cmds={cmds} id="edit.trimEnd" />
              <C cmds={cmds} id="edit.delete" />
            </Rows>
          </Group>
          <Group label={t('video.playback')}>
            <C cmds={cmds} id="play.toggle" tall icon={playing ? 'pause' : 'play'} label={playing ? t('common.pause') : t('common.play')} />
            <Rows>
              <C cmds={cmds} id="play.start" />
              <C cmds={cmds} id="play.prevFrame" />
              <C cmds={cmds} id="play.nextFrame" />
            </Rows>
          </Group>
          <Group label={t('common.record')}>
            <C cmds={cmds} id="record.screen" tall />
            <C cmds={cmds} id="record.voice" tall />
            <C cmds={cmds} id="intro.new" tall />
          </Group>
          <Group label={t('video.project')}>
            <C cmds={cmds} id="insert.media" tall />
            <C cmds={cmds} id="file.export" tall />
            <Rows>
              <C cmds={cmds} id="file.open" />
              <C cmds={cmds} id="file.save" />
              <C cmds={cmds} id="file.saveAs" />
            </Rows>
          </Group>
        </>
      ) : null}

      {tab === 'intro' && intro ? (
        <>
          <Group label={t('video.scene')}>
            <C cmds={cmds} id="intro.record" tall />
            <Rows>
              <C cmds={cmds} id="intro.up" />
              <C cmds={cmds} id="intro.down" />
              <C cmds={cmds} id="intro.remove" />
            </Rows>
          </Group>
          <Group label={t('common.layout')}>
            {LAYOUTS.map((l) => {
              const c = cmds[`intro.layout.${l}`];
              return (
                <Button key={l} tall className="vi-layout-btn" label={c.label} title={`${c.label}: ${c.tip}`} pressed={c.pressed} disabled={c.disabled} onClick={() => c.run()}>
                  <LayoutGlyph layout={l} size="ribbon" />
                </Button>
              );
            })}
          </Group>
          <Group label={t('video.ribbon.brand')}>
            <C cmds={cmds} id="intro.brand" tall />
          </Group>
          <Group label={t('common.finish')}>
            <C cmds={cmds} id="captions.import" tall />
            <C cmds={cmds} id="file.export" tall />
          </Group>
        </>
      ) : null}

      {tab === 'insert' ? (
        <>
          <Group label={t('common.media')}>
            <C cmds={cmds} id="insert.media" tall />
            <Rows>
              <C cmds={cmds} id="insert.mediaLayer" />
              <C cmds={cmds} id="insert.mediaHere" />
            </Rows>
            <C cmds={cmds} id="record.screen" tall />
            <C cmds={cmds} id="record.voice" tall />
          </Group>
          <Group label={t('video.titles')}>
            {Object.keys(TITLE_PRESETS).map((p) => {
              const c = cmds[`insert.title.${p}`];
              return (
                <Button key={p} tall className="vd-title-btn" label={names[p]} title={`${names[p]}: ${c.tip}`} disabled={c.disabled} onClick={() => c.run()}>
                  <TitleTile preset={p} />
                </Button>
              );
            })}
          </Group>
          <Group label={t('common.captions')}>
            <C cmds={cmds} id="captions.import" tall />
            <Rows>
              <C cmds={cmds} id="captions.export" />
            </Rows>
          </Group>
          <Group label={t('video.markers')}>
            <C cmds={cmds} id="marker.add" tall />
            <Rows>
              <C cmds={cmds} id="marker.prev" />
              <C cmds={cmds} id="marker.next" />
            </Rows>
          </Group>
        </>
      ) : null}

      {tab === 'edit' ? (
        <>
          <Group label={t('common.cut')}>
            <C cmds={cmds} id="edit.split" tall />
            <Rows>
              <C cmds={cmds} id="edit.splitAll" />
              <C cmds={cmds} id="edit.trimStart" />
              <C cmds={cmds} id="edit.trimEnd" />
            </Rows>
            <C cmds={cmds} id="clip.freeze" tall />
            <C cmds={cmds} id="clip.reverse" tall />
          </Group>
          <Group label={t('video.ripple')}>
            <C cmds={cmds} id="edit.ripple" tall />
            <Rows>
              <C cmds={cmds} id="edit.rippleDelete" />
              <C cmds={cmds} id="edit.closeGap" />
              <C cmds={cmds} id="edit.delete" />
            </Rows>
          </Group>
          <Group label={t('common.link')}>
            <C cmds={cmds} id="edit.detach" tall />
            <Rows>
              <C cmds={cmds} id="edit.link" />
              <C cmds={cmds} id="edit.duplicate" />
            </Rows>
          </Group>
          <Group label={t('video.ribbon.in-and-out')}>
            <Rows>
              <C cmds={cmds} id="range.in" />
              <C cmds={cmds} id="range.out" />
              <C cmds={cmds} id="range.clear" />
            </Rows>
          </Group>
          <Group label={t('video.ribbon.tracks')}>
            <C cmds={cmds} id="lane.split" tall />
            <Rows>
              <C cmds={cmds} id="lane.insert" />
              <C cmds={cmds} id="lane.up" />
              <C cmds={cmds} id="lane.down" />
            </Rows>
          </Group>
          <Group label={t('video.ribbon.history')}>
            <C cmds={cmds} id="edit.undo" tall />
            <C cmds={cmds} id="edit.redo" tall />
          </Group>
        </>
      ) : null}

      {tab === 'transitions' ? (
        <>
          <Group label={t('video.ribbon.transition')}>
            {['none', 'fade', 'dip', 'wipe', 'slide', 'push', 'zoom'].map((k) => (
              <C key={k} cmds={cmds} id={`transition.${k}`} tall />
            ))}
          </Group>
          <Group label={t('common.timing')}>
            <label className="vd-ribbon-field">
              <span>{t('common.duration')}</span>
              <Select value={String(transitionSeconds)} onChange={(e) => setTransitionSeconds(Number(e.target.value))}>
                {[0.25, 0.5, 0.75, 1, 1.5, 2, 3].map((s) => (
                  <option key={s} value={String(s)}>{t('common.seconds-s', { seconds: s })}</option>
                ))}
              </Select>
            </label>
          </Group>
        </>
      ) : null}

      {tab === 'clip' ? (
        <>
          <Group label={t('common.media')}>
            <C cmds={cmds} id="clip.replace" tall />
          </Group>
          <Group label={t('common.fit')}>
            <C cmds={cmds} id="clip.fit.cover" tall />
            <C cmds={cmds} id="clip.fit.contain" tall />
            <C cmds={cmds} id="clip.fit.blur" tall />
          </Group>
          <Group label={t('common.recording')}>
            <C cmds={cmds} id="clip.autoZoom" tall />
          </Group>
          <Group label={t('common.rotate-and')}>
            <Rows>
              <C cmds={cmds} id="clip.rotateLeft" />
              <C cmds={cmds} id="clip.rotateRight" />
            </Rows>
            <Rows>
              <C cmds={cmds} id="clip.flipX" />
              <C cmds={cmds} id="clip.flipY" />
            </Rows>
          </Group>
          <Group label={t('common.sound')}>
            <C cmds={cmds} id="clip.mute" tall />
            <Rows>
              <C cmds={cmds} id="clip.fadeIn" />
              <C cmds={cmds} id="clip.fadeOut" />
              <C cmds={cmds} id="edit.detach" />
            </Rows>
          </Group>
        </>
      ) : null}

      {tab === 'view' ? (
        <>
          <Group label={t('video.ribbon.timeline')}>
            <C cmds={cmds} id="view.zoomIn" tall />
            <C cmds={cmds} id="view.zoomOut" tall />
            <C cmds={cmds} id="view.fit" tall />
          </Group>
          <Group label={t('common.panels')}>
            <C cmds={cmds} id="view.library" tall />
            <C cmds={cmds} id="view.properties" tall />
            <C cmds={cmds} id="view.monitor" tall />
          </Group>
          <Group label={t('video.ribbon.helpers')}>
            <C cmds={cmds} id="view.snap" tall />
            <C cmds={cmds} id="view.follow" tall />
          </Group>
          <Group label={t('common.frame')}>
            <C cmds={cmds} id="view.aspect.landscape" tall />
            <C cmds={cmds} id="view.aspect.vertical" tall />
            <C cmds={cmds} id="view.aspect.square" tall />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );
}
