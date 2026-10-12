// The Video ribbon: Home, Insert, Edit, Transitions, Clip and View, each
// button a command from the window's table (its label, its key, whether it
// can run now), so a button, its shortcut and the menu can never disagree.

import React from 'react';
import { Ribbon, Group, Button, Rows, Select, prettyKey, t } from '@rutba/office-ui';
import { TITLE_PRESETS } from '@rutba/studio/project';

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
function TitleTile({ preset }) {
  return (
    <span className={`vd-title-tile p-${preset}`} aria-hidden="true">
      <span className="band" />
      <span className="words" />
    </span>
  );
}

export const TITLE_NAMES = () => ({ title: t('Title'), lowerThird: t('Lower Third'), endCard: t('End Card'), callout: t('Callout') });

export default function VideoRibbon({ tab, setTab, cmds, playing, transitionSeconds, setTransitionSeconds, selection }) {
  const tabs = [
    { id: 'home', label: t('Home') },
    { id: 'insert', label: t('Insert') },
    { id: 'edit', label: t('Edit') },
    { id: 'transitions', label: t('Transitions') },
    { id: 'clip', label: t('Clip') },
    { id: 'view', label: t('View') },
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
          <Group label={t('Clipboard')}>
            <C cmds={cmds} id="edit.paste" tall />
            <Rows>
              <C cmds={cmds} id="edit.cut" />
              <C cmds={cmds} id="edit.copy" />
              <C cmds={cmds} id="edit.duplicate" />
            </Rows>
          </Group>
          <Group label={t('Editing')}>
            <C cmds={cmds} id="edit.split" tall />
            <Rows>
              <C cmds={cmds} id="edit.trimStart" />
              <C cmds={cmds} id="edit.trimEnd" />
              <C cmds={cmds} id="edit.delete" />
            </Rows>
          </Group>
          <Group label={t('Playback')}>
            <C cmds={cmds} id="play.toggle" tall icon={playing ? 'pause' : 'play'} label={playing ? t('Pause') : t('Play')} />
            <Rows>
              <C cmds={cmds} id="play.start" />
              <C cmds={cmds} id="play.prevFrame" />
              <C cmds={cmds} id="play.nextFrame" />
            </Rows>
          </Group>
          <Group label={t('Record')}>
            <C cmds={cmds} id="record.screen" tall />
          </Group>
          <Group label={t('Project')}>
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

      {tab === 'insert' ? (
        <>
          <Group label={t('Media')}>
            <C cmds={cmds} id="insert.media" tall />
            <C cmds={cmds} id="record.screen" tall />
          </Group>
          <Group label={t('Titles')}>
            {Object.keys(TITLE_PRESETS).map((p) => {
              const c = cmds[`insert.title.${p}`];
              return (
                <Button key={p} tall className="vd-title-btn" label={names[p]} title={`${names[p]}: ${c.tip}`} disabled={c.disabled} onClick={() => c.run()}>
                  <TitleTile preset={p} />
                </Button>
              );
            })}
          </Group>
          <Group label={t('Markers')}>
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
          <Group label={t('Cut')}>
            <C cmds={cmds} id="edit.split" tall />
            <Rows>
              <C cmds={cmds} id="edit.splitAll" />
              <C cmds={cmds} id="edit.trimStart" />
              <C cmds={cmds} id="edit.trimEnd" />
            </Rows>
          </Group>
          <Group label={t('Ripple')}>
            <C cmds={cmds} id="edit.ripple" tall />
            <Rows>
              <C cmds={cmds} id="edit.rippleDelete" />
              <C cmds={cmds} id="edit.closeGap" />
              <C cmds={cmds} id="edit.delete" />
            </Rows>
          </Group>
          <Group label={t('Link')}>
            <C cmds={cmds} id="edit.detach" tall />
            <Rows>
              <C cmds={cmds} id="edit.link" />
              <C cmds={cmds} id="edit.duplicate" />
            </Rows>
          </Group>
          <Group label={t('In and Out')}>
            <Rows>
              <C cmds={cmds} id="range.in" />
              <C cmds={cmds} id="range.out" />
              <C cmds={cmds} id="range.clear" />
            </Rows>
          </Group>
          <Group label={t('History')}>
            <C cmds={cmds} id="edit.undo" tall />
            <C cmds={cmds} id="edit.redo" tall />
          </Group>
        </>
      ) : null}

      {tab === 'transitions' ? (
        <>
          <Group label={t('Transition to this clip')}>
            {['none', 'fade', 'dip', 'wipe', 'slide', 'push', 'zoom'].map((k) => (
              <C key={k} cmds={cmds} id={`transition.${k}`} tall />
            ))}
          </Group>
          <Group label={t('Timing')}>
            <label className="vd-ribbon-field">
              <span>{t('Duration')}</span>
              <Select value={String(transitionSeconds)} onChange={(e) => setTransitionSeconds(Number(e.target.value))}>
                {[0.25, 0.5, 0.75, 1, 1.5, 2, 3].map((s) => (
                  <option key={s} value={String(s)}>{t('{seconds} s', { seconds: s })}</option>
                ))}
              </Select>
            </label>
          </Group>
        </>
      ) : null}

      {tab === 'clip' ? (
        <>
          <Group label={t('Fit')}>
            <C cmds={cmds} id="clip.fit.cover" tall />
            <C cmds={cmds} id="clip.fit.contain" tall />
            <C cmds={cmds} id="clip.fit.blur" tall />
          </Group>
          <Group label={t('Rotate and flip')}>
            <Rows>
              <C cmds={cmds} id="clip.rotateLeft" />
              <C cmds={cmds} id="clip.rotateRight" />
            </Rows>
            <Rows>
              <C cmds={cmds} id="clip.flipX" />
              <C cmds={cmds} id="clip.flipY" />
            </Rows>
          </Group>
          <Group label={t('Sound')}>
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
          <Group label={t('Timeline')}>
            <C cmds={cmds} id="view.zoomIn" tall />
            <C cmds={cmds} id="view.zoomOut" tall />
            <C cmds={cmds} id="view.fit" tall />
          </Group>
          <Group label={t('Helpers')}>
            <C cmds={cmds} id="view.snap" tall />
            <C cmds={cmds} id="view.follow" tall />
          </Group>
          <Group label={t('Frame')}>
            <C cmds={cmds} id="view.aspect.landscape" tall />
            <C cmds={cmds} id="view.aspect.vertical" tall />
            <C cmds={cmds} id="view.aspect.square" tall />
          </Group>
        </>
      ) : null}
    </Ribbon>
  );
}
