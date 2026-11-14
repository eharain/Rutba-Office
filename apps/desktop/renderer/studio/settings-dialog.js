// The studio's settings, behind the gear on the first line, split by what they
// belong to (A6). This setup: what is made and where it goes, kept with the
// setup and in its file. Every setup: the countdown and Take Picture, the same
// whichever setup is open. Back soon and Auto are in Set up; My streams opens
// from the Live panel.

import React from 'react';
import { Dialog, Select, Button, t } from '@rutba/office-ui';
import { Toggle } from './toggle.js';
import { Seg } from './parts.js';

const Row = ({ label, children }) => <div className="st-set-row"><span className="st-set-label">{label}</span><div className="st-set-value">{children}</div></div>;

const Section = ({ title, note, children }) => (
  <section className="st-set-section" aria-label={title}>
    <div className="st-set-head"><h2>{title}</h2><span>{note}</span></div>
    {children}
  </section>
);

/**
 * @param {object} p
 * @param {object} p.settings  normaliseStudio's
 * @param {string} p.setupName  the setup open now
 * @param {boolean} p.busy  recording or live: what is made cannot change
 * @param {string} [p.pictureKey]  the key that takes a picture
 * @param {boolean} [p.live]  Live Broadcast's: the copy kept while live
 * @param {boolean} [p.apart]  the Recorder on its own: each source as its own file can be asked for
 */
export function StudioSettings({ settings, setupName, busy, pictureKey = '', live = false, apart = false, title = null, change, chooseFolder, onClose }) {
  return (
    <Dialog title={title || t('studio.settings-dialog.recorder-settings')} className="st-settings-dialog" width={620} onClose={onClose} actions={<Button primary label={t('common.done')} className="st-settings-done" onClick={onClose} />}>
      <div className="st-settings">
        <Section title={t('studio.settings-dialog.this-setup')} note={t('studio.settings-dialog.name-kept-setup', { name: setupName })}>
          <Row label={t('common.size')}>
            <Select value={String(settings.size)} disabled={busy} className="rw-select st-size" aria-label={t('common.size')} onChange={(e) => change({ size: Number(e.target.value) })}>
              <option value="1080">{t('studio.settings-dialog.1920-1080-full')}</option>
              <option value="720">{t('studio.settings-dialog.1280-720-hd')}</option>
            </Select>
          </Row>
          <Row label={t('studio.settings-dialog.frame-rate')}>
            <Select value={String(settings.fps)} disabled={busy} className="rw-select st-fps" aria-label={t('studio.settings-dialog.frame-rate')} onChange={(e) => change({ fps: Number(e.target.value) })}>
              {[24, 30, 60].map((n) => <option key={n} value={String(n)}>{t('studio.n-frames-second', { n })}</option>)}
            </Select>
          </Row>
          <Row label={t('common.quality')}>
            <Seg value={settings.quality} disabled={busy} className="st-quality" label={t('common.quality')} options={[['small', t('studio.smaller-files')], ['standard', t('studio.standard')], ['high', t('studio.sharper')]]} onChange={(v) => change({ quality: v })} />
          </Row>
          <Row label={t('studio.settings-dialog.recordings-saved')}>
            <div className="st-settings-row"><span title={settings.folder}>{settings.folder || '…'}</span><Button label={t('common.change-2')} disabled={busy} onClick={chooseFolder} /></div>
          </Row>
          {apart ? <Row label={t('studio.settings-dialog.each-source-own')}><Toggle on={settings.isolated} className="st-isolated" label={t('studio.settings-dialog.also-keep-screen')} disabled={busy} onChange={(v) => change({ isolated: v })} /></Row> : null}
          {live ? <Row label=""><Toggle on={settings.liveCopy} label={t('studio.settings-dialog.record-copy-computer')} onChange={(v) => change({ liveCopy: v })} /></Row> : null}
        </Section>
        <Section title={t('studio.settings-dialog.every-setup')} note={t('studio.settings-dialog.same-whichever-setup')}>
          <Row label={t('studio.settings-dialog.countdown')}>
            <Seg value={settings.countdown} disabled={busy} className="st-countdown" label={t('studio.settings-dialog.countdown')} options={[[0, t('common.none')], ...[3, 5, 10].map((n) => [n, t('common.seconds-s', { seconds: n })])]} onChange={(v) => change({ countdown: v })} />
          </Row>
          <Row label=""><Toggle on={settings.minimise} label={t('studio.settings-dialog.minimise-window-while')} disabled={busy} onChange={(v) => change({ minimise: v })} /></Row>
          <Row label={t('studio.settings-dialog.take-picture-of')}>
            <Seg value={settings.shot === 'program' ? 'program' : 'selected'} className="st-shot-of" label={t('studio.settings-dialog.take-picture-of')} options={[['program', t('studio.settings-dialog.what-is-shown')], ['selected', t('studio.settings-dialog.the-selected-source')]]} onChange={(v) => change({ shot: v })} />
          </Row>
          <Row label={t('studio.settings-dialog.picture-timer')}>
            <Seg value={settings.shotTimer} className="st-shot-timer" label={t('studio.settings-dialog.picture-timer')} options={[[0, t('studio.settings-dialog.at-once')], ...[3, 5, 10].map((n) => [n, t('studio.settings-dialog.after-seconds-s', { seconds: n })])]} onChange={(v) => change({ shotTimer: v })} />
          </Row>
          {pictureKey ? <Row label={t('studio.settings-dialog.take-pictures-key')}><kbd className="st-set-key">{pictureKey}</kbd></Row> : null}
        </Section>
      </div>
    </Dialog>
  );
}
