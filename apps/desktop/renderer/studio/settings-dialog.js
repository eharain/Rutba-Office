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
    <Dialog title={title || t('Recorder settings')} className="st-settings-dialog" width={620} onClose={onClose} actions={<Button primary label={t('Done')} className="st-settings-done" onClick={onClose} />}>
      <div className="st-settings">
        <Section title={t('This setup')} note={t('{name}: kept with the setup, and in its file', { name: setupName })}>
          <Row label={t('Size')}>
            <Select value={String(settings.size)} disabled={busy} className="rw-select st-size" aria-label={t('Size')} onChange={(e) => change({ size: Number(e.target.value) })}>
              <option value="1080">{t('1920 × 1080 (Full HD)')}</option>
              <option value="720">{t('1280 × 720 (HD)')}</option>
            </Select>
          </Row>
          <Row label={t('Frame rate')}>
            <Select value={String(settings.fps)} disabled={busy} className="rw-select st-fps" aria-label={t('Frame rate')} onChange={(e) => change({ fps: Number(e.target.value) })}>
              {[24, 30, 60].map((n) => <option key={n} value={String(n)}>{t('{n} frames a second', { n })}</option>)}
            </Select>
          </Row>
          <Row label={t('Quality')}>
            <Seg value={settings.quality} disabled={busy} className="st-quality" label={t('Quality')} options={[['small', t('Smaller files')], ['standard', t('Standard')], ['high', t('Sharper')]]} onChange={(v) => change({ quality: v })} />
          </Row>
          <Row label={t('Recordings are saved in')}>
            <div className="st-settings-row"><span title={settings.folder}>{settings.folder || '…'}</span><Button label={t('Change…')} disabled={busy} onClick={chooseFolder} /></div>
          </Row>
          {apart ? <Row label={t('Each source as its own file')}><Toggle on={settings.isolated} className="st-isolated" label={t('Also keep the screen, camera and sound apart, to edit as layers in Video')} disabled={busy} onChange={(v) => change({ isolated: v })} /></Row> : null}
          {live ? <Row label=""><Toggle on={settings.liveCopy} label={t('Record a copy on this computer while live')} onChange={(v) => change({ liveCopy: v })} /></Row> : null}
        </Section>
        <Section title={t('Every setup')} note={t('The same whichever setup is open')}>
          <Row label={t('Countdown')}>
            <Seg value={settings.countdown} disabled={busy} className="st-countdown" label={t('Countdown')} options={[[0, t('None')], ...[3, 5, 10].map((n) => [n, t('{seconds} s', { seconds: n })])]} onChange={(v) => change({ countdown: v })} />
          </Row>
          <Row label=""><Toggle on={settings.minimise} label={t('Minimise this window while recording')} disabled={busy} onChange={(v) => change({ minimise: v })} /></Row>
          <Row label={t('Take Picture of')}>
            <Seg value={settings.shot === 'program' ? 'program' : 'selected'} className="st-shot-of" label={t('Take Picture of')} options={[['program', t('What is shown')], ['selected', t('The selected source')]]} onChange={(v) => change({ shot: v })} />
          </Row>
          <Row label={t('Picture timer')}>
            <Seg value={settings.shotTimer} className="st-shot-timer" label={t('Picture timer')} options={[[0, t('At once')], ...[3, 5, 10].map((n) => [n, t('After {seconds} s', { seconds: n })])]} onChange={(v) => change({ shotTimer: v })} />
          </Row>
          {pictureKey ? <Row label={t('Take Picture\'s key')}><kbd className="st-set-key">{pictureKey}</kbd></Row> : null}
        </Section>
      </div>
    </Dialog>
  );
}
