// The studio's settings, behind the gear on the first line: what is made
// (size, frame rate, quality), where it goes, the countdown, Take Picture's
// subject and timer, the placeholder picture, words and music, and, in Live
// Broadcast, Go Live's destinations.

import React from 'react';
import { Dialog, Field, Select, Input, Button, basename, t } from '@rutba/office-ui';
import { Toggle } from './toggle.js';

/**
 * @param {object} p
 * @param {object} p.settings  normaliseStudio's
 * @param {object} p.stack
 * @param {boolean} p.busy  recording or live: what is made cannot change
 * @param {boolean} [p.live]  Live Broadcast's: the copy kept while live, and the destinations
 */
export function StudioSettings({ settings, stack, busy, live = false, title = null, change, changeStack, chooseFolder, choosePlaceholder, chooseMusic, openDestinations, onClose }) {
  return (
    <Dialog title={title || t('Recorder settings')} className="st-settings-dialog" width={560} onClose={onClose} actions={<Button primary label={t('Done')} className="st-settings-done" onClick={onClose} />}>
      <div className="st-settings">
        <Field label={t('Size')}>
          <Select value={String(settings.size)} disabled={busy} className="st-size" onChange={(e) => change({ size: Number(e.target.value) })}>
            <option value="1080">1080p (1920 × 1080)</option>
            <option value="720">720p (1280 × 720)</option>
          </Select>
        </Field>
        <Field label={t('Frame rate')}>
          <Select value={String(settings.fps)} disabled={busy} onChange={(e) => change({ fps: Number(e.target.value) })}>
            {[24, 30, 60].map((n) => <option key={n} value={String(n)}>{t('{n} frames a second', { n })}</option>)}
          </Select>
        </Field>
        <Field label={t('Quality')}>
          <Select value={settings.quality} disabled={busy} onChange={(e) => change({ quality: e.target.value })}>
            <option value="small">{t('Smaller files')}</option>
            <option value="standard">{t('Standard')}</option>
            <option value="high">{t('Sharper')}</option>
          </Select>
        </Field>
        <Field label={t('Countdown')}>
          <Select value={String(settings.countdown)} disabled={busy} onChange={(e) => change({ countdown: Number(e.target.value) })}>
            <option value="0">{t('None')}</option>
            {[3, 5, 10].map((n) => <option key={n} value={String(n)}>{t('{seconds} s', { seconds: n })}</option>)}
          </Select>
        </Field>
        <div className="wide">
          <Field label={t('Recordings are saved in')}>
            <div className="st-settings-row"><span title={settings.folder}>{settings.folder || '…'}</span><Button label={t('Change…')} disabled={busy} onClick={chooseFolder} /></div>
          </Field>
        </div>
        <div className="wide"><Toggle on={settings.minimise} label={t('Minimise this window while recording')} disabled={busy} onChange={(v) => change({ minimise: v })} /></div>
        <Field label={t('Take Picture of')}>
          <Select value={settings.shot === 'program' ? 'program' : 'selected'} className="st-shot-of" onChange={(e) => change({ shot: e.target.value })}>
            <option value="program">{t('What is shown')}</option>
            <option value="selected">{t('The selected source')}</option>
          </Select>
        </Field>
        <Field label={t('Picture timer')}>
          <Select value={String(settings.shotTimer)} className="st-shot-timer" onChange={(e) => change({ shotTimer: Number(e.target.value) })}>
            <option value="0">{t('At once')}</option>
            {[3, 5, 10].map((n) => <option key={n} value={String(n)}>{t('After {seconds} s', { seconds: n })}</option>)}
          </Select>
        </Field>
        <div className="wide">
          <Field label={t('Placeholder, shown when Video is on Placeholder')}>
            <div className="st-settings-row">
              <span>{stack.placeholder ? basename(stack.placeholder) : t('A card with words')}</span>
              <Button label={t('Picture…')} onClick={choosePlaceholder} />
              {stack.placeholder ? <Button label={t('Use words')} onClick={() => changeStack((s) => ({ ...s, placeholder: '' }))} /> : null}
            </div>
          </Field>
        </div>
        {!stack.placeholder ? (
          <div className="wide">
            <Field label={t('The card says')}>
              <Input className="st-placeholder-text" value={stack.placeholderText} placeholder={t('Back soon')} onChange={(e) => changeStack((s) => ({ ...s, placeholderText: e.target.value }))} />
            </Field>
          </div>
        ) : null}
        <div className="wide">
          <Field label={t('Music, played when Sound is on Music')}>
            <div className="st-settings-row">
              <span>{stack.music ? basename(stack.music) : t('A quiet chord, made here')}</span>
              <Button label={t('Music…')} onClick={chooseMusic} />
              {stack.music ? <Button label={t('Use the chord')} onClick={() => changeStack((s) => ({ ...s, music: '' }))} /> : null}
            </div>
          </Field>
        </div>
        {live ? (
          <>
            <div className="wide"><Toggle on={settings.liveCopy} label={t('Record a copy on this computer while live')} onChange={(v) => change({ liveCopy: v })} /></div>
            <div className="wide"><Button icon="globe" label={t('Live Destinations…')} className="st-dests-btn" onClick={openDestinations} /></div>
          </>
        ) : null}
      </div>
    </Dialog>
  );
}
