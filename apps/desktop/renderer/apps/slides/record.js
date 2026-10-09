// Record: the microphone as the show goes — each slide's stretch of voice
// kept apart, with how long the slide was up — and Insert → Audio's Record
// Audio box. The voice is kept as a WAV, the one sound every PowerPoint
// plays, made here from the microphone's own samples.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Dialog, t } from '@rutba/office-ui';
import { wavOfPcm } from './sounds.js';
import { Downsampler } from './downsample.js';

const RATE = 22050;


/**
 * The microphone, recording: `mark(slide)` closes the stretch so far (for
 * the slide that was up) and begins one for `slide`; `pause()`/`resume()`
 * hold it; `stop()` answers each slide's stretch — `{ slide, wav, ms }` —
 * its voice as WAV bytes and how long the slide was up.
 */
export class NarrationRecorder {
  static async open() {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true }, video: false });
    return new NarrationRecorder(stream);
  }

  constructor(stream) {
    this.stream = stream;
    this.ctx = new AudioContext();
    this.source = this.ctx.createMediaStreamSource(stream);
    this.node = this.ctx.createScriptProcessor(4096, 1, 1);
    this.voice = new Downsampler(this.ctx.sampleRate, RATE);
    this.paused = false;
    this.parts = [];
    this.current = null;
    this.node.onaudioprocess = (e) => {
      if (this.paused || !this.current) return;
      this.voice.push(e.inputBuffer.getChannelData(0));
    };
    this.source.connect(this.node);
    // A processor runs only while connected onward; nothing reaches the speakers.
    this.mute = this.ctx.createGain();
    this.mute.gain.value = 0;
    this.node.connect(this.mute);
    this.mute.connect(this.ctx.destination);
    this.heldMs = 0;
    this.heldAt = null;
  }

  /** The slide now up: what was recorded for the one before is closed off. */
  mark(slide) {
    const now = performance.now();
    this.close(now);
    this.current = { slide, at: now };
    this.heldMs = 0;
  }

  close(now = performance.now()) {
    if (!this.current) return;
    const held = this.heldMs + (this.heldAt != null ? now - this.heldAt : 0);
    const ms = Math.max(0, now - this.current.at - held);
    this.parts.push({ slide: this.current.slide, samples: this.voice.take(), ms });
    this.current = null;
  }

  pause() { if (!this.paused) { this.paused = true; this.heldAt = performance.now(); } }

  resume() {
    if (!this.paused) return;
    this.paused = false;
    this.heldMs += performance.now() - (this.heldAt ?? performance.now());
    this.heldAt = null;
  }

  async stop() {
    this.close();
    this.node.disconnect();
    this.source.disconnect();
    this.stream.getTracks().forEach((t) => t.stop());
    await this.ctx.close().catch(() => {});
    // A slide shown twice keeps its last stretch, as PowerPoint records over it.
    const last = new Map();
    for (const p of this.parts) last.set(p.slide, p);
    return [...last.values()].map((p) => ({ slide: p.slide, ms: Math.round(p.ms), wav: wavOfPcm(p.samples, RATE) }));
  }
}

const clock = (ms) => `${Math.floor(ms / 60000)}:${String(Math.floor((ms / 1000) % 60)).padStart(2, '0')}`;

/** The show's recording bar: the red dot, the time on this slide and in all, Pause and Stop. */
export function RecordingBar({ startedAt, slideAt, paused, onPause, onStop }) {
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    const t = setInterval(() => setNow(Date.now()), 250);
    return () => clearInterval(t);
  }, []);
  return (
    <div className="sl-recbar" role="status" aria-label={t('Recording')}>
      <span className={`sl-recbar-dot${paused ? ' paused' : ''}`} />
      <span>{paused ? t('Paused') : t('Recording')}</span>
      <b className="sl-recbar-slide">{clock(now - slideAt)}</b>
      <span>{t('this slide ·')}</span>
      <b className="sl-recbar-total">{clock(now - startedAt)}</b>
      <button type="button" className="sl-recbar-btn sl-recbar-pause" onClick={(e) => { e.stopPropagation(); onPause(); }}>{paused ? t('Resume') : t('Pause')}</button>
      <button type="button" className="sl-recbar-btn sl-recbar-stop" onClick={(e) => { e.stopPropagation(); onStop(); }}>{t('Stop')}</button>
    </div>
  );
}

/** Insert → Audio → Record Audio: record, stop, hear it back, and put it on the slide as a sound. */
export function RecordAudioDialog({ onInsert, onClose }) {
  const [state, setState] = useState('idle');
  const [error, setError] = useState(null);
  const [clip, setClip] = useState(null);
  const [started, setStarted] = useState(0);
  const [now, setNow] = useState(0);
  const rec = useRef(null);
  const url = useRef(null);
  const alive = useRef(true);
  const opening = useRef(false);
  useEffect(() => () => {
    alive.current = false;
    rec.current?.stop().catch(() => {});
    if (url.current) URL.revokeObjectURL(url.current);
  }, []);
  useEffect(() => {
    if (state !== 'recording') return undefined;
    const t = setInterval(() => setNow(Date.now()), 250);
    return () => clearInterval(t);
  }, [state]);
  const start = async () => {
    if (opening.current || rec.current) return;
    opening.current = true;
    try {
      setError(null);
      const opened = await NarrationRecorder.open();
      // Closed while the microphone was opening: it is turned off again.
      if (!alive.current) { opened.stop().catch(() => {}); return; }
      rec.current = opened;
      rec.current.mark(0);
      setStarted(Date.now());
      setNow(Date.now());
      setState('recording');
    } catch (err) {
      setError(t('The microphone could not be opened: {reason}', { reason: err.message || err }));
    } finally {
      opening.current = false;
    }
  };
  const stop = async () => {
    const [part] = await rec.current.stop();
    rec.current = null;
    if (url.current) URL.revokeObjectURL(url.current);
    url.current = URL.createObjectURL(new Blob([part.wav], { type: 'audio/wav' }));
    setClip(part);
    setState('done');
  };
  return (
    <Dialog
      title={t('Record Sound')}
      width={380}
      onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('OK')} className="sl-recaudio-ok" disabled={!clip} onClick={() => onInsert(clip)} /></>}
    >
      <div className="sl-recaudio">
        <div className="sl-recaudio-time">{state === 'recording' ? clock(now - started) : clip ? clock(clip.ms) : '0:00'}</div>
        <div className="sl-recaudio-row">
          {state === 'recording'
            ? <Button label={t('Stop')} className="sl-recaudio-stop" onClick={stop} />
            : <Button primary={!clip} label={clip ? t('Record again') : t('Record')} className="sl-recaudio-record" onClick={start} />}
          <Button label={t('Play')} disabled={!clip || state === 'recording'} onClick={() => { if (url.current) new Audio(url.current).play().catch(() => {}); }} />
        </div>
        {error ? <div className="sl-recaudio-error">{error}</div> : null}
      </div>
    </Dialog>
  );
}

export const RECORD_CSS = `
.sl-recbar { position: fixed; left: 50%; top: 16px; transform: translateX(-50%); z-index: 50; display: flex; align-items: center; gap: 8px; padding: 7px 10px 7px 14px; background: rgba(20, 22, 26, 0.9); color: #fff; border-radius: 10px; font: 12.5px/1 var(--font-ui, sans-serif); }
.sl-recbar-dot { width: 10px; height: 10px; border-radius: 50%; background: #e5484d; animation: sl-rec-blink 1.2s infinite; }
.sl-recbar-dot.paused { animation: none; background: #f5a524; }
@keyframes sl-rec-blink { 50% { opacity: 0.35; } }
.sl-recbar b { font-variant-numeric: tabular-nums; }
.sl-recbar-btn { border: 1px solid rgba(255,255,255,0.35); background: transparent; color: #fff; border-radius: 6px; padding: 4px 10px; font: inherit; cursor: pointer; }
.sl-recbar-btn:hover { background: rgba(255,255,255,0.12); }
.sl-recaudio { display: flex; flex-direction: column; gap: 12px; align-items: center; }
.sl-recaudio-time { font: 600 28px/1 var(--font-ui, sans-serif); font-variant-numeric: tabular-nums; }
.sl-recaudio-row { display: flex; gap: 8px; }
.sl-recaudio-error { color: var(--bad); font-size: 12.5px; }
`;
