// The sound's level while the preview plays, beside the playhead's time
// (the 2026-10-16 plan: "meters while playing"). It reads the mix's peak
// from the preview (@rutba/studio startAudioPreview's level) a frame at a
// time, holds a peak for a moment, and is red near full scale.

import React, { useEffect, useRef, useState } from 'react';
import { t } from '@rutba/office-ui';

const FLOOR = -60;
const widthOf = (db) => (Number.isFinite(db) ? Math.max(0, Math.min(1, (db - FLOOR) / -FLOOR)) : 0);

/** @param {{ playback: { current: { audio?: { level?: () => number } } | null }, playing: boolean }} p */
export function PlayMeter({ playback, playing }) {
  const [db, setDb] = useState(-Infinity);
  const peak = useRef({ db: -Infinity, at: 0 });
  useEffect(() => {
    if (!playing) { setDb(-Infinity); peak.current = { db: -Infinity, at: 0 }; return undefined; }
    let raf = 0;
    const tick = () => {
      const now = playback.current?.audio?.level?.() ?? -Infinity;
      const t0 = performance.now();
      if (now > peak.current.db || t0 - peak.current.at > 900) peak.current = { db: now, at: t0 };
      setDb(now);
      raf = requestAnimationFrame(tick);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [playing, playback]);
  const hot = db > -1;
  const shown = Number.isFinite(db) ? Math.round(db) : null;
  return (
    <span className={`vd-meter${hot ? ' hot' : ''}`} data-db={shown ?? ''} title={shown == null ? t('Sound level') : t('Sound level {db} dB', { db: shown })} aria-hidden="true">
      <i style={{ width: `${widthOf(db) * 100}%` }} />
      <b style={{ left: `${widthOf(peak.current.db) * 100}%` }} />
    </span>
  );
}

export const METER_CSS = `
.vd-meter { position: relative; width: 70px; height: 8px; border-radius: 4px; background: #2a2e36; overflow: hidden; flex: none; margin: 0 6px; }
.vd-meter i { position: absolute; left: 0; top: 0; bottom: 0; background: linear-gradient(90deg, #2bb673 0%, #2bb673 70%, #f5a524 88%, #e5484d 100%); transition: width 60ms linear; }
.vd-meter b { position: absolute; top: 0; bottom: 0; width: 2px; background: #fff; opacity: 0.8; }
.vd-meter.hot { box-shadow: 0 0 0 1px #e5484d; }
`;
