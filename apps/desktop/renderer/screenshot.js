// Insert → Screenshot, shared by Documents and Presentations: Office's
// gallery of Available Windows — every window open but this one — and the
// screens, each as a thumbnail; the one picked is taken at its own size and
// handed back as a PNG for the app to insert like any picture.

import React, { useEffect, useState } from 'react';
import { Button, Dialog, Spinner, t } from '@rutba/office-ui';

/**
 * @param {{ shell, onPick, onClose, record? }} props
 *   onPick({ bytes, width, height, name }) — the picture, taken; or, with
 *   `record` (Screen Recording), onPick(source) — the window or screen to record
 */
export function ScreenshotDialog({ shell, onPick, onClose, record = false }) {
  const [sources, setSources] = useState(null);
  const [error, setError] = useState(null);
  const [taking, setTaking] = useState(null);

  useEffect(() => {
    let live = true;
    Promise.resolve(shell.capture?.sources())
      .then((list) => { if (live) setSources(list || []); })
      .catch((err) => { if (live) setError(err.message || String(err)); });
    return () => { live = false; };
  }, [shell]);

  const take = async (source) => {
    if (taking) return;
    setTaking(source.id);
    try {
      await onPick(record ? source : await shell.capture.grab({ id: source.id }));
    } catch (err) {
      setError(err.message || String(err));
      setTaking(null);
    }
  };

  const group = (kind, title) => {
    const list = (sources || []).filter((s) => s.kind === kind);
    if (!list.length) return null;
    return (
      <section className="ss-group">
        <div className="ss-title">{title}</div>
        <div className="ss-grid">
          {list.map((s) => (
            <button key={s.id} type="button" className={`ss-source${taking === s.id ? ' busy' : ''}`} data-source={s.id} title={s.name} disabled={Boolean(taking)} onClick={() => take(s)}>
              <img src={s.thumbnail} alt="" draggable={false} />
              <span>{s.name}</span>
            </button>
          ))}
        </div>
      </section>
    );
  };

  return (
    <Dialog title={record ? t('Screen Recording') : t('Screenshot')} width={620} onClose={onClose} actions={<Button label={t('Cancel')} onClick={onClose} />}>
      <style>{SCREENSHOT_CSS}</style>
      <div className="ss">
        {error ? <div className="ss-error">{error}</div> : null}
        {sources == null && !error ? <div className="ss-wait"><Spinner /> {t('Looking at what is open…')}</div> : null}
        {sources && !sources.length ? <div className="ss-wait">{t('No other window is open to take a picture of.')}</div> : null}
        {group('window', t('Available Windows'))}
        {group('screen', t('Screens'))}
      </div>
    </Dialog>
  );
}

const SCREENSHOT_CSS = `
.ss { display: flex; flex-direction: column; gap: 10px; max-height: 60vh; overflow: auto; }
.ss-title { font-size: 11px; text-transform: uppercase; letter-spacing: .06em; color: var(--ink-3); font-weight: 600; margin: 0 0 6px; }
.ss-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(170px, 1fr)); gap: 10px; }
.ss-source { display: flex; flex-direction: column; gap: 5px; border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); padding: 6px; cursor: pointer; font: inherit; font-size: 12px; color: var(--ink); text-align: left; }
.ss-source:hover:not(:disabled) { border-color: var(--accent-line); background: var(--selected); }
.ss-source:disabled { cursor: default; opacity: .6; }
.ss-source.busy { opacity: 1; border-color: var(--accent); }
.ss-source img { width: 100%; height: 100px; object-fit: contain; background: var(--sunken); border-radius: var(--r-1); }
.ss-source span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ss-wait { display: flex; align-items: center; gap: 8px; color: var(--ink-2); padding: 12px 2px; }
.ss-error { color: var(--bad); padding: 4px 2px; }
`;
