// The strip under On air for a web page source (owner, 5 Nov: "user able to flip
// next and previous product from web CMS group as the display live").
//
//   Previous · 3 of 12 · Blue lawn kurta · Next
//
// The strip speaks the store's show-page contract (consumer's show-deck.ts):
// messages, never keys, so the page's Space never meets Live Broadcast's Space
// for Take. Page Up and Page Down do the same, so a presentation clicker works.
// A page that does not report its state is sent the arrow keys instead, and the
// strip then has only Previous and Next. A flip is on air at once: the page is
// one page, shared by Next and On air, and the strip says so.
//
// Also the box Set up › Add › Web page opens.

import React, { useEffect, useState } from 'react';
import { Button, Dialog, Field, Icon, Input, t } from '@rutba/office-ui';

/** What the strip offers for a page's state: pure, so a check can read it. */
export function flipOffer(state) {
  if (!state || state.status === 'off') return { show: false };
  if (state.status === 'unavailable') return { show: true, unavailable: true, canPrevious: false, canNext: false };
  const reported = Boolean(state.reported);
  const count = reported ? state.count : 0;
  const index = reported ? state.index : 0;
  const empty = reported && count === 0;
  return {
    show: true,
    unavailable: false,
    reported,
    empty,
    index,
    count,
    title: reported ? state.title : '',
    canPrevious: !empty && (!reported || state.loop || index > 0),
    canNext: !empty && (!reported || state.loop || index < count - 1),
  };
}

function Strip({ item, web, primary }) {
  const offer = flipOffer(web.stateOf(item));
  const flip = (verb) => web.flip(item, verb);
  const sound = web.stateOf(item)?.sound;
  useEffect(() => {
    if (!primary) return undefined;
    const onKey = (e) => {
      if (e.key !== 'PageDown' && e.key !== 'PageUp') return;
      const el = e.target;
      if (el?.tagName === 'INPUT' || el?.tagName === 'TEXTAREA' || el?.isContentEditable) return;
      e.preventDefault();
      if (e.key === 'PageDown' && offer.canNext) web.flip(item, 'next');
      else if (e.key === 'PageUp' && offer.canPrevious) web.flip(item, 'previous');
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [primary, offer.canNext, offer.canPrevious, web, item]);
  if (!offer.show) return null;
  return (
    <div className="st-flip" role="group" aria-label={t('studio.flip-strip.flip-the-page')} data-id={item.id} data-status={offer.unavailable ? 'unavailable' : 'ready'} data-reported={offer.reported ? '1' : '0'}>
      <button type="button" className="st-flip-btn st-flip-prev" disabled={!offer.canPrevious} title={t('common.previous')} onClick={() => flip('previous')}>
        <Icon name="chevronLeft" size={14} />{t('common.previous')} <kbd>PgUp</kbd> {/* words-ok: a key's name */}
      </button>
      <div className="st-flip-now" aria-live="polite">
        {offer.unavailable ? <b className="wn">{t('studio.not-available')}</b>
          : offer.reported ? (
            <>
              <small>{offer.empty ? t('studio.flip-strip.nothing-show-group') : t('studio.flip-strip.n-of-count', { n: offer.index + 1, count: offer.count })}</small>
              <b>{offer.title}</b>
            </>
          ) : null}
      </div>
      <button type="button" className="st-flip-btn st-flip-next" disabled={!offer.canNext} title={t('common.next')} onClick={() => flip('next')}>
        {t('common.next')} <kbd>PgDn</kbd> {/* words-ok: a key's name */}<Icon name="chevronRight" size={14} />
      </button>
      <button type="button" className="st-flip-sound" aria-pressed={Boolean(sound)} title={sound ? t('studio.flip-strip.turn-pages-sound') : t('studio.flip-strip.turn-pages')} onClick={() => web.sound(item, !sound)}>
        <Icon name="speaker" size={16} />
      </button>
    </div>
  );
}

/**
 * The strip for every web page on air.
 * @param {object} p
 * @param {object} p.stack  the stack on air
 * @param {import('./web-source.js').WebSources} p.web  media.web
 */
export function FlipStrip({ stack, web }) {
  const pages = stack.items.filter((it) => it.on && it.kind === 'web' && it.live);
  if (!pages.length) return null;
  return (
    <div className="st-flips">
      {pages.map((it, i) => <Strip key={it.id} item={it} web={web} primary={i === 0} />)}
      <span className="st-flip-note">{t('studio.flip-strip.flip-goes-air')}</span>
    </div>
  );
}

const nameFrom = (url) => {
  try {
    const u = new URL(url);
    return `${u.hostname}${u.pathname === '/' ? '' : u.pathname}`.slice(0, 120);
  } catch {
    return '';
  }
};

/**
 * Add › Web page: the address, pasted. https only; the store's Product Groups
 * screen has a copy button for each group's show address.
 * @param {{ onAdd: (o: { url: string, name: string }) => void, onClose: () => void, allowLoopback?: boolean }} p
 */
export function WebPageDialog({ onAdd, onClose, allowLoopback = false }) {
  const [url, setUrl] = useState('');
  const clean = url.trim();
  const ok = clean.startsWith('https://') || (allowLoopback && clean.startsWith('http://127.0.0.1'));
  const add = () => { if (ok) onAdd({ url: clean, name: nameFrom(clean) }); };
  return (
    <Dialog title={t('studio.flip-strip.add-web-page')} className="st-web-dialog" onClose={onClose} actions={(
      <>
        <Button label={t('common.cancel')} onClick={onClose} />
        <Button primary label={t('common.add')} className="st-web-add" disabled={!ok} onClick={add} />
      </>
    )}>
      <Field label={t('common.address')} hint={t('studio.flip-strip.page-starts-https')}>
        <Input className="st-web-url" value={url} placeholder="https://" autoFocus onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') add(); }} aria-label={t('common.address')} />
      </Field>
    </Dialog>
  );
}

export const FLIP_CSS = `
.st-flips { display: flex; flex-direction: column; align-items: center; gap: 4px; flex: none; }
.st-flip { display: flex; align-items: center; gap: 12px; padding: 8px 10px; background: var(--surface); border: 1px solid var(--line-soft); border-radius: 12px; }
.st-flip-btn { height: 44px; padding: 0 16px; border-radius: 8px; border: 1px solid var(--line); background: var(--surface); font: inherit; font-size: 14px; font-weight: 600; color: var(--ink); display: inline-flex; align-items: center; gap: 8px; cursor: pointer; }
.st-flip-btn:disabled { opacity: 0.45; cursor: default; }
.st-flip-next:not(:disabled) { background: var(--accent); border-color: var(--accent); color: #fff; }
.st-flip-btn kbd { display: inline-grid; place-items: center; height: 18px; padding: 0 5px; border-radius: 4px; background: var(--n-10, #eceef1); font-size: 10.5px; font-weight: 700; color: var(--ink-2); font-family: inherit; }
.st-flip-next:not(:disabled) kbd { background: rgba(255,255,255,0.25); color: #fff; }
.st-flip-now { display: flex; flex-direction: column; align-items: center; min-width: 220px; max-width: 360px; }
.st-flip-now small { font-size: 12px; color: var(--ink-2); }
.st-flip-now b { font-size: 14px; max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.st-flip-now b.wn { color: #8a4b05; }
.st-flip-sound { width: 36px; height: 36px; border-radius: 8px; border: 1px solid var(--line); background: var(--surface); color: var(--ink-3); display: grid; place-items: center; cursor: pointer; }
.st-flip-sound[aria-pressed="true"] { color: var(--accent); border-color: var(--accent); }
.st-flip-note { font-size: 12px; color: var(--ink-3); }
.st-web-dialog { width: min(480px, 92vw); }
.st-web-url { height: 34px; padding: 0 10px; border: 1px solid var(--line); border-radius: 7px; box-sizing: border-box; width: 100%; font: inherit; background: var(--surface); color: var(--ink); }
`;
