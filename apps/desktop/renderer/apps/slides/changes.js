// Review → Show Changes: the pane that lists what is different about this
// deck since it was last open on this computer, slide by slide — a click goes
// to the slide — and Dismiss, which clears the list once read.

import React from 'react';
import { Button, Icon, t } from '@rutba/office-ui';

const WORDS = { added: t('Added'), removed: t('Removed'), moved: t('Moved'), changed: t('Changed') };
const ICONS = { added: 'plus', removed: 'trash', moved: 'sort', changed: 'textbox' };

/** `changes`: `{ first }` the first time this deck is open here, else `{ list, savedBy, saved }`. */
export function ChangesPane({ changes, onGo, onDismiss }) {
  if (!changes) return <div className="sl-pane-empty">{t('Looking for changes…')}</div>;
  if (changes.first) return <div className="sl-pane-empty">{t('This is the first time this deck has been open on this computer, so there is nothing to compare it with yet. Next time, whatever is new in the file is listed here.')}</div>;
  const list = changes.list || [];
  const when = changes.saved ? new Date(changes.saved) : null;
  return (
    <div className="sl-changes">
      {changes.savedBy || when ? (
        <div className="sl-changes-by">{changes.savedBy
          ? (when && !Number.isNaN(when.getTime()) ? t('Last saved by {person}, {when}', { person: changes.savedBy, when: when.toLocaleString() }) : t('Last saved by {person}', { person: changes.savedBy }))
          : (when && !Number.isNaN(when.getTime()) ? t('Last saved, {when}', { when: when.toLocaleString() }) : t('Last saved'))}</div>
      ) : null}
      {list.length ? list.map((c) => (
        <button key={`${c.kind}-${c.id}`} type="button" className={`sl-change sl-change-${c.kind}`} data-kind={c.kind} disabled={c.index == null} onClick={() => c.index != null && onGo(c.index)} title={c.index != null ? t('Go to the slide') : t('This slide is no longer in the deck')}>
          <Icon name={ICONS[c.kind]} size={14} />
          <span className="sl-change-text">
            <span className="sl-change-head">{t('{change}: slide {number}', { change: WORDS[c.kind], number: c.index != null ? c.index + 1 : c.was + 1 })}{c.title ? ` · ${c.title}` : ''}</span>
            {c.details?.length ? <span className="sl-change-details">{c.details.join('; ')}</span> : null}
          </span>
        </button>
      )) : <div className="sl-pane-empty">{t('Nothing has changed since this deck was last open here.')}</div>}
      {list.length ? <Button icon="check" label={t('Dismiss')} className="sl-changes-dismiss" title={t('Dismiss — these changes are read; the list clears')} onClick={onDismiss} /> : null}
    </div>
  );
}

export const CHANGES_CSS = `
.sl-changes { display: flex; flex-direction: column; gap: 4px; padding: 6px; }
.sl-changes-by { font-size: 12px; opacity: .75; padding: 2px 4px 6px; }
.sl-change { display: flex; gap: 8px; align-items: flex-start; text-align: start; background: none; border: 1px solid var(--line); border-radius: 6px; padding: 6px 8px; color: inherit; font: inherit; cursor: pointer; }
.sl-change:disabled { cursor: default; opacity: .7; }
.sl-change:hover:not(:disabled) { border-color: var(--accent); }
.sl-change-text { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
.sl-change-head { font-weight: 600; font-size: 12.5px; }
.sl-change-details { font-size: 12px; opacity: .8; overflow-wrap: anywhere; }
.sl-change-added .rw-icon { color: #2e9d4a; }
.sl-change-removed .rw-icon { color: #d13438; }
.sl-thumb-chg { display: inline-block; width: 7px; height: 7px; border-radius: 50%; background: #c239b3; margin-top: 3px; }
`;
