// What changed, told in the window.
//
// An update installs on quit and the next launch looks exactly like the
// last one; the owner called that silence (2026-09-20). The first window of
// a new version says the version arrived and offers the notes — the
// release's own, bundled at build time from docs/releases/v<version>.md, so
// a copy with the network switched off tells the same story the release
// page does.

import React, { useEffect, useState } from 'react';
import { Button } from '@rutba/office-ui';
import { SITE } from '@rutba/office-formats/registry';
import { parseNotes } from './whatsnew/notes.js';

const Runs = ({ runs }) => runs.map((r, i) => (r.bold ? <strong key={i}>{r.text}</strong> : r.code ? <code key={i}>{r.text}</code> : <React.Fragment key={i}>{r.text}</React.Fragment>));

/** The blocks drawn: bullets gathered into lists, the first paragraph as the lead. */
function Notes({ blocks }) {
  const out = [];
  let list = null;
  let lead = true;
  blocks.forEach((b, i) => {
    if (b.kind === 'li') {
      if (!list) out.push((list = { key: `l${i}`, items: [] }));
      list.items.push(<li key={i}><Runs runs={b.runs} /></li>);
      return;
    }
    list = null;
    if (b.kind === 'h') out.push(<h3 key={i}>{b.text}</h3>);
    else {
      out.push(<p key={i} className={lead ? 'lead' : ''}><Runs runs={b.runs} /></p>);
      lead = false;
    }
  });
  return out.map((n) => (n.items ? <ul key={n.key}>{n.items}</ul> : n));
}

/**
 * The release's notes for one version, in a dialog. `version` is the one
 * to tell about; the bundled notes are shown only when they are its own,
 * otherwise the releases page is offered instead of a wrong story.
 */
export function WhatsNew({ shell, version, onClose }) {
  const [notes, setNotes] = useState(undefined);
  useEffect(() => {
    let live = true;
    fetch('rutba://app/whatsnew.json')
      .then((r) => r.json())
      .then((j) => live && setNotes(j?.notes && (!version || j.version === version) ? { version: j.version, blocks: parseNotes(j.notes) } : null))
      .catch(() => live && setNotes(null));
    return () => {
      live = false;
    };
  }, [version]);
  const open = (url) => shell.shell.openExternal({ url }).catch(() => {});
  return (
    <div className="rw-scrim" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <style>{CSS}</style>
      <div className="rw-dialog rw-whatsnew" style={{ width: 580 }}>
        <div className="rw-dialog-head">What’s new in Rutba Office {version || notes?.version || ''}</div>
        <div className="rw-dialog-body rw-whatsnew-body">
          {notes === undefined ? (
            <p className="rw-hint">Reading the notes…</p>
          ) : notes === null ? (
            <p>The notes for this version are on the releases page.</p>
          ) : (
            <Notes blocks={notes.blocks} />
          )}
        </div>
        <div className="rw-dialog-foot">
          <Button label="All releases" onClick={() => open(SITE.releases)} />
          <Button label="Close" primary className="rw-whatsnew-close" onClick={onClose} />
        </div>
      </div>
    </div>
  );
}

const CSS = `
.rw-whatsnew-body { max-height: 62vh; overflow: auto; font-size: 13px; line-height: 1.5; color: var(--ink-2); }
.rw-whatsnew-body p { margin: 8px 0; }
.rw-whatsnew-body p.lead { font-size: 14px; color: var(--ink); margin: 10px 0 4px; }
.rw-whatsnew-body h3 { font-size: 13.5px; font-weight: 650; color: var(--ink); margin: 18px 0 6px; }
.rw-whatsnew-body ul { margin: 6px 0; padding-left: 18px; }
.rw-whatsnew-body li { margin: 5px 0; }
.rw-whatsnew-body strong { color: var(--ink); font-weight: 600; }
.rw-whatsnew-body code { font-family: var(--font-mono, ui-monospace, monospace); font-size: 12px; background: var(--sunken); padding: 0 4px; border-radius: 3px; }
`;
