// File → Info → Check for Issues → Inspect Document (Inspect Workbook,
// Inspect Presentation): the Document Inspector of Word, Excel and
// PowerPoint. The kinds of hidden or personal content are ticked, Inspect
// looks, each kind says what it found, and Remove All takes it out of the
// open file, which is then shown as it now is. As in Office, a removal is
// not undone with Undo; the dialog says so before anything is removed.

import React, { useState } from 'react';
import { Button, Dialog, Icon, Spinner, t, tn } from '@rutba/office-ui';

/** Each kind of content, with Office's heading and what it looks for. */
const KINDS = {
  comments: { label: () => t('Comments'), detail: () => t('Inspects for comments and notes.'), found: (n) => tn(n, '{count} comment was found.', '{count} comments were found.') },
  revisions: { label: () => t('Revisions'), detail: () => t('Inspects for tracked changes. Removing them accepts them.'), found: (n) => tn(n, '{count} tracked change was found.', '{count} tracked changes were found.') },
  personal: { label: () => t('Personal Information'), detail: () => t('Inspects for the author, the last person to save, the company and the names of reviewers and commenters.'), found: (n) => tn(n, '{count} name or detail was found.', '{count} names or details were found.') },
  properties: { label: () => t('Document Properties'), detail: () => t('Inspects for the title, subject, keywords and properties of the file\'s own.'), found: (n) => tn(n, '{count} property was found.', '{count} properties were found.') },
  customXml: { label: () => t('Custom XML Data'), detail: () => t('Inspects for custom XML data stored with the file.'), found: (n) => tn(n, '{count} custom XML part was found.', '{count} custom XML parts were found.') },
  headers: { label: () => t('Headers, Footers, and Watermarks'), detail: () => t('Inspects for information in headers, footers and watermarks.'), found: (n) => tn(n, '{count} header or footer was found.', '{count} headers or footers were found.') },
  hidden: { label: () => t('Hidden Text'), detail: () => t('Inspects for text formatted as hidden.'), found: (n) => tn(n, '{count} piece of hidden text was found.', '{count} pieces of hidden text were found.') },
  invisible: { label: () => t('Invisible Content'), detail: () => t('Inspects for objects formatted as invisible.'), found: (n) => tn(n, '{count} invisible object was found.', '{count} invisible objects were found.') },
  ink: { label: () => t('Ink'), detail: () => t('Inspects for ink drawn on the file.'), found: (n) => tn(n, '{count} ink drawing was found.', '{count} ink drawings were found.') },
  embedded: { label: () => t('Embedded Documents'), detail: () => t('Inspects for documents embedded in the file, which may hold information that is not shown.'), found: (n) => tn(n, '{count} embedded document was found. Delete it from the file to remove it.', '{count} embedded documents were found. Delete them from the file to remove them.') },
  notes: { label: () => t('Presentation Notes'), detail: () => t('Inspects for words in the speaker notes.'), found: (n) => tn(n, '{count} slide has notes.', '{count} slides have notes.') },
  offSlide: { label: () => t('Off-Slide Content'), detail: () => t('Inspects for objects off the edge of the slide.'), found: (n) => tn(n, '{count} object is off the slide.', '{count} objects are off the slide.') },
  hiddenSheets: { label: () => t('Hidden Worksheets'), detail: () => t('Inspects for worksheets that are hidden.'), found: (n) => tn(n, '{count} hidden worksheet was found.', '{count} hidden worksheets were found.') },
  hiddenRowsColumns: { label: () => t('Hidden Rows and Columns'), detail: () => t('Inspects for hidden rows and columns, which may hold data.'), found: (n) => tn(n, '{count} hidden row or column was found.', '{count} hidden rows or columns were found.') },
  hiddenNames: { label: () => t('Hidden Names'), detail: () => t('Inspects for defined names that are hidden.'), found: (n) => tn(n, '{count} hidden name was found.', '{count} hidden names were found.') },
};
export const inspectLabel = (id) => (KINDS[id]?.label || (() => id))();

const INSPECT_CSS = `
.di-lead { margin: 0 0 10px; color: var(--ink-2); font-size: 12.5px; line-height: 1.5; }
.di-list { display: flex; flex-direction: column; border: 1px solid var(--line); border-radius: 8px; max-height: 440px; overflow: auto; }
.di-row { display: grid; grid-template-columns: 22px 1fr auto; gap: 10px; align-items: start; padding: 10px 12px; border-bottom: 1px solid var(--line-soft); }
.di-row:last-child { border-bottom: 0; }
.di-row h5 { margin: 0 0 2px; font-size: 13px; font-weight: 600; }
.di-row p { margin: 0; color: var(--ink-2); font-size: 12px; line-height: 1.45; }
.di-row .di-items { margin-top: 4px; color: var(--ink-3); font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 400px; }
.di-mark { width: 20px; height: 20px; border-radius: 50%; display: grid; place-items: center; margin-top: 1px; }
.di-mark.clean { background: color-mix(in srgb, var(--good, #1e8e3e) 14%, transparent); color: var(--good, #1e8e3e); }
.di-mark.found { background: color-mix(in srgb, var(--warn, #b26a00) 16%, transparent); color: var(--warn, #b26a00); font-weight: 700; font-size: 13px; }
.di-mark.gone { background: color-mix(in srgb, var(--accent) 14%, transparent); color: var(--accent-text); }
.di-row input[type="checkbox"] { margin-top: 3px; }
.di-warn { display: flex; gap: 8px; align-items: flex-start; margin-top: 10px; padding: 8px 10px; border-radius: 8px; font-size: 12px; color: var(--ink-2); background: color-mix(in srgb, var(--warn, #b26a00) 8%, var(--surface)); border: 1px solid color-mix(in srgb, var(--warn, #b26a00) 24%, transparent); }
.di-busy { display: flex; align-items: center; gap: 10px; padding: 24px; justify-content: center; color: var(--ink-2); }
`;

/**
 * The Inspector for one window: `inspect` asks the document service what
 * the open file holds; `remove(ids)` takes kinds out and hands back the
 * window's new model.
 */
export function InspectDialog({ app, onClose, inspect, remove }) {
  const noun = app === 'sheets' ? t('Workbook') : app === 'slides' ? t('Presentation') : t('Document');
  const [stage, setStage] = useState('choose');
  const [items, setItems] = useState(null);
  const [ticked, setTicked] = useState(null);
  const [gone, setGone] = useState(new Set());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const run = async () => {
    setBusy(true);
    setError(null);
    try {
      const r = await inspect();
      setItems(r.items);
      setTicked((was) => was || new Set(r.items.filter((k) => KINDS[k.id]).map((k) => k.id)));
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  };
  React.useEffect(() => { run(); /* what can be inspected, before anything is ticked */ }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const removeKind = async (id) => {
    setBusy(true);
    try {
      const r = await remove([id]);
      setItems(r.items);
      setGone((g) => new Set([...g, id]));
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  };
  const shown = (items || []).filter((k) => KINDS[k.id] && (stage === 'choose' || ticked?.has(k.id)));
  return (
    <Dialog
      title={app === 'sheets' ? t('Document Inspector: Workbook') : app === 'slides' ? t('Document Inspector: Presentation') : t('Document Inspector')}
      width={600}
      onClose={onClose}
      className="di-dialog"
      actions={stage === 'choose'
        ? <><Button label={t('Close')} onClick={onClose} /><Button primary className="di-inspect" label={t('Inspect')} disabled={busy || !ticked?.size} onClick={() => { setStage('results'); run(); }} /></>
        : <><Button className="di-reinspect" label={t('Reinspect')} disabled={busy} onClick={() => { setGone(new Set()); setStage('choose'); }} /><Button primary className="di-close" label={t('Close')} onClick={onClose} /></>}
    >
      <style>{INSPECT_CSS}</style>
      <p className="di-lead">{stage === 'choose'
        ? t('To check the {noun} for the content selected, click Inspect.', { noun })
        : t('Review the inspection results.')}</p>
      {busy && !items ? <div className="di-busy"><Spinner />{t('Inspecting…')}</div> : null}
      {error ? <div className="di-warn"><Icon name="info" size={14} /><span>{error}</span></div> : null}
      <div className="di-list">
        {shown.map((k) => {
          const kind = KINDS[k.id];
          if (stage === 'choose') {
            return (
              <label key={k.id} className="di-row" data-kind={k.id}>
                <input type="checkbox" className={`di-tick-${k.id}`} checked={Boolean(ticked?.has(k.id))} onChange={(e) => setTicked((s) => { const n = new Set(s); if (e.target.checked) n.add(k.id); else n.delete(k.id); return n; })} />
                <div><h5>{kind.label()}</h5><p>{kind.detail()}</p></div>
                <span />
              </label>
            );
          }
          const removed = gone.has(k.id) && !k.found;
          return (
            <div key={k.id} className="di-row" data-kind={k.id} data-found={k.found}>
              <span className={`di-mark ${removed ? 'gone' : k.found ? 'found' : 'clean'}`}>{removed || !k.found ? <Icon name="check" size={13} /> : '!'}</span>
              <div>
                <h5>{kind.label()}</h5>
                <p>{removed ? t('All items were removed.') : k.found ? kind.found(k.found) : t('No items were found.')}</p>
                {k.found && k.items?.length ? <div className="di-items" data-tip={k.items.join(', ')}>{k.items.join(' · ')}</div> : null}
              </div>
              {k.found && !k.keep ? <Button className={`di-remove-${k.id}`} label={t('Remove All')} disabled={busy} onClick={() => removeKind(k.id)} /> : <span />}
            </div>
          );
        })}
      </div>
      {stage === 'results' && shown.some((k) => k.found && !k.keep) ? (
        <div className="di-warn"><Icon name="info" size={14} /><span>{t('Some removed content cannot be put back with Undo. Save a copy first if you may want it again.')}</span></div>
      ) : null}
    </Dialog>
  );
}
