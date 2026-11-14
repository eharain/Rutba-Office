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
  comments: { label: () => t('common.comments'), detail: () => t('inspect.inspects-comments-notes'), found: (n) => tn(n, 'inspect.count-comment-found') },
  revisions: { label: () => t('common.revisions'), detail: () => t('inspect.inspects-tracked-changes'), found: (n) => tn(n, 'inspect.count-tracked-change') },
  personal: { label: () => t('inspect.personal-information'), detail: () => t('inspect.inspects-author-last'), found: (n) => tn(n, 'inspect.count-name-detail') },
  properties: { label: () => t('inspect.document-properties'), detail: () => t('inspect.inspects-title-subject'), found: (n) => tn(n, 'inspect.count-property-found') },
  customXml: { label: () => t('inspect.custom-xml-data'), detail: () => t('inspect.inspects-custom-xml'), found: (n) => tn(n, 'inspect.count-custom-xml') },
  headers: { label: () => t('inspect.headers-footers'), detail: () => t('inspect.inspects-information'), found: (n) => tn(n, 'inspect.count-header-footer') },
  hidden: { label: () => t('inspect.hidden-text'), detail: () => t('inspect.inspects-text-formatted'), found: (n) => tn(n, 'inspect.count-piece-hidden') },
  invisible: { label: () => t('inspect.invisible-content'), detail: () => t('inspect.inspects-objects'), found: (n) => tn(n, 'inspect.count-invisible-object') },
  ink: { label: () => t('common.ink'), detail: () => t('inspect.inspects-ink-drawn'), found: (n) => tn(n, 'inspect.count-ink-drawing') },
  embedded: { label: () => t('inspect.embedded-documents'), detail: () => t('inspect.inspects-documents'), found: (n) => tn(n, 'inspect.count-embedded-document') },
  notes: { label: () => t('inspect.presentation-notes'), detail: () => t('inspect.inspects-words-speaker'), found: (n) => tn(n, 'inspect.count-slide-notes') },
  offSlide: { label: () => t('inspect.off-slide-content'), detail: () => t('inspect.inspects-objects-off'), found: (n) => tn(n, 'inspect.count-object-off') },
  hiddenSheets: { label: () => t('inspect.hidden-worksheets'), detail: () => t('inspect.inspects-worksheets'), found: (n) => tn(n, 'inspect.count-hidden-worksheet') },
  hiddenRowsColumns: { label: () => t('inspect.hidden-rows-columns'), detail: () => t('inspect.inspects-hidden-rows'), found: (n) => tn(n, 'inspect.count-hidden-row') },
  hiddenNames: { label: () => t('inspect.hidden-names'), detail: () => t('inspect.inspects-defined-names'), found: (n) => tn(n, 'inspect.count-hidden-name') },
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
  const noun = app === 'sheets' ? t('common.workbook') : app === 'slides' ? t('common.presentation') : t('inspect.document');
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
      title={app === 'sheets' ? t('inspect.document-inspector-workbook') : app === 'slides' ? t('inspect.document-inspector-presentation') : t('inspect.document-inspector')}
      width={600}
      onClose={onClose}
      className="di-dialog"
      actions={stage === 'choose'
        ? <><Button label={t('common.close')} onClick={onClose} /><Button primary className="di-inspect" label={t('inspect.inspect')} disabled={busy || !ticked?.size} onClick={() => { setStage('results'); run(); }} /></>
        : <><Button className="di-reinspect" label={t('inspect.reinspect')} disabled={busy} onClick={() => { setGone(new Set()); setStage('choose'); }} /><Button primary className="di-close" label={t('common.close')} onClick={onClose} /></>}
    >
      <style>{INSPECT_CSS}</style>
      <p className="di-lead">{stage === 'choose'
        ? t('inspect.check-noun-content', { noun })
        : t('inspect.review-inspection')}</p>
      {busy && !items ? <div className="di-busy"><Spinner />{t('inspect.inspecting')}</div> : null}
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
                <p>{removed ? t('inspect.all-items-removed') : k.found ? kind.found(k.found) : t('inspect.no-items-found')}</p>
                {k.found && k.items?.length ? <div className="di-items" data-tip={k.items.join(', ')}>{k.items.join(' · ')}</div> : null}
              </div>
              {k.found && !k.keep ? <Button className={`di-remove-${k.id}`} label={t('common.remove-all')} disabled={busy} onClick={() => removeKind(k.id)} /> : <span />}
            </div>
          );
        })}
      </div>
      {stage === 'results' && shown.some((k) => k.found && !k.keep) ? (
        <div className="di-warn"><Icon name="info" size={14} /><span>{t('inspect.some-removed-content')}</span></div>
      ) : null}
    </Dialog>
  );
}
