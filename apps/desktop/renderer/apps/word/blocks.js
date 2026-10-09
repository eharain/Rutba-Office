// Insert → Quick Parts: building blocks.
//
// Words worth keeping — a terms-of-payment clause, an address block, a
// sign-off — saved from the selection into a gallery and put in again with
// a click, in this document or any other. Word's two galleries are here:
// Quick Parts, at the top of the menu, and AutoText, under its heading;
// the Building Blocks Organizer lists both, to put one in or delete it.
//
// The galleries live outside any one document, in the suite's own store,
// as Word keeps them in a template rather than in the file being written.
// What a block holds is its paragraphs' own XML, made to travel by the
// engine (see @rutba/ooxml/blocks): words and their look, nothing that
// needs the document it came from.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Dialog, Field, Input, Empty, t } from '@rutba/office-ui';
import { blockText } from '@rutba/ooxml/blocks';

const KEY = 'word.buildingBlocks';
const GALLERIES = [['quickParts', t('Quick Parts')], ['autoText', t('AutoText')]];
const galleryName = (id) => (GALLERIES.find(([g]) => g === id) || [])[1] || id;

/**
 * @param {object} o
 * @param {object} o.shell the window's shell (its store keeps the galleries)
 * @param {(...ops) => Promise<object>} o.apply runs an op on the document
 * @param {(text: string, opts?: object) => void} o.toast
 * @returns {{ head: () => Array, tail: () => Array, node: React.ReactNode }}
 *   the menu items Quick Parts opens with and ends with, and the dialogs.
 */
export function useBuildingBlocks({ shell, apply, toast }) {
  const [blocks, setBlocks] = useState([]);
  const [dialog, setDialog] = useState(null);

  useEffect(() => {
    shell.store.get({ key: KEY, fallback: [] }).then((b) => setBlocks(Array.isArray(b) ? b : [])).catch(() => {});
  }, [shell]);

  const persist = useCallback(async (next) => {
    setBlocks(next);
    await shell.store.set({ key: KEY, value: next });
  }, [shell]);

  const insert = useCallback(async (b) => {
    await apply({ op: 'insertBuildingBlock', block: { paragraphs: b.paragraphs, inline: Boolean(b.inline) } });
  }, [apply]);

  // Save Selection: the selection read as a block keeps it, then named.
  const startSave = useCallback(async (gallery) => {
    const next = await apply({ op: 'buildingBlock' });
    let block = null;
    try {
      block = next?.opResult ? JSON.parse(next.opResult) : null;
    } catch {
      block = null;
    }
    if (!block?.paragraphs?.length) {
      toast(t('Select the words to keep first.'), { tone: 'bad' });
      return;
    }
    const preview = blockText(block.paragraphs);
    const name = preview.split(/\s+/).filter(Boolean).slice(0, 4).join(' ').replace(/[.,;:]+$/, '') || t('Building block');
    setDialog({ kind: 'save', block, preview, name, gallery, category: 'General', description: '' });
  }, [apply, toast]);

  const save = useCallback(async (form) => {
    const name = form.name.trim();
    if (!name) return;
    const entry = {
      id: `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`,
      name,
      gallery: form.gallery,
      category: form.category.trim() || 'General',
      description: form.description.trim(),
      paragraphs: form.block.paragraphs,
      inline: Boolean(form.block.inline),
      preview: form.preview,
    };
    // A name already in that gallery is redefined, as Word asks to.
    const replaced = blocks.some((b) => b.gallery === entry.gallery && b.name.toLowerCase() === name.toLowerCase());
    await persist([...blocks.filter((b) => !(b.gallery === entry.gallery && b.name.toLowerCase() === name.toLowerCase())), entry]);
    setDialog(null);
    toast(replaced ? t('Redefined “{name}” in {gallery}', { name, gallery: galleryName(entry.gallery) }) : t('Saved “{name}” in {gallery}', { name, gallery: galleryName(entry.gallery) }), { tone: 'good' });
  }, [blocks, persist, toast]);

  const remove = useCallback(async (id) => {
    await persist(blocks.filter((b) => b.id !== id));
  }, [blocks, persist]);

  const byName = (a, b) => a.category.localeCompare(b.category) || a.name.localeCompare(b.name);
  const entries = (gallery) => blocks.filter((b) => b.gallery === gallery).sort(byName)
    .map((b) => ({ label: b.name, icon: 'textbox', title: b.preview, run: () => insert(b) }));

  /** What Quick Parts opens with: its own gallery, then AutoText's. */
  const head = () => {
    const quick = entries('quickParts');
    return [
      ...(quick.length ? [{ heading: true, label: t('Quick Parts') }, ...quick, '-'] : []),
      { heading: true, label: t('AutoText') },
      ...entries('autoText'),
      { label: t('Save Selection to AutoText Gallery…'), run: () => startSave('autoText') },
      '-',
    ];
  };
  /** What it ends with. */
  const tail = () => [
    '-',
    { label: t('Building Blocks Organizer…'), run: () => setDialog({ kind: 'organizer' }) },
    { label: t('Save Selection to Quick Part Gallery…'), icon: 'save', run: () => startSave('quickParts') },
  ];

  const node = dialog?.kind === 'save' ? (
    <SaveBlockDialog initial={dialog} categories={[...new Set(['General', ...blocks.map((b) => b.category)])]} onSave={save} onClose={() => setDialog(null)} />
  ) : dialog?.kind === 'organizer' ? (
    <OrganizerDialog blocks={[...blocks].sort((a, b) => a.gallery.localeCompare(b.gallery) || byName(a, b))} onInsert={async (b) => { await insert(b); setDialog(null); }} onDelete={remove} onClose={() => setDialog(null)} />
  ) : null;

  return { head, tail, node };
}

/** Create New Building Block: its name, gallery, category and description. */
function SaveBlockDialog({ initial, categories, onSave, onClose }) {
  const [form, setForm] = useState(initial);
  const set = (patch) => setForm((f) => ({ ...f, ...patch }));
  return (
    <Dialog
      title={t('Create New Building Block')}
      width={480}
      onClose={onClose}
      actions={<><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('OK')} disabled={!form.name.trim()} onClick={() => onSave(form)} /></>}
    >
      <Field label={t('Name')}>
        <Input className="wd-bb-name" value={form.name} autoFocus onChange={(e) => set({ name: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter' && form.name.trim()) onSave(form); }} />
      </Field>
      <Field label={t('Gallery')}>
        <select className="rw-input wd-bb-gallery" value={form.gallery} onChange={(e) => set({ gallery: e.target.value })}>
          {GALLERIES.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
        </select>
      </Field>
      <Field label={t('Category')}>
        <Input value={form.category} list="wd-bb-categories" onChange={(e) => set({ category: e.target.value })} />
        <datalist id="wd-bb-categories">{categories.map((c) => <option key={c} value={c} />)}</datalist>
      </Field>
      <Field label={t('Description')}>
        <Input value={form.description} onChange={(e) => set({ description: e.target.value })} placeholder={t('Optional')} />
      </Field>
      <div className="wd-bb-preview" title={t('What it holds')}>{form.preview}</div>
    </Dialog>
  );
}

/** Building Blocks Organizer: every block, to put one in or delete it. */
function OrganizerDialog({ blocks, onInsert, onDelete, onClose }) {
  const [picked, setPicked] = useState(blocks[0]?.id ?? null);
  const current = blocks.find((b) => b.id === picked) || null;
  return (
    <Dialog
      title={t('Building Blocks Organizer')}
      width={620}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Delete')} disabled={!current} onClick={() => { onDelete(current.id); setPicked(blocks.find((b) => b.id !== current.id)?.id ?? null); }} />
          <Button label={t('Close')} onClick={onClose} />
          <Button primary label={t('Insert')} disabled={!current} onClick={() => onInsert(current)} />
        </>
      }
    >
      {blocks.length ? (
        <div className="wd-bb-organizer">
          <div className="wd-bb-list" role="listbox" aria-label={t('Building blocks')}>
            <div className="wd-bb-row head"><span>{t('Name')}</span><span>{t('Gallery')}</span><span>{t('Category')}</span></div>
            {blocks.map((b) => (
              <button key={b.id} type="button" role="option" aria-selected={b.id === picked} className={`wd-bb-row${b.id === picked ? ' on' : ''}`} onClick={() => setPicked(b.id)} onDoubleClick={() => onInsert(b)}>
                <span>{b.name}</span><span>{galleryName(b.gallery)}</span><span>{b.category}</span>
              </button>
            ))}
          </div>
          <div className="wd-bb-preview">{current ? (current.description ? `${current.description}\n\n` : '') + current.preview : ''}</div>
        </div>
      ) : (
        <Empty icon="textbox" title={t('No building blocks yet')}>{t('Select some words, then Insert → Quick Parts → Save Selection to Quick Part Gallery.')}</Empty>
      )}
    </Dialog>
  );
}

export const BLOCKS_CSS = `
.wd-bb-preview { margin-top: 10px; padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface-2, var(--window)); font-size: 12.5px; color: var(--ink-2); white-space: pre-wrap; max-height: 140px; overflow: auto; }
.wd-bb-organizer { display: grid; grid-template-columns: minmax(0, 1.4fr) minmax(0, 1fr); gap: 12px; }
.wd-bb-list { border: 1px solid var(--line); border-radius: var(--r-2); max-height: 300px; overflow: auto; }
.wd-bb-row { display: grid; grid-template-columns: 1.4fr 1fr 1fr; gap: 8px; width: 100%; padding: 6px 10px; border: 0; background: none; text-align: left; font: inherit; font-size: 12.5px; color: var(--ink); cursor: pointer; }
.wd-bb-row span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.wd-bb-row.head { font-weight: 600; color: var(--ink-2); cursor: default; border-bottom: 1px solid var(--line); position: sticky; top: 0; background: var(--surface); }
.wd-bb-row.on { background: var(--accent-soft); }
.wd-bb-organizer .wd-bb-preview { margin-top: 0; max-height: 300px; }
`;
