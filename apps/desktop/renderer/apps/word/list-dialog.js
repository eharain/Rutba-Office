// Home → Multilevel List → Define New Multilevel List, as Word's: a list of
// one's own, a level at a time — its number style, the number's own words
// with the levels above it drawn in, where it starts, where it sits, whether
// its numbers are all figures (legal), and the paragraph style it is linked
// to, so every heading of that style is numbered by it. The preview is drawn
// by the same counting the page uses.

import React, { useMemo, useState } from 'react';
import { Button, Dialog, Field, Input, Select, t } from '@rutba/office-ui';
import { computeListLabels } from '@rutba/doc-view/lists';

const STYLES = [
  ['decimal', '1, 2, 3, …'],
  ['decimalZero', '01, 02, 03, …'],
  ['lowerLetter', 'a, b, c, …'],
  ['upperLetter', 'A, B, C, …'],
  ['lowerRoman', 'i, ii, iii, …'],
  ['upperRoman', 'I, II, III, …'],
  ['bullet', t('Bullet •')],
  ['none', t('(none)')],
];

/** Twentieths of a point to centimetres and back, as the dialog speaks. */
const TW_PER_CM = 566.93;
const cm = (tw) => Math.round((tw / TW_PER_CM) * 100) / 100;
const tw = (value) => Math.max(0, Math.round(Number(value) * TW_PER_CM) || 0);

/** Word's 1. 1.1. 1.1.1. to start from: each level's number carrying the ones above it. */
const startingLevels = () => Array.from({ length: 9 }, (_, i) => ({
  format: 'decimal',
  text: Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.') + '.',
  start: 1,
  legal: false,
  alignedCm: cm(720 * (i + 1) - 360),
  indentCm: cm(720 * (i + 1)),
  style: '',
}));

/** Each level's first number as the page would label it. */
function previewLabels(levels) {
  const defs = { preview: levels.map((l) => ({ format: l.format, lvlText: l.format === 'bullet' ? (l.text || '•') : l.text, start: Number(l.start) || 0, legal: l.legal, indentPx: 0, font: null })) };
  const blocks = levels.map((_, i) => ({ index: i, numbering: { numId: 'preview', level: i } }));
  const flow = levels.map((_, i) => ({ kind: 'paragraph', paragraphIndex: i }));
  const out = computeListLabels(flow, blocks, defs);
  return levels.map((_, i) => out.get(i)?.label ?? '');
}

export function DefineListDialog({ styles = [], onClose, onApply }) {
  const [levels, setLevels] = useState(startingLevels);
  const [at, setAt] = useState(0);
  const [name, setName] = useState('');
  const level = levels[at];
  const set = (patch) => setLevels((all) => all.map((l, i) => (i === at ? { ...l, ...patch } : l)));
  const labels = useMemo(() => previewLabels(levels), [levels]);
  // A style may be linked to one level only: the one picked here leaves any other.
  const linkStyle = (style) => setLevels((all) => all.map((l, i) => (i === at ? { ...l, style } : style && l.style === style ? { ...l, style: '' } : l)));
  const indentOk = levels.every((l) => Number(l.indentCm) >= Number(l.alignedCm) && Number(l.alignedCm) >= 0);
  // A numbered level that has lost its own number (%2 in level 2's words).
  const missing = levels.findIndex((l, i) => l.format !== 'none' && l.format !== 'bullet' && !String(l.text).includes(`%${i + 1}`));
  const textOk = missing < 0;
  const apply = () => onApply({
    name: name.trim() || null,
    levels: levels.map((l) => ({
      format: l.format,
      text: l.format === 'bullet' ? (l.text || '•') : l.text,
      start: Number(l.start) || 0,
      legal: l.legal,
      indentTw: tw(l.indentCm),
      hangingTw: Math.max(0, tw(l.indentCm) - tw(l.alignedCm)),
      style: l.style || null,
    })),
  });
  return (
    <Dialog
      title={t('Define new Multilevel list')}
      width={620}
      onClose={onClose}
      className="wd-define-list"
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="wd-define-list-ok" disabled={!indentOk || !textOk} onClick={apply} />
        </>
      }
    >
      <div className="ml-form">
        <div style={{ display: 'grid', gridTemplateColumns: '64px 1fr', gap: 12 }}>
          <div role="listbox" aria-label={t('Level to change')} style={{ display: 'grid', gap: 2 }}>
            {levels.map((_, i) => (
              <button key={i} type="button" role="option" aria-selected={i === at} className={`wd-define-level${i === at ? ' on' : ''}`} onClick={() => setAt(i)}
                style={{ padding: '3px 0', border: '1px solid var(--line, #ccc)', background: i === at ? 'var(--accent-soft, #dbe6fb)' : 'transparent', borderRadius: 4, cursor: 'pointer' }}>
                {i + 1}
              </button>
            ))}
          </div>
          <div className="wd-define-preview" aria-label={t('Preview')} style={{ border: '1px solid var(--line, #ccc)', borderRadius: 4, padding: '6px 10px', fontSize: 12.5, lineHeight: '17px', overflow: 'hidden' }}>
            {levels.map((l, i) => (
              <div key={i} data-level={i} style={{ paddingLeft: `${Math.min(14, Number(l.alignedCm) || 0) * 14}px`, fontWeight: i === at ? 600 : 400, whiteSpace: 'nowrap' }}>
                <span className="wd-define-label">{labels[i]}</span>
                <span style={{ display: 'inline-block', width: 120, height: 5, marginLeft: 8, background: 'var(--line, #ccc)', verticalAlign: 'middle' }} />
              </div>
            ))}
          </div>
        </div>
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
          <Field label={t('Number style for level {level}', { level: at + 1 })}>
            <Select data-role="define-format" value={level.format} onChange={(e) => {
              const format = e.target.value;
              set(format === 'bullet' ? { format, text: '•' } : level.format === 'bullet' ? { format, text: `%${at + 1}.` } : { format });
            }}>
              {STYLES.map(([value, label]) => <option key={value} value={value}>{label}</option>)}
            </Select>
          </Field>
          <Field label={t('Include level number from')}>
            <Select value="" disabled={at === 0 || level.format === 'bullet'} onChange={(e) => { if (e.target.value) set({ text: `${e.target.value}${level.text}` }); }}>
              <option value="">{at === 0 ? t('(the first level has none above it)') : t('Choose a level above…')}</option>
              {Array.from({ length: at }, (_, k) => <option key={k} value={`%${k + 1}.`}>{t('Level {level}', { level: k + 1 })}</option>)}
            </Select>
          </Field>
        </div>
        <Field label={t('Enter formatting for number')} hint={level.format === 'bullet' ? t('The bullet character.') : at === 0 ? t('%{level} is this level\'s number.', { level: at + 1 }) : at === 1 ? t('%2 is this level\'s number; %1 the number of the level above it.') : t('%{level} is this level\'s number; %1 to %{above} the numbers of the levels above it.', { level: at + 1, above: at })}>
          <Input data-role="define-text" value={level.text} maxLength={60} onChange={(e) => set({ text: e.target.value })} />
        </Field>
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr 1fr' }}>
          <Field label={t('Start at')}>
            <Input data-role="define-start" type="number" min="0" max="32767" value={level.start} disabled={level.format === 'bullet' || level.format === 'none'} onChange={(e) => set({ start: e.target.value })} />
          </Field>
          <Field label={t('Aligned at (cm)')}>
            <Input type="number" min="0" max="50" step="0.01" value={level.alignedCm} onChange={(e) => set({ alignedCm: e.target.value })} />
          </Field>
          <Field label={t('Text indent at (cm)')}>
            <Input type="number" min="0" max="50" step="0.01" value={level.indentCm} onChange={(e) => set({ indentCm: e.target.value })} />
          </Field>
        </div>
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
          <Field label={t('Link level to style')}>
            <Select data-role="define-style" value={level.style} onChange={(e) => linkStyle(e.target.value)}>
              <option value="">{t('(no style)')}</option>
              {styles.map((s) => <option key={s.id} value={s.id}>{s.name || s.id}</option>)}
            </Select>
          </Field>
          <Field label={t('List name (optional)')}>
            <Input data-role="define-name" value={name} maxLength={60} placeholder={t('As Word lists it in the file')} onChange={(e) => setName(e.target.value)} />
          </Field>
        </div>
        <label className="about-auto">
          <input type="checkbox" className="wd-define-legal" checked={level.legal} disabled={level.format === 'bullet'} onChange={(e) => set({ legal: e.target.checked })} />
          <span>{t('Legal style numbering: every number in this level shown in figures (Section 1.01 under Article I)')}</span>
        </label>
        {!indentOk ? <p role="alert" style={{ margin: 0, fontSize: 12.5 }}>{t('A level\'s text sits at or past its number: Text indent at is not less than Aligned at.')}</p> : null}
        {!textOk ? <p role="alert" style={{ margin: 0, fontSize: 12.5 }}>{t('Each numbered level shows its own number: keep %{level} in level {level}\'s number.', { level: missing + 1 })}</p> : null}
      </div>
    </Dialog>
  );
}
