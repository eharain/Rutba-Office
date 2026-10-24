// Home → Styles → right-click a style → Modify, as Word's Modify Style: the
// style's font, size, bold and italic, colour, and the space above and below
// its paragraphs. What is changed here changes the style itself, so every
// paragraph that has it follows, and the file keeps it in styles.xml.

import React, { useState } from 'react';
import { Button, Dialog, Field, Input, Select, t } from '@rutba/office-ui';

const FONTS = ['Calibri', 'Calibri Light', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Tahoma', 'Garamond', 'Consolas', 'Courier New'];

/** Pixels (as the resolved styles carry them) to points, to the half point. */
const pt = (px) => (px == null ? null : Math.round(px * 0.75 * 2) / 2);

/** The style as the dialog starts it: its face, size, look and spacing now. */
function startingLook(style = {}) {
  return {
    fontName: style.fontName || 'Calibri',
    sizePt: pt(style.sizePx) ?? 11,
    bold: Boolean(style.bold),
    italic: Boolean(style.italic),
    colour: /^#[0-9a-f]{6}$/i.test(style.colour || '') ? style.colour.toLowerCase() : '#000000',
    spaceBeforePt: pt(style.spaceBeforePx) ?? 0,
    spaceAfterPt: pt(style.spaceAfterPx) ?? 0,
  };
}

export function ModifyStyleDialog({ name, style, onClose, onApply }) {
  const [start] = useState(() => startingLook(style));
  const [look, setLook] = useState(start);
  const set = (patch) => setLook((l) => ({ ...l, ...patch }));
  const size = Number(look.sizePt);
  const sizeOk = Number.isFinite(size) && size >= 1 && size <= 1638;
  const spaceOk = [look.spaceBeforePt, look.spaceAfterPt].every((v) => Number.isFinite(Number(v)) && Number(v) >= 0 && Number(v) <= 1584);
  // Only what was changed goes to the style; the rest it keeps as it was.
  const apply = () => {
    const spec = {};
    if (look.fontName !== start.fontName) spec.fontName = look.fontName;
    if (size !== start.sizePt) spec.sizePt = size;
    if (look.bold !== start.bold) spec.bold = look.bold;
    if (look.italic !== start.italic) spec.italic = look.italic;
    if (look.colour !== start.colour) spec.colour = look.colour;
    if (Number(look.spaceBeforePt) !== start.spaceBeforePt) spec.spaceBeforePt = Number(look.spaceBeforePt);
    if (Number(look.spaceAfterPt) !== start.spaceAfterPt) spec.spaceAfterPt = Number(look.spaceAfterPt);
    onApply(spec);
  };
  const fonts = FONTS.includes(look.fontName) ? FONTS : [look.fontName, ...FONTS];
  return (
    <Dialog
      title={t('Modify Style: {style}', { style: name })}
      width={460}
      onClose={onClose}
      className="wd-modify-style"
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} disabled={!sizeOk || !spaceOk} onClick={apply} />
        </>
      }
    >
      <div className="ml-form">
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 110px', gap: 12 }}>
          <Field label={t('Font')}>
            <Select data-role="style-font" value={look.fontName} onChange={(e) => set({ fontName: e.target.value })}>
              {fonts.map((f) => <option key={f} value={f}>{f}</option>)}
            </Select>
          </Field>
          <Field label={t('Size (pt)')}>
            <Input data-role="style-size" type="number" min="1" max="1638" step="0.5" value={look.sizePt} onChange={(e) => set({ sizePt: e.target.value })} />
          </Field>
        </div>
        <div style={{ display: 'flex', gap: 18, alignItems: 'center', flexWrap: 'wrap' }}>
          <label className="about-auto">
            <input data-role="style-bold" type="checkbox" checked={look.bold} onChange={(e) => set({ bold: e.target.checked })} />
            <span>{t('Bold')}</span>
          </label>
          <label className="about-auto">
            <input data-role="style-italic" type="checkbox" checked={look.italic} onChange={(e) => set({ italic: e.target.checked })} />
            <span>{t('Italic')}</span>
          </label>
          <label className="about-auto">
            <input data-role="style-colour" type="color" value={look.colour} onChange={(e) => set({ colour: e.target.value })} aria-label={t('Font colour')} />
            <span>{t('Font colour')}</span>
          </label>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <Field label={t('Space before (pt)')}>
            <Input data-role="style-before" type="number" min="0" max="1584" step="1" value={look.spaceBeforePt} onChange={(e) => set({ spaceBeforePt: e.target.value })} />
          </Field>
          <Field label={t('Space after (pt)')}>
            <Input data-role="style-after" type="number" min="0" max="1584" step="1" value={look.spaceAfterPt} onChange={(e) => set({ spaceAfterPt: e.target.value })} />
          </Field>
        </div>
        <p
          className="wd-modify-preview"
          data-role="style-preview"
          style={{
            margin: 0, padding: '10px 12px', border: '1px solid var(--rw-border, #ccc)', borderRadius: 4,
            fontFamily: look.fontName, fontSize: `${sizeOk ? Math.min(size, 36) : 11}pt`,
            fontWeight: look.bold ? 700 : 400, fontStyle: look.italic ? 'italic' : 'normal', color: look.colour,
          }}
        >
          {t('The quick brown fox jumps over the lazy dog')}
        </p>
        <p className="rw-hint" style={{ margin: 0 }}>{t('Every paragraph with this style changes with it.')}</p>
      </div>
    </Dialog>
  );
}
