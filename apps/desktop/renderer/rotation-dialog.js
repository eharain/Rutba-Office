// Arrange → Rotate → More Rotation Options…, shared by Documents,
// Worksheets and Presentations: the Rotation box of Office's Size settings,
// the angle the picked object stands at in whole degrees clockwise, typed or
// stepped. OK turns it to that angle; the frame turns with it.

import React, { useState } from 'react';
import { Button, Dialog, Field, t } from '@rutba/office-ui';

/** An angle as Office keeps one: 0 to 359.9, clockwise. */
export const normalAngle = (deg) => Math.round(((((Number(deg) || 0) % 360) + 360) % 360) * 10) / 10;

export function RotationDialog({ initial = 0, onClose, onApply }) {
  const [value, setValue] = useState(String(normalAngle(initial)));
  const ok = Number.isFinite(Number(value)) && value.trim() !== '';
  const apply = () => { if (ok) { onApply(normalAngle(value)); onClose(); } };
  return (
    <Dialog
      title={t('Rotation')}
      width={300}
      onClose={onClose}
      className="rotation-dialog"
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="rotation-ok" disabled={!ok} onClick={apply} />
        </>
      }
    >
      <Field label={t('Rotation')} hint={t('Degrees clockwise; 0 stands it upright.')}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <input
            className="rw-input rotation-angle"
            type="number"
            min={-360}
            max={360}
            step={1}
            value={value}
            autoFocus
            style={{ width: 96 }}
            onFocus={(e) => e.target.select()}
            onChange={(e) => setValue(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); apply(); } }}
          />
          <span aria-hidden="true">°</span>
        </div>
      </Field>
    </Dialog>
  );
}
