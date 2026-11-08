// A labelled on/off switch, the studio's own (the recorder dialog's Toggle
// goes with that dialog).

import React from 'react';
import { Icon } from '@rutba/office-ui';

export function Toggle({ on, onChange, label, icon, disabled, className = '' }) {
  return (
    <label className={`rec-toggle${disabled ? ' off' : ''}${className ? ` ${className}` : ''}`}>
      {icon ? <Icon name={icon} size={16} /> : null}
      <span className="rec-toggle-label">{label}</span>
      <input type="checkbox" checked={Boolean(on)} disabled={disabled} onChange={(e) => onChange(e.target.checked)} />
      <span className="rec-switch" aria-hidden="true" />
    </label>
  );
}
