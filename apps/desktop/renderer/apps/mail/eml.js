// Messages as .eml files in the window (Mail 2.25): files dropped or pasted
// in, the bar over a message opened from Explorer, and the names it is saved
// under. The main process does the reading and writing (main/mail-eml.js).

import React from 'react';
import { Button, Icon, t } from '@rutba/office-ui';
import { pathOf } from '@rutba/office-shell/client';

/** A message opened from Explorer is kept under an id of this kind, never an account's. */
export const isOpenedId = (id) => typeof id === 'string' && id.startsWith('opened_');

export const isEmlName = (name) => /\.eml$/i.test(String(name || ''));

/** The file an .eml window was opened on, or null for an ordinary Mail window. */
export function openedFileOf(boot) {
  const file = boot && typeof boot.file === 'string' ? boot.file : null;
  return file && isEmlName(file) ? file : null;
}

/** Whether a drag carries files from outside the window, rather than messages from its list. */
export const carriesFiles = (e) => [...(e?.dataTransfer?.types || [])].includes('Files');

/**
 * The .eml files among those dropped or pasted, as the service takes them:
 * by path where Windows gives one, else their bytes. Answers null when none
 * of them is an .eml.
 */
export async function emlSources(list) {
  const emls = [...(list || [])].filter((f) => isEmlName(f?.name));
  if (!emls.length) return null;
  const paths = [];
  const files = [];
  for (const f of emls) {
    const p = pathOf(f);
    if (p) paths.push(p);
    else files.push({ name: f.name, bytes: new Uint8Array(await f.arrayBuffer()) });
  }
  return { paths, files, count: emls.length };
}

/** The name a message is offered to be saved under: its subject, made safe for a file name. */
export function emlNameFor(message) {
  const base = String(message?.subject || '')
    .replace(/[\u0000-\u001f<>:"/\\|?*]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
    .slice(0, 80);
  return `${base || t('mail.eml.message')}.eml`;
}

/** Over a message opened from a file: where it came from, and the way to keep it. */
export function OpenedBar({ name, onAdd, onSave }) {
  return (
    <div className="ml-opened-bar" role="note">
      <Icon name="file" size={14} />
      <span>{t('mail.eml.opened-name-read', { name })}</span>
      <Button primary icon="folder" label={t('mail.eml.add-folder')} className="ml-opened-add" onClick={onAdd} />
      {onSave ? <Button icon="download" label={t('mail.save-as-eml')} className="ml-opened-save" onClick={onSave} /> : null}
    </div>
  );
}

const EML_CSS = `
.ml-opened-bar { display: flex; align-items: center; gap: 10px; padding: 8px 14px; border-bottom: 1px solid var(--line); background: var(--panel-2, var(--panel)); color: var(--ink-2); font-size: 12.5px; flex-wrap: wrap; }
.ml-opened-bar > span { flex: 1 1 240px; min-width: 0; }
`;

let installed = false;
export function installEmlStyles() {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const style = document.createElement('style');
  style.id = 'rutba-mail-eml-css';
  style.textContent = EML_CSS;
  document.head.appendChild(style);
}
