// Mail's Safety switch, View source, and the menu in a message (owner,
// 2026-10-20: "let the people see full raw message and let them copy URLs
// and images from mail beside download attachments. please keep an extra
// safety button on top and let this copy and download happen when its
// turned off").
//
// Safety is on for every message as it opens. While it is on, a link's
// address and a picture cannot be copied, a picture cannot be saved, and an
// attachment cannot be saved or handed to another program; words chosen can
// still be copied, and an attachment still opens in Office's own apps. Off
// is for this message only: the next one opens with it on again.

import React, { useState } from 'react';
import { Dialog, Button, Icon, Spinner, formatBytes, basename, t } from '@rutba/office-ui';

/** The bytes and type of a data: address. */
export function dataUrlBytes(url) {
  const m = /^data:([^;,]*)((?:;[^;,]*)*),(.*)$/s.exec(String(url));
  if (!m) return null;
  const type = m[1] || 'application/octet-stream';
  const base64 = /;base64/i.test(m[2]);
  const text = base64 ? atob(m[3]) : decodeURIComponent(m[3]);
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i) & 0xff;
  return { bytes, type };
}

const EXT = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/bmp': 'bmp', 'image/avif': 'avif' };

/**
 * A picture in a message, as its bytes: a picture the message carries as it
 * came; one from the web fetched now, by the main process, with no cookies
 * and no referrer.
 */
export async function pictureBytes({ shell, src, message }) {
  let bytes;
  let type = '';
  let name = 'picture';
  if (/^https?:/i.test(src)) {
    ({ bytes, type, name } = await shell.mail.fetchPicture({ url: src }));
  } else if (/^data:/i.test(src)) {
    ({ bytes, type } = dataUrlBytes(src) || {});
  } else {
    const res = await fetch(src);
    bytes = new Uint8Array(await res.arrayBuffer());
    type = (res.headers.get('content-type') || '').split(';')[0];
    const cid = Object.entries(message?.inlineUrls || {}).find(([, u]) => u === src)?.[0];
    name = (message?.attachments || []).find((a) => a.contentId === cid)?.filename || name;
  }
  if (!bytes?.length) throw new Error(t('common.picture-could-not'));
  return { bytes, type, name };
}

/** A picture in a message, saved where the person chooses. */
export async function savePicture({ shell, src, message }) {
  const { bytes, type, name } = await pictureBytes({ shell, src, message });
  const ext = EXT[type] || (/\.([a-z0-9]{2,5})$/i.exec(name)?.[1] || 'png');
  const stem = String(name).replace(/\.[^.]+$/, '') || 'picture';
  const folder = (await shell.app.paths().catch(() => ({}))).pictures || '';
  const target = await shell.dialog.save({
    title: t('mail.safety.save-picture-as'),
    defaultPath: folder ? `${folder}/${stem}.${ext}` : `${stem}.${ext}`,
    filters: [{ name: ext.toUpperCase(), extensions: [ext] }],
  });
  if (!target) return null;
  await shell.fs.write({ path: target, bytes });
  return target;
}

/** The menu for a right-click in the message: what was under the pointer, as Safety allows. */
export function frameMenuItems({ at, safe, onCopy, onOpenLink, onCopyLink, onCopyPicture, onSavePicture, onSafetyOff }) {
  const items = [];
  if (at.selection) items.push({ label: t('common.copy'), icon: 'copy', run: onCopy });
  if (at.link) {
    items.push({ label: t('mail.safety.open-link'), icon: 'link', run: () => onOpenLink(at.link) });
    items.push({ label: t('mail.copy-link'), icon: 'copy', disabled: safe, run: () => onCopyLink(at.link) });
  }
  if (at.media === 'image') {
    items.push({ label: t('mail.safety.copy-picture'), icon: 'picture', disabled: safe, run: () => onCopyPicture(at) });
    items.push({ label: t('mail.safety.save-picture-as-2'), icon: 'save', disabled: safe, run: () => onSavePicture(at.src) });
  }
  if (safe && (at.link || at.media === 'image')) items.push('-', { label: t('mail.safety.turn-safety-off'), icon: 'unlock', run: onSafetyOff });
  return items;
}

/** The Safety switch at the top of a message. */
export function SafetySwitch({ on, onChange }) {
  return (
    <button
      type="button"
      className={`ml-safety${on ? ' on' : ' off'}`}
      aria-pressed={on}
      data-tip={on
        ? t('mail.safety.safety-links-pictures')
        : t('mail.safety.safety-off-message')}
      onClick={() => onChange(!on)}
    >
      <Icon name={on ? 'shield' : 'unlock'} size={14} />
      <span>{on ? t('mail.safety.safety-on') : t('mail.safety.safety-off')}</span>
    </button>
  );
}

/** View source: the whole message as it came, or rebuilt and saying so. */
export function SourceDialog({ shell, source, subject, safe, onClose, toast }) {
  const [saving, setSaving] = useState(false);
  const copyAll = async () => {
    await shell.clipboard.writeText({ text: source.text });
    toast(t('mail.safety.source-copied'), { ms: 2500 });
  };
  const saveEml = async () => {
    setSaving(true);
    try {
      const stem = String(subject || t('mail.message')).replace(/[\\/:*?"<>|]+/g, ' ').trim().slice(0, 80) || t('mail.message');
      const target = await shell.dialog.save({ title: t('mail.save-as-eml'), defaultPath: `${stem}.eml`, filters: [{ name: t('mail.safety.email-message'), extensions: ['eml'] }] });
      if (target) {
        await shell.fs.write({ path: target, bytes: source.bytes });
        toast(t('common.saved-name', { name: basename(target) }), { tone: 'good', ms: 3000 });
      }
    } catch (err) {
      toast(err?.message || String(err), { tone: 'bad', ms: 6000 });
    } finally {
      setSaving(false);
    }
  };
  return (
    <Dialog
      title={t('common.source')}
      width={900}
      className="ml-source"
      onClose={onClose}
      actions={
        <>
          {safe ? <span className="ml-source-safe"><Icon name="shield" size={13} /> {t('mail.safety.safety-turn-off')}</span> : null}
          <span style={{ flex: 1 }} />
          <Button icon="copy" label={t('common.copy-all')} className="ml-source-copy" disabled={safe || !source?.text} onClick={copyAll} />
          <Button icon="save" label={t('mail.save-as-eml-2')} className="ml-source-save" disabled={safe || !source?.bytes || saving} onClick={saveEml} />
          <Button primary label={t('common.close')} onClick={onClose} />
        </>
      }
    >
      {source?.loading ? (
        <div className="ml-source-wait"><Spinner /> {t('mail.safety.asking-server-message')}</div>
      ) : (
        <>
          <p className={`ml-source-note${source.from === 'rebuilt' ? ' rebuilt' : ''}`} data-from={source.from}>
            {source.from === 'server'
              ? t('mail.safety.message-exactly-came', { size: formatBytes(source.bytes?.length || 0) })
              : t('mail.safety.rebuilt-mail-kept', { reason: source.reason || '' })}
          </p>
          <pre className="ml-source-text selectable">{source.text}</pre>
        </>
      )}
    </Dialog>
  );
}

export const SAFETY_CSS = `
.ml-safety { display: inline-flex; align-items: center; gap: 5px; height: 26px; padding: 0 10px; border-radius: 13px; font: inherit; font-size: 12px; font-weight: 600; cursor: pointer; border: 1px solid transparent; }
.ml-safety.on { background: color-mix(in srgb, var(--good, #2bb673) 14%, transparent); color: var(--good-text, #157a47); border-color: color-mix(in srgb, var(--good, #2bb673) 40%, transparent); }
.ml-safety.off { background: color-mix(in srgb, #f5a524 16%, transparent); color: #a15c00; border-color: color-mix(in srgb, #f5a524 45%, transparent); }
:root[data-theme='dark'] .ml-safety.on { color: #7ee2ad; }
:root[data-theme='dark'] .ml-safety.off { color: #ffc56b; }
:root[data-theme='dark'] .ml-source-note.rebuilt { color: #ffc56b; }
.ml-source .rw-dialog-body { padding-top: 6px; }
.ml-source-note { margin: 0 0 8px; font-size: 12.5px; color: var(--ink-2); }
.ml-source-note.rebuilt { color: #a15c00; }
.ml-source-text { margin: 0; max-height: 60vh; overflow: auto; padding: 10px 12px; border-radius: var(--r-2); background: var(--sunken); color: var(--ink); font: 12px/1.45 ui-monospace, "Cascadia Mono", Consolas, monospace; white-space: pre-wrap; overflow-wrap: anywhere; }
.ml-source-wait { display: flex; align-items: center; gap: 10px; padding: 24px; color: var(--ink-2); }
.ml-source-safe { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-3); }
`;
