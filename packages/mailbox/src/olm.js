// Outlook for Mac archives — .olm.
//
// Mercifully simple next to a .pst: a zip whose entries are XML documents, one
// per message, under paths that mirror the folder tree
// (`Local/Inbox/Messages/message_0001.xml`). Attachments are base64 inside the
// same XML rather than separate entries.
//
// The element names are Outlook's own — `OPFMessageCopySubject`,
// `OPFMessageCopySenderAddress` — and are stable across the versions that
// wrote these files.

import { readZip } from '@rutba/ooxml/zip';
import { parse, first, all, textOf } from '@rutba/office-formats/xml';
import { decodeBase64, stripHtml } from './mime.js';

const dec = new TextDecoder('utf-8', { fatal: false });

function text(node, name) {
  const el = first(node, name);
  return el ? textOf(el).trim() : '';
}

function addresses(node, name) {
  const holder = first(node, name);
  if (!holder) return [];
  return all(holder, 'emailAddress')
    .map((el) => ({
      name: el.attrs.OPFContactEmailAddressName || null,
      address: el.attrs.OPFContactEmailAddressAddress || '',
    }))
    .filter((a) => a.address || a.name)
    .map((a) => ({ ...a, display: a.name && a.address ? `${a.name} <${a.address}>` : a.name || a.address }));
}

/** Outlook writes dates as `2024-03-11T09:14:00`, without a zone. */
function when(value) {
  if (!value) return null;
  const d = new Date(/Z|[+-]\d{2}:?\d{2}$/.test(value) ? value : `${value}Z`);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

function messageFrom(el) {
  const html = text(el, 'OPFMessageCopyHTMLBody');
  const body = text(el, 'OPFMessageCopyBody');
  const attachments = all(el, 'messageAttachment').map((a, i) => {
    const b64 = textOf(a).trim();
    const bytes = b64 ? decodeBase64(b64) : null;
    return {
      filename: a.attrs.OPFAttachmentName || `attachment-${i + 1}`,
      type: a.attrs.OPFAttachmentContentType || 'application/octet-stream',
      contentId: a.attrs.OPFAttachmentContentID || null,
      size: Number(a.attrs.OPFAttachmentContentLength || bytes?.length || 0),
      inline: Boolean(a.attrs.OPFAttachmentContentID),
      bytes,
    };
  });

  return {
    subject: text(el, 'OPFMessageCopySubject'),
    from: addresses(el, 'OPFMessageCopyFromAddresses'),
    to: addresses(el, 'OPFMessageCopyToAddresses'),
    cc: addresses(el, 'OPFMessageCopyCCAddresses'),
    bcc: addresses(el, 'OPFMessageCopyBCCAddresses'),
    date: when(text(el, 'OPFMessageCopySentTime') || text(el, 'OPFMessageCopyReceivedTime')),
    messageId: text(el, 'OPFMessageCopyMessageID') || null,
    text: body || null,
    html: html || null,
    preview: (body || stripHtml(html)).replace(/\s+/g, ' ').slice(0, 220),
    unread: el.attrs.OPFMessageIsRead === '0',
    hasAttachments: attachments.length > 0,
    attachments,
    source: 'olm',
  };
}

/** Folder path from an entry name: `Local/Inbox/Messages/x.xml` → `Inbox`. */
function folderOf(entryName) {
  const parts = entryName.split('/').filter(Boolean);
  const trimmed = parts.filter((p) => p !== 'Local' && p !== 'Messages' && !p.endsWith('.xml'));
  return trimmed.join('/') || 'Archive';
}

export class OlmArchive {
  constructor(bytes) {
    const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
    const { entries } = readZip(buf);
    this.entries = entries.filter((e) => e.name.toLowerCase().endsWith('.xml') && !e.name.endsWith('/'));
  }

  static is(bytes) {
    try {
      const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
      const { entries } = readZip(buf);
      return entries.some((e) => /^Local\//i.test(e.name) && e.name.endsWith('.xml'));
    } catch {
      return false;
    }
  }

  /** Folder names and how many messages each holds. */
  folders() {
    const counts = new Map();
    for (const entry of this.entries) {
      const folder = folderOf(entry.name);
      const xml = dec.decode(entry.data);
      const n = (xml.match(/<email\b/g) || []).length;
      if (!n) continue;
      counts.set(folder, (counts.get(folder) || 0) + n);
    }
    return [...counts.entries()].map(([name, messages]) => ({ name, path: name, messages, depth: name.split('/').length - 1 }));
  }

  /** Every message, folder by folder. Yields rather than collecting. */
  *messages(folderName = null) {
    for (const entry of this.entries) {
      const folder = folderOf(entry.name);
      if (folderName && folder !== folderName) continue;
      const root = parse(dec.decode(entry.data));
      for (const el of all(root, 'email')) {
        yield { folder, ...messageFrom(el) };
      }
    }
  }

  summary() {
    const folders = this.folders();
    return {
      kind: 'olm',
      folders,
      messages: folders.reduce((n, f) => n + f.messages, 0),
    };
  }
}

export function readOlm(bytes) {
  return new OlmArchive(bytes);
}
