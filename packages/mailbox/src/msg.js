// Outlook .msg — [MS-OXMSG].
//
// A single message saved out of Outlook, stored as a compound file. Properties
// live in two places and you need both: fixed-length values sit inline in the
// `__properties_version1.0` stream, and variable-length values (every string
// over a few characters, every attachment's bytes) sit in their own stream
// named after the tag — `__substg1.0_0037001F` is the subject.
//
// Recipients and attachments are sub-storages numbered from zero, each with the
// same property layout, which is why one property reader serves all three.

import { CompoundFile } from '@rutba/office-formats/cfb';
import { PT, PROPS, RECIPIENT_TYPE, MSG_FLAG, decodeValue, filetimeToIso, describeTag } from './props.js';
import { parseMessage, stripHtml } from './mime.js';

const SUBSTG = '__substg1.0_';
const PROPERTIES = '__properties_version1.0';

function tagOf(name) {
  // `__substg1.0_0037001F` → { id: 0x0037, type: 0x001f }
  const hex = name.slice(SUBSTG.length, SUBSTG.length + 8);
  const id = parseInt(hex.slice(0, 4), 16);
  const type = parseInt(hex.slice(4, 8), 16);
  return { id, type, hex };
}

/** Read every property of one storage (the message, a recipient, an attachment). */
function readProperties(cfb, storage) {
  const out = new Map();
  const children = cfb.childrenOf(storage);

  // Variable-length values, one stream each.
  for (const child of children) {
    if (child.type !== 2 || !child.name.startsWith(SUBSTG)) continue;
    const { id, type } = tagOf(child.name);
    const bytes = cfb.read(child);
    out.set(id, { id, type, name: PROPS[id] || null, value: decodeValue(type, bytes, true), bytes });
  }

  // Fixed-length values, packed 16 bytes each after a header.
  const propStream = children.find((c) => c.name === PROPERTIES);
  if (propStream) {
    const bytes = cfb.read(propStream);
    // The header is 32 bytes on a message, 24 on a recipient or attachment,
    // 8 on an embedded message. Sniffing it: entries are 16 bytes and start
    // with a tag whose type is a known PT value, so try both and take the one
    // that parses cleanly.
    for (const header of [32, 24, 8]) {
      if ((bytes.length - header) % 16 !== 0) continue;
      const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      let sane = 0;
      const found = [];
      for (let at = header; at + 16 <= bytes.length; at += 16) {
        const tag = dv.getUint32(at, true);
        const { id, type } = describeTag(tag);
        if (Object.values(PT).includes(type)) sane++;
        const valueBytes = bytes.subarray(at + 8, at + 16);
        found.push({ id, type, valueBytes });
      }
      if (!found.length || sane / found.length < 0.6) continue;
      for (const f of found) {
        if (out.has(f.id)) continue; // a substg stream is the fuller value
        if (f.type === PT.BINARY || f.type === PT.UNICODE || f.type === PT.STRING8 || f.type === PT.OBJECT) continue;
        const value =
          f.type === PT.SYSTIME
            ? filetimeToIso(
                new DataView(f.valueBytes.buffer, f.valueBytes.byteOffset, 8).getUint32(0, true),
                new DataView(f.valueBytes.buffer, f.valueBytes.byteOffset, 8).getUint32(4, true)
              )
            : decodeValue(f.type, f.valueBytes, true);
        out.set(f.id, { id: f.id, type: f.type, name: PROPS[f.id] || null, value });
      }
      break;
    }
  }
  return out;
}

const get = (props, id) => props.get(id)?.value ?? null;

function addressOf(props) {
  const smtp = get(props, 0x39fe) || get(props, 0x3003);
  const type = get(props, 0x3002);
  const name = get(props, 0x3001) || get(props, 0x5ff6);
  const address = type === 'EX' && !String(smtp || '').includes('@') ? (get(props, 0x39fe) || smtp) : smtp;
  return {
    name: name || null,
    address: address || '',
    display: name && address ? `${name} <${address}>` : name || address || '',
  };
}

/** Outlook writes an EX address for internal senders; prefer the SMTP one. */
function senderOf(props) {
  const name = get(props, 0x0c1a) || get(props, 0x0042);
  const smtp = get(props, 0x5d01) || get(props, 0x5d02);
  const legacy = get(props, 0x0c1f) || get(props, 0x0065);
  const address = smtp || (String(legacy || '').includes('@') ? legacy : '');
  return { name: name || null, address: address || legacy || '', display: name ? `${name} <${address || legacy || ''}>` : address || '' };
}

export class OutlookMessage {
  constructor(cfb, storage) {
    this.cfb = cfb;
    this.storage = storage || cfb.root;
    this.props = readProperties(cfb, this.storage);
  }

  static open(bytes) {
    const cfb = new CompoundFile(bytes);
    return new OutlookMessage(cfb);
  }

  static is(bytes) {
    if (!CompoundFile.is(bytes)) return false;
    try {
      return new CompoundFile(bytes).application() === 'msg';
    } catch {
      return false;
    }
  }

  get(idOrName) {
    const id = typeof idOrName === 'number' ? idOrName : Number(Object.entries(PROPS).find(([, n]) => n === idOrName)?.[0]);
    return get(this.props, id);
  }

  recipients() {
    return this.cfb
      .childrenOf(this.storage)
      .filter((c) => c.type === 1 && c.name.startsWith('__recip_version1.0_'))
      .map((s) => {
        const p = readProperties(this.cfb, s);
        return { ...addressOf(p), type: RECIPIENT_TYPE[get(p, 0x0c15)] || 'to' };
      });
  }

  attachments() {
    return this.cfb
      .childrenOf(this.storage)
      .filter((c) => c.type === 1 && c.name.startsWith('__attach_version1.0_'))
      .map((s, index) => {
        const p = readProperties(this.cfb, s);
        const method = get(p, 0x3705);
        const dataProp = p.get(0x3701);
        // An embedded message is an attachment whose data is a storage, not bytes.
        const embeddedStorage =
          method === 5 ? this.cfb.childrenOf(s).find((c) => c.type === 1 && c.name === '__substg1.0_3701000D') : null;
        return {
          index,
          filename: get(p, 0x3707) || get(p, 0x3704) || `attachment-${index + 1}`,
          extension: get(p, 0x3703) || null,
          type: get(p, 0x370e) || 'application/octet-stream',
          contentId: get(p, 0x3712) || null,
          size: get(p, 0x0e20) || dataProp?.bytes?.length || 0,
          inline: Boolean(get(p, 0x3712)),
          method,
          bytes: dataProp?.bytes instanceof Uint8Array ? dataProp.bytes : null,
          embedded: embeddedStorage ? new OutlookMessage(this.cfb, embeddedStorage) : null,
        };
      });
  }

  /** The message in the shape the rest of the client uses. */
  toMessage({ withBytes = true } = {}) {
    const p = this.props;
    const flags = get(p, 0x0e07) || 0;
    const headers = get(p, 0x007d);
    const recipients = this.recipients();
    const html = get(p, 0x1013);
    const body = get(p, 0x1000);

    // Some senders store HTML as PT_BINARY; decode it as UTF-8 in that case.
    const htmlText =
      typeof html === 'string' ? html : html instanceof Uint8Array ? new TextDecoder('utf-8', { fatal: false }).decode(html) : null;

    const attachments = this.attachments().map((a) => ({
      filename: a.filename,
      type: a.type,
      size: a.size,
      contentId: a.contentId,
      inline: a.inline,
      ...(withBytes && a.bytes ? { bytes: a.bytes } : {}),
      ...(a.embedded ? { embedded: a.embedded.toMessage({ withBytes: false }) } : {}),
    }));

    return {
      subject: get(p, 0x0037) || get(p, 0x0e1d) || '',
      from: [senderOf(p)].filter((x) => x.address || x.name),
      to: recipients.filter((r) => r.type === 'to'),
      cc: recipients.filter((r) => r.type === 'cc'),
      bcc: recipients.filter((r) => r.type === 'bcc'),
      date: get(p, 0x0e06) || get(p, 0x0039) || get(p, 0x3007) || null,
      messageId: get(p, 0x1035) || null,
      inReplyTo: get(p, 0x1042) || null,
      references: String(get(p, 0x1039) || '').split(/\s+/).filter(Boolean),
      text: typeof body === 'string' ? body : null,
      html: htmlText,
      preview: (typeof body === 'string' ? body : stripHtml(htmlText || '')).replace(/\s+/g, ' ').slice(0, 220),
      unread: !(flags & MSG_FLAG.READ),
      flagged: false,
      hasAttachments: Boolean(flags & MSG_FLAG.HAS_ATTACH) || attachments.length > 0,
      attachments,
      messageClass: get(p, 0x001a) || 'IPM.Note',
      internetHeaders: headers || null,
      size: get(p, 0x0e08) || 0,
      source: 'msg',
    };
  }

  /**
   * The original internet message, when Outlook kept it. Otherwise null — and
   * the caller should use `toMessage()`, which is built from the properties.
   */
  originalRfc822() {
    const headers = this.get('transportMessageHeaders');
    if (!headers) return null;
    const body = this.get('body');
    if (typeof body !== 'string') return null;
    return parseMessage(`${headers}\r\n\r\n${body}`);
  }
}

/** Convenience: bytes in, message out. */
export function readMsg(bytes) {
  return OutlookMessage.open(bytes).toMessage();
}
