// Mbox and Outlook for Mac archives of any size, written a message at a time.
//
// For the checks that an archive past 2 GB opens where it lies: the seeded
// corpus (mail-corpus.js) written straight to the file, so a 2.5 GB archive
// costs the test a few seconds and no memory. The last message is one the
// tests look for - written after everything else, so it sits past the 2 GB
// mark - and `zip64` writes an .olm's directory with the zip64 records an
// archive past 4 GB has, however small it is.

import fs from 'node:fs';
import zlib from 'node:zlib';
import { corpus } from './mail-corpus.js';

/** The message written last, past the 2 GB mark in a big archive. */
export const LAST = {
  subject: 'Harbour lease: the final signed copy',
  from: { name: 'Amina Rahman', address: 'amina@northwind.example' },
  to: [{ name: 'Archive Owner', address: 'owner@example.com' }],
  date: '2024-12-30T09:30:00.000Z',
  messageId: 'harbour-final@northwind.example',
  body: 'The quartzberry clause is signed on both sides. The lease runs from January.',
  attachments: [{ filename: 'harbour-lease-signed.pdf', type: 'application/pdf', bytes: Buffer.from('%PDF-1.4\n% the signed lease\n') }],
};

/** Writes in blocks of a few megabytes rather than a call per piece. */
function sink(filePath) {
  const fd = fs.openSync(filePath, 'w');
  const block = Buffer.allocUnsafe(8 * 1048576);
  let used = 0;
  let at = 0;
  const flush = () => {
    if (used) fs.writeSync(fd, block, 0, used);
    used = 0;
  };
  return {
    get at() {
      return at;
    },
    write(buf) {
      if (typeof buf === 'string') buf = Buffer.from(buf, 'utf8');
      at += buf.length;
      if (buf.length > block.length) {
        flush();
        fs.writeSync(fd, buf);
        return;
      }
      if (used + buf.length > block.length) flush();
      buf.copy(block, used);
      used += buf.length;
    },
    close() {
      flush();
      fs.closeSync(fd);
    },
  };
}

const addr = (p) => (p.name ? `${p.name} <${p.address}>` : p.address);
const base64Lines = (bytes) => Buffer.from(bytes).toString('base64').replace(/.{76}/g, '$&\r\n');

/** One message as RFC 5322, an attachment as a MIME part. */
function rfc822(m) {
  const head = [`From: ${addr(m.from)}`, `To: ${m.to.map(addr).join(', ')}`, ...(m.cc?.length ? [`Cc: ${m.cc.map(addr).join(', ')}`] : []), `Subject: ${m.subject}`, `Date: ${new Date(m.date).toUTCString()}`, `Message-ID: <${m.messageId}>`, 'MIME-Version: 1.0'];
  if (!m.attachments?.length) return `${[...head, 'Content-Type: text/plain; charset=utf-8', '', m.body, ''].join('\r\n')}`;
  const b = `b-${m.messageId.replace(/[^a-z0-9]/gi, '')}`;
  const parts = [`--${b}`, 'Content-Type: text/plain; charset=utf-8', '', m.body, ''];
  for (const a of m.attachments) parts.push(`--${b}`, `Content-Type: ${a.type}; name="${a.filename}"`, 'Content-Transfer-Encoding: base64', `Content-Disposition: attachment; filename="${a.filename}"`, '', base64Lines(a.bytes), '');
  parts.push(`--${b}--`, '');
  return [...head, `Content-Type: multipart/mixed; boundary="${b}"`, '', ...parts].join('\r\n');
}

/**
 * An mboxrd of `count` corpus messages and LAST.
 * @returns {{ bytes: number, messages: number }}
 */
export function writeCorpusMbox(filePath, { count, attachBytes = 190_000, attachEvery = 1, seed } = {}) {
  const out = sink(filePath);
  let n = 0;
  const put = (m) => {
    out.write(`From ${m.from.address} ${new Date(m.date).toUTCString()}\n`);
    out.write(rfc822(m).replace(/\r\n/g, '\n').replace(/^(>*From )/gm, '>$1'));
    out.write('\n');
    n++;
  };
  for (const { message } of corpus({ count, attachBytes, attachEvery, ...(seed ? { seed } : {}) })) put(message);
  put(LAST);
  out.close();
  return { bytes: out.at, messages: n };
}

const xml = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const people = (tag, list) => (list?.length ? `<${tag}>${list.map((p) => `<emailAddress OPFContactEmailAddressAddress="${xml(p.address)}" OPFContactEmailAddressName="${xml(p.name)}"/>`).join('')}</${tag}>` : '');

function olmXml(m) {
  const attachments = (m.attachments || []).map((a) => `<messageAttachment OPFAttachmentName="${xml(a.filename)}" OPFAttachmentContentType="${xml(a.type)}" OPFAttachmentContentLength="${a.bytes.length}">${Buffer.from(a.bytes).toString('base64')}</messageAttachment>`).join('');
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n<emails><email OPFMessageIsRead="${m.unread ? 0 : 1}">` +
    `<OPFMessageCopySubject>${xml(m.subject)}</OPFMessageCopySubject>` +
    people('OPFMessageCopyFromAddresses', [m.from]) +
    people('OPFMessageCopyToAddresses', m.to) +
    people('OPFMessageCopyCCAddresses', m.cc) +
    `<OPFMessageCopySentTime>${m.date.replace(/\.\d{3}Z$/, '')}</OPFMessageCopySentTime>` +
    `<OPFMessageCopyMessageID>&lt;${xml(m.messageId)}&gt;</OPFMessageCopyMessageID>` +
    `<OPFMessageCopyBody>${xml(m.body)}</OPFMessageCopyBody>` +
    (attachments ? `<OPFMessageCopyAttachmentList>${attachments}</OPFMessageCopyAttachmentList>` : '') +
    '</email></emails>\n'
  );
}

/**
 * An .olm of `count` corpus messages and LAST (in Inbox), a zip entry a
 * message: a message with an attachment stored, one without deflated, as an
 * archive's mix is. `zip64` writes the zip64 records for every entry and the
 * directory, as an archive past 4 GB must; past 4 GB they are written anyway.
 * @returns {{ bytes: number, messages: number }}
 */
export function writeCorpusOlm(filePath, { count, attachBytes = 190_000, attachEvery = 1, seed, zip64 = false } = {}) {
  const out = sink(filePath);
  const central = [];
  let n = 0;
  const put = (folder, m) => {
    n++;
    const name = Buffer.from(`Local/${folder}/Messages/message_${String(n).padStart(6, '0')}.xml`, 'utf8');
    const data = Buffer.from(olmXml(m), 'utf8');
    const method = m.attachments?.length ? 0 : 8;
    const payload = method === 8 ? zlib.deflateRawSync(data, { level: 1 }) : data;
    const crc = zlib.crc32(data) >>> 0;
    const offset = out.at;
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(45, 4);
    local.writeUInt16LE(0x800, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(payload.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    out.write(local);
    out.write(name);
    out.write(payload);
    central.push({ name, method, crc, compressed: payload.length, size: data.length, offset });
  };
  for (const { folder, message } of corpus({ count, attachBytes, attachEvery, ...(seed ? { seed } : {}) })) put(folder, message);
  put('Inbox', LAST);

  const cdStart = out.at;
  for (const e of central) {
    const big = zip64 || e.offset >= 0xffffffff;
    const extra = big ? Buffer.alloc(28) : Buffer.alloc(0);
    if (big) {
      extra.writeUInt16LE(0x0001, 0);
      extra.writeUInt16LE(24, 2);
      extra.writeBigUInt64LE(BigInt(e.size), 4);
      extra.writeBigUInt64LE(BigInt(e.compressed), 12);
      extra.writeBigUInt64LE(BigInt(e.offset), 20);
    }
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0);
    c.writeUInt16LE(45, 4);
    c.writeUInt16LE(45, 6);
    c.writeUInt16LE(0x800, 8);
    c.writeUInt16LE(e.method, 10);
    c.writeUInt32LE(e.crc, 16);
    c.writeUInt32LE(big ? 0xffffffff : e.compressed, 20);
    c.writeUInt32LE(big ? 0xffffffff : e.size, 24);
    c.writeUInt16LE(e.name.length, 28);
    c.writeUInt16LE(extra.length, 30);
    c.writeUInt32LE(big ? 0xffffffff : e.offset, 42);
    out.write(c);
    out.write(e.name);
    out.write(extra);
  }
  const cdSize = out.at - cdStart;
  const big = zip64 || cdStart >= 0xffffffff || central.length >= 0xffff;
  if (big) {
    const at = out.at;
    const rec = Buffer.alloc(56);
    rec.writeUInt32LE(0x06064b50, 0);
    rec.writeBigUInt64LE(44n, 4);
    rec.writeUInt16LE(45, 12);
    rec.writeUInt16LE(45, 14);
    rec.writeBigUInt64LE(BigInt(central.length), 24);
    rec.writeBigUInt64LE(BigInt(central.length), 32);
    rec.writeBigUInt64LE(BigInt(cdSize), 40);
    rec.writeBigUInt64LE(BigInt(cdStart), 48);
    out.write(rec);
    const loc = Buffer.alloc(20);
    loc.writeUInt32LE(0x07064b50, 0);
    loc.writeBigUInt64LE(BigInt(at), 8);
    loc.writeUInt32LE(1, 16);
    out.write(loc);
  }
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(big ? 0xffff : central.length, 8);
  end.writeUInt16LE(big ? 0xffff : central.length, 10);
  end.writeUInt32LE(big ? 0xffffffff : cdSize, 12);
  end.writeUInt32LE(big ? 0xffffffff : cdStart, 16);
  out.write(end);
  out.close();
  return { bytes: out.at, messages: n };
}
