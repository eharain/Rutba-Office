// Two small mail archives and the mail of an account, about one matter.
//
// The archive search, its export for legal hold and their window checks all
// need the same thing: a few messages spread over an Outlook data file, an
// Outlook for Mac archive and an account, so that one search finds some in
// each and the answer is known. The matter is a contract with a landlord
// called Halvorsen; the other messages are there to be left out.
//
// `writeArchiveFixtures(dir)` writes the .pst and the .olm (and an mbox of
// the account's messages, for tests that want a third kind) and returns
// their paths with the messages as written. `EXPECT` is what the searches
// over them must find.

import fs from 'node:fs';
import path from 'node:path';
import { writePst } from './pst-writer.js';
import { ZipEntry, writeZip } from '../../packages/ooxml/src/zip.js';

const ejaz = { name: 'Ejaz Arain', address: 'ejaz@example.com' };
const amina = { name: 'Amina Rahman', address: 'amina@northwind.example' };
const bo = { name: 'Bo Lindqvist', address: 'bo@northwind.example' };
const it = { name: 'IT Desk', address: 'it@northwind.example' };
const carlos = { name: 'Carlos Mendes', address: 'carlos@contoso.example' };
const dana = { name: 'Dana Okafor', address: 'dana@contoso.example' };

/** Bytes that look like a file of a kind and are the same every time. */
function fileBytes(kind, size, seed) {
  const out = Buffer.alloc(size);
  let x = seed >>> 0 || 1;
  for (let i = 0; i < size; i++) {
    x ^= x << 13;
    x ^= x >>> 17;
    x ^= x << 5;
    out[i] = x & 0xff;
  }
  const head = kind === 'pdf' ? '%PDF-1.4\n' : kind === 'xlsx' ? 'PK\u0003\u0004' : '';
  out.write(head, 0, 'latin1');
  return out;
}

const at = (iso) => new Date(iso).toISOString();

export const PST_MESSAGES = {
  Inbox: [
    {
      subject: 'Halvorsen contract: signed copy',
      from: amina,
      to: [ejaz],
      cc: [bo],
      date: at('2019-03-14T09:30:00Z'),
      body: 'Attached is the signed Halvorsen contract. The indemnity clause on page four was struck out as agreed.\n\nAmina',
      messageId: 'halvorsen-1@northwind.example',
      attachments: [{ filename: 'Halvorsen contract signed.pdf', type: 'application/pdf', bytes: fileBytes('pdf', 48_000, 11) }],
    },
    {
      subject: 'Re: Halvorsen contract: signed copy',
      from: bo,
      to: [amina],
      cc: [ejaz],
      date: at('2019-03-15T10:05:00Z'),
      body: 'Thanks. Please keep the indemnity clause wording on file for the audit.\n\nBo',
      messageId: 'halvorsen-2@northwind.example',
      inReplyTo: 'halvorsen-1@northwind.example',
      unread: true,
    },
    {
      subject: 'Mailbox quota warning',
      from: it,
      to: [ejaz],
      date: at('2019-06-01T07:00:00Z'),
      body: 'Your mailbox is almost full. Archive old mail to stay under the quota.',
      messageId: 'quota-7@northwind.example',
    },
    {
      subject: 'Quarterly figures',
      from: amina,
      to: [ejaz],
      date: at('2019-09-30T16:45:00Z'),
      body: 'Revenue rose in the third quarter; the spreadsheet is attached.',
      messageId: 'q3-figures@northwind.example',
      attachments: [{ filename: 'Q3 figures.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', bytes: fileBytes('xlsx', 1_300_000, 23) }],
    },
  ],
  'Inbox/Projects': [
    {
      subject: 'Warehouse lease renewal',
      from: carlos,
      to: [ejaz],
      date: at('2019-11-02T12:00:00Z'),
      body: 'The lease renewal for the Leeds warehouse is due in January. Halvorsen are the landlords, and their agent will call.',
      messageId: 'lease-renewal@contoso.example',
    },
  ],
  'Sent Items': [
    {
      subject: 'Re: Quarterly figures',
      from: ejaz,
      to: [amina],
      date: at('2019-10-01T08:15:00Z'),
      body: 'Thank you, the figures look right.',
      messageId: 'q3-thanks@example.com',
      inReplyTo: 'q3-figures@northwind.example',
    },
  ],
};

export const OLM_MESSAGES = {
  Inbox: [
    {
      subject: 'Halvorsen site visit',
      from: carlos,
      to: [ejaz],
      date: at('2020-02-10T09:00:00Z'),
      body: 'The site visit with Halvorsen is on Thursday at ten. Bring the floor plans.',
      messageId: 'site-visit@contoso.example',
    },
    {
      subject: 'Invoice 4471',
      from: dana,
      to: [ejaz],
      date: at('2020-04-01T11:20:00Z'),
      body: 'Invoice 4471 for the warehouse fit-out is attached.',
      messageId: 'invoice-4471@contoso.example',
      attachments: [{ filename: 'invoice-4471.pdf', type: 'application/pdf', bytes: fileBytes('pdf', 21_000, 37) }],
    },
  ],
  Archive: [
    {
      subject: 'Old lease papers',
      from: carlos,
      to: [ejaz],
      date: at('2018-05-05T15:30:00Z'),
      body: 'Scans of the old lease papers, signed by Halvorsen in 2012.',
      messageId: 'old-lease@contoso.example',
      attachments: [{ filename: 'lease scans.pdf', type: 'application/pdf', bytes: fileBytes('pdf', 30_000, 41) }],
    },
  ],
};

/** The account's own messages, as RFC 5322 text a store can take. */
export const ACCOUNT_MESSAGES = [
  {
    subject: 'Halvorsen: next steps',
    from: amina,
    to: [ejaz],
    date: at('2021-01-12T09:10:00Z'),
    body: 'Next steps on the indemnity clause: their solicitor has a revised draft.',
    messageId: 'halvorsen-next@northwind.example',
  },
  {
    subject: 'Lunch on Friday',
    from: bo,
    to: [ejaz],
    date: at('2021-01-14T12:00:00Z'),
    body: 'Are you free for lunch on Friday?',
    messageId: 'lunch-friday@northwind.example',
  },
];

const addr = (p) => `${p.name} <${p.address}>`;

/** A message as it would arrive, for a store or an mbox. */
export function rfc822(m) {
  const lines = [
    `From: ${addr(m.from)}`,
    `To: ${m.to.map(addr).join(', ')}`,
    ...(m.cc?.length ? [`Cc: ${m.cc.map(addr).join(', ')}`] : []),
    `Subject: ${m.subject}`,
    `Date: ${new Date(m.date).toUTCString()}`,
    `Message-ID: <${m.messageId}>`,
    ...(m.inReplyTo ? [`In-Reply-To: <${m.inReplyTo}>`] : []),
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset=utf-8',
    '',
    m.body,
    '',
  ];
  return lines.join('\r\n');
}

const xml = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** An Outlook for Mac archive: a zip of XML, a file per message. */
export function writeOlm(filePath, folders) {
  const entries = [];
  let n = 0;
  const people = (tag, list) =>
    list?.length ? `<${tag}>${list.map((p) => `<emailAddress OPFContactEmailAddressAddress="${xml(p.address)}" OPFContactEmailAddressName="${xml(p.name)}"/>`).join('')}</${tag}>` : '';
  for (const [folder, messages] of Object.entries(folders)) {
    for (const m of messages) {
      n++;
      const attachments = (m.attachments || [])
        .map((a) => `<messageAttachment OPFAttachmentName="${xml(a.filename)}" OPFAttachmentContentType="${xml(a.type)}" OPFAttachmentContentLength="${a.bytes.length}">${Buffer.from(a.bytes).toString('base64')}</messageAttachment>`)
        .join('');
      const body =
        `<?xml version="1.0" encoding="UTF-8"?>\n<emails><email OPFMessageIsRead="${m.unread ? 0 : 1}">` +
        `<OPFMessageCopySubject>${xml(m.subject)}</OPFMessageCopySubject>` +
        people('OPFMessageCopyFromAddresses', [m.from]) +
        people('OPFMessageCopyToAddresses', m.to) +
        people('OPFMessageCopyCCAddresses', m.cc) +
        `<OPFMessageCopySentTime>${m.date.replace(/\.\d{3}Z$/, '')}</OPFMessageCopySentTime>` +
        `<OPFMessageCopyMessageID>&lt;${xml(m.messageId)}&gt;</OPFMessageCopyMessageID>` +
        `<OPFMessageCopyBody>${xml(m.body)}</OPFMessageCopyBody>` +
        (attachments ? `<OPFMessageCopyAttachmentList>${attachments}</OPFMessageCopyAttachmentList>` : '') +
        '</email></emails>\n';
      const entry = new ZipEntry({ name: `Local/${folder}/Messages/message_${String(n).padStart(4, '0')}.xml`, method: 8 });
      entry.data = Buffer.from(body, 'utf8');
      entries.push(entry);
    }
  }
  fs.writeFileSync(filePath, writeZip(entries));
  return filePath;
}

/** Write the fixtures into `dir`; returns their paths. */
export function writeArchiveFixtures(dir) {
  fs.mkdirSync(dir, { recursive: true });
  const pst = path.join(dir, 'Northwind 2019.pst');
  const olm = path.join(dir, 'Contoso archive.olm');
  const mbox = path.join(dir, 'Account export.mbox');
  writePst(
    pst,
    [
      { name: 'Inbox', messages: PST_MESSAGES.Inbox, children: [{ name: 'Projects', messages: PST_MESSAGES['Inbox/Projects'] }] },
      { name: 'Sent Items', messages: PST_MESSAGES['Sent Items'] },
    ],
    { name: 'Northwind 2019' }
  );
  writeOlm(olm, OLM_MESSAGES);
  fs.writeFileSync(
    mbox,
    ACCOUNT_MESSAGES.map((m) => `From ${m.from.address} ${new Date(m.date).toUTCString()}\n${rfc822(m).replace(/\r\n/g, '\n')}\n`).join('')
  );
  return { pst, olm, mbox };
}

/**
 * What the searches must find, over the .pst, the .olm and the account.
 * Each is the subjects, in no order.
 */
export const EXPECT = {
  halvorsen: ['Halvorsen contract: signed copy', 'Re: Halvorsen contract: signed copy', 'Warehouse lease renewal', 'Halvorsen site visit', 'Old lease papers', 'Halvorsen: next steps'],
  'halvorsen after:2019-01-01 before:2021-01-01': ['Halvorsen contract: signed copy', 'Re: Halvorsen contract: signed copy', 'Warehouse lease renewal', 'Halvorsen site visit'],
  '"indemnity clause"': ['Halvorsen contract: signed copy', 'Re: Halvorsen contract: signed copy', 'Halvorsen: next steps'],
  'from:amina has:attachment': ['Halvorsen contract: signed copy', 'Quarterly figures'],
  'filename:pdf': ['Halvorsen contract: signed copy', 'Invoice 4471', 'Old lease papers'],
  'size>1mb': ['Quarterly figures'],
  'cc:bo': ['Halvorsen contract: signed copy'],
  'halvorsen -from:bo': ['Halvorsen contract: signed copy', 'Warehouse lease renewal', 'Halvorsen site visit', 'Old lease papers', 'Halvorsen: next steps'],
  'to:amina': ['Re: Halvorsen contract: signed copy', 'Re: Quarterly figures'],
  'invoice OR quota': ['Invoice 4471', 'Mailbox quota warning'],
  'in:projects': ['Warehouse lease renewal'],
  'subject:"halvorsen contract" is:unread': ['Re: Halvorsen contract: signed copy'],
  'received:2019-03': ['Halvorsen contract: signed copy', 'Re: Halvorsen contract: signed copy'],
};
