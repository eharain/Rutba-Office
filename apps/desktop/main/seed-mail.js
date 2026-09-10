// A believable mailbox, generated.
//
// Every check and every screenshot needs mail to look at, and the one thing
// that must never be used for either is somebody's actual inbox. So this builds
// one: a conversation that spans four messages and two people, a newsletter
// carrying a tracking pixel and an unsubscribe header, a message that passed
// its sender checks and one that failed them, an HTML message, and one with a
// file attached.
//
// Between them they exercise every part of the client that is not simply a
// list: threading by Message-Id, threading by subject when the chain is
// missing, the tracker report, the unsubscribe offer, the authentication badge
// and the attachment view. A fixture that only proves "a row appeared" would
// let all six of those rot unnoticed.

import fs from 'node:fs';
import path from 'node:path';

const DAY = 86_400_000;
const HOUR = 3_600_000;

/** RFC 5322 on the wire: CRLF everywhere, headers then a blank line. */
function raw({ from, to = 'You <you@example.com>', subject, date, id, inReplyTo, references, headers = [], body, html, attachment }) {
  const lines = [
    `From: ${from}`,
    `To: ${to}`,
    `Subject: ${subject}`,
    `Date: ${new Date(date).toUTCString()}`,
    `Message-ID: <${id}>`,
    'MIME-Version: 1.0',
  ];
  if (inReplyTo) lines.push(`In-Reply-To: <${inReplyTo}>`);
  if (references?.length) lines.push(`References: ${references.map((r) => `<${r}>`).join(' ')}`);
  for (const [name, value] of headers) lines.push(`${name}: ${value}`);

  if (attachment) {
    const boundary = `----rutba-${id.replace(/[^a-z0-9]/gi, '')}`;
    lines.push(`Content-Type: multipart/mixed; boundary="${boundary}"`, '', '');
    return (
      `${lines.join('\r\n')}--${boundary}\r\n` +
      `Content-Type: text/plain; charset=utf-8\r\n\r\n${body}\r\n\r\n` +
      `--${boundary}\r\n` +
      `Content-Type: ${attachment.type}\r\n` +
      `Content-Disposition: attachment; filename="${attachment.filename}"\r\n` +
      `Content-Transfer-Encoding: base64\r\n\r\n` +
      `${Buffer.from(attachment.content).toString('base64')}\r\n` +
      `--${boundary}--\r\n`
    );
  }

  if (html) {
    lines.push('Content-Type: text/html; charset=utf-8', '', html, '');
  } else {
    lines.push('Content-Type: text/plain; charset=utf-8', '', body, '');
  }
  return lines.join('\r\n');
}

const PASSED = ['Authentication-Results', 'mx.example.com; spf=pass smtp.mailfrom=northwind.example; dkim=pass header.d=northwind.example; dmarc=pass'];
const FAILED = ['Authentication-Results', 'mx.example.com; spf=fail smtp.mailfrom=northbank-secure.example; dkim=none; dmarc=fail'];

export function seedMessages(now = Date.now()) {
  const messages = [];
  const push = (fields) => messages.push({ from: { address: fields.address }, date: new Date(fields.date).toISOString(), raw: raw(fields) });

  /* An invitation, the way a calendar sends one: a text/calendar part with METHOD:REQUEST. */
  const meeting = new Date(now + 3 * 86400000);
  meeting.setHours(14, 0, 0, 0);
  const stamp = (d) => d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
  push({
    address: 'amina@northwind.example',
    from: 'Amina Yusuf <amina@northwind.example>',
    subject: 'Invitation: Quarterly numbers',
    date: now - 2 * 3600000,
    id: 'invite-1@northwind.example',
    headers: [PASSED],
    body: 'You have been invited to Quarterly numbers.',
    attachment: {
      filename: 'invite.ics',
      type: 'text/calendar; method=REQUEST; charset=utf-8',
      content: ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Northwind//Calendar//EN', 'METHOD:REQUEST', 'BEGIN:VEVENT', 'UID:invite-1@northwind.example', `DTSTAMP:${stamp(new Date(now))}`, `DTSTART:${stamp(meeting)}`, `DTEND:${stamp(new Date(meeting.getTime() + 1800000))}`, 'SUMMARY:Quarterly numbers', 'LOCATION:Room 4', 'ORGANIZER;CN=Amina Yusuf:mailto:amina@northwind.example', 'ATTENDEE;CN=You;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:you@example.com', 'ATTENDEE;CN=Amina Yusuf;PARTSTAT=ACCEPTED:mailto:amina@northwind.example', 'SEQUENCE:0', 'STATUS:CONFIRMED', 'END:VEVENT', 'END:VCALENDAR', ''].join('\r\n'),
    },
  });

  /* A four-message conversation, chained properly. Two people, one subject. */
  push({
    address: 'amina@northwind.example',
    from: 'Amina Yusuf <amina@northwind.example>',
    subject: 'Q3 numbers are in',
    date: now - 3 * DAY,
    id: 'thread-1a@northwind.example',
    headers: [PASSED],
    body: 'The northern region closed 18% up. Full workbook attached to the board pack — the totals are computed, not typed.',
  });
  push({
    address: 'you@example.com',
    from: 'You <you@example.com>',
    to: 'Amina Yusuf <amina@northwind.example>',
    subject: 'Re: Q3 numbers are in',
    date: now - 3 * DAY + 2 * HOUR,
    id: 'thread-1b@example.com',
    inReplyTo: 'thread-1a@northwind.example',
    references: ['thread-1a@northwind.example'],
    body: 'That is better than the forecast. Is the 18% like for like, or does it include the two sites we opened in July?',
  });
  push({
    address: 'amina@northwind.example',
    from: 'Amina Yusuf <amina@northwind.example>',
    subject: 'Re: Q3 numbers are in',
    date: now - 2 * DAY,
    id: 'thread-1c@northwind.example',
    inReplyTo: 'thread-1b@example.com',
    references: ['thread-1a@northwind.example', 'thread-1b@example.com'],
    headers: [PASSED],
    body: 'Like for like. The new sites are in a separate line — they contributed 3.1% on top of that.',
  });
  push({
    address: 'tomas@berg-partners.example',
    from: 'Tomas Berg <tomas@berg-partners.example>',
    // No In-Reply-To: forwarded by hand, which is how most threads lose their
    // chain. Subject matching has to catch this one.
    subject: 'RE: Q3 numbers are in',
    date: now - 2 * DAY + 4 * HOUR,
    id: 'thread-1d@berg-partners.example',
    body: 'Adding Tomas. The board pack goes out Friday, so anything you want changed should come to me by Thursday lunchtime.',
  });

  /* A newsletter: tracking pixel, unsubscribe, HTML. */
  push({
    address: 'news@retail-weekly.example',
    from: 'Retail Weekly <news@retail-weekly.example>',
    subject: 'This week: five shops that got pricing right',
    date: now - 1 * DAY,
    id: 'news-1@retail-weekly.example',
    headers: [
      ['List-Id', 'Retail Weekly <news.retail-weekly.example>'],
      ['List-Unsubscribe', '<https://retail-weekly.example/u/8821>, <mailto:unsubscribe@retail-weekly.example?subject=unsubscribe%208821>'],
      ['List-Unsubscribe-Post', 'List-Unsubscribe=One-Click'],
      ['Precedence', 'bulk'],
    ],
    html:
      '<html><body><h1>Retail Weekly</h1>' +
      '<p>Five shops that got pricing right, and one that did not.</p>' +
      '<p><a href="https://retail-weekly.example/read/8821">Read the issue</a></p>' +
      '<img src="https://track.list-manage.com/open?id=8821" width="1" height="1" alt="">' +
      '<img src="https://cdn.retail-weekly.example/logo.png" width="180" alt="Retail Weekly">' +
      '<img src="https://www.google-analytics.com/collect?v=1&t=event" width="1" height="1" alt="">' +
      '</body></html>',
  });

  /* One that failed its checks. Worth seeing before it is opened.
     The brand is invented, like every other company in this file. It was a
     real payment brand with a digit swapped into
     the domain, which is a good imitation of the genuine article and a bad
     thing to publish: these messages reach the website through make-screens,
     and a marketing screenshot should not carry somebody else's trademark on
     a phishing example. The lookalike domain, the failed SPF and DKIM and the
     bait link all still do their job. */
  push({
    address: 'security@northbank-secure.example',
    from: '"Northbank Pay Security" <security@northbank-secure.example>',
    subject: 'Your account has been limited',
    date: now - 20 * HOUR,
    id: 'phish-1@northbank-secure.example',
    headers: [FAILED],
    html: '<html><body><p>We have limited your account. <a href="https://northbank-secure.example/verify">Verify now</a> to restore access.</p></body></html>',
  });

  /* Ordinary mail, one with a file on it. */
  push({
    address: 'billing@hosting.example',
    from: 'Accounts <billing@hosting.example>',
    subject: 'Invoice 88421',
    date: now - 14 * HOUR,
    id: 'invoice-1@hosting.example',
    body: 'Your invoice for September is attached. No action is needed — payment is by direct debit on the 4th.',
    attachment: {
      filename: 'invoice-88421.csv',
      type: 'text/csv',
      content: 'Description,Quantity,Unit,Amount\nHosting — September,1,42.00,42.00\nBackups,1,8.00,8.00\nTotal,,,50.00\n',
    },
  });
  push({
    address: 'priya@lattice.example',
    from: 'Priya Raman <priya@lattice.example>',
    subject: 'Design review Thursday',
    date: now - 9 * HOUR,
    id: 'design-1@lattice.example',
    headers: [PASSED],
    body: 'I have put the three options in the deck. My preference is the second, but I want to hear the room first.',
  });
  push({
    address: 'dan@okafor.example',
    from: 'Dan Okafor <dan@okafor.example>',
    subject: 'Lunch?',
    date: now - 5 * HOUR,
    id: 'lunch-1@okafor.example',
    body: 'Free on Friday if you are. There is a new place near the station that is supposed to be good.',
  });
  push({
    address: 'ci@build.example',
    from: 'Release bot <ci@build.example>',
    subject: 'Build 4471 passed',
    date: now - 2 * HOUR,
    id: 'ci-1@build.example',
    headers: [['Precedence', 'auto-generated']],
    body: 'All 569 tests green in 4.1 s. Artefacts are on the usual share.',
  });

  return messages;
}

/**
 * Write the fixture as an mbox and put it through the real import path, so the
 * seeded store is produced by the same code a person's archive goes through.
 */
export async function seedMail({ stores, mail, name = 'smoke-seed' }) {
  const { writeMbox } = await import('@rutba/mailbox/mbox');
  // The import names the account and the folder after the file, so this string
  // ends up in the sidebar, the window title, the search box and the status
  // bar. `smoke-seed` is the honest name for a smoke run and the wrong one for
  // a picture of a product — a screenshot captioned "smoke-seed — smoke-seed"
  // reads as a test artefact, which is what it is and not what it is for.
  // `make-screens.js` passes `Archive`.
  const file = path.join(stores.dir, `${name}.mbox`);
  fs.writeFileSync(file, Buffer.from(writeMbox(seedMessages())));
  try {
    return mail.import({ path: file, folders: null });
  } finally {
    fs.rmSync(file, { force: true });
  }
}
