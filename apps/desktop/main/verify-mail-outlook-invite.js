// Mail: an invitation as Outlook sends it (iMIP, Mail 2.25). The calendar is
// a third alternative beside the plain and HTML text, with no file name and
// no Content-Disposition. Mail lists it with the attachments, so the reading
// pane shows the calendar's card for it (calendar-invite-card.js) with
// Accept, Tentative and Decline. Nothing is answered here.
// Run alone with RUTBA_VERIFY_ONLY=mail-outlook-invite.

export async function verifyMailOutlookInvite({ open, check, until, wait, errorsIn }) {
  const stamp = Date.now();
  const subject = `Outlook invite ${stamp}`;
  const pad = (n) => String(n).padStart(2, '0');
  const utc = (ms) => { const d = new Date(ms); return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}00Z`; };
  const start = stamp + 3 * 86_400_000;
  const ics = [
    'BEGIN:VCALENDAR', 'METHOD:REQUEST', 'PRODID:Microsoft Exchange Server 2010', 'VERSION:2.0', 'BEGIN:VEVENT',
    'ORGANIZER;CN=Amina Yusuf:mailto:amina@work.example',
    'ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=You:mailto:you@example.com',
    `SUMMARY;LANGUAGE=en-GB:${subject}`, `DTSTART:${utc(start)}`, `DTEND:${utc(start + 3_600_000)}`,
    `UID:040000008200E00074C5B7101A82E008${stamp}`, 'SEQUENCE:0', `DTSTAMP:${utc(stamp)}`, 'LOCATION:Room 3', 'END:VEVENT', 'END:VCALENDAR', '',
  ].join('\r\n');
  const raw = [
    'From: Amina Yusuf <amina@work.example>', 'To: You <you@example.com>', `Subject: ${subject}`,
    `Date: ${new Date(stamp + 86_400_000).toUTCString()}`, `Message-ID: <outlook-${stamp}@work.example>`, 'MIME-Version: 1.0',
    'Content-Type: multipart/alternative; boundary="_000_alt_"', '',
    '--_000_alt_', 'Content-Type: text/plain; charset="us-ascii"', '', 'Shall we go through the numbers?', '',
    '--_000_alt_', 'Content-Type: text/html; charset="us-ascii"', '', '<p>Shall we go through the numbers?</p>', '',
    '--_000_alt_', 'Content-Type: text/calendar; charset="utf-8"; method=REQUEST', 'Content-Transfer-Encoding: base64', '',
    Buffer.from(ics).toString('base64').replace(/.{76}/g, '$&\r\n'), '',
    '--_000_alt_--', '',
  ].join('\r\n');
  let win = null;
  let accountId = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    await js(`window.rutbaOffice.mail.deliverTest(${JSON.stringify({ accountId, folder: 'Inbox', raw })})`);
    const inbox = (await js(`window.rutbaOffice.mail.folders({ accountId: ${JSON.stringify(accountId)} })`)).find((f) => /^inbox$/i.test(f.name || f.path))?.path || 'Inbox';
    const openInbox = () => js(`(() => { document.querySelector('.ml-drop[data-folder="${inbox}"] button')?.click(); return 1; })()`);
    const row = `[...document.querySelectorAll('.ml-list > .ml-row')].find((r) => r.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(subject)}))`;
    await openInbox();
    await until(async () => (await js(`Boolean(${row})`)) || (await openInbox(), await wait(400), false), 'the invitation in the list', 12000).catch(() => {});
    await js(`${row}?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-reader .ml-invite'))`), 'the invitation card', 10000).catch(() => {});
    const card = await js(`(() => { const c = document.querySelector('.ml-reader .ml-invite'); return c ? { text: c.textContent, buttons: [...c.querySelectorAll('.rw-btn')].map((b) => b.textContent.trim()) } : null; })()`);
    check('mail: an invitation as Outlook sends it, its calendar an alternative with no file name, shows the calendar\'s card with Accept, Tentative and Decline',
      Boolean(card) && card.text.includes(subject) && ['Accept', 'Tentative', 'Decline'].every((b) => card.buttons.includes(b)), JSON.stringify(card));
    const complaints = await errorsIn(win);
    check('mail: the Outlook invitation reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the Outlook invitation checks ran to the end', false, err?.message || String(err));
  } finally {
    try {
      const run = (code) => win?.webContents.executeJavaScript(code);
      if (accountId) {
        for (const f of await run(`window.rutbaOffice.mail.folders({ accountId: ${JSON.stringify(accountId)} })`)) {
          const rows = (await run(`window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: ${JSON.stringify(f.path)} })`)).rows.filter((r) => r.subject === subject);
          if (rows.length) await run(`window.rutbaOffice.mail.delete({ accountId: ${JSON.stringify(accountId)}, folder: ${JSON.stringify(f.path)}, ids: ${JSON.stringify(rows.map((r) => r.id))} })`);
        }
      }
    } catch { /* the window is gone */ }
  }
}
