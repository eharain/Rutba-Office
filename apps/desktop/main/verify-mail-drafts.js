// Mail: a draft opens to be finished, not read (owner, 2026-10-20: "a draft
// email is not editable at all it opens it in read view only").
//
// A draft is saved with To, Cc, Bcc, a subject, formatted words and an
// attachment. Opened from Drafts, it is in Compose with all of them, not in
// the reading pane. Its subject is changed and it is saved: Drafts still has
// one draft, with the new subject and the attachment, not two. Opened again
// and sent, it leaves Drafts. Run alone with RUTBA_VERIFY_ONLY=mail-drafts.

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 */
export async function verifyMailDrafts(h) {
  const { open, check, until, wait, errorsIn } = h;
  const stamp = Date.now();
  const subject = `Draft: the quarter's numbers ${stamp}`;
  let win = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account', 8000);
    const accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    // The draft, as Save draft keeps one: everything a person wrote.
    const file = await js(`window.rutbaOffice.fs.temp({ ext: 'txt', bytes: new TextEncoder().encode('Figures attached.') }).then((r) => r.path)`);
    const draft = { to: 'Amina Yusuf <amina@work.example>', cc: 'tomas@berg-partners.example', bcc: 'files@work.example', subject, text: 'The board pack goes out Friday.', html: '<p>The board pack goes out <b>Friday</b>.</p>', attachments: [{ filename: 'figures.txt', path: file, size: 17 }] };
    await js(`window.rutbaOffice.mail.saveDraft({ accountId: ${JSON.stringify(accountId)}, draft: ${JSON.stringify(draft)} })`);
    // The window reads its folders again, as it does when mail arrives.
    await win.webContents.reload();
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the account again', 8000);
    const drafts = () => js(`window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: 'Drafts', query: '', limit: 200 }).then((r) => (r.rows || r).filter((m) => (m.subject || '').includes(${JSON.stringify(String(stamp))})).map((m) => ({ id: m.id, subject: m.subject })))`);
    const openDrafts = () => js(`(() => { const f = [...document.querySelectorAll('[title="Drafts"]')].find((n) => n.textContent.trim().startsWith('Drafts')); f?.click(); return Boolean(f); })()`);
    const openRow = (words) => js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(String(stamp))}) && row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(words)})); if (!r) return false; r.click(); return true; })()`);
    const field = (name) => `[...document.querySelectorAll('.rw-dialog .rw-field, .ml-compose-page .rw-field')].find((f) => f.querySelector('label')?.textContent === ${JSON.stringify(name)})?.querySelector('input')?.value ?? null`;

    await until(async () => (await openDrafts()) && (await wait(400), await openRow('Draft:')), 'the draft in Drafts', 10000).catch(() => {});
    const composing = await until(() => js(`Boolean(document.querySelector('.rw-dialog .ml-form'))`), 'Compose', 8000).then(() => true).catch(() => false);
    const shown = await js(`({ to: ${field('To')}, cc: ${field('Cc')}, bcc: ${field('Bcc')}, subject: ${field('Subject')}, files: [...document.querySelectorAll('.rw-dialog .ml-attachment')].map((n) => n.textContent) })`);
    check('mail drafts: a draft opened from Drafts is in Compose, not the reading pane, with its To, Cc, Bcc, subject and attachment',
      composing && /amina@work\.example/.test(shown.to || '') && /tomas@/.test(shown.cc || '') && /files@/.test(shown.bcc || '') && shown.subject === subject && shown.files.some((x) => x.includes('figures.txt')),
      JSON.stringify({ composing, shown }));

    // Full view (owner, 2026-10-20: "it should have conversion to full view as normal email"):
    // the message fills the window with what was written, Smaller puts it back in its box, and the choice is kept.
    await js(`document.querySelector('.ml-compose-full')?.click(), 1`);
    const fullShown = await until(() => js(`Boolean(document.querySelector('.ml-compose-page .ml-form')) && !document.querySelector('.rw-dialog .ml-form')`), 'the full view', 4000).then(() => true).catch(() => false);
    const inFull = await js(`({ subject: ${field('Subject')}, cc: ${field('Cc')}, body: document.querySelector('.ml-compose-page .ml-rich')?.textContent || '', files: document.querySelectorAll('.ml-compose-page .ml-attachment').length, wide: Math.round(document.querySelector('.ml-compose-page')?.getBoundingClientRect().width || 0), window: window.innerWidth })`);
    const keptFull = await js(`window.rutbaOffice.store.get({ key: 'mail.composeFull', fallback: null })`);
    await js(`document.querySelector('.ml-compose-small')?.click(), 1`);
    const backInBox = await until(() => js(`Boolean(document.querySelector('.rw-dialog .ml-form')) && !document.querySelector('.ml-compose-page')`), 'the box again', 4000).then(() => true).catch(() => false);
    const body = await js(`document.querySelector('.rw-dialog .ml-rich')?.textContent || ''`);
    const keptBox = await js(`window.rutbaOffice.store.get({ key: 'mail.composeFull', fallback: null })`);
    check('mail drafts: Full view puts the message in the whole window with everything written in it, Smaller puts it back in its box, and the choice is kept for the next message',
      fullShown && inFull.subject === subject && /tomas@/.test(inFull.cc || '') && /board pack/.test(inFull.body) && inFull.files === 1 && inFull.wide >= inFull.window - 2 && keptFull === true && backInBox && /board pack/.test(body) && keptBox === false,
      JSON.stringify({ fullShown, inFull, keptFull, backInBox, body: body.slice(0, 40), keptBox }));

    // A new subject, saved: one draft, the new subject, the attachment kept.
    await js(`(() => { const n = [...document.querySelectorAll('.rw-dialog .rw-field')].find((f) => f.querySelector('label')?.textContent === 'Subject')?.querySelector('input'); if (!n) return 0; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(n, ${JSON.stringify(`Draft, edited: the quarter's numbers ${stamp}`)}); n.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(200);
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn, .rw-dialog button')].find((b) => b.textContent.trim() === 'Save draft')?.click(), 1`);
    await until(() => js(`!document.querySelector('.rw-dialog .ml-form')`), 'Compose closed', 6000).catch(() => {});
    await wait(500);
    const after = await drafts();
    const kept = after.length === 1 ? await js(`window.rutbaOffice.mail.message({ accountId: ${JSON.stringify(accountId)}, folder: 'Drafts', id: ${JSON.stringify(after[0].id)} })`) : null;
    check('mail drafts: saved again it replaces itself: one draft, with the new subject, its Cc and its attachment',
      after.length === 1 && /edited/.test(after[0].subject) && (kept?.cc || []).length === 1 && (kept?.attachments || []).some((a) => a.filename === 'figures.txt'),
      JSON.stringify({ after, cc: kept?.cc, files: (kept?.attachments || []).map((a) => a.filename) }));

    // Opened again and sent: it leaves Drafts.
    await openDrafts();
    await wait(400);
    await until(() => openRow('edited'), 'the edited draft', 8000).catch(() => {});
    await until(() => js(`Boolean(document.querySelector('.rw-dialog .ml-form'))`), 'Compose again', 8000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn, .rw-dialog button')].find((b) => b.textContent.trim() === 'Send' && !b.disabled)?.click(), 1`);
    const gone = await until(async () => (await drafts()).length === 0, 'the draft gone', 8000).then(() => true).catch(() => false);
    check('mail drafts: a draft sent leaves Drafts', gone, JSON.stringify(await drafts()));

    const complaints = await errorsIn(win);
    check('mail drafts: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail drafts: the check ran to the end', false, err?.message || String(err));
  }
}
