// Mail: delete beside each message in the list (Mail 2.25). The control sits
// at the row's end, after the star; it shows on the row under the pointer and
// on the selected row, and not on the rest. Pressed, it deletes that message
// (to the Trash, or for good where there is none), whatever else is
// selected, and the selection stays where it was. The Delete key works as it
// did. Run alone with RUTBA_VERIFY_ONLY=mail-row-delete.

import { labelKeyed } from './verify-keys.js';

export async function verifyMailRowDelete({ open, check, until, wait, errorsIn }) {
  const stamp = Date.now();
  const subjects = ['alpha', 'bravo', 'charlie'].map((w) => `Row delete: ${w} ${stamp}`);
  const raw = (subject, n) => [
    'From: Amina Yusuf <amina@work.example>', 'To: You <you@example.com>', `Subject: ${subject}`,
    // A day ahead, so the three are the top of the list together.
    `Date: ${new Date(stamp + 86_400_000 + n * 60_000).toUTCString()}`,
    `Message-ID: <row-delete-${n}-${stamp}@example.org>`, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=utf-8', '',
    `Message ${n}.`, '',
  ].join('\r\n');
  let win = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    const accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);
    for (const [n, s] of subjects.entries()) await mail('deliverTest', { accountId, folder: 'Inbox', raw: raw(s, n) });
    const folders = await mail('folders', { accountId });
    const inbox = folders.find((f) => /^inbox$/i.test(f.name || f.path))?.path || 'Inbox';
    const openInbox = () => js(`(() => { const f = ((list) => list.find((n) => ${labelKeyed(inbox, { start: true })}(n)) || list.find((n) => n.textContent.trim().startsWith(${JSON.stringify(inbox)})))([...document.querySelectorAll('[title="${inbox}"]')]); f?.click(); return 1; })()`);
    const rowOf = (s) => `[...document.querySelectorAll('.ml-list > .ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(s)}))`;
    await openInbox();
    await until(async () => (await js(`Boolean(${subjects.map(rowOf).join(' && ')})`)) || (await openInbox(), await wait(400), false), 'the three messages in the list', 12000).catch(() => {});

    // Off the pointer and not selected: no control shown. Selected: shown, at the row's end.
    await win.webContents.sendInputEvent({ type: 'mouseMove', x: 2, y: 2 });
    await js(`${rowOf(subjects[1])}?.click(), 1`);
    await until(() => js(`Boolean(${rowOf(subjects[1])}?.classList.contains('selected'))`), 'bravo selected', 5000).catch(() => {});
    await wait(200);
    const seen = (s) => js(`(() => { const b = ${rowOf(s)}?.querySelector('.ml-row-delete'); return b ? getComputedStyle(b).visibility : 'none'; })()`);
    const shown = { alpha: await seen(subjects[0]), bravo: await seen(subjects[1]) };
    const atEnd = await js(`(() => { const r = ${rowOf(subjects[1])}; const marks = r?.querySelector('.ml-row-side .ml-marks'); const b = marks?.lastElementChild; return { last: Boolean(b?.classList.contains('ml-row-delete')), label: b?.getAttribute('aria-label') || '', right: b ? Math.round(r.getBoundingClientRect().right - b.getBoundingClientRect().right) : null }; })()`);
    check('mail: the delete control sits at the end of the row and shows on the selected message, not on another one off the pointer',
      shown.alpha === 'hidden' && shown.bravo === 'visible' && atEnd.last && atEnd.label === 'Delete this message' && atEnd.right !== null && atEnd.right < 40,
      JSON.stringify({ shown, atEnd }));

    // Under the pointer: shown.
    const at = await js(`(() => { const r = ${rowOf(subjects[2])}; if (!r) return null; r.scrollIntoView({ block: 'center' }); const b = r.getBoundingClientRect(); return { x: Math.round(b.left + b.width / 2), y: Math.round(b.top + b.height / 2) }; })()`);
    if (at) win.webContents.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
    await wait(300);
    const hovered = await seen(subjects[2]);
    check('mail: the delete control shows on the message under the pointer', hovered === 'visible', hovered);

    // Pressed on charlie, with bravo selected: charlie goes to the Trash, bravo stays selected.
    await js(`${rowOf(subjects[2])}?.querySelector('.ml-row-delete')?.click(), 1`);
    await until(() => js(`!${rowOf(subjects[2])}`), 'charlie gone from the list', 6000).catch(() => {});
    await wait(500);
    const trash = (await mail('folders', { accountId })).find((f) => f.role === 'trash')?.path || null;
    const inTrash = trash ? (await mail('messages', { accountId, folder: trash })).rows.some((r) => r.subject === subjects[2]) : null;
    const stillInbox = (await mail('messages', { accountId, folder: inbox })).rows.map((r) => r.subject).filter((s) => s.startsWith('Row delete:'));
    const selectedNow = await js(`document.querySelector('.ml-list > .ml-row.selected .ml-subject')?.textContent || ''`);
    check('mail: the control deletes that message alone, to the Trash where there is one, and the selection stays where it was',
      !stillInbox.includes(subjects[2]) && stillInbox.includes(subjects[0]) && stillInbox.includes(subjects[1]) && (trash ? inTrash : true) && selectedNow.includes(subjects[1]),
      JSON.stringify({ trash, inTrash, stillInbox, selectedNow }));

    // The Delete key, as before, on the selected message.
    await js(`(${rowOf(subjects[1])})?.focus(), 1`);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Delete' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Delete' });
    await until(() => js(`!${rowOf(subjects[1])}`), 'bravo gone with the Delete key', 6000).catch(() => {});
    const keyed = await js(`!${rowOf(subjects[1])}`);
    check('mail: the Delete key still deletes the selected message', keyed, String(keyed));

    const complaints = await errorsIn(win);
    check('mail: delete beside each message reported no error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail: the delete beside each message checks ran to the end', false, err?.message || String(err));
  } finally {
    // What is left of the three is taken away, so the checks after these see the Inbox they expect.
    try {
      const accountId = await win?.webContents.executeJavaScript(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
      if (accountId) {
        for (const f of await win.webContents.executeJavaScript(`window.rutbaOffice.mail.folders({ accountId: ${JSON.stringify(accountId)} })`)) {
          const rows = (await win.webContents.executeJavaScript(`window.rutbaOffice.mail.messages({ accountId: ${JSON.stringify(accountId)}, folder: ${JSON.stringify(f.path)} })`)).rows.filter((r) => subjects.includes(r.subject));
          if (rows.length) await win.webContents.executeJavaScript(`window.rutbaOffice.mail.delete({ accountId: ${JSON.stringify(accountId)}, folder: ${JSON.stringify(f.path)}, ids: ${JSON.stringify(rows.map((r) => r.id))} })`);
        }
      }
    } catch { /* the window is gone */ }
  }
}
