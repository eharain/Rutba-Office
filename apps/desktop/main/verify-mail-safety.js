// Mail's Safety switch (owner, 2026-10-20: "let the people see full raw
// message and let them copy URLs and images from mail beside download
// attachments. please keep an extra safety button on top and let this copy
// and download happen when its turned off").
//
// A message arrives with a link, a picture it carries and an archive
// attached. With Safety on, as every message opens: a right-click on the link
// and on the picture offers Open Link but not Copy Link, Copy Picture or Save
// Picture As, and saving the attachment is refused with a word about Safety.
// View source shows the whole message, rebuilt here since the check's account
// has no server, with Copy All and Save as .eml waiting for Safety to be
// off. With Safety off: Copy Link puts the address on the clipboard, Copy
// Picture the picture, Save Picture As writes it, and the source can be
// copied. The next message opens with Safety on again.
// Run alone with RUTBA_VERIFY_ONLY=mail-safety.

import fs from 'node:fs';
import path from 'node:path';
import { clipboard } from 'electron';
import { readImage, writeParts } from '@rutba/office-shell/electron/clipboard';

// A 24 by 16 red picture.
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAABgAAAAQCAIAAACDRijCAAAAGklEQVR4nGO4o6FBFcQwatCoQaMGjRpEGQIA3r3CEG0ftIUAAAAASUVORK5CYII=';

export async function verifyMailSafety(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const ok = (p) => p.then(() => true).catch(() => false);
  const stamp = Date.now();
  const saveTo = path.join(dir, `safety-picture-${stamp}.png`);
  const raw = (n, subject) => [
    'From: Sam <sam@work.example>', 'To: You <you@example.com>', `Subject: ${subject}`, `Date: ${new Date(stamp + n * 60000).toUTCString()}`,
    `Message-ID: <safety-${stamp}-${n}@example.org>`, 'MIME-Version: 1.0', 'Content-Type: multipart/mixed; boundary="mix"', '',
    '--mix', 'Content-Type: multipart/related; boundary="rel"', '',
    '--rel', 'Content-Type: text/html; charset=utf-8', '',
    '<html><body style="margin:0"><a href="https://safety.example/report" style="display:block;height:120px;font-size:28px">The quarterly report</a>'
      + '<img src="cid:pic1" width="240" height="160" style="display:block"><p>The rest of the message.</p></body></html>',
    '--rel', 'Content-Type: image/png; name="chart.png"', 'Content-Transfer-Encoding: base64', 'Content-ID: <pic1>', 'Content-Disposition: inline; filename="chart.png"', '', PNG,
    '--rel--',
    '--mix', 'Content-Type: application/zip; name="figures.zip"', 'Content-Transfer-Encoding: base64', 'Content-Disposition: attachment; filename="figures.zip"', '', Buffer.from('PK\u0005\u0006' + '\u0000'.repeat(18), 'latin1').toString('base64'),
    '--mix--', '',
  ].join('\r\n');
  const answers = () => globalThis.__rutbaCheckDialogAnswers || [];
  let win = null;
  let kept = null;
  try {
    win = await open('mail');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.ml-accounts button').length > 0`), 'the seeded account in the sidebar', 8000);
    const accountId = await js(`(async () => (await window.rutbaOffice.mail.accounts())[0]?.id || null)()`);
    const mail = (method, args) => js(`window.rutbaOffice.mail.${method}(${JSON.stringify(args)})`);
    await mail('deliverTest', { accountId, folder: 'Inbox', raw: raw(0, 'Safety: the report') });
    await mail('deliverTest', { accountId, folder: 'Inbox', raw: raw(1, 'Safety: the next one') });
    const folders = await mail('folders', { accountId });
    const inbox = folders.find((f) => /^inbox$/i.test(f.name || f.path))?.name || 'Inbox';
    const openFolder = (name) => js(`(() => { const f = [...document.querySelectorAll('[title="${name}"]')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(name)})); if (!f) return 'no folder'; f.click(); return 'opened'; })()`);
    const openRow = (subject) => js(`(() => { const r = [...document.querySelectorAll('.ml-row')].find((row) => row.querySelector('.ml-subject')?.textContent.includes(${JSON.stringify(subject)})); if (!r) return 'no row'; r.click(); return 'opened'; })()`);
    const read = async (subject) => {
      await until(async () => (await openFolder(inbox)) === 'opened', 'the inbox', 5000).catch(() => {});
      await until(async () => (await openRow(subject)) === 'opened', `"${subject}" in the list`, 6000).catch(() => {});
      await until(() => js(`(document.querySelector('.ml-head h2')?.textContent || '').includes(${JSON.stringify(subject)})`), `"${subject}" in the reading pane`, 5000).catch(() => {});
      await wait(1200);
    };
    const dbg = win.webContents.debugger;
    if (!dbg.isAttached()) dbg.attach('1.3');
    /** A real right-click in the message frame, `dy` below its top. */
    const rightClick = async (dy) => {
      const at = await js(`(() => { const f = document.querySelector('.ml-body iframe'); if (!f) return null; const r = f.getBoundingClientRect(); return { x: Math.round(r.left + 60), y: Math.round(r.top + ${dy}) }; })()`);
      if (!at) return 'no frame';
      await dbg.sendCommand('Input.dispatchMouseEvent', { type: 'mouseMoved', x: at.x, y: at.y });
      await dbg.sendCommand('Input.dispatchMouseEvent', { type: 'mousePressed', x: at.x, y: at.y, button: 'right', clickCount: 1 });
      await dbg.sendCommand('Input.dispatchMouseEvent', { type: 'mouseReleased', x: at.x, y: at.y, button: 'right', clickCount: 1 });
      return 'clicked';
    };
    const menuItems = () => js(`[...document.querySelectorAll('.rw-menu button')].map((b) => ({ label: b.textContent.trim(), disabled: b.disabled || b.getAttribute('aria-disabled') === 'true' }))`);
    const menuFor = async (dy) => {
      await js(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
      await wait(150);
      await rightClick(dy);
      await until(() => js(`document.querySelectorAll('.rw-menu button').length > 0`), 'the message menu', 4000).catch(() => {});
      return menuItems();
    };
    const choose = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no item'; b.click(); return 'chosen'; })()`);
    const item = (items, label) => items.find((i) => i.label === label) || null;

    // Safety on, as the message opens.
    await read('Safety: the report');
    const sw = await js(`(() => { const b = document.querySelector('.ml-safety'); return b ? { on: b.classList.contains('on'), text: b.textContent.trim() } : null; })()`);
    const onLink = await menuFor(60);
    const onPicture = await menuFor(200);
    check('mail safety: a message opens with Safety on at its top, and a right-click on its link or its picture offers Open Link but not Copy Link, Copy Picture or Save Picture As',
      sw?.on === true && item(onLink, 'Open Link') && !item(onLink, 'Open Link').disabled && item(onLink, 'Copy Link')?.disabled === true && item(onLink, 'Turn Safety Off for This Message')
        && item(onPicture, 'Copy Picture')?.disabled === true && item(onPicture, 'Save Picture As…')?.disabled === true,
      JSON.stringify({ sw, onLink, onPicture }));
    await js(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 1`);
    const refused = await js(`(async () => {
      const a = [...document.querySelectorAll('.ml-attachment')].find((n) => n.textContent.includes('figures.zip'));
      if (!a) return 'no attachment';
      a.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
      await new Promise((r) => setTimeout(r, 400));
      return [...document.querySelectorAll('.rw-toast')].map((t) => t.textContent).join(' | ');
    })()`);
    const answeredBefore = answers().length;
    check('mail safety: saving an attachment waits for Safety to be off, and says so; no save dialog opens',
      /Safety is on/.test(refused) && answers().length === answeredBefore, refused.slice(0, 160));

    // View source, with Safety on: shown whole, copying and saving waiting.
    await js(`document.querySelector('.ml-source-btn')?.click(), 1`);
    const shown = await ok(until(() => js(`Boolean(document.querySelector('.ml-source .ml-source-text'))`), 'the source', 15000));
    const source = await js(`(() => { const d = document.querySelector('.ml-source'); return { from: d?.querySelector('.ml-source-note')?.dataset.from, text: (d?.querySelector('.ml-source-text')?.textContent || '').slice(0, 4000), copy: d?.querySelector('.ml-source-copy')?.disabled, save: d?.querySelector('.ml-source-save')?.disabled }; })()`);
    check('mail safety: View source shows the whole message (rebuilt here, and saying so: the check\'s account has no server), with Copy All and Save as .eml waiting for Safety',
      shown && source.from === 'rebuilt' && /Subject: Safety: the report/.test(source.text) && /X-Rutba-Rebuilt:/.test(source.text) && /figures\.zip/.test(source.text) && source.copy === true && source.save === true,
      JSON.stringify({ shown, from: source.from, copy: source.copy, save: source.save, head: source.text.slice(0, 120) }));
    await js(`[...document.querySelectorAll('.ml-source button')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);
    await wait(300);

    // Safety off: the link, the picture and the source can be copied, and the picture saved.
    await js(`document.querySelector('.ml-safety')?.click(), 1`);
    const off = await ok(until(() => js(`document.querySelector('.ml-safety')?.classList.contains('off')`), 'Safety off', 3000));
    // The person's clipboard, put back at the end.
    kept = await clipboard.readText().catch(() => '');
    await clipboard.writeText('before');
    // The system clipboard is the whole machine's: another program holding it
    // open as the copy lands leaves it as it was. A copy that did not land is
    // made once more, and the toast says whether Mail did its part.
    const toasts = () => js(`[...document.querySelectorAll('.rw-toast')].map((n) => n.textContent).join(' | ')`);
    const copyTo = async (dy, label, landed) => {
      for (let attempt = 0; attempt < 2; attempt++) {
        await menuFor(dy);
        await choose(label);
        if (await ok(until(landed, `${label} on the clipboard`, 3000))) return { landed: true, attempt };
      }
      return { landed: false, toast: await toasts() };
    };
    const link = await copyTo(60, 'Copy Link', async () => (await clipboard.readText()) === 'https://safety.example/report');
    const copiedLink = await clipboard.readText();
    await writeParts({ text: '' });
    const picture = await copyTo(200, 'Copy Picture', async () => !(await readImage()).isEmpty());
    const pictureCopied = picture.landed;
    const size = (await readImage()).getSize();
    if (!Array.isArray(globalThis.__rutbaCheckDialogAnswers)) globalThis.__rutbaCheckDialogAnswers = [];
    globalThis.__rutbaCheckDialogAnswers.push({ kind: 'save', answer: saveTo });
    await menuFor(200);
    await choose('Save Picture As…');
    const saved = await ok(until(() => fs.existsSync(saveTo) && fs.statSync(saveTo).size > 50, 'the picture on disk', 5000));
    const magic = saved ? fs.readFileSync(saveTo).subarray(0, 4).toString('hex') : '';
    await js(`document.querySelector('.ml-source-btn')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.ml-source .ml-source-text'))`), 'the source again', 15000).catch(() => {});
    const enabled = await js(`(() => { const d = document.querySelector('.ml-source'); return { copy: d?.querySelector('.ml-source-copy')?.disabled, save: d?.querySelector('.ml-source-save')?.disabled }; })()`);
    await js(`[...document.querySelectorAll('.ml-source button')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);
    check('mail safety: with Safety off, Copy Link puts the address on the clipboard, Copy Picture the picture, Save Picture As writes it as a PNG, and the source can be copied and saved',
      off && copiedLink === 'https://safety.example/report' && pictureCopied && size.width === 24 && size.height === 16 && saved && magic === '89504e47' && enabled.copy === false && enabled.save === false,
      JSON.stringify({ off, copiedLink, link, pictureCopied, picture, size, saved, magic, enabled }));

    // The next message: Safety on again.
    await read('Safety: the next one');
    const next = await js(`document.querySelector('.ml-safety')?.classList.contains('on')`);
    check('mail safety: the next message opens with Safety on again', next === true, String(next));

    const complaints = (await errorsIn(win)).filter((m) => !/Safety is on/.test(m));
    check('mail safety: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('mail safety: the check ran to the end', false, err?.message || String(err));
  } finally {
    try { win?.webContents.debugger.detach(); } catch { /* not attached */ }
    if (kept != null) await clipboard.writeText(kept).catch(() => {});
    if (Array.isArray(globalThis.__rutbaCheckDialogAnswers)) globalThis.__rutbaCheckDialogAnswers.length = 0;
  }
}
