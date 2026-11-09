// The Locker's window, end to end, in a build made with its switch: setting up,
// adding a folder, a file dropped in waiting to be locked by hand, a file opened
// into Documents and its Save written back to the locked copy, exporting a
// copy and to share, locking and unlocking (a wrong password changing nothing),
// and unlocking for good. It runs only in a run started with RUTBA_LOCKER=1;
// the release gate's full run has no Locker to check.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';

const js2 = (w, code) => w.webContents.executeJavaScript(code);
const LOCKED_ROW = (rel) => `tr[data-rel=${JSON.stringify(rel)}][data-state=locked]`;

export async function verifyLocker({ open, check, until, wait, errorsIn, doc, capture: shot = async () => {} }, { dir, files }) {
  // An off-screen window paints only when asked, so each capture asks first.
  const capture = async (w, name) => {
    w.webContents.invalidate();
    await wait(700);
    await shot(w, name);
  };
  if (process.env.RUTBA_LOCKER === '0') return;
  const answers = globalThis.__rutbaCheckDialogAnswers;
  const plan = (kind, answer) => answers.push({ kind, answer });
  const root = fs.mkdtempSync(path.join(dir, 'locker-folder-'));
  const out = fs.mkdtempSync(path.join(dir, 'locker-out-'));
  const pdfFile = path.join(root, 'Lease, signed.pdf');
  const pdfBytes = crypto.randomBytes(120000);
  fs.writeFileSync(pdfFile, pdfBytes);
  const words = path.join(root, 'Report.docx');
  fs.copyFileSync(files.docx, words);
  fs.mkdirSync(path.join(root, 'Contracts', '2026'), { recursive: true });
  fs.writeFileSync(path.join(root, 'Contracts', '2026', 'Terms.txt'), 'the terms');
  fs.writeFileSync(path.join(root, '~$Report.docx'), 'lock file');

  let win = null;
  try {
    win = await open('locker');
    const js = (code) => win.webContents.executeJavaScript(code);
    const setField = (label, value) => js(`(() => {
      const el = [...document.querySelectorAll('input')].find((i) => (i.getAttribute('aria-label') || '') === ${JSON.stringify(label)});
      if (!el) return 'no input ' + ${JSON.stringify(label)};
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'ok';
    })()`);
    const clickText = async (text) => {
      // A button that follows a typed field is enabled once the window has drawn the typing.
      await until(() => js(`[...document.querySelectorAll('button')].some((x) => x.textContent.trim() === ${JSON.stringify(text)} && !x.disabled)`), `the button ${text}`, 8000);
      return js(`(() => {
      const b = [...document.querySelectorAll('button')].find((x) => x.textContent.trim() === ${JSON.stringify(text)});
      if (!b) return 'no button ' + ${JSON.stringify(text)};
      b.click();
      return 'clicked';
    })()`);
    };
    const clickSel = (sel) => js(`(() => { const b = document.querySelector(${JSON.stringify(sel)}); if (!b) return 'no ' + ${JSON.stringify(sel)}; b.click(); return 'clicked'; })()`);
    const bodyText = () => js(`document.querySelector('.lk').textContent`);

    // First start: a password, and the recovery key shown once.
    await until(() => js(`Boolean(document.querySelector('.lk-card'))`), 'the setup card', 15000);
    await setField('Locker password', 'correct horse battery');
    await setField('The same password again', 'correct horse battery');
    check('locker: a file someone sent can be opened by a person with no locker of their own', await js(`Boolean(document.querySelector('.lk-open-sent'))`));
    const made = await clickText('Make the locker');
    await until(() => js(`Boolean(document.querySelector('[data-recovery]'))`), 'the recovery key', 15000);
    const recovery = await js(`document.querySelector('[data-recovery]').textContent`);
    check('locker: a locker is made with a password, and the recovery key is shown', made === 'clicked' && /^[A-Z2-9-]+$/.test(recovery));
    await capture(win, 'locker-setup.png');
    await clickText('I have kept it');
    await until(() => js(`Boolean(document.querySelector('.lk-add'))`), 'the main window', 8000);
    check('locker: after setup it is unlocked, and asks for a folder', /locked folders/i.test(await bodyText()) && !(await js(`document.querySelectorAll('tbody tr').length`)));
    check('locker: with no folder yet it offers a ready-made one', await js(`/Make my Locker folder/.test(document.querySelector('.lk-make')?.textContent || '')`));

    // Add a folder: the choices are asked, a suggestion on each, and the files are locked.
    plan('open', [root]);
    await clickSel('.lk-add');
    await until(() => js(`Boolean(document.querySelector('.lk-add-dialog'))`), 'the add-folder dialog', 6000);
    await capture(win, 'locker-add-folder.png');
    check('locker: adding a folder asks nothing until the person asks to change how', await js(`!document.querySelector('.lk-add-dialog .lk-choice') && Boolean(document.querySelector('.lk-change'))`));
    await clickSel('.lk-change');
    await until(() => js(`Boolean(document.querySelector('.lk-add-dialog .lk-choice'))`), 'the choices', 4000);
    await capture(win, 'locker-add-folder-choices.png');
    const asked = await js(`document.querySelector('.lk-add-dialog').textContent`);
    check('locker: adding a folder asks what to lock, names in the cloud and the plain copies, each with a suggestion',
      /Every file/.test(asked) && /Hidden/.test(asked) && /Delete the plain copy/.test(asked) && /suggested/i.test(asked), asked.slice(0, 200));
    // Visible names, to read them on disk; delete the plain copy, as suggested.
    await js(`(() => { const r = [...document.querySelectorAll('.lk-add-dialog .lk-choice')].find((x) => x.textContent.startsWith('Visible')); r.querySelector('input').click(); return 1; })()`);
    await wait(100);
    await clickSel('.lk-add-go');
    await until(() => js(`Boolean(document.querySelector('tr[data-rel="Report.docx"][data-state="locked"]'))`), 'the files to be locked', 30000);
    const sealed = fs.existsSync(`${words}.rlock`) && !fs.existsSync(words) && !fs.existsSync(pdfFile) && fs.existsSync(path.join(root, '.rutbalocker-keys'));
    check('locker: the folder\'s files are locked where they are, each its own .rlock, the plain copies gone', sealed, fs.readdirSync(root).join(', '));
    await clickSel('tr[data-rel="~$Report.docx"]');
    await until(() => js(`Boolean(document.querySelector('.lk-detail'))`), 'the line about the file', 4000);
    const tmpFileRow = await js(`document.querySelector('tr[data-rel="~$Report.docx"]')?.textContent + ' ' + document.querySelector('.lk-detail').textContent`);
    check('locker: Office\'s own temporary file stays plain, and says why', /Not locked/.test(tmpFileRow || '') && /temporary/.test(tmpFileRow || ''), tmpFileRow);
    check('locker: the list says Locked or Not locked, and the rules stay out of the way', await js(`!document.querySelector('.lk-rules') && !document.querySelector('.lk-folder-settings') && document.querySelectorAll('thead th').length === 4`));
    await capture(win, 'locker-main.png');

    // A subfolder has its own highlighted heading, and every file its own buttons without being selected first.
    const headed = await js(`(() => { const h = document.querySelector('tr.lk-dir[data-dir="Contracts/2026"]'); const r = document.querySelector('tr[data-rel="Contracts/2026/Terms.txt"]'); return Boolean(h) && h.textContent.includes('Contracts / 2026') && Boolean(r.querySelector('.lk-unlock')) && Boolean(r.querySelector('.lk-open')) && r.querySelector('td').textContent === 'Terms.txt'; })()`);
    check('locker: a subfolder has its own heading and each file its own buttons', headed);

    // Searching narrows the list, by any part of a file's path.
    const visibleRows = () => js(`[...document.querySelectorAll('tbody tr[data-rel]')].map((r) => r.dataset.rel).join('|')`);
    await setField('Search files', 'contracts terms');
    const narrowed = await until(async () => (await visibleRows()) === 'Contracts/2026/Terms.txt', 'the search to narrow the list', 4000).then(() => true, () => false);
    check('locker: searching shows only the files that match, under their subfolder', narrowed && await js(`Boolean(document.querySelector('tr.lk-dir[data-dir="Contracts/2026"]'))`), await visibleRows());
    await capture(win, 'locker-search.png');
    await setField('Search files', 'nothing is called this');
    const none = await until(() => js(`document.querySelector('.lk-table-wrap').textContent.includes('No files match')`), 'the no-match message', 4000).then(() => true, () => false);
    check('locker: a search that finds nothing says so', none);
    await setField('Search files', '');
    await until(() => js(`document.querySelectorAll('tbody tr[data-rel]').length > 1`), 'the whole list back', 4000);

    // A file arriving is listed as not locked, and stays so until the person locks it.
    fs.writeFileSync(path.join(root, 'Dropped in.txt'), 'dropped in by the person');
    const listedDrop = await until(() => js(`Boolean(document.querySelector('tr[data-rel="Dropped in.txt"][data-state="waiting"]'))`), 'the dropped file to be listed', 20000).then(() => true, () => false);
    await wait(2500);
    check('locker: a file dropped into the folder is listed as not locked and nothing locks it by itself', listedDrop && fs.existsSync(path.join(root, 'Dropped in.txt')) && !fs.existsSync(path.join(root, 'Dropped in.txt.rlock')), fs.readdirSync(root).join(', '));
    await clickSel('.lk-lockall');
    const lockedDrop = await until(() => js(`Boolean(document.querySelector('tr[data-rel="Dropped in.txt"][data-state="locked"]'))`), 'the dropped file to lock', 20000).then(() => true, () => false);
    check('locker: the Lock button locks it where it is', lockedDrop && fs.existsSync(path.join(root, 'Dropped in.txt.rlock')) && !fs.existsSync(path.join(root, 'Dropped in.txt')), fs.readdirSync(root).join(', '));

    // Export a copy: the locked file is untouched and the copy is the original.
    const dest = path.join(out, 'copy.pdf');
    const lockedBefore = fs.readFileSync(`${pdfFile}.rlock`);
    await clickSel(LOCKED_ROW('Lease, signed.pdf'));
    await until(() => js(`document.querySelector('tr[data-rel="Lease, signed.pdf"]')?.classList.contains('sel') && !document.querySelector('.lk-export').disabled`), 'the row selected', 4000);
    plan('save', dest);
    await clickSel('.lk-export');
    await until(() => fs.existsSync(dest), 'the exported copy', 8000);
    check('locker: Export copy writes the original and leaves the locked file as it was', fs.readFileSync(dest).equals(pdfBytes) && fs.readFileSync(`${pdfFile}.rlock`).equals(lockedBefore));

    // Open into Documents: a temporary copy beside the file, and its Save reaches the locked file.
    await clickSel(LOCKED_ROW('Report.docx'));
    await until(() => js(`Boolean(document.querySelector('tr.sel .lk-open'))`), 'Open offered', 4000);
    const before = fs.readFileSync(`${words}.rlock`);
    await clickSel('tr.sel .lk-open');
    // What until() waited for is not what it returns, so the window is kept here.
    let opened = null;
    await until(() => {
      opened = BrowserWindow.getAllWindows().find((x) => /locker/.test(decodeURIComponent(x.webContents.getURL())) && /[?&]file=/.test(x.webContents.getURL())) || null;
      return Boolean(opened);
    }, 'the document window', 15000);
    const url = decodeURIComponent(opened.webContents.getURL());
    const tempPath = /[?&]file=([^&]+)/.exec(url)?.[1] || '';
    check('locker: Open puts a temporary copy in a hidden folder beside the file and opens it in Documents', tempPath.startsWith(path.join(root, '.rutba-open')) && fs.readFileSync(tempPath).length > 0, tempPath);
    const rowsWhileOpen = await js(`[...document.querySelectorAll('tbody tr')].map((r) => r.dataset.rel).join()`);
    check('locker: the temporary copy is not listed as a file of the folder, and is not locked by the watcher', !/rutba-open/.test(rowsWhileOpen) && !fs.existsSync(`${tempPath}.rlock`), rowsWhileOpen);
    await until(() => doc.sessions().some((s) => s.path === tempPath), 'the document session', 15000);
    await capture(opened, 'locker-open-in-documents.png');
    const session = doc.sessions().find((s) => s.path === tempPath);
    // An unchanged save is rightly no new version, so this one has a change in it.
    await doc.apply({ id: session.id, ops: [{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 1, offset: 0 } }, { op: 'insertText', text: 'Edited while locked. ' }] });
    await doc.save({ id: session.id, path: tempPath });
    const savedBack = await until(() => !fs.readFileSync(`${words}.rlock`).equals(before), 'the save written back', 20000).then(() => true, () => false);
    const roundTrip = path.join(out, 'saved.docx');
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.exportCopy({ id: s.folders[0].id, rel: 'Report.docx', dest: ${JSON.stringify(roundTrip)}, overwrite: true }); return 1; })()`);
    check('locker: Save from Documents reaches the locked file, and what comes back out is what was saved', savedBack && fs.readFileSync(roundTrip).equals(fs.readFileSync(tempPath)));
    try { await doc.close?.({ id: session.id }); } catch { /* the window goes next */ }
    opened.destroy();
    const cleaned = await until(() => !fs.existsSync(path.join(root, '.rutba-open')), 'the temporary copy to be cleared', 15000).then(() => true, () => false);
    check('locker: closing the document clears its temporary copy and the hidden folder', cleaned, fs.readdirSync(root, { withFileTypes: false }).join(', '));

    // Lock: a wrong password changes nothing and says so; the right one opens it.
    await clickSel('.lk-lock');
    await until(() => js(`Boolean(document.querySelector('.lk-card'))`), 'the unlock card', 8000);
    await wait(200);
    await capture(win, 'locker-unlock.png');
    const allBytes = () => fs.readdirSync(root).filter((n) => n.endsWith('.rlock')).sort().map((n) => fs.readFileSync(path.join(root, n)).toString('base64')).join();
    const beforeWrong = allBytes();
    await setField('Locker password', 'wrong password');
    await clickText('Unlock with password');
    let said = '';
    await until(async () => {
      said = await js(`document.querySelector('.lk-error')?.textContent || ''`);
      return Boolean(said);
    }, 'the wrong-password message', 15000).catch(() => {});
    check('locker: a wrong password is refused and changes nothing', /password/i.test(said) && !/Error/.test(said) && allBytes() === beforeWrong, said);
    await setField('Locker password', 'correct horse battery');
    await clickText('Unlock with password');
    await until(() => js(`Boolean(document.querySelector('.lk-add'))`), 'unlocked again', 15000);
    await until(() => js(`document.querySelectorAll('tbody tr').length > 0`), 'the files listed again', 8000);
    check('locker: the right password opens it again and the files are back', true);

    // Windows Hello, with the stand-in for the person's PIN: on, and each way the prompt can end.
    const helloUnlocked = () => until(() => js(`Boolean(document.querySelector('.lk-add'))`), 'unlocked by Windows Hello', 8000).then(() => true, () => false);
    await clickSel('.lk-settings');
    await until(() => js(`Boolean(document.querySelector('.lk-advanced input'))`), 'the settings dialog', 5000);
    await clickText('Turn on Windows Hello');
    const helloOn = await until(() => js(`[...document.querySelectorAll('button')].some((x) => x.textContent.trim() === 'Turn off Windows Hello')`), 'Windows Hello on', 6000).then(() => true, () => false);
    check('locker: Windows Hello can be turned on from the settings', helloOn);
    await clickText('Cancel');
    await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the settings to close', 5000);
    const helloButton = 'Use Windows Hello (PIN, fingerprint or face)';
    await clickSel('.lk-lock');
    await until(() => js(`[...document.querySelectorAll('button')].some((x) => x.textContent.trim() === ${JSON.stringify(helloButton)})`), 'the Windows Hello button', 8000);
    await capture(win, 'locker-unlock-hello.png');
    globalThis.__rutbaFakeHello = 'UserCanceled';
    await clickText(helloButton);
    await until(() => js(`[...document.querySelectorAll('button')].some((x) => x.textContent.trim() === ${JSON.stringify(helloButton)} && !x.disabled)`), 'the prompt to end', 6000);
    await wait(300);
    check('locker: cancelling the Windows Hello prompt leaves it locked and shows no fault', !(await js(`document.querySelector('.lk-error')?.textContent || ''`)) && !(await js(`Boolean(document.querySelector('.lk-add'))`)));
    globalThis.__rutbaFakeHello = 'NotFound';
    await clickText(helloButton);
    let goneSaid = '';
    await until(async () => {
      goneSaid = await js(`document.querySelector('.lk-error')?.textContent || ''`);
      return Boolean(goneSaid);
    }, 'the missing-key message', 6000).catch(() => {});
    check('locker: a key Windows has lost points the person to the password, in plain words', /Use your password/.test(goneSaid) && !/NotFound/.test(goneSaid), goneSaid);
    await clickText(helloButton);
    check('locker: Windows Hello opens it', await helloUnlocked());
    await until(() => js(`document.querySelectorAll('tbody tr').length > 0`), 'the files listed after Windows Hello', 8000);
    await clickSel('.lk-settings');
    await until(() => js(`Boolean(document.querySelector('.lk-advanced input'))`), 'the settings dialog', 5000);
    await clickText('Turn off Windows Hello');
    const helloOff = await until(() => js(`[...document.querySelectorAll('button')].some((x) => x.textContent.trim() === 'Turn on Windows Hello')`), 'Windows Hello off', 6000).then(() => true, () => false);
    check('locker: Windows Hello can be turned off again', helloOff);
    await clickText('Cancel');
    await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the settings to close', 5000);

    // The other windows a person meets, each looked at; sharing one file for real.
    const look = async (button, name) => {
      await clickSel(button);
      await until(() => js(`Boolean(document.querySelector('.lk-dialog-body'))`), name, 5000);
      await capture(win, name);
      await clickText('Cancel');
      await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the dialog to close', 5000);
    };
    await clickSel('.lk-settings');
    await until(() => js(`Boolean(document.querySelector('.lk-advanced input'))`), 'the settings dialog', 5000);
    await clickSel('.lk-advanced input');
    const advancedOn = await until(() => js(`Boolean(document.querySelector('.lk-rules')) && Boolean(document.querySelector('.lk-folder-settings')) && document.querySelectorAll('thead th').length === 6`), 'the advanced options', 6000).then(() => true, () => false);
    check('locker: Show advanced options brings the rules, the folder settings and the extra columns', advancedOn);
    const asked2 = await js(`document.querySelector('.lk-dialog-body').textContent`);
    check('locker: the settings ask what happens to temporary copies and unlocked files when it locks, each with a suggestion', /Remove them \(suggested\)/.test(asked2) && /Leave them unlocked \(suggested\)/.test(asked2), asked2.slice(0, 200));
    await clickText('Cancel');
    await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the settings to close', 5000);
    await look('.lk-rules', 'locker-rules.png');
    await look('.lk-folder-settings', 'locker-folder-settings.png');
    await look('.lk-settings', 'locker-settings.png');
    await clickSel(LOCKED_ROW('Lease, signed.pdf'));
    await until(() => js(`Boolean(document.querySelector('.lk-share'))`), 'Share offered', 4000);
    await clickSel('.lk-share');
    await until(() => js(`Boolean(document.querySelector('.lk-dialog-body'))`), 'the share dialog', 5000);
    await setField('A password for this one file', 'one file only');
    await capture(win, 'locker-share.png');
    const shared = path.join(out, 'shared.rlock');
    plan('save', shared);
    await clickText('Save the locked copy…');
    await until(() => fs.existsSync(shared), 'the shared file', 8000);
    const opener = await js(`window.rutbaOffice.locker.openShared({ file: ${JSON.stringify(shared)}, password: 'one file only', dest: ${JSON.stringify(path.join(out, 'shared.pdf'))} }).then(() => 'opened', (e) => e.message)`);
    check('locker: Export to share makes a file that opens with its own password and holds the original', opener === 'opened' && fs.readFileSync(path.join(out, 'shared.pdf')).equals(pdfBytes), opener);
    await until(() => js(`Boolean(document.querySelector('.lk-share-done'))`), 'the ready-to-send step', 5000);
    await capture(win, 'locker-share-done.png');
    check('locker: after saving, the copy can go to an email or be shown in Explorer', await js(`Boolean(document.querySelector('.lk-share-mail') && document.querySelector('.lk-share-show'))`));
    await clickSel('.lk-share-mail');
    let mailWin = null;
    await until(() => { mailWin = BrowserWindow.getAllWindows().find((x) => /[?&]app=mail(&|$)/.test(x.webContents.getURL())) || null; return Boolean(mailWin); }, 'the mail window', 15000);
    const attached = await until(() => mailWin.webContents.executeJavaScript(`/shared\.rlock/.test(document.body.textContent)`), 'the attachment in the message', 15000).then(() => true, () => false);
    check('locker: Attach to an email opens a new message with the locked copy attached', attached, mailWin.webContents.getURL());
    mailWin.close();
    await clickSel('.lk-share-close');
    await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the share dialog to close', 5000);

    // The email offered at the start, with nothing to save first.
    await clickSel('.lk-share');
    await until(() => js(`Boolean(document.querySelector('.lk-share-email'))`), 'the email button on the first step', 5000);
    await setField('A password for this one file', 'one file only');
    await capture(win, 'locker-share-first.png');
    await clickSel('.lk-share-email');
    let mailNow = null;
    await until(() => { mailNow = BrowserWindow.getAllWindows().find((x) => /[?&]app=mail(&|$)/.test(x.webContents.getURL())) || null; return Boolean(mailNow); }, 'the mail window from the first step', 15000);
    const attachedNow = await until(() => mailNow.webContents.executeJavaScript(`/Lease, signed\.pdf\.rlock/.test(document.body.textContent)`), 'the attachment from the first step', 15000).then(() => true, () => false);
    check('locker: Attach to an email is offered at the start and opens a message with the locked copy', attachedNow && /outgoing/.test(decodeURIComponent(mailNow.webContents.getURL())), mailNow.webContents.getURL());
    mailNow.close();
    await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the share dialog to close after the email', 5000);

    // The other side: a sent file, its password, then a plain copy or into a folder.
    plan('open', [shared]);
    await clickSel('.lk-bar .lk-open-sent');
    await until(() => js(`Boolean(document.querySelector('input[aria-label="The password for this file"]'))`), 'the password box for a sent file', 5000);
    await setField('The password for this file', 'not this one');
    plan('save', path.join(out, 'not-saved.pdf'));
    await clickText('Save a plain copy…');
    await until(() => js(`Boolean(document.querySelector('.lk-dialog-body .lk-error'))`), 'the refusal', 8000);
    check('locker: a sent file refuses the wrong password and writes nothing', !fs.existsSync(path.join(out, 'not-saved.pdf')));
    await setField('The password for this file', 'one file only');
    await capture(win, 'locker-open-sent.png');
    plan('save', path.join(out, 'received.pdf'));
    await clickText('Save a plain copy…');
    await until(() => fs.existsSync(path.join(out, 'received.pdf')), 'the plain copy of the sent file', 8000);
    check('locker: a sent file opens to a plain copy with its own password', fs.readFileSync(path.join(out, 'received.pdf')).equals(pdfBytes));
    await until(() => js(`Boolean(document.querySelector('.lk-sent-close'))`), 'the finished step', 5000);
    await clickSel('.lk-sent-close');
    await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the sent-file dialog to close', 5000);
    plan('open', [shared]);
    await clickSel('.lk-bar .lk-open-sent');
    await until(() => js(`Boolean(document.querySelector('input[aria-label="The password for this file"]'))`), 'the password box again', 5000);
    await setField('The password for this file', 'one file only');
    await clickText('Add to this folder');
    await until(() => fs.existsSync(path.join(root, 'shared.rlock')), 'the sent file in the folder', 8000);
    check('locker: a sent file can be kept in a locked folder, locked, and nothing plain is left there', !fs.existsSync(path.join(root, 'shared')) && fs.existsSync(shared), fs.readdirSync(root).join(', '));
    await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the sent-file dialog to close', 5000);

    // Unlock for good.
    await clickSel(LOCKED_ROW('Lease, signed.pdf'));
    await until(() => js(`Boolean(document.querySelector('tr.sel .lk-unlock'))`), 'Unlock offered', 4000);
    plan('message', { response: 0, checked: false });
    await clickSel('tr.sel .lk-unlock');
    const backOnDisk = await until(() => fs.existsSync(pdfFile) && !fs.existsSync(`${pdfFile}.rlock`), 'the plain file back', 8000).then(() => true, () => false);
    check('locker: Unlock for good puts the plain file back and removes the locked one', backOnDisk && fs.readFileSync(pdfFile).equals(pdfBytes));
    await wait(2500);
    check('locker: a file unlocked stays plain until the person locks it again', fs.existsSync(pdfFile) && !fs.existsSync(`${pdfFile}.rlock`));

    // And back again, with the one button.
    await until(() => js(`Boolean(document.querySelector('tr[data-rel="Lease, signed.pdf"][data-state=waiting]'))`), 'the file listed as not locked', 8000);
    await clickSel('tr[data-rel="Lease, signed.pdf"][data-state=waiting] .lk-lockfile');
    const putBack = await until(() => !fs.existsSync(pdfFile) && fs.existsSync(`${pdfFile}.rlock`), 'the file locked again', 8000).then(() => true, () => false);
    check('locker: the Lock button on a file\'s own row locks it again, where it is', putBack);

    // A subfolder's own button locks only what is in it.
    const more = path.join(root, 'Contracts', '2026', 'More.txt');
    fs.writeFileSync(more, 'more terms');
    await until(() => js(`Boolean(document.querySelector('tr.lk-dir[data-dir="Contracts/2026"] .lk-lockdir'))`), 'the subfolder button', 20000);
    await clickSel('tr.lk-dir[data-dir="Contracts/2026"] .lk-lockdir');
    const dirLocked = await until(() => !fs.existsSync(more) && fs.existsSync(`${more}.rlock`), 'the subfolder file locked', 8000).then(() => true, () => false);
    check('locker: a subfolder\'s Lock button locks the files not locked in it', dirLocked);

    // The idle lock: 4 seconds here, in place of the minutes a person chooses.
    const stateNow = () => js(`window.rutbaOffice.locker.state()`);
    const isLockedNow = () => js(`Boolean(document.querySelector('.lk-card')) && !document.querySelector('.lk-add')`);
    const unlockAgain = async () => {
      await setField('Locker password', 'correct horse battery');
      await clickText('Unlock with password');
      await until(() => js(`Boolean(document.querySelector('.lk-add'))`), 'unlocked after the idle lock', 15000);
    };
    await js(`window.rutbaOffice.store.set({ key: 'locker.idleMinutes', value: 0.07 }).then(() => window.rutbaOffice.locker.touch()).then(() => 1)`);

    // A folder that keeps changing by itself is not somebody using the Locker: each change refreshes the list.
    const noise = path.join(root, 'changing by itself.txt');
    let n = 0;
    const noisy = setInterval(() => { try { fs.writeFileSync(noise, String(n += 1)); } catch { /* the folder is being read */ } }, 2000);
    const idleLocked = await until(isLockedNow, 'the idle lock', 14000).then(() => true, () => false);
    clearInterval(noisy);
    check('locker: left alone for its set time it locks by itself, even while a file in the folder keeps changing', idleLocked && n >= 2, `changes ${n}`);
    await capture(win, 'locker-idle-locked.png');
    fs.rmSync(noise, { force: true });
    await unlockAgain();

    // Clicking in the window is using it.
    const firstDue = (await stateNow()).lockAt;
    await wait(1200);
    await js(`window.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true })), 1`);
    const putOff = await until(async () => (await stateNow()).lockAt > firstDue + 800, 'the lock put off', 3000).then(() => true, () => false);
    check('locker: a click in the window puts off the idle lock', putOff);

    // A document from it open in an Office window is the Locker in use.
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Report.docx'))}))`), 'the document row', 8000);
    await clickSel(LOCKED_ROW('Report.docx'));
    await until(() => js(`Boolean(document.querySelector('tr.sel .lk-open'))`), 'Open offered for the idle check', 3000);
    await clickSel('tr.sel .lk-open');
    let docWin = null;
    await until(() => {
      docWin = BrowserWindow.getAllWindows().find((x) => /locker/.test(decodeURIComponent(x.webContents.getURL())) && /[?&]file=/.test(x.webContents.getURL())) || null;
      return Boolean(docWin);
    }, 'the document window for the idle check', 15000);
    await wait(7000);
    check('locker: the idle lock waits while a document from it is open in an Office window', !(await isLockedNow()) && (await stateNow()).open === true);
    docWin.close();
    const lockedAfterDoc = await until(isLockedNow, 'the idle lock after the document closed', 14000).then(() => true, () => false);
    check('locker: once that document is closed the idle lock goes ahead', lockedAfterDoc);
    await js(`window.rutbaOffice.store.set({ key: 'locker.idleMinutes', value: 10 }).then(() => 1)`);
    await unlockAgain();

    // Locking under an open document must not take its file away: Save still works, and is locked in at the next unlock.
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Report.docx'))}))`), 'the document row again', 8000);
    await clickSel(LOCKED_ROW('Report.docx'));
    await until(() => js(`Boolean(document.querySelector('tr.sel .lk-open'))`), 'Open offered again', 3000);
    await clickSel('tr.sel .lk-open');
    let heldWin = null;
    await until(() => {
      heldWin = BrowserWindow.getAllWindows().find((x) => /locker/.test(decodeURIComponent(x.webContents.getURL())) && /[?&]file=/.test(x.webContents.getURL())) || null;
      return Boolean(heldWin);
    }, 'the document window to lock under', 15000);
    const heldPath = /[?&]file=([^&]+)/.exec(decodeURIComponent(heldWin.webContents.getURL()))?.[1] || '';
    await until(() => doc.sessions().some((s) => s.path === heldPath), 'that document\'s session', 15000);
    const heldBefore = fs.readFileSync(`${words}.rlock`);
    await js(`window.rutbaOffice.locker.lock().then(() => 1)`);
    await until(isLockedNow, 'locked under the open document', 8000);
    check('locker: locking removes the temporary copy of a document still open in Office, and leaves its folder', !fs.existsSync(heldPath) && fs.existsSync(path.dirname(heldPath)), heldPath);
    const heldSession = doc.sessions().find((s) => s.path === heldPath);
    await doc.apply({ id: heldSession.id, ops: [{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 1, offset: 0 } }, { op: 'insertText', text: 'Saved while the Locker was locked. ' }] });
    const savedWhileLocked = await (async () => { await doc.save({ id: heldSession.id, path: heldPath }); return true; })().catch((err) => String(err?.message || err));
    check('locker: Save still works in a document left open when the Locker locked', savedWhileLocked === true, String(savedWhileLocked));
    check('locker: what is saved while it is locked does not reach the locked file without the keys', fs.readFileSync(`${words}.rlock`).equals(heldBefore));
    await unlockAgain();
    const heldIn = await until(() => !fs.readFileSync(`${words}.rlock`).equals(heldBefore), 'the save locked in at unlock', 15000).then(() => true, () => false);
    const heldOut = path.join(out, 'held.docx');
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.exportCopy({ id: s.folders[0].id, rel: 'Report.docx', dest: ${JSON.stringify(heldOut)}, overwrite: true }); return 1; })()`);
    check('locker: at the next unlock that save is locked in, and the document can go on saving', heldIn && fs.readFileSync(heldOut).equals(fs.readFileSync(heldPath)) && fs.existsSync(heldPath));
    try { await doc.close?.({ id: heldSession.id }); } catch { /* the window goes next */ }
    heldWin.destroy();
    const heldCleared = await until(() => !fs.existsSync(heldPath), 'its copy cleared after closing', 15000).then(() => true, () => false);
    check('locker: closing that document clears the copy', heldCleared);

    // Renaming and deleting from the Locker's own list.
    const renameMe = path.join(root, 'Rename me.txt');
    const deleteMe = path.join(root, 'Delete me.txt');
    fs.writeFileSync(renameMe, 'keep these words');
    fs.writeFileSync(deleteMe, 'and lose these');
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.lockWaiting({ id: s.folders[0].id }); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Rename me.txt'))}) && document.querySelector(${JSON.stringify(LOCKED_ROW('Delete me.txt'))}))`), 'the two files locked', 15000);
    await clickSel(LOCKED_ROW('Rename me.txt'));
    await until(() => js(`Boolean(document.querySelector('.lk-rename'))`), 'Rename offered', 4000);
    await clickSel('.lk-rename');
    await until(() => js(`Boolean(document.querySelector('.lk-dialog-body'))`), 'the rename dialog', 5000);
    await setField('New name', 'Renamed.txt');
    await clickText('Rename');
    const renamed = await until(() => fs.existsSync(path.join(root, 'Renamed.txt.rlock')) && !fs.existsSync(`${renameMe}.rlock`), 'the file renamed on the disk', 10000).then(() => true, () => false);
    const renamedRow = await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Renamed.txt'))}))`), 'the renamed row', 8000).then(() => true, () => false);
    const renamedOut = path.join(out, 'renamed.txt');
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.exportCopy({ id: s.folders[0].id, rel: 'Renamed.txt', dest: ${JSON.stringify(renamedOut)}, overwrite: true }); return 1; })()`);
    check('locker: Rename gives a locked file a new name, on the disk and in the list, with the same contents', renamed && renamedRow && fs.readFileSync(renamedOut, 'utf8') === 'keep these words');
    await clickSel(LOCKED_ROW('Delete me.txt'));
    await until(() => js(`Boolean(document.querySelector('.lk-delete'))`), 'Delete offered', 4000);
    await capture(win, 'locker-file-actions.png');
    plan('message', { response: 0, checked: false });
    await clickSel('.lk-delete');
    const deleted = await until(() => !fs.existsSync(`${deleteMe}.rlock`), 'the file deleted', 10000).then(() => true, () => false);
    const deletedRow = await until(() => js(`!document.querySelector('tr[data-rel="Delete me.txt"]')`), 'its row gone', 8000).then(() => true, () => false);
    check('locker: Delete removes a locked file for good, and its row', deleted && deletedRow);
    await capture(win, 'locker-renamed-deleted.png');
    check('locker: closing a document raises no false alarm about its copy', !(await js(`document.body.textContent.includes('not opened from the locker')`)));

    const home = await open('home');
    const tile = await until(() => home.webContents.executeJavaScript(`Boolean(document.querySelector('.home-card[data-app="locker"]'))`), 'the launcher card', 10000).then(() => true, () => false);
    check('locker: a build made with the switch shows the card on the launcher', tile);
    await capture(home, 'locker-home-tile.png');

    const errors = await errorsIn(win).catch(() => []);
    check('locker: no errors in the window', !errors.length, JSON.stringify(errors).slice(0, 200));

    // Closing the Locker's window locks it. The launcher stays open, or the last window closing would end the run.
    win.close();
    await until(() => win.isDestroyed(), 'the Locker window to close', 8000);
    await wait(800);
    win = await open('locker');
    const lockedAgain = await until(() => js2(win, `Boolean(document.querySelector('.lk-card')) && !document.querySelector('.lk-add')`), 'the unlock card after closing', 10000).then(() => true, () => false);
    check('locker: closing its window locks it', lockedAgain);
    home.destroy();
  } catch (err) {
    const seen = await win?.webContents.executeJavaScript('document.body.textContent').catch(() => '');
    await capture(win, 'locker-failed.png').catch(() => {});
    check('locker: the whole flow ran', false, `${err?.message || err} | the window said: ${String(seen).replace(/s+/g, ' ').slice(0, 400)}`);
  } finally {
    await win?.webContents.executeJavaScript(`window.rutbaOffice.locker.lock().then(() => 1)`).catch(() => {});
    await wait(300);
  }
}
