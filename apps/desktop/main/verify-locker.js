// The Locker's window, end to end, in a build made with its switch: setting up,
// adding a folder, a file dropped in waiting to be locked by hand, a file opened
// into Documents in memory and its Save written back to the locked copy, printing
// one from memory, exporting a copy and to share, locking and unlocking (a wrong
// password changing nothing, a lock under an open document sealing its work and
// offering it back), and unlocking for good. It runs only in a run started with
// RUTBA_LOCKER=1; the release gate's full run has no Locker to check.

import crypto from 'node:crypto';
import os from 'node:os';
import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { readZip } from '@rutba/ooxml/zip';
import { buildXlsx } from '@rutba/ooxml/build';
import { createLocker } from './extras/locker/locker.js';
import { addressParam } from './extras/locker/window-address.js';
import { memoryFiles } from './memory-files.js';
import { createPrintService } from './print.js';
import { labelKeyed } from './verify-keys.js';

const js2 = (w, code) => w.webContents.executeJavaScript(code);
const LOCKED_ROW = (rel) => `tr[data-rel=${JSON.stringify(rel)}][data-state=locked]`;
// The file a window shows, read from its address as a parameter: the address writes a space as "+".
const fileShown = (w) => addressParam(w.webContents.getURL(), 'file') || '';
/** The Office window showing a file the Locker opened, in memory (2.26.0) or from a plain copy, or null. */
const copyWindow = () => BrowserWindow.getAllWindows().find((x) => !x.isDestroyed() && [path.join('locker', 'mem'), path.join('locker', 'open')].some((d) => fileShown(x).includes(d))) || null;
/** Whether the table of files opened in memory still holds this one. */
const heldInMemory = (p) => Boolean(memoryFiles()?.has(p));

/** Every file under a folder. */
function everyFile(dirPath) {
  if (!fs.existsSync(dirPath)) return [];
  const out = [];
  for (const e of fs.readdirSync(dirPath, { withFileTypes: true })) {
    const p = path.join(dirPath, e.name);
    if (e.isDirectory()) out.push(...everyFile(p));
    else if (e.isFile()) out.push(p);
  }
  return out;
}

/**
 * The files under these folders that hold these plain bytes: a stretch from inside them (not the zip
 * header every Word file starts with), or, for words, a Word file with them in its text.
 */
const plainIn = (dirs, plain) => dirs.flatMap(everyFile).filter((f) => {
  const b = fs.readFileSync(f);
  if (plain.length <= 64) {
    if (b.includes(plain)) return true;
    if (b.subarray(0, 2).toString('latin1') !== 'PK') return false;
    try { return wordsIn(f).includes(plain.toString('utf8')); } catch { return false; }
  }
  for (const at of [Math.floor(plain.length / 3), Math.floor((plain.length * 2) / 3)]) if (b.includes(plain.subarray(at, at + 48))) return true;
  return false;
});

/** Whether any stretch of the plain bytes can be found in the locked file. */
function leaks(lockedFile, plain) {
  const locked = fs.readFileSync(lockedFile);
  const size = Math.min(32, plain.length);
  for (let at = 0; at + size <= plain.length; at += 4096) if (locked.includes(plain.subarray(at, at + size))) return true;
  return false;
}

/** The words in a Word file, read from its own XML and not through the app that wrote it. */
function wordsIn(file) {
  const part = readZip(fs.readFileSync(file)).entries.find((e) => e.name.replace(/^\//, '') === 'word/document.xml');
  return part ? Buffer.from(part.data).toString('utf8').replace(/<[^>]+>/g, '') : '';
}

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
  const wordsBytes = fs.readFileSync(words);
  fs.mkdirSync(path.join(root, 'Contracts', '2026'), { recursive: true });
  fs.writeFileSync(path.join(root, 'Contracts', '2026', 'Terms.txt'), 'the terms');
  fs.writeFileSync(path.join(root, '~$Report.docx'), 'lock file');

  let win = null;
  // The window kept up while the Locker's own is closed under a document, and what the first Locker window showed as errors before it closed.
  let keeper = null;
  let earlyErrors = [];
  const live = () => (win && !win.isDestroyed() ? win : null);
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
      const b = ((list) => list.find((x) => ${labelKeyed(text)}(x)) || list.find((x) => x.textContent.trim() === ${JSON.stringify(text)}))([...document.querySelectorAll('button')]);
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
    check('locker: the first screen offers to open a file someone sent', await js(`Boolean(document.querySelector('.lk-open-sent'))`));
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
    await js(`(() => { const r = ((list) => list.find((x) => ${labelKeyed("Visible", { start: true })}(x)) || list.find((x) => x.textContent.startsWith('Visible')))([...document.querySelectorAll('.lk-add-dialog .lk-choice')]); r.querySelector('input').click(); return 1; })()`);
    await wait(100);
    await clickSel('.lk-add-go');
    await until(() => js(`Boolean(document.querySelector('tr[data-rel="Report.docx"][data-state="locked"]'))`), 'the files to be locked', 30000);
    const sealed = fs.existsSync(`${words}.rlock`) && !fs.existsSync(words) && !fs.existsSync(pdfFile) && fs.existsSync(path.join(root, '.rutbalocker-keys'));
    check('locker: the folder\'s files are locked where they are, each its own .rlock, the plain copies gone', sealed, fs.readdirSync(root).join(', '));
    const terms = path.join(root, 'Contracts', '2026', 'Terms.txt.rlock');
    await until(() => fs.existsSync(terms), 'the subfolder file locked', 15000);
    check('locker: none of the plain bytes can be found in the locked files', !leaks(`${pdfFile}.rlock`, pdfBytes) && !leaks(`${words}.rlock`, wordsBytes) && !leaks(terms, Buffer.from('the terms')));
    // A person's first folder keeps a safety copy of what was there, and the window says plainly what it is.
    const safetyShown = await until(() => js(`Boolean(document.querySelector('.lk-safety-kept'))`), 'the safety copy notice', 8000).then(() => true, () => false);
    const safetyWords = await js(`document.querySelector('.lk-safety-kept')?.textContent || ''`);
    const safetyList = await js(`window.rutbaOffice.locker.state().then((s) => s.safety)`);
    check('locker: the first folder keeps a safety copy of its three plain files, and says it is not locked with the password and who can open it',
      safetyShown && safetyList.length === 1 && safetyList[0].files === 3 && /not locked with your Locker password/.test(safetyWords) && /Windows account|Office profile/.test(safetyWords) && /protects? it if the disk is taken/.test(safetyWords) && /All good, remove the safety copy/.test(safetyWords), safetyWords.slice(0, 300));
    await capture(win, 'locker-safety-copy.png');
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

    // Open into Documents (2.26.0): the document in memory, no plain copy on the disk, and its Save reaches the locked file.
    await clickSel(LOCKED_ROW('Report.docx'));
    await until(() => js(`Boolean(document.querySelector('tr.sel .lk-open'))`), 'Open offered', 4000);
    const before = fs.readFileSync(`${words}.rlock`);
    // The list and the folder on the disk just before the open, to hold the open to: it adds no row and no file.
    // (The folder already holds ~$Report.docx, the owner file set up above and listed as Office's own temporary file.)
    const listRows = () => js(`[...document.querySelectorAll('tbody tr')].map((r) => r.dataset.rel || '/' + (r.dataset.dir || '')).join('|')`);
    const onDisk = () => fs.readdirSync(root, { recursive: true }).map(String).sort().join('|');
    const rowsBeforeOpen = await listRows();
    const diskBeforeOpen = onDisk();
    await clickSel('tr.sel .lk-open');
    // What until() waited for is not what it returns, so the window is kept here.
    let opened = null;
    await until(() => {
      opened = copyWindow();
      return Boolean(opened);
    }, 'the document window', 15000);
    const memPath = fileShown(opened);
    // <profile>/locker/mem/<token>/Report.docx: the token is a 0-byte file, so nothing can be written under it.
    const lockerDir = path.dirname(path.dirname(path.dirname(memPath)));
    const token = path.dirname(memPath);
    const inMemory = memPath.includes(`${path.sep}${path.join('locker', 'mem')}${path.sep}`) && !memPath.startsWith(root) && heldInMemory(memPath)
      && !fs.existsSync(memPath) && fs.statSync(token).isFile() && fs.statSync(token).size === 0 && !fs.existsSync(path.join(root, '.rutba-open'));
    const plainOpened = plainIn([lockerDir, root], wordsBytes).filter((f) => !f.startsWith(path.join(lockerDir, 'safety')));
    check('locker: Open puts a document in memory and opens it in Documents, with no plain copy of it on the disk', inMemory && !plainOpened.length, JSON.stringify({ memPath, plainOpened }));
    const rowsWhileOpen = await listRows();
    const diskWhileOpen = onDisk();
    const ownerFile = fs.readFileSync(path.join(root, '~$Report.docx'), 'utf8');
    check('locker: a document open in memory adds no row to the folder\'s list', rowsWhileOpen === rowsBeforeOpen && !/rutba-open/.test(rowsWhileOpen) && diskWhileOpen === diskBeforeOpen && ownerFile === 'lock file',
      JSON.stringify({ rowsBeforeOpen, rowsWhileOpen, added: diskWhileOpen.split('|').filter((f) => !diskBeforeOpen.split('|').includes(f)), ownerFile }));
    await until(() => doc.sessions().some((s) => s.path === memPath), 'the document session', 15000);
    await capture(opened, 'locker-open-in-documents.png');
    const session = doc.sessions().find((s) => s.path === memPath);
    // An unchanged save is rightly no new version, so this one has a change in it.
    await doc.apply({ id: session.id, ops: [{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 1, offset: 0 } }, { op: 'insertText', text: 'Edited while locked. ' }] });
    await doc.save({ id: session.id, path: memPath });
    const savedBack = await until(() => !fs.readFileSync(`${words}.rlock`).equals(before), 'the save written back', 20000).then(() => true, () => false);
    const roundTrip = path.join(out, 'saved.docx');
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.exportCopy({ id: s.folders[0].id, rel: 'Report.docx', dest: ${JSON.stringify(roundTrip)}, overwrite: true }); return 1; })()`);
    const heldNow = heldInMemory(memPath) ? await memoryFiles().read(memPath) : Buffer.alloc(0);
    check('locker: Save from Documents reaches the locked file, and what comes back out is what was saved', savedBack && fs.readFileSync(roundTrip).equals(heldNow) && wordsIn(roundTrip).includes('Edited while locked.'), wordsIn(roundTrip).slice(0, 120));
    // Past one round of the Locker's 3 second check, which is what sees the window showing it.
    await wait(3500);
    try { await doc.close?.({ id: session.id }); } catch { /* the window goes next */ }
    opened.destroy();
    const cleaned = await until(() => !heldInMemory(memPath), 'the document let go', 15000).then(() => true, () => false);
    check('locker: closing the document lets it go from memory', cleaned, memPath);

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
    // After a failed sign-in, while the safety copy is there, it can be restored with no password into a new folder.
    const restoreTo = fs.mkdtempSync(path.join(dir, 'locker-restore-'));
    const already = path.join(restoreTo, 'Already here.txt');
    fs.writeFileSync(already, 'already here');
    check('locker: after a failed sign-in the password screen offers a restore from the safety copy', await js(`Boolean(document.querySelector('.lk-unlock-safety .lk-safety-restore-open'))`));
    await clickSel('.lk-unlock-safety .lk-safety-restore-open');
    await until(() => js(`Boolean(document.querySelector('.lk-safety-restore'))`), 'the restore dialog', 6000);
    await capture(win, 'locker-safety-restore.png');
    const restoreWords = await js(`document.querySelector('.lk-safety-restore').textContent`);
    plan('open', [restoreTo]);
    await clickSel('.lk-safety-restore-go');
    const restoredDir = path.join(restoreTo, `${path.basename(root)} (safety copy)`);
    const restored = await until(() => js(`Boolean(document.querySelector('.lk-safety-restored'))`), 'the files restored', 20000).then(() => true, () => false);
    const sameAsBefore = restored && fs.readFileSync(path.join(restoredDir, 'Lease, signed.pdf')).equals(pdfBytes) && fs.readFileSync(path.join(restoredDir, 'Report.docx')).equals(wordsBytes)
      && fs.readFileSync(path.join(restoredDir, 'Contracts', '2026', 'Terms.txt'), 'utf8') === 'the terms' && fs.readFileSync(already, 'utf8') === 'already here';
    check('locker: the restore needs no password, says so, writes the files as they were before locking into a new folder and touches nothing else',
      /without your password/.test(restoreWords) && sameAsBefore && !(await js(`Boolean(document.querySelector('.lk-add'))`)), fs.existsSync(restoredDir) ? fs.readdirSync(restoredDir).join(', ') : 'no folder');
    await clickText('Close');
    await until(() => js(`!document.querySelector('.lk-safety-restored')`), 'the restore dialog to close', 5000);
    await setField('Locker password', 'correct horse battery');
    await clickText('Unlock with password');
    await until(() => js(`Boolean(document.querySelector('.lk-add'))`), 'unlocked again', 15000);
    await until(() => js(`document.querySelectorAll('tbody tr').length > 0`), 'the files listed again', 8000);
    const namedRows = await js(`[${JSON.stringify(LOCKED_ROW('Report.docx'))}, ${JSON.stringify(LOCKED_ROW('Lease, signed.pdf'))}, ${JSON.stringify(LOCKED_ROW('Contracts/2026/Terms.txt'))}].every((s) => document.querySelector(s))`);
    check('locker: the right password opens it again and the files are back, each listed as locked', namedRows, await visibleRows());

    // All good: the folder is checked against the safety copy first. Report.docx was saved since, so the person is told and asked again.
    await until(() => js(`Boolean(document.querySelector('.lk-safety-remove'))`), 'All good offered', 8000);
    plan('message', { response: 0, checked: false });
    plan('message', { response: 0, checked: false });
    await clickSel('.lk-safety-remove');
    const safetyGone = await until(() => js(`window.rutbaOffice.locker.state().then((s) => !s.safety.length && !document.querySelector('.lk-safety-kept'))`), 'the safety copy removed', 20000).then(() => true, () => false);
    check('locker: All good checks the folder against the safety copy, names a file changed since, and removes the copy once the person accepts that', safetyGone && fs.existsSync(`${words}.rlock`) && fs.existsSync(`${pdfFile}.rlock`));

    // Windows Hello, with the stand-in for the person's PIN: on, and each way the prompt can end.
    const helloUnlocked = () => until(() => js(`Boolean(document.querySelector('.lk-add'))`), 'unlocked by Windows Hello', 8000).then(() => true, () => false);
    await clickSel('.lk-settings');
    await until(() => js(`Boolean(document.querySelector('.lk-advanced input'))`), 'the settings dialog', 5000);
    await clickText('Turn on Windows Hello');
    const helloOn = await until(() => js(`[...document.querySelectorAll('button')].some((x) => x.textContent.trim() === 'Turn off Windows Hello')`), 'Windows Hello on', 6000).then(() => true, () => false);
    check('locker: Windows Hello can be turned on from the settings (with a stand-in for the prompt, no real Windows Hello here)', helloOn);
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
    check('locker: the stand-in for Windows Hello saying yes opens it', await helloUnlocked());
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
    await setField('A password for this one file', 'seven c');
    await wait(150);
    check('locker: a password under eight characters cannot be used to send a file', await js(`[...document.querySelectorAll('button')].filter((x) => /Save the locked copy|Attach to an email/.test(x.textContent)).every((x) => x.disabled)`));
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
    check('locker: the step after saving offers an email and Explorer', await js(`Boolean(document.querySelector('.lk-share-mail') && document.querySelector('.lk-share-show'))`));
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
    check('locker: Attach to an email is offered at the start and opens a message with the locked copy', attachedNow && [...new URL(mailNow.webContents.getURL()).searchParams.values()].some((v) => v.includes(path.join('locker', 'outgoing'))), mailNow.webContents.getURL());
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
      docWin = copyWindow();
      return Boolean(docWin);
    }, 'the document window for the idle check', 15000);
    // Nothing should happen here, so it is watched for one and three quarter times the idle time, which is how long is long enough.
    const idleSpan = 0.07 * 60000;
    let lockedUnder = false;
    for (const until_ = Date.now() + idleSpan * 1.75; Date.now() < until_ && !lockedUnder;) {
      lockedUnder = await isLockedNow();
      await wait(200);
    }
    // And waits because it sees the document's window, not only because a copy handed over gets 30 seconds to open in.
    const idleSeen = await stateNow();
    check('locker: the idle lock waits while a document from it is open in an Office window, and knows that window shows it', !lockedUnder && idleSeen.open === true && idleSeen.shownCopies === 1, JSON.stringify({ lockedUnder, shownCopies: idleSeen.shownCopies, file: fileShown(docWin) }));
    docWin.close();
    const lockedAfterDoc = await until(isLockedNow, 'the idle lock after the document closed', 14000).then(() => true, () => false);
    check('locker: once that document is closed the idle lock goes ahead', lockedAfterDoc);
    await js(`window.rutbaOffice.store.set({ key: 'locker.idleMinutes', value: 10 }).then(() => 1)`);
    await unlockAgain();

    // Locking under an open document (owner, 13 Nov, D-2026-11-13-22): its unsaved work is sealed as a locked
    // recovery copy, its window closes, nothing plain is left, and the next unlock offers it back.
    // What earlier locks offered is turned down first, so the one offer here is this document's.
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.dismissReopen({ keys: (s.reopen || []).map((o) => o.key) }); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Report.docx'))}))`), 'the document row again', 8000);
    await clickSel(LOCKED_ROW('Report.docx'));
    await until(() => js(`Boolean(document.querySelector('tr.sel .lk-open'))`), 'Open offered again', 3000);
    await clickSel('tr.sel .lk-open');
    let heldWin = null;
    await until(() => {
      heldWin = copyWindow();
      return Boolean(heldWin);
    }, 'the document window to lock under', 15000);
    const heldPath = fileShown(heldWin);
    await until(() => doc.sessions().some((s) => s.path === heldPath), 'that document\'s session', 15000);
    const heldBefore = fs.readFileSync(`${words}.rlock`);
    check('locker: a document opened from it is known to be showing in its Office window, a space in its path or not', (await stateNow()).shownCopies === 1 && /\s/.test(heldPath), heldPath);
    const heldSession = doc.sessions().find((s) => s.path === heldPath);
    const unsavedWords = 'Unsaved when the Locker locked.';
    await doc.apply({ id: heldSession.id, ops: [{ op: 'setSelection', anchor: { block: 1, offset: 0 }, focus: { block: 1, offset: 0 } }, { op: 'insertText', text: `${unsavedWords} ` }] });
    await js(`window.rutbaOffice.locker.lock().then(() => 1)`);
    await until(isLockedNow, 'locked under the open document', 8000);
    const closedUnder = await until(() => heldWin.isDestroyed(), 'its window to close', 8000).then(() => true, () => false);
    const plainUnsaved = plainIn([lockerDir, root], Buffer.from(unsavedWords));
    check('locker: locking under an open document closes its window and lets it go from memory, with nothing plain of its unsaved work on the disk',
      closedUnder && !heldInMemory(heldPath) && !doc.sessions().some((s) => s.path === heldPath) && !plainUnsaved.length, JSON.stringify({ closedUnder, plainUnsaved }));
    check('locker: what was unsaved when it locked does not reach the locked file', fs.readFileSync(`${words}.rlock`).equals(heldBefore));
    await unlockAgain();
    const reopenOffered = await until(() => js(`/Report\\.docx, with unsaved changes/.test(document.querySelector('.lk-reopen')?.textContent || '')`), 'the offer to reopen', 10000).then(() => true, () => false);
    check('locker: the next unlock offers to reopen it, with its unsaved changes', reopenOffered, await js(`document.querySelector('.lk-reopen')?.textContent || ''`));
    await capture(win, 'locker-reopen-offer.png');
    await clickSel('.lk-reopen-yes');
    let backWin = null;
    await until(() => { backWin = copyWindow(); return Boolean(backWin); }, 'the reopened document window', 15000);
    const backPath = fileShown(backWin);
    await until(() => doc.sessions().some((s) => s.path === backPath), 'the reopened session', 15000);
    const back = doc.sessions().find((s) => s.path === backPath);
    const hits = await doc.search({ id: back.id, query: unsavedWords });
    check('locker: Reopen opens it again with that work, marked unsaved', hits.length > 0 && back.dirty === true, JSON.stringify({ hits: hits.length, dirty: back.dirty }));
    await doc.save({ id: back.id, path: backPath });
    const heldIn = await until(() => !fs.readFileSync(`${words}.rlock`).equals(heldBefore), 'the reopened work saved', 15000).then(() => true, () => false);
    const heldOut = path.join(out, 'held.docx');
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.exportCopy({ id: s.folders[0].id, rel: 'Report.docx', dest: ${JSON.stringify(heldOut)}, overwrite: true }); return 1; })()`);
    check('locker: saved after the reopen, that work reaches the locked file', heldIn && wordsIn(heldOut).includes(unsavedWords), wordsIn(heldOut).slice(0, 120));
    // Past one round of the Locker's 3 second check, which is what sees the window showing it.
    await wait(3500);
    try { await doc.close?.({ id: back.id }); } catch { /* the window goes next */ }
    backWin.destroy();
    const heldCleared = await until(() => !heldInMemory(backPath), 'the reopened document let go', 15000).then(() => true, () => false);
    check('locker: closing that document lets it go from memory', heldCleared);

    // Closing the Locker's window while a document from it is open: the lock waits for the document, as the idle lock does.
    // This run's profile has a space in its path (verify-apps-run.js), and so has every copy the Locker opens in it.
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Report.docx'))}))`), 'the document row for closing', 8000);
    await clickSel(LOCKED_ROW('Report.docx'));
    await until(() => js(`Boolean(document.querySelector('tr.sel .lk-open'))`), 'Open offered for closing', 3000);
    await clickSel('tr.sel .lk-open');
    let keptWin = null;
    await until(() => { keptWin = copyWindow(); return Boolean(keptWin); }, 'the document window to close the Locker under', 15000);
    const keptPath = fileShown(keptWin);
    await until(() => doc.sessions().some((s) => s.path === keptPath), 'the session to close the Locker under', 15000);
    // The run ends when its last window closes, so one more stays up meanwhile.
    keeper = await open('home');
    const lockerOpen = async () => (await js2(keeper, `window.rutbaOffice.locker.state()`)).open === true;
    earlyErrors = [...await errorsIn(win), ...await js(`[...document.querySelectorAll('.lk-error')].map((n) => n.textContent)`)];
    win.close();
    await until(() => win.isDestroyed(), 'the Locker window to close under the document', 8000);
    // Past two rounds of the 3 second check that locks a closed Locker.
    await wait(7000);
    check('locker: closing its window while a document from it is open waits for the document, and keeps it open in memory', await lockerOpen() && heldInMemory(keptPath) && /\s/.test(keptPath), keptPath);
    const keptSession = doc.sessions().find((s) => s.path === keptPath);
    try { await doc.close?.({ id: keptSession.id }); } catch { /* the window goes next */ }
    keptWin.destroy();
    const lockedAfterKept = await until(async () => !(await lockerOpen()) && !heldInMemory(keptPath), 'the lock once the document closed', 15000).then(() => true, () => false);
    check('locker: once that document is closed the closed Locker locks, and lets it go from memory', lockedAfterKept);
    win = await open('locker');
    await until(isLockedNow, 'the unlock card after closing under a document', 10000);
    await unlockAgain();
    keeper.destroy();
    keeper = null;

    // Printing a document from it (2.26.0): each job held in memory for the hidden window, never a file in the
    // temp folder, and Export as PDF of one refused. A document (a PDF job) and a workbook (a page of HTML),
    // each opened from the Locker in its own window, as a person would print it.
    fs.writeFileSync(path.join(root, 'Budget.xlsx'), buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Budget', 1200]] }] }));
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.lockWaiting({ id: s.folders[0].id }); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Budget.xlsx'))}))`), 'the workbook locked', 15000);
    const printer = createPrintService({ docs: doc });
    const jobs = {};
    const tempBefore = new Set(fs.readdirSync(os.tmpdir()).filter((n) => n.startsWith('rutba-print-')));
    for (const rel of ['Report.docx', 'Budget.xlsx']) {
      await clickSel(LOCKED_ROW(rel));
      await until(() => js(`Boolean(document.querySelector('tr.sel .lk-open'))`), `Open offered for ${rel}`, 4000);
      await clickSel('tr.sel .lk-open');
      let w = null;
      await until(() => { w = copyWindow(); return Boolean(w); }, `the window for ${rel}`, 15000);
      const p = fileShown(w);
      await until(() => doc.sessions().some((s) => s.path === p), `the session for ${rel}`, 15000);
      const s = doc.sessions().find((x) => x.path === p);
      jobs[rel] = await printer.jobProbe({ id: s.id }).catch((err) => ({ failed: String(err?.message || err) }));
      if (rel === 'Report.docx') jobs.exported = await printer.pdf({ id: s.id, path: path.join(out, 'Report.pdf') }).then(() => 'written', (err) => err?.code || String(err?.message || err));
      // Past one round of the Locker's 3 second check, which is what sees the window showing it.
      await wait(3500);
      try { await doc.close?.({ id: s.id }); } catch { /* the window goes next */ }
      w.destroy();
      await until(() => !heldInMemory(p), `${rel} let go`, 15000).catch(() => {});
    }
    jobs.newTemp = fs.readdirSync(os.tmpdir()).filter((n) => n.startsWith('rutba-print-') && !tempBefore.has(n));
    const fromMemory = (j, kind) => Boolean(j) && !j.failed && j.kind === kind && j.locked === true && /^rutba:\/\/file\//.test(j.url || '') && j.temp === null && j.held === true;
    check('locker: printing a document from it loads the job from memory, a document and a workbook alike, with no file in the temp folder',
      fromMemory(jobs['Report.docx'], 'pdf') && fromMemory(jobs['Budget.xlsx'], 'html') && jobs['Budget.xlsx'].printed > 0 && !jobs.newTemp.length, JSON.stringify(jobs).slice(0, 400));
    check('locker: Export as PDF of a document from it is refused, and writes nothing', jobs.exported === 'LOCKER_ONLY' && !fs.existsSync(path.join(out, 'Report.pdf')), String(jobs.exported));
    check('locker: each print job is let go when it ends', !memoryFiles()?.entries().some((e) => e.info?.job));

    // Renaming and deleting from the Locker's own list.
    const renameMe = path.join(root, 'Rename me.txt');
    const deleteMe = path.join(root, 'Delete me.txt');
    fs.writeFileSync(renameMe, 'keep these words');
    fs.writeFileSync(deleteMe, 'and lose these');
    // Its own file for the two versions, so that part does not hang on Rename.
    const twoVersions = path.join(root, 'Two versions.txt');
    fs.writeFileSync(twoVersions, 'keep these words');
    await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.lockWaiting({ id: s.folders[0].id }); return 1; })()`);
    const threeRows = JSON.stringify(['Rename me.txt', 'Delete me.txt', 'Two versions.txt'].map(LOCKED_ROW));
    await until(() => js(`${threeRows}.every((q) => document.querySelector(q))`), 'the three files locked', 15000);
    // Each is opened in Documents and closed again first, with the Locker still unlocked.
    // 2.25.0 never saw such a copy shown, so kept it until the next lock, and
    // Rename, Delete and deleting a version said "This file is open".
    const saidOpen = [];
    // Read where each step says it, at the time: a refused Delete is only a toast, gone in 3 s.
    const saysOpen = async (where) => /This file is open/.test(await js(`[...document.querySelectorAll(${JSON.stringify(where)})].map((n) => n.textContent).join(' ')`));
    const notOpen = async (step, where) => {
      if (await saysOpen(where)) saidOpen.push(step);
    };
    const openAndClose = async (rel) => {
      await clickSel(LOCKED_ROW(rel));
      const openIt = `tr.sel[data-rel=${JSON.stringify(rel)}] .lk-open`;
      await until(() => js(`Boolean(document.querySelector(${JSON.stringify(openIt)}))`), `Open offered for ${rel}`, 4000);
      await clickSel(openIt);
      let w = null;
      await until(() => { w = copyWindow(); return Boolean(w); }, `${rel} in Documents`, 15000);
      const copy = fileShown(w);
      await until(() => doc.sessions().some((s) => s.path === copy), `the session of ${rel}`, 15000);
      // 2.25.0 never knew; the run goes on, to show what that did to Rename and Delete.
      await until(async () => (await stateNow()).shownCopies === 1, `${rel} known to be showing`, 10000).catch(() => {});
      // Open past one round of the Locker's 3 second check, which is what sees it showing.
      await wait(4000);
      const s = doc.sessions().find((x) => x.path === copy);
      try { await doc.close?.({ id: s.id }); } catch { /* the window goes next */ }
      w.destroy();
      const cleared = await until(() => !heldInMemory(copy) && !fs.existsSync(path.dirname(copy)), `${rel} let go`, 15000).then(() => true, () => false);
      return cleared ? '' : copy;
    };
    const lingering = [await openAndClose('Rename me.txt'), await openAndClose('Delete me.txt')].filter(Boolean);
    check('locker: a file opened in Documents and closed while it is unlocked is let go, not kept until the next lock', !lingering.length, lingering.join(', '));
    await clickSel(LOCKED_ROW('Rename me.txt'));
    await until(() => js(`Boolean(document.querySelector('.lk-rename'))`), 'Rename offered', 4000);
    await clickSel('.lk-rename');
    await until(() => js(`Boolean(document.querySelector('.lk-dialog-body'))`), 'the rename dialog', 5000);
    await setField('New name', 'Renamed.txt');
    await clickText('Rename');
    const renamed = await until(() => fs.existsSync(path.join(root, 'Renamed.txt.rlock')) && !fs.existsSync(`${renameMe}.rlock`), 'the file renamed on the disk', 10000).then(() => true, () => false);
    await notOpen('Rename', '.lk-dialog-body .lk-error');
    // A refused rename leaves its dialog up.
    if (!renamed) await js(`((list) => list.find((b) => ${labelKeyed("Cancel")}(b)) || list.find((b) => b.textContent.trim() === 'Cancel'))([...document.querySelectorAll('button')])?.click(), 1`);
    const renamedRow = await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Renamed.txt'))}))`), 'the renamed row', 8000).then(() => true, () => false);
    const renamedOut = path.join(out, 'renamed.txt');
    // With no Renamed.txt (Rename refused) there is nothing to export, and the run goes on.
    if (renamed) await js(`(async () => { const s = await window.rutbaOffice.locker.state(); await window.rutbaOffice.locker.exportCopy({ id: s.folders[0].id, rel: 'Renamed.txt', dest: ${JSON.stringify(renamedOut)}, overwrite: true }); return 1; })()`);
    check('locker: Rename gives a locked file a new name, on the disk and in the list, with the same contents', renamed && renamedRow && fs.readFileSync(renamedOut, 'utf8') === 'keep these words');
    await clickSel(LOCKED_ROW('Delete me.txt'));
    await until(() => js(`Boolean(document.querySelector('.lk-delete'))`), 'Delete offered', 4000);
    await capture(win, 'locker-file-actions.png');
    plan('message', { response: 0, checked: false });
    await clickSel('.lk-delete');
    let deleteRefused = false;
    const deleted = await until(async () => {
      if (await saysOpen('.rw-toast')) deleteRefused = true;
      return deleteRefused || !fs.existsSync(`${deleteMe}.rlock`);
    }, 'the file deleted', 10000).then(() => !fs.existsSync(`${deleteMe}.rlock`), () => false);
    if (deleteRefused) saidOpen.push('Delete');
    const deletedRow = await until(() => js(`!document.querySelector('tr[data-rel="Delete me.txt"]')`), 'its row gone', 8000).then(() => true, () => false);
    check('locker: Delete removes a locked file for good, and its row', deleted && deletedRow);
    await capture(win, 'locker-renamed-deleted.png');
    check('locker: closing a document raises no false alarm about its copy', !(await js(`document.body.textContent.includes('not opened from the locker')`)));

    // Two versions from the cloud: each saved on its own, then one removed from the window.
    const firstVersion = `${twoVersions}.rlock`;
    const secondVersion = path.join(root, 'Two versions (conflicted copy).txt.rlock');
    const versionLingers = await openAndClose('Two versions.txt');
    fs.copyFileSync(firstVersion, secondVersion);
    const offered = await until(() => js(`Boolean(document.querySelector('.lk-compare'))`), 'the two versions to be noticed', 20000).then(() => true, () => false);
    check('locker: two versions of a file from the cloud are noticed and offered to compare', offered);
    await clickSel('.lk-compare');
    await until(() => js(`document.querySelectorAll('.lk-version').length === 2`), 'both versions in the dialog', 5000);
    await capture(win, 'locker-two-versions.png');
    const savedVersion = path.join(out, 'one version.txt');
    plan('save', savedVersion);
    await js(`[...document.querySelectorAll('.lk-version')][1].querySelector('button').click(), 1`);
    await until(() => fs.existsSync(savedVersion), 'one version saved', 8000);
    check('locker: each of the two versions can be saved', fs.readFileSync(savedVersion, 'utf8') === 'keep these words');
    plan('message', { response: 0, checked: false });
    await js(`[...document.querySelectorAll('.lk-version-delete')][1].click(), 1`);
    const oneLeft = await until(() => js(`!document.querySelector('.lk-compare') && Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Two versions.txt'))}))`), 'one version left', 10000).then(() => true, () => false);
    await notOpen('Delete this version', '.lk-error');
    check('locker: deleting one of two versions leaves the other as the file', oneLeft && [firstVersion, secondVersion].filter((p) => fs.existsSync(p)).length === 1, fs.readdirSync(root).join(', '));
    check('locker: after a file was opened and closed, Rename, Delete and deleting one of two versions are not refused as open', !saidOpen.length && !versionLingers, JSON.stringify({ saidOpen, versionLingers }));

    const home = await open('home');
    const tile = await until(() => home.webContents.executeJavaScript(`Boolean(document.querySelector('.home-card[data-app="locker"]'))`), 'the launcher card', 10000).then(() => true, () => false);
    check('locker: a build made with the switch shows the card on the launcher', tile);
    await capture(home, 'locker-home-tile.png');

    // A folder locked by another locker: recognised when chosen, opened, its files as they were.
    const lockedElsewhere = fs.mkdtempSync(path.join(dir, 'locker-elsewhere-'));
    const elsewhereBytes = crypto.randomBytes(5000);
    fs.writeFileSync(path.join(lockedElsewhere, 'Elsewhere.pdf'), elsewhereBytes);
    const otherLocker = createLocker({ dir: fs.mkdtempSync(path.join(dir, 'locker-other-')), cost: { N: 1024, r: 8, p: 1 } });
    await otherLocker.setup('a password from elsewhere');
    const { id: otherId } = await otherLocker.addFolder(lockedElsewhere, { hidden: false });
    await otherLocker.lockWaiting(otherId);
    const listing = () => fs.readdirSync(lockedElsewhere).sort().join();
    const keyBytes = () => fs.readFileSync(path.join(lockedElsewhere, '.rutbalocker-keys'), 'utf8');
    const listedBefore = listing();
    const keysBefore = keyBytes();
    await until(() => js(`Boolean(document.querySelector('.lk-add'))`), 'the main window again', 10000);
    plan('open', [lockedElsewhere]);
    await clickSel('.lk-add');
    const openDialog = await until(() => js(`Boolean(document.querySelector('.lk-open-locked'))`), 'the open-a-locked-folder dialog', 6000).then(() => true, () => false);
    const openWords = await js(`document.querySelector('.lk-open-locked')?.textContent || document.querySelector('.lk-add-dialog')?.textContent || ''`);
    check('locker: a folder that is already locked is recognised when chosen, and asks for its password with no error and no offer to lock it',
      openDialog && /already locked/.test(openWords) && !(await js(`Boolean(document.querySelector('.lk-open-locked .lk-error')) || Boolean(document.querySelector('.lk-add-go'))`)), openWords.slice(0, 300));
    await capture(win, 'locker-open-locked.png');
    await setField('Password that locked this folder', 'not the right one');
    await clickSel('.lk-open-go');
    await until(() => js(`Boolean(document.querySelector('.lk-open-locked .lk-error'))`), 'the refusal', 15000).catch(() => false);
    const refused = await js(`document.querySelector('.lk-open-locked .lk-error')?.textContent || ''`);
    check('locker: a wrong password says so and changes nothing in the folder', /does not open/.test(refused || '') && listing() === listedBefore && keyBytes() === keysBefore, String(refused));
    await js(`document.querySelector('.lk-open-switch')?.click(), 1`);
    const recoveryField = await until(() => js(`[...document.querySelectorAll('.lk-open-locked input')].some((i) => i.getAttribute('aria-label') === 'Recovery key that opens this folder')`), 'the recovery key field', 4000).then(() => true, () => false);
    check('locker: the same dialog takes the recovery key instead', recoveryField);
    await js(`document.querySelector('.lk-open-switch')?.click(), 1`);
    await until(() => js(`[...document.querySelectorAll('.lk-open-locked input')].some((i) => i.getAttribute('aria-label') === 'Password that locked this folder')`), 'the password field again', 4000);
    await setField('Password that locked this folder', 'a password from elsewhere');
    await wait(200);
    await clickSel('.lk-open-go');
    const added = await until(() => js(`!document.querySelector('.lk-open-locked')`), 'the folder to open', 20000).then(() => true, () => false);
    const openedRow = await until(() => js(`Boolean(document.querySelector(${JSON.stringify(LOCKED_ROW('Elsewhere.pdf'))}))`), 'its locked file listed', 10000).then(() => true, () => false);
    check('locker: the right password opens the folder, its files listed as locked, and the folder on disk as it was',
      added && openedRow && listing() === listedBefore && keyBytes() === keysBefore, JSON.stringify({ added, openedRow, listing: listing() }));

    // A folder with something locked inside it says so before anything happens to it.
    const parent = fs.mkdtempSync(path.join(dir, 'locker-parent-'));
    const inner = path.join(parent, 'Inner');
    fs.mkdirSync(inner);
    fs.writeFileSync(path.join(inner, 'Inner.pdf'), crypto.randomBytes(2000));
    fs.writeFileSync(path.join(parent, 'Outer.txt'), 'outer');
    const { id: innerId } = await otherLocker.addFolder(inner, { hidden: false });
    await otherLocker.lockWaiting(innerId);
    plan('open', [parent]);
    await clickSel('.lk-add');
    await until(() => js(`Boolean(document.querySelector('.lk-add-dialog'))`), 'the add-folder dialog for the parent', 6000);
    check('locker: adding a folder with a locked folder inside it warns first, and has changed nothing', await js(`/key of its own/.test(document.querySelector('.lk-locked-below')?.textContent || '')`) && fs.existsSync(path.join(parent, 'Outer.txt')) && !fs.existsSync(path.join(parent, '.rutbalocker-keys')));
    await clickText('Cancel');
    await until(() => js(`!document.querySelector('.lk-add-dialog')`), 'the add-folder dialog to close', 5000);

    const errors = await errorsIn(win);
    const leftError = await js(`[...document.querySelectorAll('.lk-error')].map((n) => n.textContent)`);
    check('locker: no errors in the window', !errors.length && !leftError.length && !earlyErrors.length, JSON.stringify([earlyErrors, errors, leftError]).slice(0, 200));

    // Closing the Locker's window locks it. The launcher stays open, or the last window closing would end the run.
    win.close();
    await until(() => win.isDestroyed(), 'the Locker window to close', 8000);
    await wait(800);
    win = await open('locker');
    const lockedAgain = await until(() => js2(win, `Boolean(document.querySelector('.lk-card')) && !document.querySelector('.lk-add')`), 'the unlock card after closing', 10000).then(() => true, () => false);
    check('locker: closing its window locks it', lockedAgain);
    home.destroy();
  } catch (err) {
    const seen = await live()?.webContents.executeJavaScript('document.body.textContent').catch(() => '');
    if (live()) await capture(live(), 'locker-failed.png').catch(() => {});
    check('locker: the whole flow ran', false, `${err?.message || err} | the window said: ${String(seen).replace(/s+/g, ' ').slice(0, 400)}`);
  } finally {
    await (live() || (keeper && !keeper.isDestroyed() ? keeper : null))?.webContents.executeJavaScript(`window.rutbaOffice.locker.lock().then(() => 1)`).catch(() => {});
    // Not the last window, or the run would end with it.
    if (keeper && !keeper.isDestroyed() && BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed()).length > 1) keeper.destroy();
    await wait(300);
  }
}
