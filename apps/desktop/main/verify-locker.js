// The Locker's window, end to end, in a build made with its switch: setting up,
// adding a folder, a file dropped in being locked where it is, a file opened
// into Documents and its Save written back to the locked copy, exporting a
// copy and to share, locking and unlocking (a wrong password changing nothing),
// and unlocking for good. It runs only in a run started with RUTBA_LOCKER=1;
// the release gate's full run has no Locker to check.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';

const LOCKED_ROW = (rel) => `tr[data-rel=${JSON.stringify(rel)}][data-state=locked]`;

export async function verifyLocker({ open, check, until, wait, errorsIn, doc, capture: shot = async () => {} }, { dir, files }) {
  // An off-screen window paints only when asked, so each capture asks first.
  const capture = async (w, name) => {
    w.webContents.invalidate();
    await wait(700);
    await shot(w, name);
  };
  if (process.env.RUTBA_LOCKER !== '1') return;
  const answers = globalThis.__rutbaCheckDialogAnswers;
  const plan = (kind, answer) => answers.push({ kind, answer });
  const root = fs.mkdtempSync(path.join(dir, 'locker-folder-'));
  const out = fs.mkdtempSync(path.join(dir, 'locker-out-'));
  const pdfFile = path.join(root, 'Lease, signed.pdf');
  const pdfBytes = crypto.randomBytes(120000);
  fs.writeFileSync(pdfFile, pdfBytes);
  const words = path.join(root, 'Report.docx');
  fs.copyFileSync(files.docx, words);
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
    const made = await clickText('Make the locker');
    await until(() => js(`Boolean(document.querySelector('[data-recovery]'))`), 'the recovery key', 15000);
    const recovery = await js(`document.querySelector('[data-recovery]').textContent`);
    check('locker: a locker is made with a password, and the recovery key is shown', made === 'clicked' && /^[A-Z2-9-]+$/.test(recovery));
    await capture(win, 'locker-setup.png');
    await clickText('I have kept it');
    await until(() => js(`Boolean(document.querySelector('.lk-add'))`), 'the main window', 8000);
    check('locker: after setup it is unlocked, and asks for a folder', /locked folders/i.test(await bodyText()) && !(await js(`document.querySelectorAll('tbody tr').length`)));

    // Add a folder: the choices are asked, a suggestion on each, and the files are locked.
    plan('open', [root]);
    await clickSel('.lk-add');
    await until(() => js(`Boolean(document.querySelector('.lk-add-dialog'))`), 'the add-folder dialog', 6000);
    await capture(win, 'locker-add-folder.png');
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
    const tmpFileRow = await js(`document.querySelector('tr[data-rel="~$Report.docx"]')?.textContent`);
    check('locker: Office\'s own temporary file stays plain, and says why', /Not locked/.test(tmpFileRow || '') && /temporary/.test(tmpFileRow || ''), tmpFileRow);
    await capture(win, 'locker-main.png');

    // A file arriving is locked by the watcher while the locker is open.
    fs.writeFileSync(path.join(root, 'Dropped in.txt'), 'dropped in by the person');
    const lockedDrop = await until(() => js(`Boolean(document.querySelector('tr[data-rel="Dropped in.txt"][data-state="locked"]'))`), 'the dropped file to lock', 20000).then(() => true, () => false);
    check('locker: a file dropped into the folder is locked as it arrives', lockedDrop && fs.existsSync(path.join(root, 'Dropped in.txt.rlock')) && !fs.existsSync(path.join(root, 'Dropped in.txt')), fs.readdirSync(root).join(', '));

    // Export a copy: the locked file is untouched and the copy is the original.
    const dest = path.join(out, 'copy.pdf');
    const lockedBefore = fs.readFileSync(`${pdfFile}.rlock`);
    await clickSel(LOCKED_ROW('Lease, signed.pdf'));
    await until(() => js(`document.querySelector('tr[data-rel="Lease, signed.pdf"]')?.classList.contains('sel') && !document.querySelector('.lk-export').disabled`), 'the row selected', 4000);
    plan('save', dest);
    await clickSel('.lk-export');
    await until(() => fs.existsSync(dest), 'the exported copy', 8000);
    check('locker: Export copy writes the original and leaves the locked file as it was', fs.readFileSync(dest).equals(pdfBytes) && fs.readFileSync(`${pdfFile}.rlock`).equals(lockedBefore));

    // Open into Documents: a temporary copy in the profile, and its Save reaches the locked file.
    await clickSel(LOCKED_ROW('Report.docx'));
    await until(() => js(`!document.querySelector('.lk-open').disabled`), 'Open enabled', 4000);
    const before = fs.readFileSync(`${words}.rlock`);
    await clickSel('.lk-open');
    // What until() waited for is not what it returns, so the window is kept here.
    let opened = null;
    await until(() => {
      opened = BrowserWindow.getAllWindows().find((x) => /locker/.test(decodeURIComponent(x.webContents.getURL())) && /[?&]file=/.test(x.webContents.getURL())) || null;
      return Boolean(opened);
    }, 'the document window', 15000);
    const url = decodeURIComponent(opened.webContents.getURL());
    const tempPath = /[?&]file=([^&]+)/.exec(url)?.[1] || '';
    check('locker: Open puts a temporary copy in the profile folder and opens it in Documents', tempPath && !tempPath.startsWith(root) && /locker[\\/]open/.test(tempPath), tempPath);
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

    // The other windows a person meets, each looked at; sharing one file for real.
    const look = async (button, name) => {
      await clickSel(button);
      await until(() => js(`Boolean(document.querySelector('.lk-dialog-body'))`), name, 5000);
      await capture(win, name);
      await clickText('Cancel');
      await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the dialog to close', 5000);
    };
    await look('.lk-rules', 'locker-rules.png');
    await look('.lk-folder-settings', 'locker-folder-settings.png');
    await look('.lk-settings', 'locker-settings.png');
    await clickSel(LOCKED_ROW('Lease, signed.pdf'));
    await until(() => js(`!document.querySelector('.lk-share').disabled`), 'Share enabled', 4000);
    await clickSel('.lk-share');
    await until(() => js(`Boolean(document.querySelector('.lk-dialog-body'))`), 'the share dialog', 5000);
    await setField('A password for this one file', 'one file only');
    await capture(win, 'locker-share.png');
    const shared = path.join(out, 'shared.rlock');
    plan('save', shared);
    await clickText('Export…');
    await until(() => fs.existsSync(shared), 'the shared file', 8000);
    const opener = await js(`window.rutbaOffice.locker.openShared({ file: ${JSON.stringify(shared)}, password: 'one file only', dest: ${JSON.stringify(path.join(out, 'shared.pdf'))} }).then(() => 'opened', (e) => e.message)`);
    check('locker: Export to share makes a file that opens with its own password and holds the original', opener === 'opened' && fs.readFileSync(path.join(out, 'shared.pdf')).equals(pdfBytes), opener);
    await until(() => js(`!document.querySelector('.lk-dialog-body')`), 'the share dialog to close', 5000);

    // Unlock for good.
    await clickSel(LOCKED_ROW('Lease, signed.pdf'));
    await until(() => js(`!document.querySelector('.lk-unlock').disabled`), 'Unlock enabled', 4000);
    plan('message', { response: 0, checked: false });
    await clickSel('.lk-unlock');
    const backOnDisk = await until(() => fs.existsSync(pdfFile) && !fs.existsSync(`${pdfFile}.rlock`), 'the plain file back', 8000).then(() => true, () => false);
    check('locker: Unlock for good puts the plain file back and removes the locked one', backOnDisk && fs.readFileSync(pdfFile).equals(pdfBytes));
    await wait(2500);
    check('locker: a file unlocked for good stays plain; the watcher does not lock it again', fs.existsSync(pdfFile) && !fs.existsSync(`${pdfFile}.rlock`));

    const home = await open('home');
    const tile = await until(() => home.webContents.executeJavaScript(`Boolean(document.querySelector('.home-card[data-app="locker"]'))`), 'the launcher card', 10000).then(() => true, () => false);
    check('locker: a build made with the switch shows the card on the launcher', tile);
    await capture(home, 'locker-home-tile.png');
    home.destroy();

    const errors = await errorsIn(win).catch(() => []);
    check('locker: no errors in the window', !errors.length, JSON.stringify(errors).slice(0, 200));
  } catch (err) {
    const seen = await win?.webContents.executeJavaScript('document.body.textContent').catch(() => '');
    await capture(win, 'locker-failed.png').catch(() => {});
    check('locker: the whole flow ran', false, `${err?.message || err} | the window said: ${String(seen).replace(/s+/g, ' ').slice(0, 400)}`);
  } finally {
    await win?.webContents.executeJavaScript(`window.rutbaOffice.locker.lock().then(() => 1)`).catch(() => {});
    await wait(300);
  }
}
