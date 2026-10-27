// Contacts' and Calendar's Import, through the Open dialog as a person uses
// it: the file chosen there comes in. Since 1.8.0 both took the first letter
// of the path chosen ("D") and said ENOENT, whichever way Import was asked
// for; the checks before this one gave the window its file directly. The
// dialog's answer is planned (ipc.js plannedAnswer), and what came in is
// removed again, so the checks after these find the books as they were. Run
// alone with RUTBA_VERIFY_ONLY=import-dialog.
import fs from 'node:fs';
import path from 'node:path';

/** Press Import on the ribbon (or the empty page's Import…) with the dialog's answer planned. */
async function importThrough(win, file) {
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [file] });
  return win.webContents.executeJavaScript(`(() => {
    const all = [...document.querySelectorAll('button:not([disabled])')];
    const b = all.find((x) => x.textContent.trim() === 'Import') || all.find((x) => x.textContent.trim() === 'Import…');
    b?.click();
    return b ? b.textContent.trim() : null;
  })()`);
}

const toastsOf = (win) => win.webContents.executeJavaScript(`[...document.querySelectorAll('.rw-toast')].map((t) => t.textContent).join(' | ')`);

/** The dialog's planned answer, taken back if Import never asked for it. */
function unplan() {
  const q = globalThis.__rutbaCheckDialogAnswers;
  if (!q) return;
  for (let i = q.length - 1; i >= 0; i--) if (q[i].kind === 'open') q.splice(i, 1);
}

export async function verifyContactsImportDialog({ open, check, until, errorsIn }, { dir }) {
  const own = path.join(dir, 'import-dialog');
  fs.mkdirSync(own, { recursive: true });
  const vcf = path.join(own, 'one card.vcf');
  fs.writeFileSync(vcf, ['BEGIN:VCARD', 'VERSION:3.0', 'N:Check;Dialog;;;', 'FN:Dialog Check', 'EMAIL:dialog.check@example.com', 'END:VCARD', ''].join('\r\n'));
  let win = null;
  let before = null;
  try {
    win = await open('contacts');
    const js = (code) => win.webContents.executeJavaScript(code);
    before = await js(`window.rutbaOffice.contacts.list({}).then((l) => l.map((c) => c.id))`);
    const pressed = await importThrough(win, vcf);
    let said = '';
    await until(async () => /added|ENOENT|error/i.test((said = await toastsOf(win))), 'what Import says', 10000).catch(() => {});
    const came = await js(`window.rutbaOffice.contacts.list({}).then((l) => l.filter((c) => !${JSON.stringify(before)}.includes(c.id)).map((c) => c.name || c.fn || c.displayName || ''))`);
    check('contacts: Import, with a .vcf chosen in the Open dialog, brings its card in, not "ENOENT ... D"',
      Boolean(pressed) && came.length === 1 && /1 added/.test(said) && !/ENOENT/.test(said), JSON.stringify({ pressed, said, came }));
    const complaints = await errorsIn(win);
    check('contacts: Import through the dialog reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('contacts: the Import-through-the-dialog check ran', false, err.message);
  } finally {
    unplan();
    if (win && !win.isDestroyed() && before) {
      await win.webContents.executeJavaScript(`(async () => {
        const added = (await window.rutbaOffice.contacts.list({})).map((c) => c.id).filter((id) => !${JSON.stringify(before)}.includes(id));
        if (added.length) await window.rutbaOffice.contacts.remove({ ids: added });
        return added.length;
      })()`).catch(() => {});
    }
  }
}

export async function verifyCalendarImportDialog({ open, check, until, errorsIn }, { dir }) {
  const own = path.join(dir, 'import-dialog');
  fs.mkdirSync(own, { recursive: true });
  const ics = path.join(own, 'one event.ics');
  const day = new Date(Date.now() + 2 * 86400000);
  const ymd = `${day.getUTCFullYear()}${String(day.getUTCMonth() + 1).padStart(2, '0')}${String(day.getUTCDate()).padStart(2, '0')}`;
  fs.writeFileSync(ics, [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Rutba Office checks//EN',
    'BEGIN:VEVENT', `UID:import-dialog-${ymd}@checks.example`, `DTSTAMP:${ymd}T090000Z`,
    `DTSTART:${ymd}T100000Z`, `DTEND:${ymd}T110000Z`, 'SUMMARY:Dialog import check', 'END:VEVENT',
    'END:VCALENDAR', '',
  ].join('\r\n'));
  let win = null;
  try {
    win = await open('calendar');
    const pressed = await importThrough(win, ics);
    let said = '';
    await until(async () => /added|ENOENT|error/i.test((said = await toastsOf(win))), 'what Import says', 10000).catch(() => {});
    check('calendar: Import, with a .ics chosen in the Open dialog, brings its event in, not "ENOENT ... D"',
      Boolean(pressed) && /1 added/.test(said) && !/ENOENT/.test(said), JSON.stringify({ pressed, said }));
    const complaints = await errorsIn(win);
    check('calendar: Import through the dialog reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('calendar: the Import-through-the-dialog check ran', false, err.message);
  } finally {
    unplan();
    if (win && !win.isDestroyed()) {
      const from = day.getTime() - 2 * 86400000;
      await win.webContents.executeJavaScript(`(async () => {
        const all = await window.rutbaOffice.calendar.events({ from: ${from}, to: ${from + 4 * 86400000} });
        for (const e of all.filter((x) => x.summary === 'Dialog import check')) await window.rutbaOffice.calendar.remove({ id: e.id });
        return 1;
      })()`).catch(() => {});
    }
  }
}
