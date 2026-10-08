// A window asking the system to open a program is warned from the main
// process, not from the window: a page calling openPath on a .exe (or a
// "run.exe " that Windows would run as one) gets the warning, and Cancel
// opens nothing. Run alone with RUTBA_VERIFY_ONLY=openprogram.

import fs from 'node:fs';
import path from 'node:path';

export async function verifyOpenProgram({ open, check, until }, { dir }) {
  try {
    const program = path.join(dir, 'invoice.exe');
    fs.writeFileSync(program, 'not really a program');
    const win = await open('home');
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(window.rutbaOffice?.shell)`), 'the window', 8000);
    const queue = globalThis.__rutbaCheckDialogAnswers;
    const before = queue?.length ?? 0;
    // The person says Cancel to the warning, for both spellings of the name.
    queue?.push({ kind: 'message', answer: { response: 1, checked: false } });
    queue?.push({ kind: 'message', answer: { response: 1, checked: false } });
    const plain = await js(`window.rutbaOffice.shell.openPath({ path: ${JSON.stringify(program)} })`);
    const spaced = await js(`window.rutbaOffice.shell.openPath({ path: ${JSON.stringify(`${program} `)} })`);
    const asked = (queue?.length ?? 0) === before;
    check('shell: a program asked to be opened from a window is warned about from the main process, and Cancel opens nothing',
      plain === false && spaced === false && asked, JSON.stringify({ plain, spaced, asked }));
  } catch (err) {
    check('shell: the program warning checks ran', false, err?.message || String(err));
  }
}
