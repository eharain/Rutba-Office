// A drag whose mouseup is lost still ends: in a Worksheets window, a press,
// the pointer out of the window and back with no button held (let go out
// there), sends the window the mouseup that never came; and so does the
// window losing focus mid-press. A move inside the window is left alone.
// Run alone with RUTBA_VERIFY_ONLY=lostmouseup.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';

export async function verifyLostMouseup({ open, check, until }, { dir }) {
  const file = path.join(dir, 'lost-mouseup.xlsx');
  try {
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['a', 1], ['b', 2]] }] }));
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the window', 8000);
    const result = await js(`(() => {
      let ups = 0;
      const count = () => { ups += 1; };
      window.addEventListener('mouseup', count);
      const out = {};
      window.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 50, clientY: 50, button: 0, buttons: 1 }));
      window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: 60, clientY: 60, buttons: 1 }));
      window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: 65, clientY: 65, buttons: 0 }));
      out.whileHeld = ups;
      document.documentElement.dispatchEvent(new MouseEvent('mouseout', { bubbles: true, relatedTarget: null }));
      window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: 70, clientY: 70, buttons: 0 }));
      out.letGoOutside = ups;
      window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: 80, clientY: 80, buttons: 0 }));
      out.once = ups;
      window.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 50, clientY: 50, button: 0, buttons: 1 }));
      window.dispatchEvent(new Event('blur'));
      out.leftTheWindow = ups;
      window.removeEventListener('mouseup', count);
      return out;
    })()`);
    check('windows: a drag whose mouseup is lost — let go outside the window, or the window left mid-press — is sent the mouseup, once',
      result.whileHeld === 0 && result.letGoOutside === 1 && result.once === 1 && result.leftTheWindow === 2, JSON.stringify(result));
  } catch (err) {
    check('windows: the lost mouseup checks ran', false, err?.message || String(err));
  }
}
