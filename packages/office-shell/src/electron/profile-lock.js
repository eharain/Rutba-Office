// The window a profile with a password opens with (D-2026-10-26-02): a small
// window of its own, before the launcher and before anything of the profile
// is read, asking for the password. The right one makes the key that opens
// the profile's sealed files (profiles.js); a wrong one is said to be wrong
// and asked again; Cancel, or closing the window, quits this Rutba Office and
// leaves the profile as it was. The password itself is tried and forgotten
// here: never kept, never written.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { BrowserWindow, ipcMain, nativeTheme, screen } from 'electron';
import { unlockProfile } from '../profiles.js';
import { speakAs, t } from './words.js';

const here = path.dirname(fileURLToPath(import.meta.url));

const escape = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

/** The profile's settings, read before anything else of it, for the window's language and theme. */
function settingsOf(folder) {
  try {
    return JSON.parse(fs.readFileSync(path.join(folder, 'settings.json'), 'utf8')) || {};
  } catch {
    return {};
  }
}

function page({ profile, dark, rtl }) {
  const initial = escape(Array.from(String(profile.name || '?'))[0]?.toUpperCase() || '?');
  return `<!doctype html><html lang="en" dir="${rtl ? 'rtl' : 'ltr'}" class="${dark ? 'dark' : ''}"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'">
<title>${escape(profile.name)}</title>
<style>
:root { --bg: #f5f6f8; --surface: #ffffff; --ink: #16181c; --ink-2: #4a515c; --line: #d2d6dd; --accent: #2b5fd9; --bad: #c8372d; }
.dark { --bg: #17191d; --surface: #202329; --ink: #eceef1; --ink-2: #b6bcc6; --line: #3d424b; --accent: #6d95ff; --bad: #ff7a70; }
* { box-sizing: border-box; }
html, body { margin: 0; height: 100%; background: var(--bg); color: var(--ink); font: 13px "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif; }
body { display: flex; align-items: center; justify-content: center; padding: 20px; }
form { width: 100%; max-width: 320px; display: grid; gap: 10px; }
.who { display: flex; align-items: center; gap: 12px; margin-bottom: 4px; }
.mark { width: 40px; height: 40px; border-radius: 50%; display: grid; place-items: center; color: #fff; font: 600 18px "Segoe UI Variable Display", "Segoe UI", sans-serif; background: ${escape(profile.colour)}; flex: none; }
h1 { font: 600 16px "Segoe UI Variable Display", "Segoe UI", sans-serif; margin: 0; }
.who p { margin: 2px 0 0; color: var(--ink-2); }
label { font-weight: 600; }
input { height: 32px; padding: 0 10px; border: 1px solid var(--line); border-radius: 7px; background: var(--surface); color: var(--ink); font: inherit; }
input:focus { outline: 2px solid var(--accent); outline-offset: -1px; border-color: transparent; }
.error { color: var(--bad); min-height: 18px; margin: 0; }
.buttons { display: flex; gap: 8px; justify-content: flex-end; }
button { height: 30px; padding: 0 16px; border-radius: 7px; border: 1px solid var(--line); background: var(--surface); color: var(--ink); font: inherit; cursor: pointer; }
button.primary { background: var(--accent); border-color: var(--accent); color: #fff; }
button:disabled { opacity: 0.6; cursor: default; }
.note { color: var(--ink-2); font-size: 12px; margin: 4px 0 0; }
</style></head><body>
<form id="f" autocomplete="off">
  <div class="who"><span class="mark" aria-hidden="true">${initial}</span><div><h1>${escape(profile.name)}</h1><p>${escape(t('This profile has a password.'))}</p></div></div>
  <label for="pw">${escape(t('Password'))}</label>
  <input id="pw" type="password" autofocus spellcheck="false" aria-describedby="err">
  <p class="error" id="err" role="alert"></p>
  <div class="buttons"><button type="button" id="cancel">${escape(t('Cancel'))}</button><button type="submit" class="primary" id="go">${escape(t('Open'))}</button></div>
  <p class="note">${escape(t('A forgotten password cannot be recovered.'))}</p>
</form>
<script>
const f = document.getElementById('f'), pw = document.getElementById('pw'), err = document.getElementById('err'), go = document.getElementById('go');
f.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (go.disabled) return;
  go.disabled = true; err.textContent = '';
  const r = await window.rutbaLock.open(pw.value);
  if (!r || !r.ok) { go.disabled = false; err.textContent = (r && r.error) || ''; pw.select(); pw.focus(); }
});
document.getElementById('cancel').addEventListener('click', () => window.rutbaLock.cancel());
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') window.rutbaLock.cancel(); });
</script></body></html>`;
}

/**
 * Asks for a protected profile's password. Resolves with { key, rekey } once the
 * right one is typed, or null when the person cancels or closes the window.
 * `onWindow` hears of the window (a second launch brings it to the front).
 */
export function askProfilePassword({ profile, iconPath = null, onWindow = null }) {
  const settings = settingsOf(profile.folder);
  speakAs({ settings: { get: (k, f) => (k in settings ? settings[k] : f) } });
  const dark = settings.theme === 'dark' || (settings.theme !== 'light' && nativeTheme.shouldUseDarkColors);
  const rtl = ['ur', 'ar'].includes(String(settings.language || '').slice(0, 2));
  const offscreen = process.env.RUTBA_WINDOW_DISPLAY === 'offscreen';
  const at = offscreen ? { x: Math.max(...screen.getAllDisplays().map((d) => d.bounds.x + d.bounds.width)) + 100, y: 100 } : {};

  return new Promise((resolve) => {
    let done = false;
    const win = new BrowserWindow({
      width: 400,
      height: 300,
      ...at,
      resizable: false,
      maximizable: false,
      fullscreenable: false,
      show: false,
      skipTaskbar: offscreen,
      autoHideMenuBar: true,
      icon: iconPath || undefined,
      title: profile.name,
      backgroundColor: dark ? '#17191d' : '#f5f6f8',
      webPreferences: { preload: path.join(here, 'profile-lock-preload.cjs'), contextIsolation: true, sandbox: true, nodeIntegration: false, spellcheck: false },
    });
    win.setMenu(null);
    onWindow?.(win);
    const finish = (opened) => {
      if (done) return;
      done = true;
      ipcMain.removeHandler('rutba-office/profile-lock:open');
      ipcMain.removeHandler('rutba-office/profile-lock:cancel');
      resolve(opened);
      if (!win.isDestroyed()) win.destroy();
    };
    // Tried one at a time: a second press while scrypt works is not a second guess.
    let trying = false;
    ipcMain.handle('rutba-office/profile-lock:open', async (event, password) => {
      if (event.sender !== win.webContents || trying) return { ok: false, error: '' };
      if (!password) return { ok: false, error: t('Type the profile’s password.') };
      trying = true;
      try {
        // A change of password a crash stopped midway opens with the old password or the new (profiles.js).
        const opened = await unlockProfile(profile, password);
        if (!opened) return { ok: false, error: t('That password is not right. Try again.') };
        finish(opened);
        return { ok: true };
      } finally {
        trying = false;
      }
    });
    ipcMain.handle('rutba-office/profile-lock:cancel', (event) => {
      if (event.sender === win.webContents) finish(null);
    });
    win.on('closed', () => finish(null));
    win.once('ready-to-show', () => (offscreen ? win.showInactive() : win.show()));
    win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(page({ profile, dark, rtl }))}`).catch(() => {});
  });
}
