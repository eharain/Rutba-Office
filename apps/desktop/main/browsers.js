// The web browsers on this computer, for a sign-in that has to happen in one.
//
// Google and Microsoft sign a mail program in through their own page, in a
// real browser (oauth.js says why). Which browser is the person's choice: the
// one that holds their password manager, their security key, the account
// they are already signed in to. So the browsers installed are listed - on
// Windows as Windows lists them for "Default apps" (the registry's
// StartMenuInternet keys, the machine's and the user's), on macOS from the
// Applications folders, on Linux from the commands on the path - with the
// default one marked, and the sign-in page opened in the one chosen. Nothing
// here goes to the network; reading the list starts no browser.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFile, spawn } from 'node:child_process';

/**
 * Browsers from the text `reg query <...>\StartMenuInternet /s` prints.
 * @returns {Array<{ id, name, exe, progId }>}
 */
export function parseRegistryBrowsers(text) {
  const blocks = String(text || '').split(/\r?\n(?=HKEY_)/);
  const keys = new Map(); // browser key -> { name, exe, progId }
  for (const block of blocks) {
    const [head, ...lines] = block.split(/\r?\n/);
    const m = /\\StartMenuInternet\\([^\\]+)(\\.*)?$/i.exec(head.trim());
    if (!m) continue;
    const id = m[1];
    const rest = (m[2] || '').toLowerCase();
    const entry = keys.get(id) || { id, name: null, exe: null, progId: null };
    for (const line of lines) {
      const v = /^\s+(\S.*?)\s{2,}REG_(?:EXPAND_)?SZ\s{2,}(.*)$/.exec(line);
      if (!v) continue;
      const [, valueName, data] = v;
      if (!rest && valueName === '(Default)') entry.name = data.trim();
      if (rest === '\\shell\\open\\command' && valueName === '(Default)') {
        const quoted = /^"([^"]+)"/.exec(data.trim());
        entry.exe = (quoted ? quoted[1] : data.trim().split(/\s+-/)[0]).trim();
      }
      if (rest === '\\capabilities\\urlassociations' && /^https$/i.test(valueName)) entry.progId = data.trim();
    }
    keys.set(id, entry);
  }
  // Internet Explorer is still listed by Windows, but the providers' sign-in pages no longer open in it.
  return [...keys.values()].filter((b) => b.exe && /\.exe$/i.test(b.exe) && !/iexplore\.exe$/i.test(b.exe)).map((b) => ({ ...b, name: b.name || b.id }));
}

const run = (file, args) =>
  new Promise((resolve) => {
    execFile(file, args, { windowsHide: true, timeout: 5000, maxBuffer: 4 * 1024 * 1024 }, (err, stdout) => resolve(err ? '' : String(stdout)));
  });

const MAC = [
  ['safari', 'Safari'],
  ['chrome', 'Google Chrome'],
  ['firefox', 'Firefox'],
  ['edge', 'Microsoft Edge'],
  ['brave', 'Brave Browser'],
  ['arc', 'Arc'],
  ['opera', 'Opera'],
  ['vivaldi', 'Vivaldi'],
];
const LINUX = [
  ['firefox', 'Firefox', 'firefox'],
  ['chrome', 'Google Chrome', 'google-chrome'],
  ['chromium', 'Chromium', 'chromium'],
  ['edge', 'Microsoft Edge', 'microsoft-edge'],
  ['brave', 'Brave', 'brave-browser'],
  ['vivaldi', 'Vivaldi', 'vivaldi'],
  ['opera', 'Opera', 'opera'],
];

/**
 * The browsers installed, the default first when it is known.
 * @param {{ platform?: string, reg?: (args) => Promise<string>, exists?: (file) => boolean }} [opts] for the tests
 * @returns {Promise<Array<{ id, name, isDefault }>>}
 */
export async function listBrowsers({ platform = process.platform, reg = (args) => run('reg', args), exists = fs.existsSync } = {}) {
  let found = [];
  if (platform === 'win32') {
    const roots = ['HKCU\\SOFTWARE\\Clients\\StartMenuInternet', 'HKLM\\SOFTWARE\\Clients\\StartMenuInternet', 'HKLM\\SOFTWARE\\WOW6432Node\\Clients\\StartMenuInternet'];
    const seen = new Set();
    for (const root of roots) {
      for (const b of parseRegistryBrowsers(await reg(['query', root, '/s']))) {
        if (seen.has(b.exe.toLowerCase())) continue;
        seen.add(b.exe.toLowerCase());
        if (exists(b.exe)) found.push(b);
      }
    }
    const choice = await reg(['query', 'HKCU\\Software\\Microsoft\\Windows\\Shell\\Associations\\UrlAssociations\\https\\UserChoice', '/v', 'ProgId']);
    const progId = /ProgId\s+REG_SZ\s+(\S+)/i.exec(choice)?.[1] || null;
    found = found.map((b) => ({ ...b, isDefault: Boolean(progId && b.progId && b.progId.toLowerCase() === progId.toLowerCase()) }));
  } else if (platform === 'darwin') {
    for (const [id, name] of MAC) {
      for (const dir of ['/Applications', path.join(os.homedir(), 'Applications')]) {
        if (fs.existsSync(path.join(dir, `${name}.app`))) {
          found.push({ id, name, app: name, isDefault: false });
          break;
        }
      }
    }
  } else {
    const dirs = String(process.env.PATH || '').split(':');
    for (const [id, name, cmd] of LINUX) {
      if (dirs.some((d) => d && fs.existsSync(path.join(d, cmd)))) found.push({ id, name, cmd, isDefault: false });
    }
  }
  return found.sort((a, b) => Number(b.isDefault) - Number(a.isDefault) || a.name.localeCompare(b.name));
}

/**
 * Open `url` in the browser chosen (by the id listBrowsers gave), or in the
 * system's default when none is chosen or it is no longer there.
 * @param {string} url
 * @param {string|null} browserId
 * @param {{ openDefault: (url) => Promise<any>, list?: Array }} deps
 */
export async function openIn(url, browserId, { openDefault, list = null } = {}) {
  if (!/^https:\/\//i.test(url)) throw new Error('Only a secure web address is opened in a browser.');
  const browsers = browserId ? list || (await listBrowsers()) : [];
  const b = browserId ? browsers.find((x) => x.id === browserId) : null;
  if (!b) return openDefault(url);
  const child = b.exe ? spawn(b.exe, [url], { detached: true, stdio: 'ignore', windowsHide: false }) : b.app ? spawn('open', ['-a', b.app, url], { detached: true, stdio: 'ignore' }) : spawn(b.cmd, [url], { detached: true, stdio: 'ignore' });
  child.on('error', () => openDefault(url));
  child.unref();
  return { opened: b.name };
}
