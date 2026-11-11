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

// reg.exe answers in a fraction of a second, but on a computer busy with
// other work starting it took over five seconds, and a limit of five left the
// list empty: the limit is for a reg that never answers, not a slow one.
// A key that is not there (most users have no browsers of their own under
// HKCU) is an answer, the empty text; a reg that timed out or could not be
// started is no answer, null, so the list it leaves short is asked again.
/**
 * A reg.exe to ask: `(args) => Promise<string|null>`.
 * @param {{ file?: string, before?: string[], timeout?: number }} [opts] another program in its place, for the tests
 */
export const regRunner = ({ file = 'reg', before = [], timeout = 30000 } = {}) => (args) =>
  new Promise((resolve) => {
    execFile(file, [...before, ...args], { windowsHide: true, timeout, maxBuffer: 4 * 1024 * 1024 }, (err, stdout) => {
      if (!err) resolve(String(stdout));
      else resolve(typeof err.code === 'number' && !err.killed ? '' : null);
    });
  });
const reg = regRunner();

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
 * @param {{ platform?: string, reg?: (args) => Promise<string|null>, exists?: (file) => boolean }} [opts] for the tests
 * @returns {Promise<Array<{ id, name, isDefault }>>}
 */
export async function listBrowsers(opts = {}) {
  return (await readBrowsers(opts)).list;
}

/**
 * The browsers installed, and whether every read answered: a reg.exe that
 * timed out (null from `reg`) leaves the list short, or empty, and
 * `complete` false.
 * @returns {Promise<{ list: Array<{ id, name, isDefault }>, complete: boolean }>}
 */
export async function readBrowsers({ platform = process.platform, reg: ask = reg, exists = fs.existsSync } = {}) {
  let found = [];
  let complete = true;
  if (platform === 'win32') {
    const roots = ['HKCU\\SOFTWARE\\Clients\\StartMenuInternet', 'HKLM\\SOFTWARE\\Clients\\StartMenuInternet', 'HKLM\\SOFTWARE\\WOW6432Node\\Clients\\StartMenuInternet'];
    const seen = new Set();
    // The four reads at once, each a reg.exe of its own, taken in the same order.
    const [choice, ...texts] = await Promise.all([
      ask(['query', 'HKCU\\Software\\Microsoft\\Windows\\Shell\\Associations\\UrlAssociations\\https\\UserChoice', '/v', 'ProgId']),
      ...roots.map((root) => ask(['query', root, '/s'])),
    ].map((p) => Promise.resolve(p).catch(() => null)));
    complete = choice !== null && texts.every((text) => text !== null);
    for (const text of texts) {
      for (const b of parseRegistryBrowsers(text)) {
        if (seen.has(b.exe.toLowerCase())) continue;
        seen.add(b.exe.toLowerCase());
        if (exists(b.exe)) found.push(b);
      }
    }
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
  return { list: found.sort((a, b) => Number(b.isDefault) - Number(a.isDefault) || a.name.localeCompare(b.name)), complete };
}

/**
 * The browsers, read once and kept: the sign-in dialog asks each time it
 * opens, and a computer's browsers seldom change. A list is kept only when
 * it is whole and not empty. A read that timed out, or found nothing, is
 * asked again the next time, so a busy moment never leaves the person with
 * no choice of browser for good. Asks made while a read is under way share it.
 * @param {{ read?: () => Promise<{ list, complete }> }} [opts] for the tests
 */
export function browserList({ read = () => readBrowsers() } = {}) {
  let kept = null;
  let reading = null;
  return {
    get() {
      if (kept) return Promise.resolve(kept);
      if (!reading) {
        reading = Promise.resolve()
          .then(read)
          .then(({ list, complete }) => {
            if (complete && list.length) kept = list;
            return list;
          })
          .catch(() => [])
          .finally(() => {
            reading = null;
          });
      }
      return reading;
    },
    /** Forget the list kept, so the next ask reads again. */
    forget() {
      kept = null;
    },
  };
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
