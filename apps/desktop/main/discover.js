// What other mail clients have left on this computer.
//
// "Import your mail" is a much better offer when the application already knows
// where your mail is. Every one of these clients keeps its data somewhere
// predictable, so this looks, reports what it found, and lets the person choose
// — it never reads a message until they say so.
//
//   Thunderbird   prefs.js is plain text and names every server and identity;
//                 the mail itself is mbox, which we already read.
//   Outlook       .pst and .ost in the two places Outlook puts them, plus the
//                 account host names from the profile registry.
//   Apple Mail    ~/Library/Mail, a tree of .emlx files.
//   Windows Mail  the old Windows Live Mail store, .eml on disk.
//
// Nothing here writes, and nothing leaves the machine.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const run = promisify(execFile);
const home = os.homedir();
const isWindows = process.platform === 'win32';
const isMac = process.platform === 'darwin';

const exists = (p) => {
  try {
    return Boolean(p) && fs.existsSync(p);
  } catch {
    return false;
  }
};

const sizeOf = (p) => {
  try {
    return fs.statSync(p).size;
  } catch {
    return 0;
  }
};

/** Directories, one level, that exist. */
function dirsIn(root) {
  try {
    return fs
      .readdirSync(root, { withFileTypes: true })
      .filter((e) => e.isDirectory())
      .map((e) => path.join(root, e.name));
  } catch {
    return [];
  }
}

function filesIn(root, match, depth = 2, out = []) {
  if (depth < 0) return out;
  let entries = [];
  try {
    entries = fs.readdirSync(root, { withFileTypes: true });
  } catch {
    return out;
  }
  for (const entry of entries) {
    const full = path.join(root, entry.name);
    if (entry.isDirectory()) filesIn(full, match, depth - 1, out);
    else if (match(entry.name)) out.push(full);
  }
  return out;
}

/* ── Thunderbird ────────────────────────────────────────────────────────── */

const THUNDERBIRD_ROOTS = [
  isWindows && path.join(process.env.APPDATA || '', 'Thunderbird'),
  isMac && path.join(home, 'Library', 'Thunderbird'),
  path.join(home, '.thunderbird'),
].filter(Boolean);

/**
 * prefs.js is a list of `user_pref("key", value);` lines. Parsing it with a
 * regular expression is exactly right here: it is generated, one pref per line,
 * and has been for twenty years.
 */
function readPrefs(file) {
  const prefs = {};
  let text = '';
  try {
    text = fs.readFileSync(file, 'utf8');
  } catch {
    return prefs;
  }
  for (const m of text.matchAll(/user_pref\("([^"]+)",\s*(.*?)\);\s*$/gm)) {
    const raw = m[2].trim();
    let value = raw;
    if (raw.startsWith('"')) value = raw.slice(1, -1).replace(/\\"/g, '"');
    else if (raw === 'true' || raw === 'false') value = raw === 'true';
    else if (/^-?\d+$/.test(raw)) value = Number(raw);
    prefs[m[1]] = value;
  }
  return prefs;
}

function thunderbirdAccounts() {
  const found = [];
  for (const root of THUNDERBIRD_ROOTS) {
    if (!exists(root)) continue;
    const profilesRoot = exists(path.join(root, 'Profiles')) ? path.join(root, 'Profiles') : root;
    for (const profile of dirsIn(profilesRoot)) {
      const prefsFile = path.join(profile, 'prefs.js');
      if (!exists(prefsFile)) continue;
      const prefs = readPrefs(prefsFile);

      // Every configured incoming server is `mail.server.serverN.*`.
      const servers = new Set();
      for (const key of Object.keys(prefs)) {
        const m = /^mail\.server\.(server\d+)\./.exec(key);
        if (m) servers.add(m[1]);
      }

      const identities = Object.keys(prefs)
        .map((k) => /^mail\.identity\.(id\d+)\.useremail$/.exec(k))
        .filter(Boolean)
        .map((m) => ({ id: m[1], email: prefs[`mail.identity.${m[1]}.useremail`], name: prefs[`mail.identity.${m[1]}.fullName`] }));

      const smtp = Object.keys(prefs)
        .map((k) => /^mail\.smtpserver\.(smtp\d+)\.hostname$/.exec(k))
        .filter(Boolean)
        .map((m) => ({
          host: prefs[`mail.smtpserver.${m[1]}.hostname`],
          port: prefs[`mail.smtpserver.${m[1]}.port`] || 465,
          user: prefs[`mail.smtpserver.${m[1]}.username`],
        }))[0] || null;

      for (const server of servers) {
        const type = prefs[`mail.server.${server}.type`];
        if (type === 'none' || type === 'rss') continue; // Local Folders, feeds
        const host = prefs[`mail.server.${server}.hostname`];
        const user = prefs[`mail.server.${server}.userName`];
        const directory = prefs[`mail.server.${server}.directory`];
        const identity = identities.find((i) => String(i.email || '').split('@')[1] === String(host || '').split('.').slice(-2).join('.')) || identities[0];

        // Thunderbird stores mail as mbox, which this suite already reads.
        const stores = directory && exists(directory) ? filesIn(directory, (n) => !n.includes('.') && n !== 'msgFilterRules.dat', 2) : [];

        found.push({
          source: 'Thunderbird',
          profile: path.basename(profile),
          email: identity?.email || user || host,
          name: identity?.name || null,
          incoming: host ? { protocol: type || 'imap', host, port: prefs[`mail.server.${server}.port`] || (type === 'pop3' ? 995 : 993), user, secure: true } : null,
          outgoing: smtp ? { host: smtp.host, port: smtp.port, user: smtp.user, secure: smtp.port === 465 } : null,
          stores: stores.map((p) => ({ path: p, bytes: sizeOf(p), format: 'mbox' })).filter((s) => s.bytes > 0),
        });
      }
    }
  }
  return found;
}

/* ── Outlook ────────────────────────────────────────────────────────────── */

const OUTLOOK_DIRS = [
  isWindows && path.join(process.env.LOCALAPPDATA || '', 'Microsoft', 'Outlook'),
  path.join(home, 'Documents', 'Outlook Files'),
  isMac && path.join(home, 'Library', 'Group Containers'),
].filter(Boolean);

/**
 * Outlook's account host names live in the profile registry as UTF-16 blobs.
 * Reading them is best-effort: a missing or unfamiliar profile simply yields
 * the data files, which are the part that matters anyway.
 */
async function outlookServers() {
  if (!isWindows) return [];
  const out = [];
  for (const version of ['16.0', '15.0', '14.0']) {
    const root = `HKCU\\Software\\Microsoft\\Office\\${version}\\Outlook\\Profiles`;
    let profiles = [];
    try {
      const { stdout } = await run('reg', ['query', root], { windowsHide: true });
      profiles = stdout.split(/\r?\n/).filter((l) => l.startsWith(root + '\\'));
    } catch {
      continue;
    }
    for (const profile of profiles.slice(0, 4)) {
      try {
        const { stdout } = await run('reg', ['query', profile, '/s', '/f', 'Server', '/t', 'REG_BINARY'], {
          windowsHide: true,
          maxBuffer: 4 * 1024 * 1024,
        });
        // The blobs are UTF-16 host names; pull anything that looks like one.
        for (const m of stdout.matchAll(/REG_BINARY\s+((?:[0-9A-F]{2})+)/g)) {
          const bytes = Buffer.from(m[1], 'hex');
          const text = bytes.toString('utf16le').replace(/\0/g, '');
          const host = /([a-z0-9-]+\.)+[a-z]{2,}/i.exec(text)?.[0];
          if (host && !out.some((o) => o.host === host)) out.push({ host, profile: path.basename(profile) });
        }
      } catch {
        // Nothing readable in this profile.
      }
    }
  }
  return out;
}

function outlookDataFiles() {
  const files = [];
  for (const dir of OUTLOOK_DIRS) {
    if (!exists(dir)) continue;
    for (const file of filesIn(dir, (n) => /\.(pst|ost|olm)$/i.test(n), 2)) {
      files.push({ path: file, bytes: sizeOf(file), format: path.extname(file).slice(1).toLowerCase() });
    }
  }
  return files.filter((f) => f.bytes > 0);
}

/* ── Apple Mail and Windows Live Mail ───────────────────────────────────── */

function appleMail() {
  if (!isMac) return [];
  const root = path.join(home, 'Library', 'Mail');
  if (!exists(root)) return [];
  const stores = dirsIn(root).filter((d) => /^V\d+$/.test(path.basename(d)));
  return stores.length
    ? [{ source: 'Apple Mail', email: 'Apple Mail', incoming: null, outgoing: null, stores: stores.map((p) => ({ path: p, bytes: 0, format: 'emlx' })) }]
    : [];
}

function windowsLiveMail() {
  if (!isWindows) return [];
  const root = path.join(process.env.LOCALAPPDATA || '', 'Microsoft', 'Windows Live Mail');
  if (!exists(root)) return [];
  return [{ source: 'Windows Live Mail', email: 'Windows Live Mail', incoming: null, outgoing: null, stores: [{ path: root, bytes: 0, format: 'eml' }] }];
}

/* ── the offer ──────────────────────────────────────────────────────────── */

export function createDiscoveryService() {
  return {
    /**
     * Everything found, without reading a single message.
     * @returns {{ accounts: [], files: [], scanned: string[] }}
     */
    scan: async () => {
      const accounts = [...thunderbirdAccounts(), ...appleMail(), ...windowsLiveMail()];
      const files = outlookDataFiles();
      const servers = await outlookServers();

      if (files.length || servers.length) {
        accounts.push({
          source: 'Outlook',
          email: servers[0]?.host ? `Outlook (${servers[0].host})` : 'Outlook',
          name: null,
          incoming: servers[0] ? { protocol: 'imap', host: servers[0].host, port: 993, secure: true } : null,
          outgoing: null,
          stores: files,
        });
      }

      return {
        accounts: accounts.map((a) => ({
          ...a,
          messagesHint: a.stores.reduce((n, s) => n + s.bytes, 0),
          storeCount: a.stores.length,
        })),
        files,
        scanned: [...THUNDERBIRD_ROOTS, ...OUTLOOK_DIRS].filter(exists),
      };
    },
  };
}
