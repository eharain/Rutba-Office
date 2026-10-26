// Add or remove apps (docs/APP-PACKS.md 5.4), the main process's side.
//
// Which apps are installed is the packs folder (app-packs.js). Adding one
// brings its pack from the release of this very version, a file per pack
// beside the installer (Rutba-Office-<version>-<id>.pack, a zip of its
// archive, its folder and its manifest), or, where a folder of packs is
// named (RUTBA_PACKS_SOURCE: the checks, or an administrator's share), from
// there. The pack is unpacked and checked in a staging folder of its own,
// with the core's keys and version, exactly as a start checks it; only then
// is it put in the packs folder, and it loads at the next start. Removing one
// is written down and done at the next start, before the loader runs:
// Windows holds a loaded pack's files open. The person's documents, mail,
// settings and profile data are never touched.
//
// A copy installed for everyone (in Program Files) cannot write its packs
// folder as the person: there the changes wait in the staging folder, and
// Restart asks Windows for the administrator's permission, once, to put them
// in place while Rutba Office is closed, then starts it again.
//
// A copy with every app built in (one run from the source, the portable
// build) has no packs to add or remove, and says so.

import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { app } from 'electron';
import { PACKS } from '@rutba/office-shell/contract';
import { packFiles, packSize, packsIn, canWriteIn, pendingRemovals, markRemoval, unmarkRemoval, stagePack, installStaged } from '@rutba/office-shell/electron/pack-install';
import { readZip } from '@rutba/ooxml/zip';
import { SITE } from '@rutba/office-formats/registry';
import { registerTypes, applyFileTypes, fileTypePlan, recordPacks } from './file-types.js';

const require = createRequire(import.meta.url);
// Archives are handled as the files they are, not as Electron's folders.
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

const nameOf = (id) => PACKS.find((p) => p.id === id)?.name || id;
const ID = /^[a-z][a-z0-9-]{0,31}$/;

/** The release's file for one pack. */
export const packFileName = (version, id) => `Rutba-Office-${version}-${id}.pack`;

/**
 * A release's pack file unpacked into a folder: only that pack's three
 * things, each name checked, nothing outside the folder.
 */
export function unpackPackFile(bytes, id, into) {
  const f = packFiles(id);
  const { entries } = readZip(Buffer.from(bytes));
  if (!entries.length) throw new Error(`The ${nameOf(id)} pack file is empty.`);
  fs.rmSync(into, { recursive: true, force: true });
  fs.mkdirSync(into, { recursive: true });
  for (const e of entries) {
    const name = e.name.replace(/\\/g, '/');
    if (name.endsWith('/')) continue;
    const parts = name.split('/');
    const allowed = name === f.manifest || name === f.archive || (parts[0] === f.unpacked && parts.length > 1);
    if (!allowed || parts.some((p) => !p || p === '.' || p === '..') || /^[a-zA-Z]:/.test(name)) {
      throw new Error(`The ${nameOf(id)} pack file holds something a pack may not: ${name}. It was not installed.`);
    }
    const to = path.join(into, ...parts);
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.writeFileSync(to, e.data);
  }
  return into;
}

/**
 * The script a copy installed for everyone runs with the administrator's
 * permission while Rutba Office is closed: the packs written down to go
 * deleted, the staged ones copied in (each manifest last), the kinds of
 * file of those added registered for every user, then Rutba Office started
 * again. PowerShell, quoted with single quotes throughout.
 */
export function elevatedScript({ pid, dir, exe, removals = [], additions = [] }) {
  const q = (s) => `'${String(s).replace(/'/g, "''")}'`;
  const lines = ['$ErrorActionPreference = \'Continue\'', `Wait-Process -Id ${Number(pid)} -Timeout 60 -ErrorAction SilentlyContinue`];
  for (const id of removals) {
    const f = packFiles(id);
    lines.push(`Remove-Item -LiteralPath ${q(path.join(dir, f.manifest))} -Force -ErrorAction SilentlyContinue`);
    lines.push(`Remove-Item -LiteralPath ${q(path.join(dir, f.archive))} -Force -ErrorAction SilentlyContinue`);
    lines.push(`Remove-Item -LiteralPath ${q(path.join(dir, f.unpacked))} -Recurse -Force -ErrorAction SilentlyContinue`);
  }
  for (const { id, staging, types = [] } of additions) {
    const f = packFiles(id);
    lines.push(`Copy-Item -LiteralPath ${q(path.join(staging, f.archive))} -Destination ${q(dir)} -Force`);
    lines.push(`if (Test-Path -LiteralPath ${q(path.join(staging, f.unpacked))}) { Copy-Item -LiteralPath ${q(path.join(staging, f.unpacked))} -Destination ${q(dir)} -Recurse -Force }`);
    lines.push(`Copy-Item -LiteralPath ${q(path.join(staging, f.manifest))} -Destination ${q(dir)} -Force`);
    for (const t of types) {
      const cls = `HKLM:\\Software\\Classes\\${t.name}`;
      lines.push(`New-Item -Path ${q(cls + '\\shell\\open\\command')} -Force | Out-Null`);
      lines.push(`Set-Item -Path ${q(cls)} -Value ${q(t.description || t.name)}`);
      lines.push(`New-Item -Path ${q(cls + '\\DefaultIcon')} -Force -Value ${q(t.icon)} | Out-Null`);
      lines.push(`Set-Item -Path ${q(cls + '\\shell\\open\\command')} -Value ${q(`"${exe}" "%1"`)}`);
      lines.push(`New-Item -Path ${q(`HKLM:\\Software\\Classes\\.${t.ext}\\OpenWithProgids`)} -Force | Out-Null`);
      lines.push(`New-ItemProperty -Path ${q(`HKLM:\\Software\\Classes\\.${t.ext}\\OpenWithProgids`)} -Name ${q(t.name)} -PropertyType None -Value ([byte[]]@()) -Force | Out-Null`);
    }
  }
  lines.push(`Start-Process -FilePath ${q(exe)}`);
  return `${lines.join('\r\n')}\r\n`;
}

/**
 * @param {object} o
 * @param {object} o.plan      this run's packs (app-packs.js planPacks)
 * @param {object} o.stores
 * @param {Function} o.broadcast
 * @param {string} o.version   the core's version: a pack must carry the same
 * @param {*[]} o.keys         the keys a pack must be signed with
 * @param {string|null} o.dir  the packs folder, or null when every app is built in
 * @param {string} [o.source]  a folder of packs to add from, rather than the release
 * @param {Function} [o.fetchImpl]
 * @param {object} [o.sizes]   each pack's size as the build measured it, for one not here yet
 */
export function createPacksService({ plan, stores, broadcast = () => {}, version, keys, dir, source = process.env.RUTBA_PACKS_SOURCE || null, fetchImpl = (...a) => fetch(...a), sizes = {}, staging = path.join(os.tmpdir(), `rutba-office-packs-${process.pid}`) }) {
  const managed = Boolean(dir);
  const registry = Boolean(process.platform === 'win32' && app?.isPackaged && !process.env.RUTBA_VERIFY && !process.env.RUTBA_OFFICE_VERIFY_APPS);
  /** Added in this run, for the next start: in the packs folder already, or staged for the administrator's step. */
  const adding = new Map();
  /** Written down to go, for a copy that cannot write its packs folder. */
  const goingElevated = new Set();
  const perMachine = () => managed && !canWriteIn(dir);
  const removals = () => (managed ? [...new Set([...pendingRemovals(dir), ...goingElevated])] : []);

  const list = () => {
    const going = removals();
    const refused = new Map(plan.refused.map((r) => [r.id, r.reason]));
    const apps = PACKS.map(({ id, name }) => {
      let state = 'absent';
      if (plan.has(id)) state = going.includes(id) ? 'removing' : 'installed';
      else if (adding.has(id)) state = 'adding';
      else if (refused.has(id)) state = 'refused';
      const size = managed && (state !== 'absent' || adding.has(id)) ? packSize(adding.get(id)?.staged ? adding.get(id).staged : dir, id) : 0;
      return { id, name, state, size: size || sizes[id] || 0, reason: state === 'refused' ? refused.get(id) : null };
    });
    return {
      apps,
      managed,
      perMachine: perMachine(),
      restart: going.some((id) => plan.has(id)) || adding.size > 0,
      source: source ? 'folder' : 'release',
    };
  };
  const changed = () => {
    const out = list();
    broadcast('packs:changed', { apps: out.apps, restart: out.restart });
    return out;
  };
  const refuseUnmanaged = () => {
    if (!managed) throw new Error('Every app is built into this copy of Rutba Office, so there is nothing to add or remove. The installed copy lets you choose.');
  };

  /** The pack's three things in a folder: the named folder of packs, or the release's file fetched and unpacked. */
  async function fetchPack(id) {
    const into = path.join(staging, 'download', id);
    if (source) {
      const f = packFiles(id);
      if (fs.existsSync(path.join(source, f.manifest))) return source;
      for (const name of [packFileName(version, id), `${id}.pack`]) {
        const file = path.join(source, name);
        if (fs.existsSync(file)) return unpackPackFile(fs.readFileSync(file), id, into);
      }
      throw new Error(`The ${nameOf(id)} pack is not in ${source}.`);
    }
    const url = `${SITE.releases}/download/v${version}/${packFileName(version, id)}`;
    let res;
    try {
      res = await fetchImpl(url, { headers: { 'user-agent': 'Rutba-Office' } });
    } catch {
      throw new Error(`The ${nameOf(id)} pack could not be fetched: this computer could not reach the release. Try again when it is online.`);
    }
    if (!res.ok) throw new Error(`The ${nameOf(id)} pack could not be fetched from the release of Rutba Office ${version} (${res.status}).`);
    return unpackPackFile(Buffer.from(await res.arrayBuffer()), id, into);
  }

  /** The kinds of file a pack opens, as the registry is told them: each icon on disk in its folder. */
  const typesOf = (pack, at) => (pack.fileTypes || []).map((t) => ({ ...t, icon: t.icon ? path.join(at, `${pack.id}.asar.unpacked`, ...String(t.icon).split('/')) : null }));

  return {
    list,

    add: async ({ id } = {}) => {
      refuseUnmanaged();
      if (!ID.test(String(id)) || !PACKS.some((p) => p.id === id)) throw new Error(`There is no app called ${id} to add.`);
      // Written down to go, and wanted after all: it simply stays.
      if (removals().includes(id)) {
        if (goingElevated.delete(id)) return changed();
        unmarkRemoval(dir, id);
        return changed();
      }
      if (plan.has(id)) throw new Error(`${nameOf(id)} is already installed.`);
      if (adding.has(id)) return list();
      const from = await fetchPack(id);
      const at = path.join(staging, 'checked', id);
      const present = [...new Set([...plan.ids(), ...adding.keys()])];
      const pack = stagePack({ from, id, staging: at, publicKey: keys, version, present, name: nameOf(id) });
      if (perMachine()) {
        adding.set(id, { staged: at, types: typesOf(pack, dir) });
      } else {
        installStaged({ staging: at, id, dir });
        adding.set(id, { staged: null });
        if (registry) await registerTypes(typesOf(pack, dir), { exe: process.execPath }).catch(() => {});
      }
      fs.rmSync(path.join(staging, 'download', id), { recursive: true, force: true });
      return changed();
    },

    remove: async ({ id } = {}) => {
      refuseUnmanaged();
      if (adding.has(id)) {
        // Added in this run and never loaded: its files can go now.
        const a = adding.get(id);
        adding.delete(id);
        if (a.staged) fs.rmSync(a.staged, { recursive: true, force: true });
        else {
          const f = packFiles(id);
          fs.rmSync(path.join(dir, f.manifest), { force: true });
          fs.rmSync(path.join(dir, f.archive), { force: true });
          fs.rmSync(path.join(dir, f.unpacked), { recursive: true, force: true });
        }
        return changed();
      }
      if (!plan.has(id) && !packsIn(dir).includes(id)) throw new Error(`${nameOf(id)} is not installed.`);
      const left = plan.ids().filter((x) => x !== id && !removals().includes(x));
      if (!left.some((x) => PACKS.find((p) => p.id === x) && !['ai', 'ocr'].includes(x))) throw new Error('Keep one app at least. To remove Rutba Office altogether, uninstall it from Windows Settings.');
      if (perMachine()) goingElevated.add(id);
      else markRemoval(dir, id);
      // Its kinds of file go back to whatever opened them before, unless an app that stays opens them too.
      if (registry) {
        const gone = plan.packs.find((p) => p.id === id)?.fileTypes || [];
        const kept = new Set(plan.packs.filter((p) => left.includes(p.id)).flatMap((p) => (p.fileTypes || []).map((t) => t.ext)));
        await applyFileTypes(fileTypePlan(gone.filter((t) => !kept.has(t.ext)), [])).catch(() => {});
      }
      return changed();
    },

    keep: async ({ id } = {}) => {
      refuseUnmanaged();
      if (goingElevated.delete(id)) return changed();
      if (pendingRemovals(dir).includes(id)) {
        unmarkRemoval(dir, id);
        return changed();
      }
      return list();
    },

    restart: async () => {
      const going = removals();
      // What the next installer opens on: the packs that will be there.
      if (registry) await recordPacks([...new Set([...plan.ids().filter((x) => !going.includes(x)), ...adding.keys()])]).catch(() => {});
      const elevated = perMachine() && (goingElevated.size || [...adding.values()].some((a) => a.staged));
      if (elevated) {
        const script = path.join(staging, `apply-${crypto.randomUUID()}.ps1`);
        fs.mkdirSync(staging, { recursive: true });
        fs.writeFileSync(script, elevatedScript({
          pid: process.pid,
          dir,
          exe: process.execPath,
          removals: [...goingElevated],
          additions: [...adding.entries()].filter(([, a]) => a.staged).map(([id, a]) => ({ id, staging: a.staged, types: a.types })),
        }));
        const asked = await new Promise((resolve) => {
          const args = `'-NoProfile','-ExecutionPolicy','Bypass','-WindowStyle','Hidden','-File','"${script.replace(/'/g, "''")}"'`;
          const child = spawn('powershell.exe', ['-NoProfile', '-Command', `Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -ArgumentList ${args}`], { windowsHide: true, stdio: 'ignore' });
          child.on('error', () => resolve(false));
          child.on('exit', (code) => resolve(code === 0));
        });
        if (!asked) throw new Error('Windows did not give permission to change the apps of a copy installed for everyone. Nothing was changed; try again, or ask whoever looks after this computer.');
        app.quit();
        return { restarting: true, elevated: true };
      }
      app.relaunch();
      app.quit();
      return { restarting: true, elevated: false };
    },
  };
}
