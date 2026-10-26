// Makes and signs app packs from a built tree (docs/APP-PACKS.md, section 3).
//
//   node tools/make-packs.mjs --tree <folder> --out <folder>
//                             [--key <file>] [--version <v>] [--only mail,calendar]
//
// The tree holds one folder per pack, named by its id: <tree>/<id>/pack.json
// says what the pack is (its name, the packs it needs, its main module, its
// IPC, tiles, file types and AI tools, which of its folders stay unpacked)
// and every other file in the folder goes into the pack. For each, this
// writes <out>/<id>.asar (and <id>.asar.unpacked/ when some files stay on
// disk) and <out>/<id>.pack.json: the manifest with every file's size and
// SHA-512, signed with Ed25519.
//
// The private key comes from --key <file>, or RUTBA_PACK_KEY (the PEM text),
// or RUTBA_PACK_KEY_FILE, never from the repository: a key file inside it is
// refused. The version is the core's (the repository's package.json) unless
// --version says otherwise; the core loads only packs of its own version.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { canonical, keyIdOf, readArchive, PACK_FORMAT, PACK_FORMAT_VERSION } from '../packages/office-shell/src/electron/packs.js';
import { insideRepository } from './make-pack-key.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const ID = /^[a-z][a-z0-9-]{0,31}$/;
/** What a pack.json may say, besides the files, which are read from the folder. */
const DESCRIBED = ['name', 'needs', 'uses', 'main', 'windows', 'ipc', 'events', 'tiles', 'fileTypes', 'tools', 'registry'];

/** The private key, from a file given on the command line or the environment; never from the repository. */
export function readPrivateKey({ file = null, env = process.env } = {}) {
  const from = file || env.RUTBA_PACK_KEY_FILE || null;
  let pem;
  if (from) {
    if (insideRepository(from)) throw new Error('The pack key is never read from inside the repository. Keep it with the release secrets and give its path.');
    pem = fs.readFileSync(from, 'utf8');
  } else if (env.RUTBA_PACK_KEY) {
    pem = env.RUTBA_PACK_KEY;
  } else {
    throw new Error('No pack key: give --key <file>, or set RUTBA_PACK_KEY or RUTBA_PACK_KEY_FILE.');
  }
  const key = crypto.createPrivateKey(pem);
  if (key.asymmetricKeyType !== 'ed25519') throw new Error(`The pack key must be an Ed25519 key, not ${key.asymmetricKeyType}.`);
  return key;
}

/** Every file in a folder, relative with forward slashes, in order; a link is refused. */
function listFiles(dir, prefix = '', out = []) {
  for (const ent of fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
    const p = prefix ? `${prefix}/${ent.name}` : ent.name;
    if (ent.isSymbolicLink()) throw new Error(`${p} is a link; a pack holds files only.`);
    if (ent.isDirectory()) listFiles(path.join(dir, ent.name), p, out);
    else out.push(p);
  }
  return out;
}

const sha512 = (file) => crypto.createHash('sha512').update(fs.readFileSync(file)).digest('base64');

/** Signs a manifest: the file the core reads beside the archive. */
export function signManifest(manifest, privateKey) {
  const publicKey = crypto.createPublicKey(privateKey);
  return {
    format: PACK_FORMAT,
    keyId: keyIdOf(publicKey),
    manifest,
    signature: crypto.sign(null, Buffer.from(canonical(manifest), 'utf8'), privateKey).toString('base64'),
  };
}

/**
 * Makes one pack from its folder: the archive, then the signed manifest.
 * @returns {Promise<{ id, archive, manifest, files: number, bytes: number }>}
 */
export async function makePack({ src, out, version, privateKey }) {
  const asar = require('@electron/asar');
  const id = path.basename(src);
  if (!ID.test(id)) throw new Error(`${id}: a pack's folder is named by its id (lower case letters, digits and hyphens).`);
  const described = JSON.parse(fs.readFileSync(path.join(src, 'pack.json'), 'utf8'));
  const unknown = Object.keys(described).filter((k) => !DESCRIBED.includes(k) && k !== 'unpack');
  if (unknown.length) throw new Error(`${id}/pack.json says what a pack cannot: ${unknown.join(', ')}.`);
  if (!described.main) throw new Error(`${id}/pack.json names no main module.`);
  const unpack = (described.unpack || []).map((p) => p.replace(/\/+$/, ''));

  const names = listFiles(src).filter((p) => p !== 'pack.json');
  if (!names.length) throw new Error(`${id} has no files.`);
  const unpacked = (p) => unpack.some((u) => p === u || p.startsWith(`${u}/`));
  const files = {};
  let bytes = 0;
  for (const p of names) {
    const abs = path.join(src, ...p.split('/'));
    const size = fs.statSync(abs).size;
    bytes += size;
    files[p] = { size, sha512: sha512(abs), ...(unpacked(p) ? { unpacked: true } : {}) };
  }
  for (const key of ['main', 'windows']) if (described[key] && !files[described[key]]) throw new Error(`${id}: its ${key} module ${described[key]} is not one of its files.`);

  // The archive, made by Electron's own tool; the folders that stay on disk named as it expects.
  fs.mkdirSync(out, { recursive: true });
  const archive = path.join(out, `${id}.asar`);
  fs.rmSync(archive, { force: true });
  fs.rmSync(`${archive}.unpacked`, { recursive: true, force: true });
  const options = unpack.length ? { unpackDir: `{${unpack.flatMap((u) => [u, `${u}/**`]).join(',')}}` } : {};
  await asar.createPackageFromFiles(src, archive, names.map((p) => path.join(src, ...p.split('/'))), undefined, options);
  // What went in is what the manifest says, packed and unpacked alike.
  const table = readArchive(archive);
  for (const p of names) {
    const entry = table.get(p);
    if (!entry || entry.unpacked !== Boolean(files[p].unpacked) || entry.size !== files[p].size) throw new Error(`${id}: ${p} did not go into the archive as listed.`);
  }
  if (table.size !== names.length) throw new Error(`${id}: the archive holds files that were not listed.`);

  const manifest = {
    format: PACK_FORMAT_VERSION,
    id,
    name: described.name || id,
    version,
    needs: described.needs || [],
    uses: described.uses || [],
    archive: `${id}.asar`,
    main: described.main,
    windows: described.windows,
    ipc: described.ipc || {},
    events: described.events || [],
    tiles: described.tiles || [],
    fileTypes: described.fileTypes || [],
    tools: described.tools || [],
    registry: described.registry || undefined,
    files,
  };
  const signed = signManifest(manifest, privateKey);
  const manifestFile = path.join(out, `${id}.pack.json`);
  fs.writeFileSync(manifestFile, `${JSON.stringify(signed, null, 2)}\n`);
  return { id, archive, manifest: manifestFile, files: names.length, bytes };
}

/** Makes every pack in a tree (or the ones named). */
export async function makePacks({ tree, out, version, privateKey, only = null }) {
  const ids = fs.readdirSync(tree, { withFileTypes: true })
    .filter((d) => d.isDirectory() && fs.existsSync(path.join(tree, d.name, 'pack.json')))
    .map((d) => d.name)
    .filter((id) => !only || only.includes(id))
    .sort();
  if (!ids.length) throw new Error(`No packs in ${tree}: each is a folder with a pack.json.`);
  const made = [];
  for (const id of ids) made.push(await makePack({ src: path.join(tree, id), out, version, privateKey }));
  return made;
}

function option(args, name) {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : null;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  const tree = option(args, '--tree');
  const out = option(args, '--out');
  if (!tree || !out) {
    console.error('Usage: node tools/make-packs.mjs --tree <folder> --out <folder> [--key <file>] [--version <v>] [--only a,b]');
    process.exit(2);
  }
  try {
    const privateKey = readPrivateKey({ file: option(args, '--key') });
    const version = option(args, '--version') || JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version;
    const only = option(args, '--only')?.split(',').map((s) => s.trim()).filter(Boolean) || null;
    const made = await makePacks({ tree, out, version, privateKey, only });
    for (const m of made) console.log(`${m.id}: ${m.files} files, ${(m.bytes / 1024).toFixed(0)} KB, signed with key ${keyIdOf(crypto.createPublicKey(privateKey))}`);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}
