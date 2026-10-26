// The app packs a release carries (docs/APP-PACKS.md 5.3): one per app,
// signed, made before electron-builder runs.
//
//   node build/make-app-packs.js [--key <file>] [--only mail,word] [--release <folder>]
//
// Every pack holds what its manifest says of the app (its channels, events,
// tiles, file types, AI tools and rows of the registry, describePack in
// main/app-packs.js) and the icons Windows reads from disk. An app whose
// code has moved (CARRY_CODE) carries it too: its main-process files, its
// window chunks and the packages only it uses (build/pack-contents.js), and
// the core's archive leaves them out (electron-builder.config.cjs reads
// build/packs/moved.json). The others' main module registers the core's own
// code, which is still in the core's archive.
//
// It writes:
//   build/packs/out/       each pack: <id>.asar, <id>.asar.unpacked/, <id>.pack.json
//   build/out/pack-keys.json  the public key the core checks packs with, and each pack's size
//   build/packs/installer-packs.nsh  what resources/installer.nsh writes, a pack at a time
//   build/packs/moved.json  the core's files each app that carries its code took with it
//   <release>/Rutba-Office-<version>-<id>.pack  with --release: the file Add or remove apps fetches
//
// The private key comes from --key, RUTBA_PACK_KEY_FILE or RUTBA_PACK_KEY,
// never from the repository (tools/make-packs.mjs). A development key is
// made with tools/make-pack-key.mjs, in a folder outside the repository.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { PACKS } from '@rutba/office-shell/contract';
import { writeZip, ZipEntry } from '@rutba/ooxml/zip';
import { makePacks, readPrivateKey } from '../../../tools/make-packs.mjs';
import { describePack } from '../main/app-packs.js';
import { packFileName } from '../main/packs-service.js';
import { packFiles, packSize } from '@rutba/office-shell/electron/pack-install';
import { TOOL_LIST } from '../main/mcp/tools.js';
import { REACH_TOOL_LIST } from '../main/mcp/reach-tools.js';
import { packContents } from './pack-contents.js';

const APP = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const PACKS_DIR = path.join(APP, 'build', 'packs');
export const KEYS_FILE = path.join(APP, 'build', 'out', 'pack-keys.json');
export const INSTALLER_FILE = path.join(PACKS_DIR, 'installer-packs.nsh');
export const MOVED_FILE = path.join(PACKS_DIR, 'moved.json');
const ROOT = path.resolve(APP, '..', '..');

/** The apps whose code is in their own pack, not in the core's archive. */
export const CARRY_CODE = ['talk', 'ocr', 'base', 'calendar', 'contacts', 'pdf', 'pictures', 'image', 'word', 'sheets', 'slides'];

/** The apps a person picks on the installer's page; Text recognition and AI Tools are packs too, but never the one app installed. */
export const ADD_ONS = ['ai', 'ocr'];

/** The main module of a pack whose code is still in the core. */
const shim = (name) => `// ${name}'s pack. Its code is still in the core's archive in this version,\n// so it registers the core's own (main/app-packs.js, builtIn).\nexport function register(core) {\n  return core.builtIn(core);\n}\n`;

/**
 * A folder copied, without the tests and notes a package keeps beside its
 * code. An npm package keeps the copies of its own dependencies it holds
 * (`nested`); a workspace package's node_modules are links to the others.
 */
function copyTree(from, to, { nested = false } = {}) {
  for (const ent of fs.readdirSync(from, { withFileTypes: true })) {
    if (ent.isDirectory() && ['test', 'tests', '.git', ...(nested ? [] : ['node_modules'])].includes(ent.name)) continue;
    if (!ent.isDirectory() && /\.(md|map)$/i.test(ent.name)) continue;
    const a = path.join(from, ent.name);
    const b = path.join(to, ent.name);
    if (ent.isDirectory()) copyTree(a, b, { nested });
    else {
      fs.mkdirSync(to, { recursive: true });
      fs.copyFileSync(a, b);
    }
  }
}

/**
 * The code of an app that carries it, into its pack's folder: its main-process
 * files under main/ as they lie in the core, its window chunks under
 * renderer/chunks/ (served at the address the core imports them by), a worker
 * its window starts by name under renderer/, and its packages under
 * node_modules/. Answers the pack's main module.
 */
function carry(folder, c, rendererOut) {
  for (const file of c.main) {
    const to = path.join(folder, 'main', ...file.slice('apps/desktop/main/'.length).split('/'));
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.copyFileSync(path.join(ROOT, file), to);
  }
  for (const chunk of [...c.chunks, ...(c.workers || [])]) {
    const to = path.join(folder, 'renderer', ...chunk.split('/'));
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.copyFileSync(path.join(rendererOut, ...chunk.split('/')), to);
  }
  for (const name of c.npm) copyTree(path.join(ROOT, 'node_modules', name), path.join(folder, 'node_modules', name), { nested: true });
  for (const p of c.packages) copyTree(path.join(ROOT, 'packages', p.dir), path.join(folder, 'node_modules', p.name));
  for (const s of c.split || []) {
    for (const f of s.files) {
      const to = path.join(folder, 'node_modules', ...s.name.split('/'), ...f.split('/'));
      fs.mkdirSync(path.dirname(to), { recursive: true });
      fs.copyFileSync(path.join(ROOT, 'packages', s.dir, ...f.split('/')), to);
    }
  }
  return `main/${c.register.slice('apps/desktop/main/'.length)}`;
}

/**
 * Folders of a carried package that stay on disk beside the pack's archive,
 * as electron-builder.config.cjs unpacks them beside the core's: text
 * recognition's engine and models, read by its worker threads.
 */
const UNPACKED = { '@rutba/ocr': ['vendor'] };

/** The folders of a pack's carried code to leave unpacked. */
function unpackedOf(c) {
  const names = [...(c.npm || []), ...(c.packages || []).map((p) => p.name), ...(c.split || []).map((s) => s.name)];
  return names.flatMap((n) => (UNPACKED[n] || []).map((d) => `node_modules/${n}/${d}`));
}

/** One pack's built folder, as tools/make-packs.mjs reads it. */
function treeFor(id, tree, toolNames, contents = null, rendererOut = path.join(APP, 'build', 'out')) {
  const d = describePack(id, { toolNames });
  const folder = path.join(tree, id);
  fs.mkdirSync(path.join(folder, 'main'), { recursive: true });
  let main = 'main/index.js';
  if (contents) main = carry(folder, contents, rendererOut);
  else fs.writeFileSync(path.join(folder, 'main', 'index.js'), shim(d.name));
  // The icons Windows reads from disk: the taskbar's tile and each kind of file's.
  const copy = (from, to) => {
    if (!fs.existsSync(from)) return false;
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.copyFileSync(from, to);
    return true;
  };
  const tiles = d.tiles.filter((t) => copy(path.join(APP, 'resources', 'apps', `${t.app}.ico`), path.join(folder, ...t.icon.split('/'))));
  for (const t of d.fileTypes) {
    if (!copy(path.join(APP, 'resources', 'filetypes', path.basename(t.icon)), path.join(folder, ...t.icon.split('/')))) throw new Error(`${id}: the icon ${t.icon} for .${t.ext} is not in resources/filetypes.`);
  }
  const { id: _id, ...said } = d;
  const unpack = [tiles.length ? 'tiles' : null, d.fileTypes.length ? 'filetypes' : null, ...(contents ? unpackedOf(contents) : [])].filter(Boolean);
  fs.writeFileSync(path.join(folder, 'pack.json'), `${JSON.stringify({ ...said, tiles, main, unpack }, null, 2)}\n`);
  return folder;
}

/** A string as NSIS reads it between double quotes. */
const q = (s) => `"${String(s).replace(/\$/g, '$$$$').replace(/"/g, '$\\"')}"`;

/**
 * What the installer writes, a pack at a time: the variables, the choice,
 * the page's boxes, the files and the kinds of file each registers, and
 * every kind of file for the uninstaller to let go of.
 */
export function installerScript({ dir, packs }) {
  const ids = packs.map((p) => p.id);
  const L = [];
  const push = (...lines) => L.push(...lines);
  push(
    '; The app packs this installer carries. Made by apps/desktop/build/make-app-packs.js',
    '; from the build, with resources/installer.nsh reading it: not edited by hand.',
    '',
    '!ifndef BUILD_UNINSTALLER',
    ...ids.flatMap((id) => [`Var rutbaPack_${id}`, `Var rutbaPackBox_${id}`]),
    'Var rutbaPacksList',
    '!endif',
    '',
    `!define RUTBA_ALL_PACKS ${q(ids.join(','))}`,
    '',
    '; One pack chosen or not, from the list in $R1.',
    '!macro rutbaPackDecide ID',
    '  StrCpy $rutbaPack_${ID} "0"',
    '  ${WordFind} ",$R1," ",${ID}," "E*" $R2',
    '  ${IfNot} ${Errors}',
    '    StrCpy $rutbaPack_${ID} "1"',
    '  ${EndIf}',
    '  ClearErrors',
    '!macroend',
    '',
    '; One pack chosen when the list in $R3 does not have it: new since that list.',
    '!macro rutbaPackNew ID',
    '  ${WordFind} ",$R3," ",${ID}," "E*" $R2',
    '  ${If} ${Errors}',
    '    StrCpy $rutbaPack_${ID} "1"',
    '  ${EndIf}',
    '  ClearErrors',
    '!macroend',
    '',
    '; Which packs to write: those named by /PACKS=a,b, else those installed',
    '; now (Packs, which the installer and Rutba Office write), else every',
    '; pack: a first install, or an update from a version with one archive.',
    '; With those installed now, a pack the copy before had no word of',
    '; (PacksKnown, every pack it knew) is new in this version and written too;',
    '; one it knew and the person left out stays out. With no PacksKnown every',
    '; pack is taken as known, and none is added.',
    '!macro rutbaPacksDecide',
    '  ClearErrors',
    '  StrCpy $R4 "0"',
    '  ${GetParameters} $R0',
    '  ${GetOptions} $R0 "/PACKS=" $R1',
    '  ${If} ${Errors}',
    '  ${OrIf} $R1 == ""',
    '    ClearErrors',
    '    ReadRegStr $R1 HKCU "Software\\Rutba Office" "Packs"',
    '    ${If} ${Errors}',
    '    ${OrIf} $R1 == ""',
    '      StrCpy $R1 "${RUTBA_ALL_PACKS}"',
    '    ${Else}',
    '      StrCpy $R4 "1"',
    '    ${EndIf}',
    '  ${EndIf}',
    '  ClearErrors',
    ...ids.map((id) => `  !insertmacro rutbaPackDecide ${id}`),
    '  ${If} $R4 == "1"',
    '    ReadRegStr $R3 HKCU "Software\\Rutba Office" "PacksKnown"',
    '    ${If} ${Errors}',
    '    ${OrIf} $R3 == ""',
    '      StrCpy $R3 "${RUTBA_ALL_PACKS}"',
    '    ${EndIf}',
    '    ClearErrors',
    ...ids.map((id) => `    !insertmacro rutbaPackNew ${id}`),
    '  ${EndIf}',
    '!macroend',
    '',
    '; The page\'s boxes, in two columns, ticked as chosen.',
    '!macro rutbaPackBox ID LABEL X Y',
    '  ${NSD_CreateCheckbox} ${X} ${Y} 46% 12u "${LABEL}"',
    '  Pop $rutbaPackBox_${ID}',
    '  ${If} $rutbaPack_${ID} == "1"',
    '    ${NSD_Check} $rutbaPackBox_${ID}',
    '  ${EndIf}',
    '!macroend',
    '!macro rutbaPackBoxes',
  );
  const half = Math.ceil(ids.length / 2);
  ids.forEach((id, i) => {
    const col = i < half ? '0' : '52%';
    const row = i < half ? i : i - half;
    push(`  !insertmacro rutbaPackBox ${id} "$rutbaW_${id}" ${col} ${24 + row * 13}u`);
  });
  push(
    '!macroend',
    '',
    '; The boxes read when the page is left; $2 counts the apps ticked (not the add-ons).',
    '!macro rutbaPackRead ID APP',
    '  ${NSD_GetState} $rutbaPackBox_${ID} $1',
    '  ${If} $1 == ${BST_CHECKED}',
    '    StrCpy $rutbaPack_${ID} "1"',
    '    IntOp $2 $2 + ${APP}',
    '  ${Else}',
    '    StrCpy $rutbaPack_${ID} "0"',
    '  ${EndIf}',
    '!macroend',
    '!macro rutbaPackReads',
    ...ids.map((id) => `  !insertmacro rutbaPackRead ${id} ${ADD_ONS.includes(id) ? 0 : 1}`),
    '!macroend',
    '',
    '; A kind of file, registered as electron-builder registers one, with the',
    '; program\'s path quoted; and let go of again.',
    '!macro rutbaAssociate EXT CLASS DESCRIPTION ICON',
    '  WriteRegStr SHELL_CONTEXT "Software\\Classes\\.${EXT}" "" "${CLASS}"',
    '  WriteRegNone SHELL_CONTEXT "Software\\Classes\\.${EXT}\\OpenWithProgids" "${CLASS}"',
    '  WriteRegStr SHELL_CONTEXT "Software\\Classes\\${CLASS}" "" "${DESCRIPTION}"',
    '  WriteRegStr SHELL_CONTEXT "Software\\Classes\\${CLASS}\\DefaultIcon" "" "${ICON}"',
    '  WriteRegStr SHELL_CONTEXT "Software\\Classes\\${CLASS}\\shell" "" "open"',
    '  WriteRegStr SHELL_CONTEXT "Software\\Classes\\${CLASS}\\shell\\open" "" "Open with Rutba Office"',
    '  WriteRegStr SHELL_CONTEXT "Software\\Classes\\${CLASS}\\shell\\open\\command" "" "$\\"$appExe$\\" $\\"%1$\\""',
    '!macroend',
    '!macro rutbaUnassociate EXT CLASS',
    '  DeleteRegValue SHELL_CONTEXT "Software\\Classes\\.${EXT}\\OpenWithProgids" "${CLASS}"',
    '  ReadRegStr $R0 SHELL_CONTEXT "Software\\Classes\\.${EXT}" ""',
    '  ${If} $R0 == "${CLASS}"',
    '    DeleteRegValue SHELL_CONTEXT "Software\\Classes\\.${EXT}" ""',
    '  ${EndIf}',
    '  DeleteRegKey SHELL_CONTEXT "Software\\Classes\\${CLASS}"',
    '!macroend',
    '',
  );
  // Each pack's files, its manifest last, and its kinds of file.
  for (const p of packs) {
    const f = packFiles(p.id);
    push(`!macro rutbaWritePack_${p.id}`);
    push(`  DetailPrint ${q(`Installing ${p.name}`)}`);
    push('  SetOutPath "$INSTDIR\\resources\\packs"');
    push(`  File ${q(path.join(dir, f.archive))}`);
    if (fs.existsSync(path.join(dir, f.unpacked))) push(`  File /r ${q(path.join(dir, f.unpacked))}`);
    push(`  File ${q(path.join(dir, f.manifest))}`);
    for (const t of p.fileTypes || []) {
      const icon = `$INSTDIR\\resources\\packs\\${f.unpacked}\\${t.icon.split('/').join('\\')}`;
      push(`  !insertmacro rutbaAssociate ${q(t.ext)} ${q(t.name || t.ext)} ${q(t.description || '')} "${icon}"`);
    }
    push('!macroend');
  }
  push(
    '',
    '; The packs chosen written, and Packs, so the next installer and Rutba Office',
    '; know them; and PacksKnown, every pack this installer has, so the next one',
    '; can tell a pack new in its version from one left out here.',
    '!macro rutbaWritePacks',
    '  StrCpy $rutbaPacksList ""',
    '  CreateDirectory "$INSTDIR\\resources\\packs"',
  );
  for (const id of ids) {
    push(`  \${If} $rutbaPack_${id} == "1"`, `    !insertmacro rutbaWritePack_${id}`, `    StrCpy $rutbaPacksList "$rutbaPacksList,${id}"`, '  ${EndIf}');
  }
  push(
    '  StrCpy $rutbaPacksList $rutbaPacksList "" 1',
    '  WriteRegStr HKCU "Software\\Rutba Office" "Packs" "$rutbaPacksList"',
    '  WriteRegStr HKCU "Software\\Rutba Office" "PacksKnown" "${RUTBA_ALL_PACKS}"',
    '  SetOutPath "$INSTDIR"',
    '  ; Windows is told the kinds of file changed.',
    '  System::Call "shell32::SHChangeNotify(i 0x08000000, i 0, i 0, i 0)"',
    '!macroend',
    '',
    '; Every kind of file any pack registers, let go of by the uninstaller.',
    '!macro rutbaForgetTypes',
  );
  const seen = new Set();
  for (const p of packs) {
    for (const t of p.fileTypes || []) {
      const key = `${t.ext}|${t.name || t.ext}`;
      if (seen.has(key)) continue;
      seen.add(key);
      push(`  !insertmacro rutbaUnassociate ${q(t.ext)} ${q(t.name || t.ext)}`);
    }
  }
  push('!macroend', '');
  return L.join('\r\n');
}

/** A release's file for one pack: a zip of its three things, as Add or remove apps unpacks it (packs-service.js). */
export function packZip(dir, id) {
  const f = packFiles(id);
  const entries = [];
  const add = (abs, name) => {
    const e = new ZipEntry({ name, method: 8, crc: 0, compressedSize: 0, uncompressedSize: 0, compressed: Buffer.alloc(0), dosTime: 0, dosDate: 0x21, flags: 0x0800, externalAttrs: 0, comment: Buffer.alloc(0) });
    e.data = fs.readFileSync(abs);
    entries.push(e);
  };
  const walk = (abs, name) => {
    for (const ent of fs.readdirSync(abs, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
      if (ent.isDirectory()) walk(path.join(abs, ent.name), `${name}/${ent.name}`);
      else add(path.join(abs, ent.name), `${name}/${ent.name}`);
    }
  };
  add(path.join(dir, f.archive), f.archive);
  if (fs.existsSync(path.join(dir, f.unpacked))) walk(path.join(dir, f.unpacked), f.unpacked);
  add(path.join(dir, f.manifest), f.manifest);
  return writeZip(entries);
}

/**
 * Makes every pack (or those named), signed, and what the core and the
 * installer read of them. `meta` and `rendererOut` are the renderer's build
 * (what went into each file, and the folder of them), when not the last one.
 */
export async function makeAppPacks({ privateKey, version, only = null, release = null, root = PACKS_DIR, keysFile = KEYS_FILE, installerFile = INSTALLER_FILE, carryCode = CARRY_CODE, movedFile = path.join(root, 'moved.json'), meta = null, rendererOut = path.join(APP, 'build', 'out') }) {
  const tree = path.join(root, 'tree');
  const out = path.join(root, 'out');
  fs.rmSync(tree, { recursive: true, force: true });
  fs.rmSync(out, { recursive: true, force: true });
  const toolNames = [...TOOL_LIST, ...REACH_TOOL_LIST].map((t) => t.name);
  const ids = PACKS.map((p) => p.id).filter((id) => !only || only.includes(id));
  const contents = await packContents(ids.filter((id) => carryCode.includes(id)), { meta });
  for (const id of ids) treeFor(id, tree, toolNames, contents[id] || null, rendererOut);
  // What the core's archive leaves out: each moved app's files, by their paths in it.
  fs.mkdirSync(path.dirname(movedFile), { recursive: true });
  fs.writeFileSync(movedFile, `${JSON.stringify({ packs: Object.fromEntries(Object.entries(contents).map(([id, c]) => [id, c.core])) }, null, 2)}\n`);
  const made = await makePacks({ tree, out, version, privateKey });
  const packs = ids.map((id) => JSON.parse(fs.readFileSync(path.join(out, `${id}.pack.json`), 'utf8')).manifest);
  const sizes = Object.fromEntries(ids.map((id) => [id, packSize(out, id)]));
  const publicPem = crypto.createPublicKey(privateKey).export({ format: 'pem', type: 'spki' });
  fs.mkdirSync(path.dirname(keysFile), { recursive: true });
  fs.writeFileSync(keysFile, `${JSON.stringify({ keys: [publicPem], sizes }, null, 2)}\n`);
  fs.writeFileSync(installerFile, installerScript({ dir: out, packs }));
  const files = [];
  if (release) {
    fs.mkdirSync(release, { recursive: true });
    for (const id of ids) {
      const file = path.join(release, packFileName(version, id));
      fs.writeFileSync(file, packZip(out, id));
      files.push(file);
    }
  }
  fs.rmSync(tree, { recursive: true, force: true });
  return { out, made, sizes, files, contents };
}

function option(args, name) {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : null;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  try {
    const privateKey = readPrivateKey({ file: option(args, '--key') });
    const version = JSON.parse(fs.readFileSync(path.join(APP, 'package.json'), 'utf8')).version;
    const only = option(args, '--only')?.split(',').map((s) => s.trim()).filter(Boolean) || null;
    const release = option(args, '--release');
    const { sizes, files } = await makeAppPacks({ privateKey, version, only, release: release ? path.resolve(release) : null });
    for (const [id, bytes] of Object.entries(sizes)) console.log(`${id}: ${(bytes / 1024).toFixed(0)} KB`);
    if (files.length) console.log(`${files.length} release pack files in ${path.dirname(files[0])}`);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}
