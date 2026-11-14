// Every place that reads or writes a document's bytes by path asks the table
// of files opened in memory first (2.26.0, D-2026-11-13-20). The shell's code
// runs only inside Electron, so these read its source: a new file call that
// takes a path and does not ask the table fails here, before it can write a
// locked file out in the clear.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { asEnglish } from './label-english.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (rel) => asEnglish(fs.readFileSync(path.join(ROOT, rel), 'utf8')).replace(/\r\n/g, '\n');

/** The body of each `name: async? ({ path: p ... }) =>` entry inside `impl.<ns> = { ... };`. */
function entriesOf(src, ns) {
  const start = src.indexOf(`impl.${ns} = {`);
  assert.ok(start >= 0, `impl.${ns} is there`);
  const end = src.indexOf('\n  };\n', start);
  const body = src.slice(start, end);
  const out = new Map();
  const re = /\n {4}(\w+): (?:async )?\(([^)]*)\)/g;
  const marks = [...body.matchAll(re)];
  marks.forEach((m, i) => {
    const to = i + 1 < marks.length ? marks[i + 1].index : body.length;
    out.set(m[1], { args: m[2], body: body.slice(m.index, to) });
  });
  return out;
}

// Calls that only ever touch the scratch files made in this run, never a person's file.
const SCRATCH_ONLY = new Set(['temp', 'append', 'patch', 'dropTemp']);

test('every shell file call that takes a path asks the table first', () => {
  const src = read('packages/office-shell/src/electron/ipc.js');
  const fsCalls = entriesOf(src, 'fs');
  assert.ok(fsCalls.size > 10);
  const missed = [];
  for (const [name, { args, body }] of fsCalls) {
    if (SCRATCH_ONLY.has(name)) continue;
    if (!/path: p\b|\bto\b/.test(args)) continue;
    const asks = name === 'keepTemp' ? /mem\(to\)/.test(body) : /mem\(p\)|mem\(path\.join/.test(body);
    if (!asks) missed.push(name);
  }
  assert.deepEqual(missed, [], `these do not ask the table: ${missed.join(', ')}`);
  for (const name of ['read', 'readHead', 'readText', 'write', 'writeText', 'stat', 'exists', 'list', 'keepTemp']) assert.ok(fsCalls.has(name), name);
});

test('the shell keeps a file opened in memory out of the recent list, the thumbnail cache, the folder view and other programs', () => {
  const src = read('packages/office-shell/src/electron/ipc.js');
  assert.match(src, /addRecent: \(p\) => \(mem\(p\?\.path \?\? p\) \|\| memoryFiles\(\)\?\.isLocked\?\.\(String\(p\?\.path \?\? p \?\? ''\)\) \? undefined : stores\.recent\.add\(p\)\)/);
  // A frame a window drew is never kept for a file from the Locker, nor from a window showing one (a locked film's plain copy).
  assert.match(src, /put: async \(\{ path: p, bytes \}, win\) => \(thumbnailer && !refusesThumb\(p, win\) \? thumbnailer\.put\(p, bytes\) : \{ stored: false \}\)/);
  assert.match(src, /const refusesThumb = \(p, win\) => !p \|\| Boolean\(mem\(p\)\) \|\| Boolean\(memoryFiles\(\)\?\.isLocked\?\.\(String\(p\)\)\) \|\| Boolean\(lockedIn\(win\)\);/);
  // The Locker's plain copies too: no thumbnail, and not listed from the command line either.
  const shellMain = read('packages/office-shell/src/electron/main.js');
  assert.match(shellMain, /if \(!memoryFiles\(\)\?\.isLocked\?\.\(p\)\) stores\.recent\.add\(\{ path: p, app: which \}\);/);
  const app = entriesOf(src, 'shell');
  assert.match(app.get('showInFolder').body, /mem\(p\)/);
  assert.match(app.get('openPath').body, /if \(mem\(p\)\) throw notHere\(p\)/);
});

test('rutba://file and rutba://thumb answer a file opened in memory from the table, never from the disk', () => {
  const src = read('packages/office-shell/src/electron/protocol.js');
  const file = src.slice(src.indexOf("if (host === 'file')"), src.indexOf("if (host === 'thumb')"));
  const held = file.indexOf('held?.isUnder(target)');
  assert.ok(held > 0 && held < file.indexOf('fs.existsSync(target)'), 'the table is asked before the disk');
  assert.ok(held < file.indexOf('pictureResponse('), 'and before anything is decoded through a path');
  const thumb = src.slice(src.indexOf("if (host === 'thumb')"), src.indexOf("if (host === 'blob')"));
  assert.ok(thumb.indexOf('memoryFiles()?.isUnder(target) || memoryFiles()?.isLocked?.(target)') > 0 && thumb.indexOf('memoryFiles()?.isLocked?.(target)') < thumb.indexOf('thumbnailer.respond'));
});

test('the shell is handed the table by the desktop app', () => {
  assert.match(read('packages/office-shell/src/electron/main.js'), /installProtocol\(\{[^}]*memoryFiles \}\)/);
  assert.match(read('packages/office-shell/src/electron/main.js'), /buildImplementations\(\{[^}]*memoryFiles \}\)/);
  assert.match(read('apps/desktop/main/main.js'), /\n {2}memoryFiles,\n/);
});

test('documents.js and pdf.js read and write a document only through the hooked helpers', () => {
  const doc = read('apps/desktop/main/documents.js');
  // The open runners read through readDocument; the two write helpers ask inMemory first.
  assert.doesNotMatch(doc, /step\.value\.read \? (await )?fs\./);
  for (const fn of ['function writeWhole(target, data, encoding) {\n    if (inMemory(target))', 'async function writeWholeAsync(target, data, encoding) {\n    if (inMemory(target))']) {
    assert.ok(doc.includes(fn), fn);
  }
  const pdf = read('apps/desktop/main/pdf.js');
  // Every finished PDF goes through writeOut, which asks the table.
  assert.equal((pdf.match(/await writeWhole\(/g) || []).length, 1, 'writeWhole is called only from writeOut');
  assert.doesNotMatch(pdf, /await fsp\.readFile\(file\)/);
});

test('a window showing a file from the Locker writes only back into it', () => {
  const src = read('packages/office-shell/src/electron/ipc.js');
  const fsCalls = entriesOf(src, 'fs');
  for (const [name, arg] of [['write', 'p'], ['writeText', 'p'], ['temp', 'null'], ['keepTemp', 'to']]) {
    assert.ok(fsCalls.get(name).body.includes(`staysIn(win, ${arg})`), name);
  }
  assert.match(src, /return shown && t\.isLocked\(shown\) \? shown : null;/);
});

test('Video refuses to edit a film from the Locker, in words that only a build with the Locker carries', async () => {
  const player = read('apps/desktop/renderer/apps/video/player.js');
  const editor = read('apps/desktop/renderer/apps/video/index.js');
  assert.match(player, /export const lockedFilmSays = process\.env\.RUTBA_LOCKER === '1' \? \(name\) => import\('\.\/locked-film\.js'\)/);
  assert.match(player, /if \(stat\?\.locked && lockedFilmSays\)/);
  assert.equal(editor.split('if (stat?.locked && lockedFilmSays)').length - 1, 2, 'adding a film and replacing a clip with one');
  assert.doesNotMatch((player + editor).replace(/process\.env\.RUTBA_LOCKER/g, ''), /locker/i, 'nothing in Video names it but the build switch');
  // The words module runs only in a window (its imports are the window's), so its source is read.
  const words = read('apps/desktop/renderer/apps/video/locked-film.js');
  assert.match(words, /LOCKED_FILM = '\{name\} is in the Locker\. Video plays it, but does not edit or export it\. To edit it, unlock it in the Locker first\.'/);
  for (const tag of ['ur', 'ar', 'hi']) assert.match(words, new RegExp(`\\n  ${tag}: \\{ \\[LOCKED_FILM\\]: '\\{name\\} [^']+' \\},`), tag);
  assert.match(words, /for \(const \[tag, words\] of Object\.entries\(LOCKED_FILM_WORDS\)\) addToCatalogue\(tag, words\);/);
});

test('the Locker window offers back what was open, says why a file opened from a plain copy, and names Touch ID on a Mac', () => {
  const win = read('apps/desktop/renderer/apps/locker.js');
  // The offer after an unlock: reopen all, or leave them closed, asking first when that drops unsaved changes.
  assert.match(win, /<ReopenOffer shell=\{shell\} offers=\{state\.reopen\} toast=\{toast\} onChanged=\{refresh\} \/>/);
  const offer = win.slice(win.indexOf('function ReopenOffer('), win.indexOf('function Settings('));
  assert.match(offer, /shell\.locker\.reopen\(\{ keys \}\)/);
  assert.match(offer, /shell\.locker\.dismissReopen\(\{ keys \}\)/);
  assert.ok(offer.indexOf('if (unsaved.length)') < offer.indexOf('dismissReopen'), 'asked before unsaved changes go');
  assert.match(offer, /t\('Reopen the files that were open'\)/);
  // A file that opened from a plain copy says why, for each reason open gives.
  assert.match(win, /if \(out\.plainWhy\) toast\?\.\(plainLine\(out\.plainWhy, e\.name\)/);
  const line = win.slice(win.indexOf('function plainLine('), win.indexOf('function ReopenOffer('));
  for (const why of ['too-large', 'by-path']) assert.ok(line.includes(`why === '${why}'`), why);
  assert.match(line, /is over 1 GB/);
  // The quick unlock is named for what it is.
  assert.match(win, /state\.quick === 'touchid' \? t\('Use Touch ID'\) : t\('Use Windows Hello/);
  assert.match(win, /state\.helloOn \? t\('Turn off Touch ID'\) : t\('Turn on Touch ID'\)/);
});

test('Word, Sheets and Slides say when a file opened as saved because its sealed work would not open, in words only a build with the Locker carries', () => {
  const says = read('apps/desktop/renderer/apps/set-aside.js');
  assert.match(says, /const says = process\.env\.RUTBA_LOCKER === '1' \? \(name, how\) => import\('\.\/locked-work\.js'\)/);
  assert.match(says, /says\(name, opened\.setAside\)/);
  assert.match(says, /if \(!opened\?\.setAside \|\| !says\) return;/);
  assert.doesNotMatch(says.replace(/process\.env\.RUTBA_LOCKER/g, ''), /locker/i, 'nothing there names it but the build switch');
  for (const app of ['word', 'sheets', 'slides']) assert.match(read(`apps/desktop/renderer/apps/${app}.js`), /sayIfSetAside\((opened|r), toast\)/, app);
  const words = read('apps/desktop/renderer/apps/locked-work.js');
  assert.match(words, /SET_ASIDE = '\{name\} opened as it was last saved\. Its unsaved work could not be opened here\. The Locker keeps it, and offers it again once this file is closed\.'/);
  // A copy that does not open is never offered, and the window does not say it will be.
  assert.match(words, /SET_ASIDE_UNREAD = '\{name\} opened as it was last saved\. Its unsaved work could not be opened\.'/);
  assert.match(words, /t\(how === 'unread' \? SET_ASIDE_UNREAD : SET_ASIDE, \{ name \}\)/);
  for (const tag of ['ur', 'ar', 'hi']) assert.match(words, new RegExp(`\\n  ${tag}: \\{\\n`), tag);
  for (const key of ['SET_ASIDE', 'SET_ASIDE_UNREAD']) {
    assert.equal(words.match(new RegExp(`\\n    \\[${key}\\]: '[^']*\\{name\\}[^']*',`, 'g'))?.length, 3, `${key} in ur, ar and hi`);
  }
  // documents.js tells the window, and never marks the saved file unsaved over that work.
  const doc = read('apps/desktop/main/documents.js');
  assert.match(doc, /\.\.\.\(work === 'offered' \|\| work === 'unread' \? \{ setAside: work \} : \{\}\)/);
  assert.match(doc, /if \(work === 'sealed'\) session\.dirty = true;/);
});

test('a window showing one of the Locker\'s plain copies saves over that copy, and a film asked for whole is streamed', () => {
  const src = read('packages/office-shell/src/electron/ipc.js');
  // Its own file, or a new one beside it in the same locked folder; never another file opened from the Locker.
  assert.match(src, /if \(p && samePath\(p, shown\)\) return;\n    if \(p && mem\(p\) && memoryFiles\(\)\?\.besides\?\.\(String\(p\), shown\)\) return;\n    throw notHere\(shown\);/);
  const proto = read('packages/office-shell/src/electron/protocol.js');
  const fn = proto.slice(proto.indexOf('export async function memoryResponse('), proto.indexOf('/* ── pictures'));
  assert.ok(fn.indexOf('if (!m && st.ranged)') < fn.indexOf('await table.read(target)'), 'a ranged file is never read whole for a request with no range');
  assert.match(fn, /new ReadableStream\(/);
});

test('the Locker window says a file in an older format is saved as a new file beside it, and its Save As asks no place', () => {
  const win = read('apps/desktop/renderer/apps/locker.js');
  assert.match(win, /else if \(out\.note === 'old-kind'\) toast\?\.\(t\('\{name\} is in an older format\. Saving it makes \{saved\} beside it, in the same locked folder\.', \{ name: e\.name, saved: savedAs\(e\.name\) \}\)/);
  assert.match(win, /const NEWER = \{ doc: 'docx', htm: 'docx', xls: 'xlsx', ppt: 'pptx' \};/);
  // Two copies of one file's work reopened together: the second opens as a new file beside it, and says so.
  assert.match(win, /if \(o\.note === 'beside'\) toast\?\.\(t\('\{name\} is already open, so these unsaved changes open as \{saved\}, a new file beside it in the same locked folder\.', \{ name, saved: o\.beside \}\)/);
  assert.match(read('apps/desktop/main/extras/locker/index.js'), /beside: done\.beside \|\| null, oldKind: Boolean\(done\.oldKind\)/);
  // Reopen goes newest first, through the opener; an older format beside a file still says what a save makes.
  assert.match(read('apps/desktop/main/extras/locker/index.js'), /for \(const done of await memory\.reopenAll\(keys\)\)/);
  assert.match(win, /if \(o\.note === 'old-kind' \|\| o\.oldKind\) toast\?\.\(t\('\{name\} is in an older format\. Saving it makes \{saved\} beside it, in the same locked folder\.', \{ name: shownAs, saved: savedAs\(shownAs\) \}\)/);
  const ipc = read('packages/office-shell/src/electron/ipc.js');
  const save = ipc.slice(ipc.indexOf('    save: async (p, win) => {'), ipc.indexOf('    message: async (p, win) => {'));
  assert.ok(save.indexOf('OLD_FORMAT.test(shown)') > 0 && save.indexOf('OLD_FORMAT.test(shown)') < save.indexOf('dialog.showSaveDialog'), 'answered before any dialog is shown');
  assert.match(save, /return path\.join\(path\.dirname\(shown\), path\.basename\(String\(p\?\.defaultPath \|\| shown\)\)\);/);
});
