// Fonts for a language's letters (L8): the table, the check that adds a
// font pack only when this computer has no face for a language, and the
// font slots a window is given.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { addonFileUrl } from '../packages/office-shell/src/electron/addon-files.js';
import {
  FONT_NEEDS, scriptOf, needsOf, faceMeets, fontsNeeded, createFontFinder, readFacesOfFile, fontFiles,
  fontPackFaces, windowFonts, fontSlots, systemFontDirs,
} from '../packages/office-shell/src/electron/language-fonts.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { facesOf } = createRequire(import.meta.url)('../packages/pdf/src/truetype.js');
const LIBERATION = path.join(ROOT, 'apps', 'desktop', 'vendor', 'pdfjs', 'standard_fonts', 'LiberationSans-Regular.ttf');
const row = (pack) => FONT_NEEDS.find((r) => r.pack === pack);
const face = (family, text) => ({ family, has: (cp) => [...text].some((c) => c.codePointAt(0) === cp) });

test('Urdu asks for Nastaliq first, then any Arabic face; Chinese by its script or region', () => {
  assert.deepEqual(needsOf('ur').map((r) => r.pack), ['font-aran', 'font-arab']);
  assert.deepEqual(needsOf('ur-PK').map((r) => r.pack), ['font-aran', 'font-arab']);
  assert.deepEqual(needsOf('lang-ur'.replace(/^lang-/, '')).map((r) => r.pack), ['font-aran', 'font-arab']);
  assert.deepEqual(needsOf('ar').map((r) => r.pack), ['font-arab']);
  assert.deepEqual(needsOf('hi').map((r) => r.pack), ['font-deva']);
  assert.equal(scriptOf('zh-Hans'), 'Hans');
  assert.equal(scriptOf('zh-Hant-TW'), 'Hant');
  assert.equal(scriptOf('zh-TW'), 'Hant');
  assert.equal(scriptOf('zh-CN'), 'Hans');
  assert.deepEqual(needsOf('fr'), []);
  assert.deepEqual(needsOf('xx'), []);
  // Each pack named once, as font-<script> in lower case.
  assert.equal(new Set(FONT_NEEDS.map((r) => r.pack)).size, FONT_NEEDS.length);
  for (const r of FONT_NEEDS) assert.match(r.pack, /^font-[a-z]{4}$/);
});

test('a face meets a row by its letters, and by its family where the row names them', () => {
  assert.equal(faceMeets(face('Arial', 'ےڈبی'), row('font-arab')), true);
  // Arial has Urdu's letters, but is not Nastaliq.
  assert.equal(faceMeets(face('Arial', 'ےڈبی'), row('font-aran')), false);
  assert.equal(faceMeets(face('Noto Nastaliq Urdu', 'ےڈ'), row('font-aran')), true);
  assert.equal(faceMeets(face('noto nastaliq urdu', 'ے'), row('font-aran')), false);
  // Simplified needs a letter only it writes.
  assert.equal(faceMeets(face('Traditional only', '漢這'), row('font-hans')), false);
});

test('a font pack is added only when nothing here meets the language, and the feed offers it', async () => {
  const has = (met) => async (r) => (met.includes(r.pack) ? 'Some Face' : null);
  assert.deepEqual(await fontsNeeded({ tags: ['lang-hi'], has: has(['font-deva']), offered: ['font-deva'] }), { add: [], lacking: [] });
  assert.deepEqual((await fontsNeeded({ tags: ['lang-hi'], has: has([]), offered: ['font-deva'] })).add, ['font-deva']);
  // Already here or staged.
  assert.deepEqual((await fontsNeeded({ tags: ['lang-hi'], has: has([]), installed: ['font-deva'], offered: ['font-deva'] })).add, []);
  // Not offered: said, not added.
  const none = await fontsNeeded({ tags: ['hi'], has: has([]), offered: [] });
  assert.deepEqual(none.add, []);
  assert.deepEqual(none.lacking.map((r) => r.pack), ['font-deva']);
  // Urdu with no Nastaliq: the Nastaliq pack when offered.
  assert.deepEqual((await fontsNeeded({ tags: ['lang-ur'], has: has(['font-arab']), offered: ['font-aran', 'font-arab'] })).add, ['font-aran']);
  // With no Nastaliq pack offered, an Arabic face here is enough; with none either, the Arabic pack.
  assert.deepEqual((await fontsNeeded({ tags: ['lang-ur'], has: has(['font-arab']), offered: ['font-arab'] })).add, []);
  assert.deepEqual((await fontsNeeded({ tags: ['lang-ur'], has: has([]), offered: ['font-arab'] })).add, ['font-arab']);
  // Urdu and Arabic together ask for the Arabic pack once.
  assert.deepEqual((await fontsNeeded({ tags: ['lang-ar', 'lang-ur'], has: has([]), offered: ['font-arab'] })).add, ['font-arab']);
  // English and languages written in Latin letters never need one.
  assert.deepEqual((await fontsNeeded({ tags: ['lang-en', 'lang-en-us', 'lang-fr'], has: has([]), offered: FONT_NEEDS.map((r) => r.pack) })).add, []);
});

test('the finder reads a file only when asked, each once, and names the face that met the row', async () => {
  const reads = [];
  const files = ['/f/arial.ttf', '/f/other.ttf', '/f/NotoNastaliqUrdu-Regular.ttf', '/f/broken.ttf'];
  const fonts = {
    '/f/arial.ttf': [face('Arial', 'бЖαאשبی')],
    '/f/other.ttf': [face('Other', 'x')],
    '/f/NotoNastaliqUrdu-Regular.ttf': [face('Noto Nastaliq Urdu', 'ےڈبی')],
  };
  const has = createFontFinder({
    files: () => files,
    faces: (file) => {
      reads.push(file);
      if (!fonts[file]) throw new Error('not a font');
      return fonts[file];
    },
  });
  assert.equal(await has(row('font-cyrl')), 'Arial');
  assert.deepEqual(reads, ['/f/arial.ttf']);
  // The file whose name suggests the family is read first.
  assert.equal(await has(row('font-aran')), 'Noto Nastaliq Urdu');
  assert.equal(reads[1], '/f/NotoNastaliqUrdu-Regular.ttf');
  // A row nothing meets reads the rest, and a broken file is passed over.
  assert.equal(await has(row('font-deva')), null);
  assert.deepEqual([...new Set(reads)].sort(), [...files].sort());
  assert.equal(reads.length, new Set(reads).size);
  // Asked again: the answer kept, nothing read.
  const before = reads.length;
  assert.equal(await has(row('font-cyrl')), 'Arial');
  assert.equal(reads.length, before);

  // A face whose lookup throws (a cmap that reads past its file) meets nothing, and the next face is asked (L6).
  const bad = { family: 'Broken', has: () => { throw new RangeError('offset is out of bounds'); } };
  const past = createFontFinder({ files: () => ['/f/a.ttf', '/f/b.ttf'], faces: (f) => (f === '/f/a.ttf' ? [bad] : [face('Arial', 'бЖ')]) });
  assert.equal(await past(row('font-cyrl')), 'Arial');
  // And a look that fails as a whole is not kept: asked again, it is tried again.
  let fail = true;
  const again = createFontFinder({ files: () => { if (fail) throw new Error('no fonts folder'); return ['/f/a.ttf']; }, faces: () => [face('Arial', 'бЖ')] });
  await assert.rejects(again(row('font-cyrl')));
  fail = false;
  assert.equal(await again(row('font-cyrl')), 'Arial');
});

test('a real font file: its family and letters, read in parts, and a collection of faces', () => {
  const read = readFacesOfFile(facesOf);
  const [f] = read(LIBERATION);
  assert.equal(f.family, 'Liberation Sans');
  assert.equal(f.has('б'.codePointAt(0)), true);
  assert.equal(f.has('α'.codePointAt(0)), true);
  assert.equal(f.has('ب'.codePointAt(0)), false);
  assert.equal(faceMeets(f, row('font-cyrl')), true);
  assert.equal(faceMeets(f, row('font-arab')), false);

  // The same face twice in a .ttc: the face's table offsets move by the header's size.
  const bytes = fs.readFileSync(LIBERATION);
  const header = Buffer.alloc(20);
  header.write('ttcf', 0, 'latin1');
  header.writeUInt32BE(0x00010000, 4);
  header.writeUInt32BE(2, 8);
  header.writeUInt32BE(20, 12);
  header.writeUInt32BE(20, 16);
  const face0 = Buffer.from(bytes);
  const count = face0.readUInt16BE(4);
  for (let i = 0; i < count; i++) face0.writeUInt32BE(face0.readUInt32BE(12 + 16 * i + 8) + 20, 12 + 16 * i + 8);
  const ttc = Buffer.concat([header, face0]);
  const both = facesOf(ttc);
  assert.equal(both.length, 2);
  assert.ok(both.every((x) => x.family === 'Liberation Sans' && x.has('б'.codePointAt(0))));
  // Something that is not a font fails to read, and the finder passes it over.
  assert.throws(() => facesOf(Buffer.from('not a font at all')));
});

test('font packs here: their faces, served from the pack, and given to a window in their language', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-fontpack-'));
  try {
    const aran = path.join(dir, 'font-aran');
    fs.mkdirSync(path.join(aran, 'fonts'), { recursive: true });
    fs.copyFileSync(LIBERATION, path.join(aran, 'fonts', 'Nastaliq.ttf'));
    fs.writeFileSync(path.join(aran, 'manifest.json'), '{}');
    const deva = path.join(dir, 'font-deva');
    fs.mkdirSync(deva);
    fs.copyFileSync(LIBERATION, path.join(deva, 'Deva.ttf'));
    const addons = [
      { id: 'font-aran', kind: 'font', dir: aran, manifest: {} },
      { id: 'font-deva', kind: 'font', dir: deva, manifest: {} },
      { id: 'lang-ur', kind: 'language', dir: aran, manifest: {} },
    ];
    const packFaces = fontPackFaces(addons, readFacesOfFile(facesOf));
    assert.deepEqual(packFaces.map((f) => [f.pack, f.rel, f.family, f.first]), [
      ['font-aran', 'fonts/Nastaliq.ttf', 'Liberation Sans', true],
      ['font-deva', 'Deva.ttf', 'Liberation Sans', false],
    ]);
    const urlOf = (f) => addonFileUrl(f.pack, f.rel);
    const ur = windowFonts('ur', { packFaces, urlOf });
    assert.deepEqual(ur, [{ family: 'Liberation Sans', url: 'rutba://app/addons/font-aran/fonts/Nastaliq.ttf', first: true }]);
    assert.deepEqual(fontSlots(ur), { first: '"Liberation Sans",', last: '' });
    const hi = windowFonts('hi', { packFaces, urlOf });
    assert.deepEqual(fontSlots(hi), { first: '', last: '"Liberation Sans",' });
    assert.deepEqual(windowFonts('fr', { packFaces, urlOf }), []);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a window in Urdu is given the computer\'s own Nastaliq first when no pack brings one', () => {
  const got = windowFonts('ur', { packFaces: [], urlOf: () => '', systemFirst: { 'font-aran': 'Jameel Noori Nastaleeq' } });
  assert.deepEqual(got, [{ family: 'Jameel Noori Nastaleeq', url: null, first: true }]);
  assert.deepEqual(fontSlots(got), { first: '"Jameel Noori Nastaleeq",', last: '' });
  assert.deepEqual(windowFonts('ar', { packFaces: [], urlOf: () => '', systemFirst: { 'font-aran': 'Jameel Noori Nastaleeq' } }), []);
  // A family name cannot end the CSS string it is put in.
  assert.equal(fontSlots([{ family: 'Bad", serif; x', first: true }]).first, '"Bad, serif; x",');
});

test('the font folders each platform keeps, and the files found in them', () => {
  assert.deepEqual(systemFontDirs({ platform: 'win32', env: { WINDIR: 'C:\\Windows', LOCALAPPDATA: 'C:\\Users\\a\\AppData\\Local' } }).map((d) => d.replace(/\\/g, '/')),
    ['C:/Windows/Fonts', 'C:/Users/a/AppData/Local/Microsoft/Windows/Fonts']);
  assert.ok(systemFontDirs({ platform: 'darwin', env: {}, home: '/Users/a' }).includes('/System/Library/Fonts'));
  assert.ok(systemFontDirs({ platform: 'linux', env: {}, home: '/home/a' }).includes('/usr/share/fonts'));
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-fontdirs-'));
  try {
    fs.mkdirSync(path.join(dir, 'truetype', 'noto'), { recursive: true });
    for (const f of ['a.ttf', 'b.OTF', 'c.ttc', 'notes.txt', path.join('truetype', 'noto', 'd.ttf')]) fs.writeFileSync(path.join(dir, f), '');
    assert.deepEqual(fontFiles([dir, path.join(dir, 'missing')]).map((f) => path.relative(dir, f).split(path.sep).join('/')).sort(), ['a.ttf', 'b.OTF', 'c.ttc', 'truetype/noto/d.ttf']);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
