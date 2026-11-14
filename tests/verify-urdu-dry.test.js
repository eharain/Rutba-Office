// The Urdu, Arabic and Hindi checks (apps/desktop/main/verify-urdu.js) run
// dry: no Electron and no windows, a stand-in harness that answers each
// question the checks ask a window. Every check runs to its end and passes
// when the windows answer as they should, and fails when they do not: a check
// that throws before it is asked (a name used twice, as `word` once was, the
// pack's words and the Documents window) is caught here, not by a gate.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import esbuild from 'esbuild';
import { keyFor } from '../apps/desktop/main/verify-keys.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MAIN = path.join(ROOT, 'apps/desktop/main');

/** verify-urdu.js bundled for node, with stand-ins for electron, the main process's words (globalThis.packTables) and the document builder. */
async function checks() {
  const stand = {
    electron: 'export const BrowserWindow = { getAllWindows: () => [] };',
    '@rutba/office-shell/electron/words': 'export const wordsFor = (tag) => (globalThis.packTables[tag] ? { tag, table: globalThis.packTables[tag] } : null);',
    '@rutba/ooxml/build': 'export const buildDocx = () => Buffer.from("docx");',
  };
  const out = await esbuild.build({
    entryPoints: [path.join(MAIN, 'verify-urdu.js')],
    bundle: true,
    format: 'esm',
    platform: 'node',
    write: false,
    logLevel: 'silent',
    plugins: [{
      name: 'stand-ins',
      setup(b) {
        b.onResolve({ filter: /^(electron|@rutba\/office-shell\/electron\/words|@rutba\/ooxml\/build)$/ }, (a) => ({ path: a.path, namespace: 'stand-in' }));
        b.onLoad({ filter: /.*/, namespace: 'stand-in' }, (a) => ({ contents: stand[a.path], loader: 'js' }));
      },
    }],
  });
  // Beside verify-urdu.js: verify-keys.js reads the labels by a path relative to itself (createRequire).
  const file = path.join(MAIN, `.verify-urdu-dry-${process.pid}.mjs`);
  fs.writeFileSync(file, out.outputFiles[0].text);
  try {
    return await import(pathToFileURL(file).href);
  } finally {
    fs.rmSync(file, { force: true });
  }
}

/** A pack's words for the labels the checks read, by key. */
const WORDS = {
  ur: { Home: 'ہوم', Insert: 'داخل کریں', Table: 'جدول', Cancel: 'منسوخ کریں', Hijri: 'ہجری' },
  ar: { Home: 'الرئيسية', Insert: 'إدراج', Table: 'جدول', Cancel: 'إلغاء', Hijri: 'هجري' },
  hi: { Home: 'होम', Insert: 'सम्मिलित करें', Table: 'तालिका', Cancel: 'रद्द करें', Hijri: 'हिजरी' },
};
const tableOf = (tag) => Object.fromEntries(Object.entries(WORDS[tag]).map(([english, w]) => [keyFor(english), w]));

/**
 * What each window answers, as the app in `tag` should: `wrong` names one
 * answer to give as an English window would, to see its check fail.
 */
function answers(tag, { wrong = null } = {}) {
  const rtl = tag !== 'hi';
  const dir = rtl ? 'rtl' : 'ltr';
  const w = (english) => (wrong === 'english' ? english : WORDS[tag][english]);
  const locale = new Intl.Locale(tag);
  const firstDay = (typeof locale.getWeekInfo === 'function' ? locale.getWeekInfo() : locale.weekInfo)?.firstDay ?? 1;
  const typed = { ur: 'اردو', ar: 'عربي', hi: 'हिन्दी' }[tag];
  const rules = [
    [/store\.set/, 1],
    [/'\.wd-page \[data-block="0"\]'/, true],
    [/^\(\{ lang:/, { lang: tag, dir: wrong === 'direction' ? (rtl ? 'ltr' : 'rtl') : dir, tabs: [w('Home'), w('Insert'), 'View'], page: 'ltr', ribbon: wrong === 'direction' ? (rtl ? 'ltr' : 'rtl') : dir }],
    [/getBoundingClientRect/, { x: 10, y: 10 }],
    [/textContent\.includes\(/, true],
    [/\[data-block="1"\]'\)\?\.textContent$/, `Another line. ${typed}`],
    [/\.rw-tab'\)\]\)\?\.click\(\)/, 1],
    [/return Boolean\(b\)/, true],
    [/Boolean\(document\.querySelector\('\.rw-dialog'\)\?\.textContent\)/, true],
    [/document\.querySelector\('\.rw-dialog'\)\?\.textContent \|\| ''/, wrong === 'english' ? 'Insert table Columns Rows' : `${w('Table')} ${w('Insert')}`],
    [/\.rw-dialog \.rw-btn/, 1],
    [/Boolean\(document\.querySelector\('\.sh-grid'\)\)/, true],
    [/^\(\{ dir: document\.documentElement\.dir, grid:/, { dir, grid: 'ltr', formula: 'ltr', tabs: [w('Home')] }],
    [/Boolean\(document\.querySelector\('\.sl-stage'\)\)/, true],
    [/^\(\{ dir: document\.documentElement\.dir, stage:/, { dir, stage: 'ltr' }],
    [/Boolean\(document\.querySelector\('\.cal-month-head div'\)\)/, true],
    [/\.cal-day-h'\)\.length >= 28/, true],
    [/const days =/, { days: 30, named: rtl ? ['١ رجب'] : [], button: rtl, first: new Date(2024, 0, 7 + (firstDay % 7)).toLocaleDateString(tag, { weekday: 'short' }) }],
  ];
  return (code) => {
    const rule = rules.find(([re]) => re.test(code.trim()));
    if (!rule) throw new Error(`a question the stand-in does not know: ${code.slice(0, 80)}`);
    return rule[1];
  };
}

/** Runs one language's checks against windows answering with `answer`; what each check said. */
async function run(verify, answer) {
  const said = [];
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'verify-urdu-dry-'));
  const win = () => ({
    id: 0,
    webContents: { executeJavaScript: async (code) => answer(code), sendInputEvent() {}, invalidate() {}, capturePage: async () => ({ toPNG: () => Buffer.alloc(0) }) },
    isDestroyed: () => false,
    destroy() {},
  });
  const harness = {
    open: async () => win(),
    check: (name, ok, detail) => said.push({ name, ok: Boolean(ok), detail }),
    until: async (fn, what) => { if (!(await fn())) throw new Error(`no ${what}`); },
    wait: async () => {},
    errorsIn: async () => [],
  };
  try {
    await verify(harness, { dir, files: { xlsx: 'book.xlsx', pptx: 'deck.pptx' } });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
  return said;
}

const LANGS = [['ur', 'verifyUrdu', 8], ['ar', 'verifyArabic', 8], ['hi', 'verifyHindi', 7]];

test('every Urdu, Arabic and Hindi check runs to its end and passes when the windows are in the language', async () => {
  const c = await checks();
  for (const [tag, fn, count] of LANGS) {
    globalThis.packTables = { [tag]: tableOf(tag) };
    const said = await run(c[fn], answers(tag));
    assert.ok(!said.some((s) => /checks ran$/.test(s.name)), `${tag}: ${JSON.stringify(said.find((s) => /checks ran$/.test(s.name)))}`);
    assert.equal(said.length, count, `${tag}: ${said.map((s) => s.name).join(' | ')}`);
    assert.deepEqual(said.filter((s) => !s.ok), [], tag);
  }
});

test('the checks fail when a window is not in the language, its frame the wrong way, or the pack missing', async () => {
  const c = await checks();
  for (const [tag, fn] of LANGS) {
    globalThis.packTables = { [tag]: tableOf(tag) };
    const english = await run(c[fn], answers(tag, { wrong: 'english' }));
    assert.ok(english.some((s) => !s.ok && /Documents window/.test(s.name)), `${tag}: the Documents check`);
    assert.ok(english.some((s) => !s.ok && /dialog/.test(s.name)), `${tag}: the dialog check`);
    assert.ok(!english.some((s) => /checks ran$/.test(s.name)), tag);
    const turned = await run(c[fn], answers(tag, { wrong: 'direction' }));
    assert.ok(turned.some((s) => !s.ok && /Documents window/.test(s.name)), `${tag}: the frame's direction`);
    // No pack installed: the checks' own words are English, and a window in English fails them.
    globalThis.packTables = {};
    const none = await run(c[fn], answers(tag, { wrong: 'english' }));
    assert.ok(none.some((s) => !s.ok && /Documents window/.test(s.name)), `${tag}: no pack`);
  }
});
