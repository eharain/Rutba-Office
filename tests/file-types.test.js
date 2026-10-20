// Choose apps on Windows: an app turned off takes its types out of the
// registry, turned on puts them back, and a type another program has taken
// meanwhile is left with it. Run against a key of its own under HKCU, the
// way the installer writes the classes, never the real Software\Classes.

import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { fileTypePlan, applyFileTypes, installerChoice, recordChoice } from '../apps/desktop/main/file-types.js';
import { fileAssociations } from '@rutba/office-formats/registry';

const windows = process.platform === 'win32';
const ROOT = `HKCU\\Software\\RutbaOfficeTest${process.pid}\\Classes`;
const reg = (...args) => { try { return execFileSync('reg', args, { windowsHide: true, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }); } catch { return null; } };
const defaultOf = (key) => {
  const out = reg('query', key, '/ve');
  if (out == null) return null;
  const v = /REG_SZ\s+(.*?)\s*$/m.exec(out)?.[1] ?? '';
  return /^\(.*\)$/.test(v) ? '' : v;
};
const listed = (ext, progId) => reg('query', `${ROOT}\\.${ext}\\OpenWithProgids`, '/v', progId) != null;

/** What the installer writes for one type (APP_ASSOCIATE). */
function install(ext, progId) {
  reg('add', `${ROOT}\\.${ext}`, '/ve', '/d', progId, '/f');
  reg('add', `${ROOT}\\.${ext}\\OpenWithProgids`, '/v', progId, '/t', 'REG_NONE', '/f');
  reg('add', `${ROOT}\\${progId}\\shell\\open\\command`, '/ve', '/d', '"C:\\Rutba Office\\Rutba Office.exe" "%1"', '/f');
}

test('file types: the plan keeps the chosen apps\' types; Worksheets off lets .xlsx go and keeps .csv for Contacts', () => {
  const chosen = ['mail', 'calendar', 'contacts', 'word', 'slides', 'pdf', 'pictures', 'image', 'video', 'recorder'];
  const plan = fileTypePlan(fileAssociations(), fileAssociations({ apps: chosen }));
  const by = Object.fromEntries(plan.map((p) => [p.ext, p.keep]));
  assert.equal(by.xlsx, false);
  assert.equal(by.ods, false);
  assert.equal(by.docx, true);
  assert.equal(by.csv, true);
  assert.equal(plan.every((p) => typeof p.progId === 'string' && p.progId), true);
});

test('file types: withdrawn when the app is turned off, back when it is turned on, and a type taken by another program stays with it', { skip: !windows && 'Windows only' }, async () => {
  try {
    install('xlsx', 'Excel Workbook Test');
    install('docx', 'Word Document Test');
    install('ods', 'OpenDocument Spreadsheet Test');
    // A type the suite registers that this copy did not install is left alone.
    reg('add', `${ROOT}\\.odp`, '/ve', '/d', 'SomebodyElse.odp', '/f');
    const plan = [
      { ext: 'xlsx', progId: 'Excel Workbook Test', keep: false },
      { ext: 'ods', progId: 'OpenDocument Spreadsheet Test', keep: false },
      { ext: 'docx', progId: 'Word Document Test', keep: true },
      { ext: 'odp', progId: 'Presentation Test', keep: false },
    ];
    const off = await applyFileTypes(plan, { root: ROOT });
    assert.deepEqual(off.withdrawn.sort(), ['ods', 'xlsx']);
    assert.deepEqual(off.notOurs, ['odp']);
    assert.equal(defaultOf(`${ROOT}\\.xlsx`), '');
    assert.equal(listed('xlsx', 'Excel Workbook Test'), false);
    assert.equal(defaultOf(`${ROOT}\\.docx`), 'Word Document Test');
    assert.equal(listed('docx', 'Word Document Test'), true);
    assert.equal(defaultOf(`${ROOT}\\.odp`), 'SomebodyElse.odp');

    // Meanwhile another program takes .ods.
    reg('add', `${ROOT}\\.ods`, '/ve', '/d', 'LibreOffice.Calc', '/f');
    const on = await applyFileTypes(plan.map((p) => ({ ...p, keep: true })), { root: ROOT });
    assert.deepEqual(on.restored.sort(), ['ods', 'xlsx']);
    assert.equal(defaultOf(`${ROOT}\\.xlsx`), 'Excel Workbook Test');
    assert.equal(listed('xlsx', 'Excel Workbook Test'), true);
    assert.equal(defaultOf(`${ROOT}\\.ods`), 'LibreOffice.Calc');
    assert.equal(listed('ods', 'OpenDocument Spreadsheet Test'), true);
    const again = await applyFileTypes(plan.map((p) => ({ ...p, keep: true })), { root: ROOT });
    assert.deepEqual(again.restored, []);
  } finally {
    reg('delete', ROOT.replace(/\\Classes$/, ''), '/f');
  }
});

test('the installer\'s choice: read only when the installer chose last; what Office records reads back as Office\'s', { skip: !windows && 'Windows only' }, async () => {
  const key = `HKCU\\Software\\RutbaOfficeChoiceTest${process.pid}`;
  try {
    assert.equal(await installerChoice({ key }), null);
    reg('add', key, '/v', 'AppsOff', '/t', 'REG_SZ', '/d', 'sheets,video,', '/f');
    reg('add', key, '/v', 'AppsChosenAt', '/t', 'REG_SZ', '/d', 'installer', '/f');
    assert.deepEqual(await installerChoice({ key }), { off: ['sheets', 'video'], addons: [] });
    await recordChoice(['pdf'], { key });
    assert.equal(await installerChoice({ key }), null);
    assert.match(reg('query', key, '/v', 'AppsOff'), /REG_SZ\s+pdf/);
    await recordChoice([], { key });
    reg('add', key, '/v', 'AppsChosenAt', '/t', 'REG_SZ', '/d', 'installer', '/f');
    assert.deepEqual(await installerChoice({ key }), { off: [], addons: [] });
  } finally {
    reg('delete', key, '/f');
  }
});
