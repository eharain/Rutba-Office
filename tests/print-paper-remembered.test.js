// The paper a print starts on, and one layout at a time (2.28.0).
//
// Until 2.28.0 the print dialog opened on A4 every time, in every app and every region. Now it opens
// on the paper this person last chose in that app, and on the region's paper (Letter or A4) until
// they have chosen one. A workbook whose file names a paper still opens on that, and one whose file
// names none is not made to say A4 by a print area set from the ribbon.
//
// And a choice changed in the dialog no longer starts a whole new layout while the older ones keep
// running: one is in hand at a time, and of the changes made meanwhile only the last is laid out.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { SheetView } from '@rutba/sheet-view';
import { readPageSetup, writePageSetup } from '@rutba/sheet-view/print';
import { buildXlsx } from '@rutba/ooxml/build';
import { PAPERS, PAPER_MM, regionPaper, paperKey, startingPaper, oneAtATime } from '../apps/desktop/renderer/print-page.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (rel) => fs.readFileSync(path.join(ROOT, rel), 'utf8');

test('a region prints on Letter where the region does, and on A4 everywhere else', () => {
  for (const r of ['US', 'CA', 'MX', 'PH', 'CL', 'CO', 'VE', 'us', ' ca ']) assert.equal(regionPaper(r), 'Letter', r);
  for (const r of ['PK', 'GB', 'DE', 'IN', 'AE', 'BR', 'AU', 'JP', '', null, undefined, 'ZZ', 'USA']) assert.equal(regionPaper(r), 'A4', String(r));
});

test('a print starts on the paper last chosen in the app, and on the region\'s until one is chosen', () => {
  assert.equal(startingPaper({ remembered: 'A5', region: 'US' }), 'A5', 'the paper chosen stands over the region');
  assert.equal(startingPaper({ remembered: null, region: 'US' }), 'Letter');
  assert.equal(startingPaper({ remembered: null, region: 'PK' }), 'A4');
  assert.equal(startingPaper({}), 'A4', 'nothing known: A4');
  assert.equal(startingPaper(), 'A4');
  // What the setting holds is a person's file: a paper the dialog does not offer is not started on.
  for (const odd of ['Tabloid', 'a4', '', 9, {}, ['Letter']]) assert.equal(startingPaper({ remembered: odd, region: 'US' }), 'Letter', JSON.stringify(odd));
  assert.deepEqual(PAPERS, Object.keys(PAPER_MM), 'the papers offered are the papers a page can be laid out on');
  assert.ok(PAPERS.includes('A4') && PAPERS.includes('Letter'));
});

test('each app remembers its own paper, under a setting named for it', () => {
  assert.equal(paperKey('sheets'), 'print.paper.sheets');
  assert.equal(paperKey('mail'), 'print.paper.mail');
  assert.notEqual(paperKey('calendar'), paperKey('contacts'));
  assert.equal(paperKey(''), 'print.paper.app');
  assert.equal(paperKey('a b/c.d'), 'print.paper.abcd', 'nothing in a name makes another setting\'s key');
});

test('one piece of work is in hand at a time, and of what is asked meanwhile only the last is done', async () => {
  const lane = oneAtATime();
  const began = [];
  let running = 0;
  let most = 0;
  const gates = {};
  const work = (name) => () => {
    began.push(name);
    running += 1;
    most = Math.max(most, running);
    return new Promise((resolve) => {
      gates[name] = () => {
        running -= 1;
        resolve(name);
      };
    });
  };
  const a = lane(work('a'));
  await new Promise((r) => setImmediate(r));
  // Four changes while the first is laid out: none begins, and only the last waits.
  const b = lane(work('b'));
  const c = lane(work('c'));
  const d = lane(work('d'));
  const e = lane(work('e'));
  assert.deepEqual(began, ['a']);
  assert.deepEqual(await b, { dropped: true });
  assert.deepEqual(await c, { dropped: true });
  assert.deepEqual(await d, { dropped: true });
  gates.a();
  assert.deepEqual(await a, { value: 'a' });
  await new Promise((r) => setImmediate(r));
  assert.deepEqual(began, ['a', 'e'], 'the last change is laid out next, and none of those between');
  gates.e();
  assert.deepEqual(await e, { value: 'e' });
  assert.equal(most, 1, 'never two at once');
  // With nothing in hand the next begins at once.
  const f = lane(work('f'));
  await new Promise((r) => setImmediate(r));
  assert.deepEqual(began, ['a', 'e', 'f']);
  gates.f();
  assert.deepEqual(await f, { value: 'f' });
});

test('work that fails says so to the one who asked, and the next piece is still done', async () => {
  const lane = oneAtATime();
  await assert.rejects(lane(() => Promise.reject(new Error('could not be laid out'))), /could not be laid out/);
  await assert.rejects(lane(() => { throw new Error('thrown at once'); }), /thrown at once/);
  assert.deepEqual(await lane(() => 7), { value: 7 });
});

test('a workbook says whether its file names a paper, and a setup sent back as read does not make it name one', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Report', rows: [['a', 'b'], [1, 2]] }] }));
  const xml = () => {
    const part = view.workbook.partNameFor('Report');
    return view.workbook.snapshotParts([part])[part];
  };
  const first = readPageSetup(view, 'Report');
  assert.equal(first.paperStated, false, 'a file with no setup names no paper');
  assert.equal(first.paper, 'A4', 'and reads as A4, as before');
  // A print area set from the ribbon reads the setup and sends it back whole.
  writePageSetup(view, 'Report', { ...first, area: 'A1:B2' });
  assert.doesNotMatch(xml(), /paperSize=/, 'the file still names no paper');
  assert.equal(readPageSetup(view, 'Report').paperStated, false);
  assert.equal(readPageSetup(new SheetView(view.save()), 'Report').paperStated, false, 'nor after a save');
  // A paper chosen is written: by the print dialog (no word either way), and by the ribbon's Size.
  writePageSetup(view, 'Report', { paper: 'Letter', orientation: 'auto' });
  assert.match(xml(), /paperSize="1"/);
  assert.equal(readPageSetup(view, 'Report').paperStated, true);
  assert.equal(readPageSetup(view, 'Report').paper, 'Letter');
  writePageSetup(view, 'Report', { ...readPageSetup(view, 'Report'), area: 'A1:B2' });
  assert.match(xml(), /paperSize="1"/, 'a paper the file names is kept through a write');
  const fresh = new SheetView(buildXlsx({ sheets: [{ name: 'Report', rows: [['a']] }] }));
  writePageSetup(fresh, 'Report', { ...readPageSetup(fresh, 'Report'), paper: 'A5', paperStated: true });
  assert.equal(readPageSetup(fresh, 'Report').paper, 'A5');
  assert.equal(readPageSetup(fresh, 'Report').paperStated, true);
});

test('the dialog starts on that paper, keeps the one chosen, and lays out through the one lane', () => {
  const dialog = read('apps/desktop/renderer/print.js');
  // The paper to start on is asked before anything is counted, and a workbook's own paper stands when its file names one.
  assert.match(dialog, /paperToStartOn\(shell, app\)/);
  assert.match(dialog, /paperStated === false \? null : fromFile\.paper/);
  assert.match(dialog, /if \(!started\) return undefined;/);
  // A paper chosen is kept for the app; a document's is the document's own.
  assert.match(dialog, /if \(patch\.paper\) rememberPaper\(patch\.paper\);/);
  assert.match(dialog, /if \(kind === 'doc' \|\| !app \|\| !PAPERS\.includes\(paper\)\) return;/);
  assert.match(dialog, /shell\.store\.set\(\{ key: paperKey\(app\), value: paper \}\)/);
  // Every layout the dialog asks for goes through a lane, and none is asked for beside it.
  for (const call of ['shell.print.pageSummary(', 'shell.print.summary(', 'shell.print.preview(']) {
    const at = dialog.indexOf(call);
    assert.ok(at > 0, call);
    assert.equal(dialog.indexOf(call, at + 1), -1, `${call} is asked for in one place`);
    assert.match(dialog.slice(Math.max(0, at - 400), at), /lanes\.current\s*\.(count|picture)\(/, `${call} goes through a lane`);
  }
  // The shell says the region, and the two exports with no dialog go on the same paper.
  assert.match(read('packages/office-shell/src/electron/ipc.js'), /region: regionOf\(app\)/);
  assert.match(read('apps/desktop/renderer/apps/sheets.js'), /said\?\.paperStated \? said\.paper : await paperToStartOn\(shell\)/);
  assert.match(read('apps/desktop/renderer/apps/sheets.js'), /next\.paper = arg\.size; next\.paperStated = true;/);
  assert.match(read('apps/desktop/renderer/apps/slides.js'), /paper: await paperToStartOn\(shell\)/);
});
