// What keeps a window check steady beside other lanes (apps/desktop/main/verify-steady.js, 2.28.0).
//
// Eleven checks failed once each in the 2.27.0 gate beside other lanes and passed alone. A check may
// now try a step again, but only on evidence that the step was undone from outside the check (the
// window lost the keyboard or was resized) or that a press landed beside a handle. These tests hold
// the rule that matters: a step that fails with no such evidence fails at once, as it did before.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { steadily, still, missed } from '../apps/desktop/main/verify-steady.js';

const MAIN = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'apps', 'desktop', 'main');
const read = (f) => fs.readFileSync(path.join(MAIN, f), 'utf8');

/** A page that answers the two questions steadily() asks: begin noting, and what was noted since. */
function page() {
  const seen = [];
  const js = async (code) => {
    if (/slice\((\d+)\)/.test(code)) return seen.slice(Number(/slice\((\d+)\)/.exec(code)[1]));
    return seen.length;
  };
  return { js, seen };
}

test('a step that passes is run once', async () => {
  const { js } = page();
  let runs = 0;
  const got = await steadily(js, async () => { runs += 1; return { ok: true, found: 7 }; });
  assert.equal(runs, 1);
  assert.deepEqual(got, { ok: true, found: 7, tries: 1, seen: [] });
});

test('a step that fails with nothing seen from outside fails at once: it is not tried again', async () => {
  const { js } = page();
  let runs = 0;
  const got = await steadily(js, async () => { runs += 1; return { ok: false, why: 'the menu had no Info' }; });
  assert.equal(runs, 1, 'no second try without evidence');
  assert.equal(got.ok, false);
  assert.equal(got.tries, 1);
  assert.deepEqual(got.seen, []);
  assert.equal(got.why, 'the menu had no Info', 'the try\'s own findings are kept for the detail');
});

test('a step undone from outside is tried again, and the answer says the tries and what was seen', async () => {
  const { js, seen } = page();
  let runs = 0;
  const got = await steadily(js, async (n) => {
    runs += 1;
    if (n === 1) { seen.push('blur'); return { ok: false }; }
    return { ok: true };
  });
  assert.equal(runs, 2);
  assert.deepEqual(got, { ok: true, tries: 2, seen: ['blur'] });
});

test('what was seen before a try began is not evidence for it', async () => {
  const { js, seen } = page();
  seen.push('resize', 'resize');
  let runs = 0;
  const got = await steadily(js, async () => { runs += 1; return { ok: false }; });
  assert.equal(runs, 1);
  assert.deepEqual(got.seen, []);
});

test('three tries at most, however often the window is disturbed', async () => {
  const { js, seen } = page();
  let runs = 0;
  const got = await steadily(js, async () => { runs += 1; seen.push('blur'); return { ok: false }; });
  assert.equal(runs, 3);
  assert.equal(got.ok, false);
  assert.equal(got.tries, 3);
  assert.deepEqual(got.seen, ['blur', 'blur', 'blur']);
});

test('a point is read until two readings in a row agree', async () => {
  const readings = [{ x: 10, y: 10 }, { x: 40, y: 10 }, { x: 52, y: 10 }, { x: 52, y: 10 }, { x: 90, y: 90 }];
  let asked = 0;
  const at = await still(async () => readings[Math.min(asked++, readings.length - 1)], async () => {}, { gap: 1, limit: 50 });
  assert.deepEqual(at, { x: 52, y: 10 });
  assert.equal(asked, 4, 'and no further once they agree');
  // A point that is not there yet, then is.
  const late = [null, null, { x: 5, y: 6 }, { x: 5, y: 6 }];
  let n = 0;
  assert.deepEqual(await still(async () => late[Math.min(n++, late.length - 1)], async () => {}, { gap: 1, limit: 50 }), { x: 5, y: 6 });
  // One that is never there answers null, within its limit.
  assert.equal(await still(async () => null, async () => {}, { gap: 1, limit: 5 }), null);
});

test('a drag that did nothing is made again only when its press missed the handle', () => {
  assert.equal(missed(null), true, 'no press reached the page');
  assert.equal(missed({ handle: false, took: 'sh-cells' }), true, 'it landed beside the handle');
  assert.equal(missed({ handle: true, took: 'sh-obj-rotate' }), false, 'the handle took it and turned nothing: the window\'s fault');
});

test('the checks that try again say so in their detail, and judge every try by the same test', () => {
  const ribbon = read('verify-ribbon.js');
  assert.match(ribbon, /const tipped = await steadily\(js, async \(n\) => \{/);
  assert.match(ribbon, /tries: tipped\.tries, seen: tipped\.seen\.join\(' '\)/);
  const inspect = read('verify-inspect.js');
  assert.match(inspect, /const way = await steadily\(js, async \(n\) => \{/);
  assert.match(inspect, /tries: way\.tries, seen: way\.seen\.join\(' '\)/);
  // The two drags: the handle read standing still, the press noted, and a press the handle took not tried again.
  for (const [file, handle] of [['verify-rotate-frame-sheet.js', '.sh-obj-rotate'], ['verify-deck-model3d.js', '.sl-model3d-handle']]) {
    const src = read(file);
    assert.ok(src.includes(`await pressWatch(js, '${handle}');`), `${file} notes the press`);
    assert.match(src, /await still\(/, `${file} reads the handle standing still`);
    assert.match(src, /if \(!missed\((landed|took)\)\) (return false|break);/, `${file} stops when the handle took the press`);
    assert.match(src, /for \(let n = 1; n <= 3; n\+\+\) \{/, `${file} tries three times at most`);
  }
  // The library waits on its own list, and no check here waits a fixed time for a file to be read.
  const video = read('verify-video-library.js');
  assert.match(video, /'the clip in the library', SIGNAL_WAIT\)/);
});
