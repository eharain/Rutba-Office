// The pages PDF draws for a print (pdf/print-draw.js): counted as they are
// drawn, stopped when asked, and nothing kept of a print that does not happen.

import test from 'node:test';
import assert from 'node:assert/strict';
import { drawForPrint, printCancelled } from '../apps/desktop/renderer/apps/pdf/print-draw.js';

const pages = (n) => Array.from({ length: n }, (_, i) => ({ page: i + 1 }));

test('every page is drawn in order, and each is told before it is drawn', async () => {
  const told = [];
  const order = [];
  const out = await drawForPrint(pages(3), async (p, i) => { order.push(`draw ${p.page}`); return { url: `u${i}` }; }, {
    onPage: (n, count) => { told.push([n, count]); order.push(`tell ${n}`); },
  });
  assert.deepEqual(out, [{ url: 'u0' }, { url: 'u1' }, { url: 'u2' }]);
  assert.deepEqual(told, [[1, 3], [2, 3], [3, 3]]);
  assert.deepEqual(order, ['tell 1', 'draw 1', 'tell 2', 'draw 2', 'tell 3', 'draw 3']);
});

test('stopped part way, no more is drawn, what was drawn is let go, and the answer is none', async () => {
  let stop = false;
  const drawn = [];
  const dropped = [];
  const out = await drawForPrint(pages(300), async (p) => { drawn.push(p.page); if (p.page === 4) stop = true; return { url: `u${p.page}` }; }, {
    stopped: () => stop,
    drop: (pic) => dropped.push(pic.url),
  });
  assert.equal(out, null);
  assert.deepEqual(drawn, [1, 2, 3, 4], 'the page being drawn is finished, and no other begun');
  assert.deepEqual(dropped, ['u1', 'u2', 'u3', 'u4']);
});

test('stopped while the last page was drawn, the print still does not happen', async () => {
  let stop = false;
  const dropped = [];
  const out = await drawForPrint(pages(2), async (p) => { if (p.page === 2) stop = true; return { url: `u${p.page}` }; }, { stopped: () => stop, drop: (pic) => dropped.push(pic.url) });
  assert.equal(out, null);
  assert.deepEqual(dropped, ['u1', 'u2']);
});

test('stopped before it began, nothing is drawn or told', async () => {
  const seen = [];
  const out = await drawForPrint(pages(2), async () => { seen.push('draw'); return {}; }, { stopped: () => true, onPage: () => seen.push('tell') });
  assert.equal(out, null);
  assert.deepEqual(seen, []);
});

test('a page that cannot be drawn fails the print, and the pages before it are let go', async () => {
  const dropped = [];
  await assert.rejects(
    drawForPrint(pages(3), async (p) => { if (p.page === 3) throw new Error('page 3 is damaged'); return { url: `u${p.page}` }; }, { drop: (pic) => dropped.push(pic.url) }),
    /page 3 is damaged/
  );
  assert.deepEqual(dropped, ['u1', 'u2']);
});

test('a file of no pages draws nothing and answers an empty list', async () => {
  assert.deepEqual(await drawForPrint([], async () => { throw new Error('never'); }), []);
});

test('a dialog the person closed is told apart from a print that failed', () => {
  assert.equal(printCancelled({ ok: false, reason: 'cancelled' }), true);
  assert.equal(printCancelled({ ok: false, reason: 'Print job canceled' }), true);
  assert.equal(printCancelled({ ok: false, reason: null }), true);
  assert.equal(printCancelled({ ok: false, reason: 'Invalid printer settings' }), false);
  assert.equal(printCancelled({ ok: false, reason: 'failed' }), false);
  assert.equal(printCancelled({ ok: true, reason: null }), false);
});
