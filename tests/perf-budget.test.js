/**
 * Performance budget — a large file must open, lay out and save in seconds.
 *
 * The numbers here are generous on purpose (fifteen seconds a step on a slow
 * build machine). They exist to catch a step that has gone quadratic, which
 * shows up as minutes, not as a few percent. Each test prints its timings so a
 * drift towards the limit is visible in the log long before it trips.
 */
import test from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { buildDocx, buildXlsx, colName, Workbook, recalculateWorkbook } from '@rutba/ooxml';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { MailStore } from '@rutba/mailbox';

const BUDGET_MS = 15000;

/**
 * Run one named step, print how long it took, and fail if it blew the budget.
 * The budget is held against the step's own CPU time, not the clock: the rest
 * of the suite runs beside this file, each file in a process of its own, and
 * on a busy machine a three-second recalculation took sixteen by the clock. A
 * step gone quadratic costs minutes of its own CPU however quiet the machine.
 */
function timed(label, steps, fn) {
  const t0 = performance.now();
  const c0 = process.cpuUsage();
  const out = fn();
  const ms = Math.round(performance.now() - t0);
  const used = process.cpuUsage(c0);
  const cpu = Math.round((used.user + used.system) / 1000);
  steps.push(`${label}=${ms}ms (cpu ${cpu}ms)`);
  assert.ok(cpu < BUDGET_MS, `${label} took ${cpu}ms of its own CPU (${ms}ms by the clock), budget ${BUDGET_MS}ms`);
  return out;
}

const WORDS = 'the quick brown fox jumps over a lazy dog while ledgers balance and totals agree'.split(' ');
const sentence = (seed, n) => Array.from({ length: n }, (_, i) => WORDS[(seed * 7 + i * 3) % WORDS.length]).join(' ');

test('a 300-page document opens, paginates and saves inside the budget', () => {
  const steps = [];
  // Each paragraph runs to ~1/3 of a page (about 260 words): three runs, mixed formatting.
  const paragraphs = Array.from({ length: 900 }, (_, i) => ({
    runs: [
      { text: sentence(i, 100) + ' ' },
      { text: sentence(i + 1, 60) + ' ', bold: i % 3 === 0 },
      { text: sentence(i + 2, 100), italic: i % 5 === 0 },
    ],
  }));
  const bytes = timed('build', steps, () => buildDocx({ paragraphs }));
  const view = timed('open+blocks', steps, () => { const v = openDocx(bytes); void v.blocks; return v; });
  const pages = timed('paginate', steps, () => view.pages);
  const count = pages?.pages?.length ?? pages?.length ?? 0;
  timed('repaginate-cached-lines', steps, () => { view._pages = null; return view.pages; });
  // An edit first, so the save really has a changed part to flush.
  timed('edit', steps, () => view.doc.setParagraphRuns(450, [{ text: 'Edited in the middle.' }]));
  const saved = timed('save', steps, () => view.doc.save());
  const again = timed('reopen', steps, () => { const v = openDocx(saved); void v.blocks; return v; });
  console.log(`# doc: ${count} pages, ${(bytes.length / 1024) | 0} KiB; ${steps.join(' ')}`);
  assert.equal(again.blocks.length, 900);
  assert.ok(count >= 250, `expected a long document, got ${count} pages`);
});

test('a 100,000-cell workbook with formulas recalculates and saves inside the budget', () => {
  const steps = [];
  // 5,000 rows x 20 columns. Column A is data; every other cell is a formula
  // over its left and upper neighbours, so the dependency graph is one wide
  // wavefront rather than a flat list.
  const ROWS = 5000, COLS = 20;
  const rows = [];
  for (let r = 0; r < ROWS; r++) {
    const row = new Array(COLS);
    for (let c = 0; c < COLS; c++) {
      if (c === 0) row[c] = r + 1;
      else if (r === 0) row[c] = `=${colName(c - 1)}1+1`;
      else row[c] = `=${colName(c - 1)}${r + 1}+${colName(c)}${r}*0.5`;
    }
    rows.push(row);
  }
  const bytes = timed('build', steps, () => buildXlsx({ sheets: [{ name: 'Big', rows }] }));
  const wb = timed('open', steps, () => Workbook.open(bytes));
  void wb;
  const result = timed('recalc+write-back', steps, () => recalculateWorkbook(bytes));
  const reopened = timed('reopen', steps, () => Workbook.open(result.output));
  void reopened;
  console.log(`# sheet: ${ROWS * COLS} cells, ${result.calculated} calculated, ${(bytes.length / 1024) | 0} KiB; ${steps.join(' ')}`);
  assert.ok(result.calculated >= ROWS * (COLS - 1));
  assert.equal(result.errors.length, 0);
  assert.equal(result.cycles.length, 0);

  // The same cell count again, but tall and thin. Write-back used to look each
  // row up with a linear scan, which is invisible on 5,000 rows and took twenty
  // seconds on 100,000; a running total down one column is the shape that
  // exposes it.
  const tall = Array.from({ length: 50000 }, (_, r) => (r === 0 ? [1, '=A1+1'] : [r + 1, `=A${r + 1}+B${r}`]));
  const tallSteps = [];
  const tallBytes = buildXlsx({ sheets: [{ name: 'Tall', rows: tall }] });
  const tallResult = timed('tall recalc+write-back', tallSteps, () => recalculateWorkbook(tallBytes));
  console.log(`# sheet (tall): 100000 cells, ${tallResult.calculated} calculated; ${tallSteps.join(' ')}`);
  assert.equal(tallResult.calculated, 50000);
});

test('a 200-slide deck opens, renders and saves inside the budget', () => {
  const steps = [];
  const slides = Array.from({ length: 200 }, (_, i) => (i === 0
    ? { layout: 'title', title: 'Annual review', body: 'Finance' }
    : { layout: 'obj', title: `Slide ${i + 1}`, body: [sentence(i, 8), sentence(i + 1, 10), sentence(i + 2, 6)] }));
  const bytes = timed('build', steps, () => buildPptx({ title: 'Big deck', slides }));
  const deck = timed('open', steps, () => Deck.open(bytes));
  assert.equal(deck.slideCount, 200);
  timed('scene-all', steps, () => { for (let i = 0; i < deck.slideCount; i++) deck.slide(i); });
  timed('render-all', steps, () => { for (let i = 0; i < deck.slideCount; i++) renderSlide(deck.slide(i)); });
  timed('edit-every-slide', steps, () => {
    for (let i = 0; i < deck.slideCount; i++) {
      const shape = deck.slide(i).shapes.find((s) => s.text !== undefined || s.paragraphs);
      if (shape) deck.setText(i, shape.id, [`Edited ${i}`]);
    }
  });
  const saved = timed('save', steps, () => deck.save());
  const again = timed('reopen', steps, () => Deck.open(saved));
  console.log(`# deck: 200 slides, ${(bytes.length / 1024) | 0} KiB; ${steps.join(' ')}`);
  assert.equal(again.slideCount, 200);
});

test('marking a 3,000-message folder read and emptying it rewrites the index once, not once per message', () => {
  // The per-message form re-parsed and rewrote the folder's whole index for
  // every message: 20 s for 2,000, growing with the square of the folder.
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'perf-mail-'));
  try {
    const steps = [];
    const store = new MailStore(root);
    const messages = Array.from({ length: 3000 }, (_, i) => ({
      messageId: `<${i}@example.test>`, subject: `Subject ${i}`, from: { address: 'a@example.test', name: 'A' }, to: [],
      date: new Date(Date.UTC(2026, 0, 1, 0, 0, i)).toISOString(), text: 'hello '.repeat(10), unread: true, attachments: [],
    }));
    timed('put', steps, () => store.putMany('acc', 'INBOX', messages));
    const ids = store.list('acc', 'INBOX', { limit: 100000 }).rows.map((r) => r.id);
    assert.equal(ids.length, 3000);
    timed('mark-read', steps, () => store.setFlagsMany('acc', 'INBOX', ids, { unread: false }));
    assert.equal(store.list('acc', 'INBOX', { limit: 100000, unreadOnly: true }).total, 0);
    assert.equal(store.get('acc', 'INBOX', ids[0]).unread, false);
    timed('remove', steps, () => store.removeMany('acc', 'INBOX', ids));
    assert.equal(store.list('acc', 'INBOX', { limit: 10 }).total, 0);
    console.log(`# mail: 3000 messages; ${steps.join(' ')}`);
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});
