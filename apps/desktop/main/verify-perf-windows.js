// The performance budget, in the windows: a 300-page document, a
// 100,000-row workbook and a 200-slide deck each open, and are scrolled to
// their end, within a stated time, their window held under a stated memory.
// The engine's own budget (tests/perf-budget.perf.js) times the work behind
// a window; this times the whole of what a window does to open and scroll,
// as the processor time Office's own processes spent (D-2026-11-13-20), the
// clock beside it. The numbers
// are generous, as the engine's are: they catch a step gone quadratic, which
// shows as minutes and gigabytes, not a few percent. Each check prints what
// it measured, so a drift shows long before it trips. Run alone with
// RUTBA_VERIFY_ONLY=perfwindows.
import fs from 'node:fs';
import path from 'node:path';
import { app } from 'electron';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import { cpuSince, machineNow, costWords } from './verify-cost.js';
import { SIGNAL_WAIT } from './verify-waits.js';

/**
 * What a window may spend, and hold. Since D-2026-11-13-20 (the owner: "Yes,
 * own processor time (Recommended)") the open and the scroll are judged on
 * the processor time Office's own processes spent on them, with the clock
 * beside it in the detail: beside other work the clock measured that work
 * too. The numbers are as before.
 */
export const WINDOW_BUDGET = { openMs: 20000, scrollMs: 8000, memoryMb: 1500 };

/** How long each wait may run by the clock: long, so the judgement is the processor time's (D-2026-11-13-20). */
const OUTER = (ms) => Math.max(SIGNAL_WAIT, 3 * ms);

const WORDS = 'the quick brown fox jumps over a lazy dog while ledgers balance and totals agree'.split(' ');
const sentence = (seed, n) => Array.from({ length: n }, (_, i) => WORDS[(seed * 7 + i * 3) % WORDS.length]).join(' ');

/** A window's renderer, in megabytes of working set. */
function memoryOf(win) {
  const pid = win.webContents.getOSProcessId();
  const metric = app.getAppMetrics().find((m) => m.pid === pid);
  return metric ? Math.round(metric.memory.workingSetSize / 1024) : null;
}

/** The scrollable element that holds most of the page. */
const SCROLLER = `[...document.querySelectorAll('*')].filter((el) => el.scrollHeight > el.clientHeight + 400 && /auto|scroll/.test(getComputedStyle(el).overflowY)).sort((a, b) => b.scrollHeight - a.scrollHeight)[0]`;

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the files are written
 */
export async function verifyPerfWindows({ open, check, until, wait, press, errorsIn }, { dir }) {
  const B = WINDOW_BUDGET;
  // A document of 900 paragraphs, each a third of a page: about 300 pages.
  try {
    const file = path.join(dir, 'perf-300-pages.docx');
    fs.writeFileSync(file, buildDocx({ paragraphs: Array.from({ length: 900 }, (_, i) => ({ runs: [{ text: `${i + 1}. ${sentence(i, 100)} ` }, { text: sentence(i + 1, 60) + ' ', bold: i % 3 === 0 }, { text: sentence(i + 2, 100), italic: i % 5 === 0 }] })) }));
    const openCost = cpuSince();
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const pagesShown = () => js(`(() => { const m = /of\\s+([\\d,]+)/.exec([...document.querySelectorAll('.rw-status *, .rw-statusbar *')].map((n) => n.textContent).find((t) => /Page \\d+ of/.test(t)) || ''); return m ? Number(m[1].replace(/,/g, '')) : 0; })()`);
    let seen = 0;
    const opened = await until(async () => { const n = await pagesShown(); const settled = n >= 250 && n === seen; seen = n; return settled; }, 'the pages laid out', OUTER(B.openMs)).then(() => true).catch(() => false);
    const workPid = win.webContents.getOSProcessId();
    const openUsed = openCost.end({ need: [workPid] });
    const pages = await pagesShown();
    const scrollCost = cpuSince();
    await js(`(() => { const s = ${SCROLLER}; if (s) s.scrollTop = s.scrollHeight; return Boolean(s); })()`);
    // At the end: the scroll run to its bottom, and the last paragraph laid out on the last page.
    const reached = await until(() => js(`(() => { const s = ${SCROLLER}; const b = document.querySelector('.wd-page [data-block="899"]'); return Boolean(s && b && b.offsetHeight > 0 && s.scrollTop + s.clientHeight >= s.scrollHeight - 2); })()`), 'the last paragraph', OUTER(B.scrollMs)).then(() => true).catch(() => false);
    const scrollUsed = scrollCost.end({ need: [workPid] });
    const machine = await machineNow();
    const mb = memoryOf(win);
    check(`perf: a 300-page document opens within ${Math.round(B.openMs / 1000)} s and scrolls to its end within ${Math.round(B.scrollMs / 1000)} s of Office's own processor time, its window under ${B.memoryMb} MB`,
      opened && reached && openUsed.measured && scrollUsed.measured && openUsed.cpuMs < B.openMs && scrollUsed.cpuMs < B.scrollMs && mb != null && mb < B.memoryMb,
      `${pages} pages; open: ${costWords(openUsed)}; scroll: ${costWords(scrollUsed, machine)}; ${mb} MB`);
    const complaints = await errorsIn(win);
    check('perf: the long document reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('perf: the long document checks ran', false, err?.message || String(err));
  }

  // A workbook of 100,000 rows and four columns.
  try {
    const file = path.join(dir, 'perf-100k-rows.xlsx');
    const rows = [['Item', 'Region', 'Units', 'Price']];
    for (let i = 1; i <= 100000; i++) rows.push([`Item ${i}`, ['East', 'West', 'North', 'South'][i % 4], i % 97, (i % 50) + 0.5]);
    fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Data', rows }] }));
    const openCost = cpuSince();
    const win = await open('sheets', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const opened = await until(() => js(`document.body.innerText.includes('Item 1') && document.body.innerText.includes('Region')`), 'the first rows', OUTER(B.openMs)).then(() => true).catch(() => false);
    const workPid = win.webContents.getOSProcessId();
    const openUsed = openCost.end({ need: [workPid] });
    const scrollCost = cpuSince();
    await js(`document.querySelector('.sh-grid, .sh-sheet, canvas, body')?.focus?.(), 1`);
    await press(wc, 'End', { modifiers: ['control'] });
    // The name box says the last cell, and the grid has drawn its row.
    const reached = await until(() => js(`document.querySelector('.sh-namebox')?.textContent.trim() === 'D100001' && document.body.innerText.includes('Item 100000')`), 'the last row', OUTER(B.scrollMs)).then(() => true).catch(() => false);
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'perf-sheet-end.png'), (await wc.capturePage()).toPNG());
    const scrollUsed = scrollCost.end({ need: [workPid] });
    const machine = await machineNow();
    const mb = memoryOf(win);
    check(`perf: a 100,000-row workbook opens within ${Math.round(B.openMs / 1000)} s and Ctrl+End reaches its last row within ${Math.round(B.scrollMs / 1000)} s of Office's own processor time, its window under ${B.memoryMb} MB`,
      opened && reached && openUsed.measured && scrollUsed.measured && openUsed.cpuMs < B.openMs && scrollUsed.cpuMs < B.scrollMs && mb != null && mb < B.memoryMb,
      `open: ${costWords(openUsed)}; scroll: ${costWords(scrollUsed, machine)}; ${mb} MB`);
    const complaints = await errorsIn(win);
    check('perf: the large workbook reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('perf: the large workbook checks ran', false, err?.message || String(err));
  }

  // A deck of 200 slides.
  try {
    const file = path.join(dir, 'perf-200-slides.pptx');
    fs.writeFileSync(file, buildPptx({ title: 'Long deck', slides: Array.from({ length: 200 }, (_, i) => ({ layout: 'obj', title: `Slide ${i + 1}`, body: sentence(i, 30) })) }));
    const openCost = cpuSince();
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const opened = await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 200 && Boolean(document.querySelector('.sl-svg, .sl-stage svg'))`), 'the slides', OUTER(B.openMs)).then(() => true).catch(() => false);
    const workPid = win.webContents.getOSProcessId();
    const openUsed = openCost.end({ need: [workPid] });
    const scrollCost = cpuSince();
    await js(`(() => { const last = [...document.querySelectorAll('.sl-thumb')].pop(); last?.scrollIntoView({ block: 'end' }); return Boolean(last); })()`);
    const reached = await until(() => js(`(() => { const last = [...document.querySelectorAll('.sl-thumb')].pop(); if (!last) return false; const r = last.getBoundingClientRect(); return r.bottom > 0 && r.top < innerHeight && Boolean(last.querySelector('svg, img')); })()`), 'the last slide drawn', OUTER(B.scrollMs)).then(() => true).catch(() => false);
    const scrollUsed = scrollCost.end({ need: [workPid] });
    const machine = await machineNow();
    const mb = memoryOf(win);
    check(`perf: a 200-slide deck opens within ${Math.round(B.openMs / 1000)} s and its last slide is drawn in the strip within ${Math.round(B.scrollMs / 1000)} s of Office's own processor time, its window under ${B.memoryMb} MB`,
      opened && reached && openUsed.measured && scrollUsed.measured && openUsed.cpuMs < B.openMs && scrollUsed.cpuMs < B.scrollMs && mb != null && mb < B.memoryMb,
      `open: ${costWords(openUsed)}; scroll: ${costWords(scrollUsed, machine)}; ${mb} MB`);
    const complaints = await errorsIn(win);
    check('perf: the long deck reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('perf: the long deck checks ran', false, err?.message || String(err));
  }
  void wait;
}
