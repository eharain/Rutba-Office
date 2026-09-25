// Worksheets: Insert → Scatter, View → Page Break Preview, View → Split.
//
// A scatter with smooth lines is inserted over a block of X and Y values
// and drawn with its curves and markers. On a report two pages wide, Page
// Break Preview greys what is not printed, edges the pages in blue, writes
// "Page 1" across the first, draws the paper's breaks dashed; a break
// dragged up becomes one put by hand, solid, and the status bar's Normal
// button goes back. Split at D6 divides the window with two bars; the
// wheel over the top pane scrolls it alone; a bar dragged moves the split;
// pressed again, the window is whole. Saved, the sheet view carries the
// preview and the split pane. A frame of a 60,000-row sheet, split and in
// the preview, is timed. Run alone with RUTBA_VERIFY_ONLY=views.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const XY = [
  ['Hours', 'Score', 'Target'],
  [1, 52, 50],
  [2, 58, 55],
  [4, 71, 60],
  [7, 83, 65],
  [9, 90, 70],
];
const REPORT = Array.from({ length: 150 }, (_, r) => Array.from({ length: 14 }, (_, c) => (r === 0 ? 'Col ' + (c + 1) : r * 10 + c)));

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixtures are written
 */
export async function verifyViews(h, { dir }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'views.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Data', rows: XY }, { name: 'Report', rows: REPORT }] }));
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="C6"]'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const menuItem = async (pattern) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => ${pattern}.test(b.textContent)))`), 'the menu', 4000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => ${pattern}.test(b.textContent)); if (!b) return 'no item'; b.click(); return 'clicked'; })()`);
    };
    const select = async (ref) => {
      await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="${ref}"]'))`), ref, 4000).catch(() => {});
      await js(`(() => { document.querySelector('.sh-cell[data-ref="${ref}"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
      return until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === '${ref}'`), ref + ' active', 4000).catch(() => false);
    };
    const refresh = async () => { win.webContents.invalidate(); await wait(700); };

    // ── Insert → Scatter with smooth lines.
    const onB2 = await select('B2');
    await tab('Insert');
    const opened = await clickIn('Scatter');
    const picked = await menuItem(/^Scatter with Smooth Lines and Markers$/);
    const drawn = await until(() => js(`(() => { const s = document.querySelector('.sh-drawing svg'); return Boolean(s) && s.querySelectorAll('ellipse').length === 10; })()`), 'the scatter chart', 6000).catch(() => false);
    const shape = await js(`(() => {
      const s = document.querySelector('.sh-drawing svg');
      if (!s) return null;
      const d = [...s.querySelectorAll('path')].map((p) => p.getAttribute('d') || '').filter((d) => /C/.test(d));
      const labels = [...s.querySelectorAll('text')].map((t) => t.textContent.trim());
      return { curves: d.length, dots: s.querySelectorAll('ellipse').length, labels };
    })()`);
    const spec = model().drawings?.find((d) => d.kind === 'chart');
    check('sheets: Insert → Scatter with Smooth Lines and Markers draws X against Y — a curve and five markers a series, numbers along the foot',
      onB2 === true && opened === 'clicked' && picked === 'clicked' && drawn === true && shape.curves === 2 && shape.dots === 10 && shape.labels.includes('10') && Boolean(spec),
      `${opened} ${picked}; ${JSON.stringify(shape)}`);
    await refresh();
    await capture(win, 'sheets-scatter.png');

    // ── Page Break Preview on the report.
    await js(`(() => { [...document.querySelectorAll('.sh-tab')].find((t) => t.textContent.trim() === 'Report')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="N20"]'))`), 'the report', 6000).catch(() => {});
    await tab('View');
    const pbPressed = await clickIn('Page Break Preview');
    // The preview opens at 60%: the frame for the wider view is the one with the first page's foot in it.
    const preview = await until(() => js(`Boolean(document.querySelector('.sh-pb .sh-pb-page[data-page="1"]')) && Boolean(document.querySelector('.sh-pb:not(.mirror) .sh-pb-break.row'))`), 'the preview', 6000).catch(() => false);
    const seen = await js(`(() => {
      const box = (n) => { if (!n) return null; const r = n.getBoundingClientRect(); return { x: Math.round(r.left), y: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height) }; };
      const breaks = [...document.querySelectorAll('.sh-pb-break')];
      const style = (n, pseudo) => { const s = getComputedStyle(n, pseudo); return { top: s.borderTopStyle, left: s.borderLeftStyle, colour: s.borderTopColor }; };
      return {
        pages: Number(document.querySelector('.sh-pb:not(.mirror)')?.dataset.pages || 0),
        page1: document.querySelector('.sh-pb-page[data-page="1"]')?.textContent || '',
        area: box(document.querySelector('.sh-pb-area')),
        areaBorder: getComputedStyle(document.querySelector('.sh-pb-area')).borderTopColor,
        shade: document.querySelectorAll('.sh-pb-shade').length,
        rowBreaks: breaks.filter((b) => b.classList.contains('row')).map((b) => ({ index: Number(b.dataset.index), manual: b.classList.contains('manual'), line: style(b, '::after').top })),
        colBreaks: breaks.filter((b) => b.classList.contains('col')).map((b) => ({ index: Number(b.dataset.index), manual: b.classList.contains('manual'), line: style(b, '::after').left })),
        statusOn: document.querySelector('.sh-viewbtn.on')?.dataset.view || '',
      };
    })()`);
    check('sheets: View → Page Break Preview greys what is not printed, edges the pages in blue, writes Page 1 across the first and draws the paper\'s breaks dashed',
      pbPressed === 'clicked' && preview === true && seen.pages >= 4 && seen.page1 === 'Page 1' && seen.areaBorder === 'rgb(33, 82, 200)' && seen.shade >= 2
        && seen.rowBreaks.length >= 1 && seen.rowBreaks.every((b) => !b.manual && b.line === 'dashed') && seen.colBreaks.length >= 1 && seen.statusOn === 'pageBreakPreview',
      JSON.stringify(seen));
    await refresh();
    await capture(win, 'sheets-pagebreak.png');

    // Drag the first row break up ten rows: it becomes a break put by hand.
    const first = seen.rowBreaks[0]?.index ?? 0;
    const target = Math.max(5, first - 10);
    const dragged = await js(`(() => {
      const b = document.querySelector('.sh-pb-break.row[data-index="${first}"]');
      const layer = document.querySelector('.sh-cells');
      if (!b || !layer) return 'no break';
      const r = b.getBoundingClientRect();
      const L = layer.getBoundingClientRect();
      const row = [...document.querySelectorAll('.sh-rowheads .sh-head')].find((n) => n.textContent.trim() === '${target + 1}');
      if (!row) return 'no row ${target + 1}';
      const to = row.getBoundingClientRect().top;
      const x = r.left + 40;
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: x, clientY: r.top + 4 }));
      for (let i = 1; i <= 6; i++) window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: x, clientY: r.top + 4 + ((to - r.top - 4) * i) / 6 }));
      window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: x, clientY: to }));
      return 'dragged';
    })()`);
    const moved = await until(() => js(`Boolean(document.querySelector('.sh-pb-break.row.manual[data-index="${target}"]'))`), 'the manual break', 6000).catch(() => false);
    const setup = doc.pageSetup({ id: session.id });
    const solid = await js(`(() => { const b = document.querySelector('.sh-pb-break.row.manual'); return b ? getComputedStyle(b, '::after').borderTopStyle : ''; })()`);
    check('sheets: a break dragged up in the preview becomes one put by hand — solid, and in the page setup the printer reads',
      dragged === 'dragged' && moved === true && (setup.rowBreaks || []).includes(target) && solid === 'solid',
      `${dragged}; from ${first} to ${target}; moved ${moved}; rowBreaks ${JSON.stringify(setup.rowBreaks)}; line ${solid}`);
    await refresh();
    await capture(win, 'sheets-pagebreak-manual.png');

    // ── Split at D6.
    const onD6 = await select('D6');
    const splitPressed = await clickIn('Split');
    const bars = await until(() => js(`Boolean(document.querySelector('.sh-splitbar.h')) && Boolean(document.querySelector('.sh-splitbar.v'))`), 'the split bars', 5000).catch(() => false);
    const panes = await js(`(() => ({
      top: [...document.querySelectorAll('.sh-pin-rows.split > .sh-cell')].map((c) => c.dataset.ref).filter((r) => /^[D-Z]/.test(r)).slice(0, 2),
      corner: document.querySelectorAll('.sh-pin-corner.split .sh-cell').length,
      left: document.querySelectorAll('.sh-pin-cols.split .sh-cell').length,
    }))()`);
    const s0 = model().split;
    check('sheets: View → Split at D6 divides the window in four with two bars — rows 1 to 5 over, columns A to C beside, each pane its own layer',
      onD6 === true && splitPressed === 'clicked' && bars === true && s0 && s0.rows.join(',') === '0,1,2,3,4' && s0.cols.join(',') === '0,1,2' && panes.corner > 0 && panes.left > 0,
      `${splitPressed}; bars ${bars}; split ${JSON.stringify(s0 && { ...s0, rows: s0.rows, cols: s0.cols })}; ${JSON.stringify(panes)}`);

    // The wheel over the top pane scrolls it alone.
    const mainBefore = await js(`document.querySelector('.sh-grid').scrollTop`);
    await js(`(() => {
      const cell = document.querySelector('.sh-pin-rows.split > .sh-cell');
      for (let i = 0; i < 4; i++) cell.dispatchEvent(new WheelEvent('wheel', { bubbles: true, cancelable: true, deltaY: 100 }));
      return 1;
    })()`);
    const scrolled = await until(() => (model().split?.top || 0) >= 6, 'the top pane to scroll', 5000).catch(() => false);
    const topRows = await until(() => js(`[...document.querySelectorAll('.sh-pin-rowheads.split .sh-head')].map((n) => n.textContent.trim())[0] !== '1'`), 'the top pane\'s headings', 4000).catch(() => false);
    const mainAfter = await js(`document.querySelector('.sh-grid').scrollTop`);
    check('sheets: the wheel over the top pane scrolls it alone — its rows and headings move, the main pane stays',
      scrolled === true && topRows === true && mainBefore === mainAfter,
      `top ${model().split?.top}; headings ${await js(`[...document.querySelectorAll('.sh-pin-rowheads.split .sh-head')].map((n) => n.textContent.trim()).join(',')`)}; main ${mainBefore} → ${mainAfter}`);
    await refresh();
    await capture(win, 'sheets-split.png');

    // The horizontal bar dragged down three rows.
    const h0 = model().split?.height || 0;
    await js(`(() => {
      const bar = document.querySelector('.sh-splitbar.h');
      const r = bar.getBoundingClientRect();
      const x = r.left + 300;
      // Three rows, in screen pixels at whatever zoom the grid is drawn.
      const z = parseFloat(getComputedStyle(document.querySelector('.sh-grid')).zoom) || 1;
      const d = 60 * z;
      bar.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: x, clientY: r.top + 2 }));
      window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: x, clientY: r.top + 2 + d / 2 }));
      window.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: x, clientY: r.top + 2 + d }));
      window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: x, clientY: r.top + 2 + d }));
      return 1;
    })()`);
    const grown = await until(() => (model().split?.height || 0) === h0 + 60, 'the split lower', 5000).catch(() => false);
    check('sheets: dragging the split bar down three rows makes the top pane three rows taller, on a row edge', grown === true, `${h0} → ${model().split?.height}`);

    // Saved: the preview and the split are in the sheet's view as Excel keeps them.
    doc.save({ id: session.id });
    const saved = new SheetView(fs.readFileSync(file));
    saved.selectSheet('Report');
    const part = saved.workbook.partNameFor('Report');
    const xml = saved.workbook.snapshotParts([part])[part];
    const pane = (/<pane\b[^>]*\/>/.exec(xml) ?? [''])[0];
    check('sheets: the saved sheet view says pageBreakPreview and carries the split pane — xSplit, ySplit, no frozen state — and the manual break',
      /<sheetView\b[^>]*view="pageBreakPreview"/.test(xml) && /xSplit="\d+"/.test(pane) && /ySplit="\d+"/.test(pane) && !/state=/.test(pane) && new RegExp('<brk id="' + target + '"[^>]*man="1"').test(xml),
      `${pane}; view ${/view="(\w+)"/.exec(xml)?.[1]}`);

    // Split again: whole. The status bar's Normal: the grid.
    await clickIn('Split');
    const whole = await until(() => js(`!document.querySelector('.sh-splitbar') && !document.querySelector('.sh-pin.split')`), 'the window whole', 5000).catch(() => false);
    await js(`(() => { document.querySelector('.sh-viewbtn[data-view="normal"]').click(); return 1; })()`);
    const normal = await until(() => js(`!document.querySelector('.sh-pb') && document.querySelector('.sh-viewbtn.on')?.dataset.view === 'normal'`), 'Normal', 5000).catch(() => false);
    check('sheets: Split pressed again makes the window whole, and the status bar\'s Normal leaves the preview', whole === true && normal === true && model().split === null && model().viewMode === 'normal', `whole ${whole}, normal ${normal}`);

    const complaints = await errorsIn(win);
    check('sheets: the scatter, page break preview and split checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the scatter, page break preview and split checks ran', false, err.message);
  }

  // ── A frame of a long sheet, split and in the preview, timed.
  try {
    const big = path.join(dir, 'views-60k.xlsx');
    const rows = Array.from({ length: 60000 }, (_, r) => [r === 0 ? 'Item' : 'Row ' + r, r, r * 2, r % 7, r * 0.5, 'x' + (r % 13)]);
    fs.writeFileSync(big, buildXlsx({ sheets: [{ name: 'Big', rows }] }));
    const win = await open('sheets', big);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="A2"]'))`), 'the long sheet', 20000);
    const session = sessionFor('sheet');
    // The engine's frame, in the preview and split, at the top and deep down.
    const apply = (ops) => doc.apply({ id: session.id, ops });
    apply([{ op: 'setViewMode', mode: 'pageBreakPreview' }, { op: 'select', row: 8, col: 2 }, { op: 'toggleSplit' }]);
    const times = [];
    for (const y of [0, 200000, 600000, 1100000, 300000]) {
      const t0 = process.hrtime.bigint();
      doc.viewport({ id: session.id, width: 1400, height: 800, x: 0, y });
      times.push(Number(process.hrtime.bigint() - t0) / 1e6);
    }
    const m = doc.model({ id: session.id });
    // The window's own frame: a scroll to the middle, painted.
    const painted = await js(`(async () => {
      const g = document.querySelector('.sh-grid');
      const t0 = performance.now();
      g.scrollTop = 500000;
      const want = (n) => [...document.querySelectorAll('.sh-rowheads .sh-head')].some((h) => Number(h.textContent) >= n);
      for (let i = 0; i < 200 && !want(24000); i++) await new Promise((r) => requestAnimationFrame(r));
      return { ms: Math.round(performance.now() - t0), ok: want(24000) };
    })()`);
    const worst = Math.max(...times);
    check('sheets: a frame of a 60,000-row sheet split and in Page Break Preview stays quick — the engine\'s frame and the window\'s scroll to row 25,000',
      Boolean(m.split) && m.viewMode === 'pageBreakPreview' && m.pageBreaks?.count > 100 && worst < 250 && painted.ok && painted.ms < 1500,
      `engine frames ${times.map((t) => t.toFixed(1)).join(', ')} ms; pages ${m.pageBreaks?.count}; window scroll ${painted.ms} ms`);
  } catch (err) {
    check('sheets: the long-sheet frame was timed', false, err.message);
  }
}
