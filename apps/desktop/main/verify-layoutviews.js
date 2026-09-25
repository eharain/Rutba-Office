// Worksheets: View → Page Layout, View → Custom Views, Page Layout → Background.
//
// A report opens in Page Layout: the sheet on white pages on a grey desk,
// A1 inside page 1's margins, rulers above and beside the pages, "Click to
// add header" in the margin and "Click to add data" on the blank page past
// the report; the page count is the printer's. A click on the header opens
// Header & Footer, and the header prints on each page with its number. The
// ruler goes off and on from the View tab. Saved, the sheet view says
// pageLayout. Custom Views: rows grouped and folded, a view added from the
// dialog, the rows opened and the zoom changed, Show folds them and puts the
// zoom back; in a workbook with a table the button is greyed with Excel's
// reason. Background: a picture tiled behind the cells, the Delete
// Background button takes it away. Run alone with RUTBA_VERIFY_ONLY=layoutviews.
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { buildXlsx } from '@rutba/ooxml/build';
import { crc32 } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { readPageSetup } from '@rutba/sheet-view/print';

const REGIONS = ['North', 'South', 'East', 'West'];
const REPS = ['Amira', 'Ben', 'Chen', 'Dana', 'Eli', 'Farah'];
const REPORT = [['Region', 'Rep', 'Month', 'Units', 'Price', 'Revenue', 'Cost', 'Margin', 'Target', 'Status']];
for (let i = 1; i <= 140; i++) {
  const units = 20 + ((i * 37) % 90);
  const price = 12 + (i % 7) * 1.5;
  const revenue = Math.round(units * price * 100) / 100;
  const cost = Math.round(revenue * (0.55 + (i % 5) * 0.04) * 100) / 100;
  REPORT.push([REGIONS[i % 4], REPS[i % 6], ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'][i % 6], units, price, revenue, cost, Math.round((revenue - cost) * 100) / 100, 1200, revenue - cost > 400 ? 'On track' : 'Watch']);
}

/** A soft paper-grain PNG, drawn here so a check needs no file of its own. */
function patternPng(size = 48) {
  const raw = Buffer.alloc(size * (size * 3 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 3 + 1)] = 0;
    for (let x = 0; x < size; x++) {
      const i = y * (size * 3 + 1) + 1 + x * 3;
      const d = (x + y) % 24 < 2 || (x - y + 48) % 24 < 2;
      raw[i] = d ? 214 : 236;
      raw[i + 1] = d ? 228 : 244;
      raw[i + 2] = d ? 246 : 252;
    }
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body) >>> 0);
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 2; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixtures are written
 */
export async function verifyLayoutViews(h, { dir }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'layoutviews.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{
    name: 'Report',
    rows: REPORT,
    styles: { 'A1:J1': { bold: true, fill: '1F4E79', colour: 'FFFFFF' }, 'E2:H141': { numFmt: '#,##0.00' }, 'I2:I141': { numFmt: '#,##0' } },
  }] }));
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    const apply = (ops) => doc.apply({ id: session.id, ops });
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="J2"]'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const setField = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      setter.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'set';
    })()`);
    const clickSel = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'no ' + ${JSON.stringify(selector)}; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const refresh = async () => { win.webContents.invalidate(); await wait(700); };
    const nudge = () => js(`(() => { const g = document.querySelector('.sh-grid'); g.scrollTop += 1; g.scrollTop -= 1; g.dispatchEvent(new Event('scroll')); return 1; })()`);

    // ── View → Page Layout.
    await tab('View');
    const pressed = await clickIn('Page Layout');
    const paper = await until(() => js(`document.querySelectorAll('.sh-pl-page').length >= 2 && Boolean(document.querySelector('.sh-pl-page[data-page="1"]')) && Boolean(document.querySelector('.sh-pl-page[data-page="1"] ~ .sh-cell, .sh-cell[data-ref="A1"]'))`), 'the pages', 8000).catch(() => false);
    await wait(400);
    const seen = await js(`(() => {
      const r = (n) => { if (!n) return null; const b = n.getBoundingClientRect(); return { x: Math.round(b.left), y: Math.round(b.top), w: Math.round(b.width), h: Math.round(b.height) }; };
      const p1 = document.querySelector('.sh-pl-page[data-page="1"]');
      const box = p1?.querySelector('.sh-pl-box');
      return {
        page1: r(p1), box: r(box), a1: r(document.querySelector('.sh-cell[data-ref="A1"]')),
        pages: Number(document.querySelector('.sh-pl')?.dataset.pages || 0),
        blank: [...document.querySelectorAll('.sh-pl-page.blank .sh-pl-blank')].map((n) => n.textContent.trim())[0] || '',
        ask: p1?.querySelector('.sh-pl-zone.head .ask')?.textContent || '',
        rulers: document.querySelectorAll('.sh-pl-ruler').length,
        desk: getComputedStyle(document.querySelector('.sh-grid')).backgroundColor,
        pageColour: p1 ? getComputedStyle(p1).backgroundColor : '',
        status: document.querySelector('.sh-viewbtn.on')?.dataset.view || '',
      };
    })()`);
    const summary = doc.printSummary({ id: session.id, options: readPageSetup(new SheetView(fs.readFileSync(file)), 'Report') });
    const inside = seen.a1 && seen.box && seen.a1.x >= seen.box.x - 1 && seen.a1.y >= seen.box.y - 1 && seen.a1.x < seen.box.x + 4 && seen.a1.y < seen.box.y + 4;
    check('sheets: View → Page Layout draws the sheet on white pages on a grey desk — A1 at the top left inside page 1\'s margins, rulers, "Click to add header", a blank page saying "Click to add data", the status bar\'s Page Layout on',
      pressed === 'clicked' && paper === true && inside && seen.rulers >= 2 && seen.ask === 'Click to add header' && seen.blank === 'Click to add data' && seen.desk !== seen.pageColour && seen.status === 'pageLayout',
      JSON.stringify(seen));
    check('sheets: Page Layout\'s pages are the printer\'s — the same number as the print dialog counts',
      seen.pages === summary.pages && model().pageLayout.count === summary.pages, `layout ${seen.pages}, print ${summary.pages}`);
    await refresh();
    await capture(win, 'sheets-page-layout.png');

    // A click on the header opens Header & Footer; the header prints on each page with its number.
    await js(`(() => { document.querySelector('.sh-pl-page[data-page="1"] .sh-pl-zone.head').click(); return 1; })()`);
    const hfOpen = await until(() => js(`Boolean(document.querySelector('.sh-hf-header'))`), 'Header & Footer', 5000).catch(() => false);
    await setField('.sh-hf-header', '&LSales report&RPage &P of &N');
    await clickSel('.sh-hf-ok');
    const headed = await until(() => js(`(document.querySelector('.sh-pl-page[data-page="1"] .sh-pl-zone.head.set')?.textContent || '').includes('Page 1 of')`), 'the header on page 1', 8000).catch(() => false);
    const headText = await js(`(() => { const z = document.querySelector('.sh-pl-page[data-page="1"] .sh-pl-zone.head.set'); return z ? [...z.querySelectorAll('span')].map((s) => s.textContent) : []; })()`);
    const a1Moved = await js(`(() => { const a = document.querySelector('.sh-cell[data-ref="A1"]').getBoundingClientRect(); const z = document.querySelector('.sh-pl-page[data-page="1"] .sh-pl-zone.head.set').getBoundingClientRect(); return a.top >= z.bottom - 1; })()`);
    check('sheets: a click on "Click to add header" opens Header & Footer; the header then sits on each page with its own number, the cells under it as they print',
      hfOpen === true && headed === true && headText[0] === 'Sales report' && /^Page 1 of \d+$/.test(headText[2]) && a1Moved === true,
      `${hfOpen} ${headed}; ${JSON.stringify(headText)}; cells under ${a1Moved}`);
    await refresh();
    await capture(win, 'sheets-page-layout-header.png');

    // View → Ruler off and on.
    await tab('View');
    await clickIn('Ruler');
    const off = await until(() => js(`document.querySelectorAll('.sh-pl-ruler').length === 0`), 'the rulers gone', 5000).catch(() => false);
    await clickIn('Ruler');
    const on = await until(() => js(`document.querySelectorAll('.sh-pl-ruler').length >= 2`), 'the rulers back', 5000).catch(() => false);
    check('sheets: View → Ruler takes the rulers off the pages and puts them back', off === true && on === true, `off ${off}, on ${on}`);

    // Scrolled to page 2: its first row sits at the top of its printable box.
    const L = model().pageLayout;
    await js(`(() => { document.querySelector('.sh-grid').scrollTop = ${Math.round(L.pageHeight + 30)}; return 1; })()`);
    const page2 = await until(() => js(`(() => {
      const p = document.querySelector('.sh-pl-page[data-page="2"]');
      if (!p) return false;
      const top = p.querySelector('.sh-pl-box').getBoundingClientRect().top;
      return [...document.querySelectorAll('.sh-rowheads .sh-head')].some((n) => Math.abs(n.getBoundingClientRect().top - top) < 30);
    })()`), 'page 2 and its rows', 8000).catch(() => false);
    const p2 = await js(`(() => {
      const p = document.querySelector('.sh-pl-page[data-page="2"]');
      const box = p.querySelector('.sh-pl-box').getBoundingClientRect();
      const heads = [...document.querySelectorAll('.sh-rowheads .sh-head')].map((n) => ({ label: n.textContent.trim(), top: n.getBoundingClientRect().top }));
      const first = heads.filter((h) => h.top >= box.top - 1).sort((a, b) => a.top - b.top)[0];
      return { first: first?.label, gap: first ? Math.round(first.top - box.top) : null };
    })()`);
    doc.save({ id: session.id });
    const savedXml = (() => { const v = new SheetView(fs.readFileSync(file)); const part = v.workbook.partNameFor('Report'); return v.workbook.snapshotParts([part])[part]; })();
    check('sheets: page 2 begins with the row the printer starts it with, and the saved sheet view says pageLayout',
      page2 === true && Number(p2.first) > 30 && p2.gap !== null && p2.gap <= 24 && /<sheetView\b[^>]*view="pageLayout"/.test(savedXml),
      `${JSON.stringify(p2)}; ${/<sheetView\b[^>]*>/.exec(savedXml)?.[0]}`);
    await js(`(() => { document.querySelector('.sh-grid').scrollTop = 0; return 1; })()`);

    // ── View → Custom Views: rows 4 to 6 grouped and folded, then a view added.
    await clickIn('Normal');
    await until(() => js(`!document.querySelector('.sh-pl')`), 'Normal', 5000).catch(() => {});
    apply([{ op: 'group', axis: 'row', from: 3, to: 5 }, { op: 'outlineToggle', axis: 'row', level: 1, start: 3 }]);
    await js(`(() => { document.querySelector('.sh').focus(); return 1; })()`);
    await nudge();
    await until(() => js(`![...document.querySelectorAll('.sh-rowheads .sh-head')].some((n) => n.textContent.trim() === '5')`), 'rows 4 to 6 folded', 6000).catch(() => {});
    const folded = model().rows.every((r) => r.index < 3 || r.index > 5);
    await tab('View');
    await clickIn('Custom Views');
    await until(() => js(`Boolean(document.querySelector('.sh-cviews-add'))`), 'the Custom Views dialog', 4000).catch(() => {});
    await clickSel('.sh-cviews-add');
    await until(() => js(`Boolean(document.querySelector('.sh-cview-name'))`), 'Add View', 4000).catch(() => {});
    await setField('.sh-cview-name', 'Summary rows');
    await refresh();
    await capture(win, 'sheets-custom-views-add.png');
    await clickSel('.sh-cview-ok');
    const listed = await until(() => js(`Boolean(document.querySelector('.sh-ranges-row[data-view="Summary rows"]'))`), 'the view listed', 6000).catch(() => false);
    await refresh();
    await capture(win, 'sheets-custom-views.png');
    await clickSel('.sh-cviews-close');
    check('sheets: View → Custom Views → Add keeps the way the workbook looks under a name, listed in the dialog',
      folded && listed === true && model().customViews.some((v) => v.name === 'Summary rows' && v.hiddenRowCol && v.printSettings), `folded ${folded}; listed ${listed}`);

    // The rows opened and the zoom changed; Show puts both back.
    apply([{ op: 'outlineToggle', axis: 'row', level: 1, start: 3 }]);
    await js(`(() => { window.dispatchEvent(new Event('focus')); return 1; })()`);
    await nudge();
    await until(() => js(`[...document.querySelectorAll('.sh-rowheads .sh-head')].some((n) => n.textContent.trim() === '5')`), 'rows open again', 6000).catch(() => {});
    await clickIn('Zoom');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === '150%'))`), 'the zoom menu', 3000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === '150%').click(); return 1; })()`);
    await until(() => js(`Math.abs((parseFloat(getComputedStyle(document.querySelector('.sh-grid')).zoom) || 1) - 1.5) < 0.01`), 'zoom 150%', 4000).catch(() => {});
    await clickIn('Custom Views');
    await until(() => js(`Boolean(document.querySelector('.sh-ranges-row[data-view="Summary rows"]'))`), 'the dialog', 4000).catch(() => {});
    await js(`(() => { document.querySelector('.sh-ranges-row[data-view="Summary rows"]').click(); return 1; })()`);
    await clickSel('.sh-cviews-show');
    const refolded = await until(() => js(`![...document.querySelectorAll('.sh-rowheads .sh-head')].some((n) => n.textContent.trim() === '5') && [...document.querySelectorAll('.sh-rowheads .sh-head')].some((n) => n.textContent.trim() === '7')`), 'rows folded again', 8000).catch(() => false);
    const zoomBack = await until(() => js(`Math.abs((parseFloat(getComputedStyle(document.querySelector('.sh-grid')).zoom) || 1) - 1) < 0.01`), 'the zoom back', 5000).catch(() => false);
    doc.save({ id: session.id });
    const wbXml = (() => { const v = new SheetView(fs.readFileSync(file)); return v.pkg.text('xl/workbook.xml'); })();
    check('sheets: Show puts the view back — rows 4 to 6 folded again and the zoom as it was kept — and the file carries customWorkbookView, customSheetView and Excel\'s .wvu.Rows name',
      refolded === true && zoomBack === true && /<customWorkbookView name="Summary rows" guid="\{[0-9A-F-]{36}\}"/.test(wbXml) && /_\.wvu\.Rows" localSheetId="0" hidden="1">Report!\$4:\$6</.test(wbXml),
      `refolded ${refolded}; zoom ${zoomBack}; ${(/<customWorkbookViews>[\s\S]*?<\/customWorkbookViews>/.exec(wbXml) || [''])[0].slice(0, 160)}`);

    // ── Page Layout → Background: a picture behind the cells, then Delete Background.
    apply([{ op: 'setBackground', contentType: 'image/png', data: patternPng() }]);
    await nudge();
    const bg = await until(() => js(`/rutba:\\/\\/blob\\//.test(document.querySelector('.sh-cells')?.style.backgroundImage || '') && document.querySelector('.sh')?.classList.contains('has-bg')`), 'the background drawn', 8000).catch(() => false);
    const through = await js(`getComputedStyle(document.querySelector('.sh-cell[data-ref="B3"]')).backgroundColor`);
    const header = await js(`getComputedStyle(document.querySelector('.sh-cell[data-ref="A1"]')).backgroundColor`);
    await tab('Page Layout');
    const labelled = await until(() => js(`Boolean([...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Delete Background'))`), 'Delete Background', 4000).catch(() => false);
    await refresh();
    await capture(win, 'sheets-background.png');
    check('sheets: Page Layout → Background tiles a picture behind the cells — a cell with no fill lets it through, the header keeps its fill — and the button becomes Delete Background',
      bg === true && /rgba\(0, 0, 0, 0\)|transparent/.test(through) && !/rgba\(0, 0, 0, 0\)/.test(header) && labelled === true,
      `bg ${bg}; B3 ${through}; A1 ${header}; label ${labelled}`);
    await clickIn('Delete Background');
    const gone = await until(() => js(`!document.querySelector('.sh-cells')?.style.backgroundImage && !document.querySelector('.sh.has-bg')`), 'the background gone', 6000).catch(() => false);
    check('sheets: Delete Background takes the picture away', gone === true && model().background === null, `gone ${gone}`);

    const complaints = await errorsIn(win);
    check('sheets: the page layout, custom views and background checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the page layout, custom views and background checks ran', false, err.message);
  }

  // ── Custom Views greyed in a workbook with a table, with Excel's reason.
  try {
    const tfile = path.join(dir, 'layoutviews-table.xlsx');
    fs.writeFileSync(tfile, buildXlsx({ sheets: [{ name: 'List', rows: REPORT.slice(0, 12) }] }));
    const win = await open('sheets', tfile);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="C4"]'))`), 'the list', 8000);
    doc.apply({ id: session.id, ops: [{ op: 'select', row: 0, col: 0 }, { op: 'select', row: 11, col: 9, extend: true }, { op: 'formatAsTable', style: 'TableStyleMedium2' }] });
    await js(`(() => { const g = document.querySelector('.sh-grid'); g.scrollTop += 1; g.dispatchEvent(new Event('scroll')); return 1; })()`);
    await js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === 'View')?.click(); return 1; })()`);
    const greyed = await until(() => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Custom Views'); return b && b.disabled ? b.dataset.tip : ''; })()`), 'Custom Views greyed', 6000).catch(() => '');
    const tip = await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'Custom Views'); return b ? (b.dataset.tip || '') : ''; })()`);
    check('sheets: in a workbook with a table View → Custom Views is greyed, its tip giving Excel\'s reason', greyed && /^Custom Views — not available in a workbook that contains a table/.test(tip), tip);
  } catch (err) {
    check('sheets: the Custom Views table check ran', false, err.message);
  }

  // ── A frame of a 60,000-row sheet in Page Layout, timed.
  try {
    const big = path.join(dir, 'layoutviews-60k.xlsx');
    const rows = Array.from({ length: 60000 }, (_, r) => [r === 0 ? 'Item' : 'Row ' + r, r, r * 2, r % 7, r * 0.5, 'x' + (r % 13)]);
    fs.writeFileSync(big, buildXlsx({ sheets: [{ name: 'Big', rows }] }));
    const win = await open('sheets', big);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="A2"]'))`), 'the long sheet', 20000);
    const session = sessionFor('sheet');
    doc.apply({ id: session.id, ops: [{ op: 'setViewMode', mode: 'pageLayout' }] });
    const times = [];
    for (const y of [0, 300000, 900000, 1500000, 200000]) {
      const t0 = process.hrtime.bigint();
      doc.viewport({ id: session.id, width: 1400, height: 800, x: 0, y });
      times.push(Number(process.hrtime.bigint() - t0) / 1e6);
    }
    const m = doc.model({ id: session.id });
    const worst = Math.max(...times.slice(1));
    check('sheets: a frame of a 60,000-row sheet in Page Layout stays quick, at the top and deep down',
      m.viewMode === 'pageLayout' && m.pageLayout?.count > 1000 && worst < 250, `frames ${times.map((t) => t.toFixed(1)).join(', ')} ms; pages ${m.pageLayout?.count}`);
    doc.apply({ id: session.id, ops: [{ op: 'setViewMode', mode: 'normal' }] });
  } catch (err) {
    check('sheets: the long-sheet page layout frame was timed', false, err.message);
  }
}
