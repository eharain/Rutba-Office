// Worksheets in the dark theme keep a file's own cell colours, as Excel
// does, and every word in them reads. A sheet with the cases a file brings:
// a navy header with white words, a total row filled #EEF2F7 with automatic
// words (the one the screenshots caught, light words on a light fill), a
// teal figure and a navy note with no fill, black set by hand, a mid grey and
// a black fill, a conditional format's red on pink, a colour scale and a
// table style's bands. Shown in the light theme and the dark one (chosen the
// way a person's choice is kept, and a window opened on it) at 1300x700 and
// 1680x1050: every fill is the file's, every cell's words are measured
// against what is under them, and the borders show. Captures go to
// RUTBA_VERIFY_CAPTURE as sheet-dark-*.png. The fixture is a file of its
// own, in a folder of its own. Run alone with RUTBA_VERIFY_ONLY=sheet-dark.

import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { parseColour, contrast, AA, AA_LARGE, INK } from '@rutba/office-ui/contrast';

const SIZES = [[1300, 700], [1680, 1050]];

/** The fixture, and what the file says each cell looks like. */
function fixture(dir) {
  const rows = [
    ['Region', 'Q1', 'Q2', 'Q3', 'Change', null, 'Item', 'Count'],
    ['North', 1200, 1350, 1500, 0.25, null, 'Pens', 40],
    ['South', 900, 860, 1100, 0.22, null, 'Paper', 12],
    ['East', 1500, 1600, 1580, 0.05, null, 'Ink', 7],
    ['Total', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)', '=D5/B5-1'],
    [],
    ['Growth', 0.32, null, 'Note in navy'],
    ['Set black', 'by hand', null, 'Grey fill', 'Black fill'],
  ];
  const styles = {
    'A1:E1': { bold: true, fill: '#1F3A5F', colour: '#FFFFFF', border: true },
    'A5:D5': { bold: true, fill: '#EEF2F7', border: true },
    'E2:E5': { numFmt: '0%' },
    'E5:E5': { bold: true, fill: '#EEF2F7', border: true, numFmt: '0%' },
    'B7:B7': { bold: true, colour: '#0E7C66', numFmt: '0.0%' },
    'D7:D7': { colour: '#1F3864' },
    'A8:B8': { colour: '#000000' },
    'D8:D8': { fill: '#7F7F7F' },
    'E8:E8': { fill: '#000000' },
  };
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Night', rows, styles }] }), { viewportWidth: 1100, viewportHeight: 620 });
  // Over 1,000 in Q3: Excel's green on green.
  view.select(1, 3); view.select(3, 3, { extend: true });
  view.addConditionalRule({ kind: 'cellIs', operator: 'greaterThan', values: [1000], look: 'green' });
  // Under 900 in Q2: red on pink.
  view.select(1, 2); view.select(3, 2, { extend: true });
  view.addConditionalRule({ kind: 'cellIs', operator: 'lessThan', values: [900], look: 'red' });
  // The change as a colour scale: fills with automatic words.
  view.select(1, 4); view.select(3, 4, { extend: true });
  view.addConditionalRule({ kind: 'colorScale' });
  // A table, banded.
  view.select(0, 6); view.select(3, 7, { extend: true });
  view.formatAsTable({ style: 'TableStyleMedium2', stripes: true });
  view.select(9, 0);
  const own = path.join(dir, 'sheet-dark-fixtures');
  fs.mkdirSync(own, { recursive: true });
  const file = path.join(own, 'night.xlsx');
  fs.writeFileSync(file, view.save());
  // A copy for the dark window, so the two are two files and two windows.
  fs.copyFileSync(file, path.join(own, 'night-dark.xlsx'));
  return { light: file, dark: path.join(own, 'night-dark.xlsx') };
}

/** Each drawn cell: its words, the colours it is painted in, and its edges. */
const CELLS = `(() => [...document.querySelectorAll('.sh-cells .sh-cell, .sh-cell')].filter((n, i, a) => a.indexOf(n) === i).map((n) => {
  const cs = getComputedStyle(n);
  return {
    ref: n.dataset.ref, text: n.textContent.trim(), color: cs.color, background: cs.backgroundColor,
    fill: n.style.background || n.style.backgroundColor || '',
    top: cs.borderTopStyle !== 'none' ? cs.borderTopColor : null,
    bottom: cs.borderBottomStyle !== 'none' && n.style.borderBottom ? cs.borderBottomColor : null,
    surface: getComputedStyle(document.querySelector('.sh-grid') || document.body).backgroundColor,
    sel: n.classList.contains('sel') || n.classList.contains('active'),
  };
}))()`;

/** 'rgb(r, g, b)' as '#rrggbb'. */
const hexOf = (css) => { const rgb = parseColour(css); return rgb ? '#' + rgb.map((n) => n.toString(16).padStart(2, '0')).join('') : null; };

export async function verifySheetDark({ open, check, until, wait }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  let js = null;
  try {
    const files = fixture(dir);
    let win = null;
    const snap = async (name) => {
      if (!shots) return;
      win.webContents.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(shots, name), (await win.webContents.capturePage()).toPNG());
    };
    const size = async (w, h) => {
      win.setContentSize(w, h);
      await until(() => js('[innerWidth, innerHeight]').then(([iw, ih]) => Math.abs(iw - w) <= 2 && Math.abs(ih - h) <= 2), `the window at ${w}x${h}`, 5000).catch(() => {});
    };
    const byRef = (cells) => new Map(cells.map((c) => [c.ref, c]));
    const seen = {};
    const bad = { fills: [], words: [], edges: [], own: [] };
    for (const theme of ['light', 'dark']) {
      // The theme as the window's own menu keeps it, then a window opened on it.
      if (js) await js(`window.rutbaOffice.store.set({ key: 'theme', value: ${JSON.stringify(theme)} }).then(() => 1)`);
      win = await open('sheets', files[theme]);
      js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(`[...document.querySelectorAll('.sh-cell')].some((n) => n.dataset.ref === 'A5' && /Total/.test(n.textContent)) && [...document.querySelectorAll('.sh-cell')].some((n) => n.dataset.ref === 'H4')`), 'the sheet with its total row and table', 12000);
      await until(() => js(`document.documentElement.dataset.theme === ${JSON.stringify(theme)}`), `the ${theme} theme`, 5000);
      for (const [w, h] of SIZES) {
        await size(w, h);
        const at = `${theme} ${w}x${h}`;
        // Wait for the painter to have drawn this theme: the total row's words on its fill in the suite's dark ink.
        await until(async () => {
          const c = byRef((await js(CELLS)) || []).get('A5');
          return c && hexOf(c.background) === '#eef2f7' && hexOf(c.color) === INK.light;
        }, `the total row painted at ${at}`, 5000).catch(() => {});
        const cells = byRef((await js(CELLS)) || []);
        seen[at] = cells.size;
        const fileFills = { A1: '#1f3a5f', A5: '#eef2f7', B5: '#eef2f7', E5: '#eef2f7', D8: '#7f7f7f', E8: '#000000' };
        for (const [ref, fill] of Object.entries(fileFills)) {
          const c = cells.get(ref);
          if (!c) { bad.fills.push(`${at}: ${ref} not drawn`); continue; }
          if (hexOf(c.background) !== fill) bad.fills.push(`${at}: ${ref} painted ${hexOf(c.background)}, the file says ${fill}`);
        }
        for (const c of cells.values()) {
          if (!c.text || c.sel) continue;
          const under = parseColour(c.background) ? c.background : c.surface;
          const ratio = contrast(c.color, under);
          if (ratio == null || ratio < AA) bad.words.push(`${at}: ${c.ref} "${c.text}" ${hexOf(c.color)} on ${hexOf(under)} at ${ratio?.toFixed(2)}`);
        }
        // The author's own pairs on fills stay theirs.
        const header = cells.get('A1');
        if (header && hexOf(header.color) !== '#ffffff') bad.own.push(`${at}: the header's white words drawn ${hexOf(header.color)}`);
        const pink = cells.get('C3');
        if (pink && (hexOf(pink.background) !== '#ffc7ce' || hexOf(pink.color) !== '#9c0006')) bad.own.push(`${at}: the red rule drawn ${hexOf(pink.color)} on ${hexOf(pink.background)}`);
        // No fill: the teal figure is the file's teal by day, a lighter teal at night.
        const teal = cells.get('B7');
        if (teal) {
          const [r, g, b] = parseColour(teal.color) || [0, 0, 0];
          if (theme === 'light' && hexOf(teal.color) !== '#0e7c66') bad.own.push(`${at}: the teal figure drawn ${hexOf(teal.color)} by day`);
          if (theme === 'dark' && !(g > b && b > r)) bad.own.push(`${at}: the teal figure lost its hue at night (${hexOf(teal.color)})`);
        }
        // The total row's edges show against the sheet: the file's light grey
        // (#BFBFBF) is faint on white by its author's choice, plain at night.
        const total = cells.get('B5');
        if (total) {
          const edge = total.top;
          const ratio = edge ? contrast(edge, total.surface) : null;
          if (ratio == null || ratio < (theme === 'dark' ? AA_LARGE : 1.5)) bad.edges.push(`${at}: the total row's top edge ${hexOf(edge)} on ${hexOf(total.surface)} at ${ratio?.toFixed(2)}`);
        }
        await snap(`sheet-dark-${w}x${h}-${theme}.png`);
      }
    }
    check('sheet-dark: every cell the file filled keeps its own fill in the light theme and the dark one, as Excel keeps it (the header, the total row, a grey and a black fill)',
      !bad.fills.length, bad.fills.slice(0, 4).join('; ') || `cells drawn ${JSON.stringify(seen)}`);
    check('sheet-dark: every cell\'s words read on what is under them (4.5:1), at night too: automatic words on the total row\'s light fill go dark, on dark fills light, a rule\'s and a table style\'s cells included',
      !bad.words.length, bad.words.slice(0, 4).join('; ') || 'all readable, both themes, both sizes');
    check('sheet-dark: a font colour the file set on a fill is kept (white on navy, a rule\'s red on pink); with no fill a dark colour is the file\'s by day and a lighter one of the same hue at night',
      !bad.own.length, bad.own.slice(0, 4).join('; ') || 'kept and lifted as said');
    check('sheet-dark: the total row\'s borders show against the sheet in either theme', !bad.edges.length, bad.edges.slice(0, 4).join('; ') || 'shown');
  } catch (err) {
    check('sheet-dark: the checks ran', false, err.stack || err.message);
  } finally {
    // The checks' own setting back: every later window opens light.
    if (js) await js(`window.rutbaOffice.store.set({ key: 'theme', value: 'light' }).then(() => 1)`).catch(() => {});
  }
}
