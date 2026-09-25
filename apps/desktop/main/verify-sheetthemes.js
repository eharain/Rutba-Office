// Worksheets: Page Layout → Themes, Colours, Fonts and Effects.
//
// A workbook with a table, a header filled in the theme's first accent, a
// chart and a shape: Themes opens a gallery of the suite's eleven themes,
// each drawn as a little sheet; picking Ember repaints the header, the
// table, the chart and the shape in Ember's accents and the default font
// becomes Ember's body face. Colours lists the palettes, and Customise
// Colours saves a palette of one's own; Fonts lists the pairs and Customise
// Fonts saves two faces; Effects offers the four format schemes. Undo puts
// the theme back. Saved, theme1.xml says it. Run alone with
// RUTBA_VERIFY_ONLY=sheetthemes.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { THEMES } from '@rutba/office-formats/themes';

function fixture() {
  const made = new SheetView(buildXlsx({ sheets: [{
    name: 'Budget',
    rows: [['Quarterly budget', ''], ['', ''], ['Area', 'Spend'], ['Rent', 4200], ['Staff', 9100], ['Travel', 1300], ['Tools', 800]],
    styles: { 'A1:B1': { bold: true, fill: '1F4E79', colour: 'FFFFFF', size: 14 } },
  }] }));
  // The title filled with the theme's first accent, as Excel writes one.
  made.pkg.write_('xl/styles.xml', made.pkg.text('xl/styles.xml').replace('<fgColor rgb="FF1F4E79"/>', '<fgColor theme="4"/>'));
  const view = SheetView.open(made.pkg.write());
  view.select(3, 0);
  view.formatAsTable({ style: 'TableStyleMedium2', stripes: true });
  view.select(4, 1);
  view.insertChart({ kind: 'column', title: 'Spend' });
  view.select(18, 0);
  view.select(21, 2, { extend: true });
  view.insertShape({ geometry: 'roundRect', text: 'On budget' });
  return view.save();
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixtures are written
 */
export async function verifySheetThemes(h, { dir }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'themes.xlsx');
  fs.writeFileSync(file, fixture());
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.sh-drawing[data-kind="chart"] svg'))`), 'the grid and chart', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('.rw-ribbon button, .rw-tabs button, button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const clickSel = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'no ' + ${JSON.stringify(selector)}; if (b.disabled) return 'disabled'; b.click(); return 'clicked'; })()`);
    const setField = (selector, v) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set.call(el, ${JSON.stringify(v)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'set';
    })()`);
    const refresh = async () => { win.webContents.invalidate(); await wait(700); };
    const look = () => js(`(() => {
      const bg = (ref) => { const n = document.querySelector('.sh-cell[data-ref="' + ref + '"]'); return n ? getComputedStyle(n).backgroundColor : ''; };
      const font = (ref) => { const n = document.querySelector('.sh-cell[data-ref="' + ref + '"]'); return n ? getComputedStyle(n).fontFamily : ''; };
      const fills = (sel) => [...document.querySelectorAll(sel + ' svg [fill]')].map((n) => n.getAttribute('fill').toLowerCase());
      return { title: bg('A1'), header: bg('A3'), font: font('A4'), chart: fills('.sh-drawing[data-kind="chart"]'), shape: fills('.sh-drawing[data-kind="shape"]') };
    })()`);
    const rgb = (hex) => { const n = parseInt(hex, 16); return `rgb(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255})`; };

    // ── Themes: the gallery of the suite's eleven, each a little sheet.
    const before = await look();
    await tab('Page Layout');
    const opened = await clickIn('Themes');
    const cards = await until(() => js(`document.querySelectorAll('.sh-dg .sl-dg-card[data-theme]').length === 11 && document.querySelectorAll('.sh-dg .sl-dg-card svg').length === 11`), 'the Themes gallery', 5000).catch(() => false);
    await refresh();
    await capture(win, 'sheets-themes-gallery.png');
    await clickSel('.sh-dg .sl-dg-card[data-theme="ember"]');
    const ember = THEMES.find((t) => t.id === 'ember');
    const changed = await until(async () => (await look()).title === rgb(ember.palette.accent1), 'the title in Ember', 8000).catch(() => false);
    const after = await look();
    const accent = '#' + ember.palette.accent1.toLowerCase();
    check('sheets: Page Layout → Themes shows the suite\'s eleven themes as little sheets; Ember repaints the theme-filled title, the table, the chart and the shape, and the default font becomes its body face',
      opened === 'clicked' && cards === true && changed === true && after.header !== before.header
        && after.chart.includes(accent) && after.shape.includes(accent) && /Segoe UI/.test(after.font) && model().design?.builtIn === 'ember',
      `${opened}; cards ${cards}; ${JSON.stringify({ before: [before.title, before.header, before.font], after: [after.title, after.header, after.font] })}; chart ${after.chart.includes(accent)}; shape ${after.shape.includes(accent)}`);
    await refresh();
    await capture(win, 'sheets-theme-ember.png');

    // ── Colours: a palette; Customise Colours saves one's own.
    await clickIn('Colours');
    await until(() => js(`document.querySelectorAll('.sh-dg .sl-dg-row[data-palette]').length >= 11`), 'the Colours list', 5000).catch(() => {});
    await refresh();
    await capture(win, 'sheets-theme-colours.png');
    await clickSel('.sh-dg .sl-dg-row[data-palette="lagoon"]');
    const lagoon = await until(() => model().design?.colorName === 'Lagoon', 'Lagoon', 6000).catch(() => false);
    await clickIn('Colours');
    await until(() => js(`Boolean(document.querySelector('.sh-dg .sl-dg-custom'))`), 'Customise', 4000).catch(() => {});
    await clickSel('.sh-dg .sl-dg-custom');
    await until(() => js(`Boolean(document.querySelector('.sl-cc-hex-accent1'))`), 'Customise Colours', 4000).catch(() => {});
    await setField('.sl-cc-hex-accent1', '#B83280');
    await setField('.sl-cc-namefield', 'House colours');
    await wait(200);
    await refresh();
    await capture(win, 'sheets-theme-customise-colours.png');
    await clickSel('.sl-cc-save');
    const own = await until(() => model().design?.colorName === 'House colours' && model().design?.colors?.accent1 === 'B83280', 'the own palette', 6000).catch(() => false);
    const titleOwn = await until(async () => (await look()).title === rgb('B83280'), 'the title in the own accent', 6000).catch(() => false);
    check('sheets: Colours applies a palette, and Customise Colours saves one\'s own — the theme-filled title follows',
      lagoon === true && own === true && titleOwn === true, `lagoon ${lagoon}; own ${own}; title ${titleOwn}`);

    // ── Fonts: a pair; Customise Fonts two faces of one's own.
    await clickIn('Fonts');
    await until(() => js(`document.querySelectorAll('.sh-dg .sl-dg-row[data-pair]').length === 11`), 'the Fonts list', 5000).catch(() => {});
    await clickSel('.sh-dg .sl-dg-row[data-pair="classic"]');
    const classic = await until(async () => /Georgia/.test((await look()).font), 'Georgia', 6000).catch(() => false);
    await clickIn('Fonts');
    await until(() => js(`Boolean(document.querySelector('.sh-dg .sl-dg-custom'))`), 'Customise Fonts', 4000).catch(() => {});
    await clickSel('.sh-dg .sl-dg-custom');
    await until(() => js(`Boolean(document.querySelector('.sl-cf-minor'))`), 'the fonts dialog', 4000).catch(() => {});
    await setField('.sl-cf-major', 'Bahnschrift');
    await setField('.sl-cf-minor', 'Verdana');
    await setField('.sl-cf-name', 'Clear');
    await wait(200);
    await clickSel('.sl-cf-save');
    const verdana = await until(async () => /Verdana/.test((await look()).font) && model().design?.fontName === 'Clear', 'Verdana', 6000).catch(() => false);
    check('sheets: Fonts applies a pair and Customise Fonts two faces of one\'s own — the default font follows the body face', classic === true && verdana === true, `classic ${classic}; verdana ${verdana}`);

    // ── Effects.
    await clickIn('Effects');
    await until(() => js(`document.querySelectorAll('.sh-dg .sl-dg-card[data-effects]').length === 4`), 'the Effects gallery', 5000).catch(() => {});
    await refresh();
    await capture(win, 'sheets-theme-effects.png');
    await clickSel('.sh-dg .sl-dg-card[data-effects="outline"]');
    const outlined = await until(() => model().design?.effects === 'outline', 'Outline', 6000).catch(() => false);
    check('sheets: Effects offers the four format schemes and applies one', outlined === true, `outline ${outlined}`);

    // ── Undo steps back through them; saved, the theme part says it.
    await js(`(() => { document.querySelector('.sh').focus(); return 1; })()`);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Z', modifiers: ['control'] });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Z', modifiers: ['control'] });
    // Ember's own format scheme is Lifted: Undo puts it back.
    const undone = await until(() => model().design?.effects === 'lifted', 'undo', 6000).catch(() => false);
    doc.save({ id: session.id });
    const saved = new SheetView(fs.readFileSync(file));
    const xml = saved.pkg.text('xl/theme/theme1.xml');
    check('sheets: Undo takes the last theme change back; saved, theme1.xml carries Ember\'s name, the own colours and the own fonts',
      undone === true && /<a:theme\b[^>]*name="Ember"/.test(xml) && /<a:clrScheme name="House colours">/.test(xml) && /<a:fontScheme name="Clear">/.test(xml),
      `undone ${undone}; ${/<a:theme\b[^>]*>/.exec(xml)?.[0]} ${/<a:clrScheme[^>]*>/.exec(xml)?.[0]} ${/<a:fontScheme[^>]*>/.exec(xml)?.[0]}`);
    await refresh();
    await capture(win, 'sheets-theme-final.png');

    const complaints = await errorsIn(win);
    check('sheets: the theme checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the theme checks ran', false, err.message);
  }
}
