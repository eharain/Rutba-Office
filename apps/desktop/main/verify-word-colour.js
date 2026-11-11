// Documents: Font colour and Text highlight colour, as Word's. Each is a
// split button: its arrow opens the palette (the document theme's ten
// colours with five lighter and darker steps under each, the ten standard
// colours, Automatic, the recent colours and More colours), its face puts on
// the colour last chosen. Checked on chosen words, at the caret for the words
// typed next, in a table cell, by a hex typed in More colours, and back to
// Automatic; the highlight's fifteen colours likewise; then the saved file
// carries them and opens with them. Run alone with RUTBA_VERIFY_ONLY=word-colour.

import fs from 'node:fs';
import path from 'node:path';
import { OoxmlPackage } from '@rutba/ooxml';
import { buildDocx } from '@rutba/ooxml/build';
import { themeGrid } from '../renderer/apps/colour/grid.js';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyWordColour(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const own = path.join(dir, 'word-colour');
  const file = path.join(own, 'colours.docx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, buildDocx({
      styles: true,
      paragraphs: [
        { text: 'A heading in colour' }, { text: 'The quick brown fox' }, { text: 'Second line here' }, { text: '' },
        { table: { rows: [['a1', 'b1'], ['a2', 'b2']] } }, { text: 'After the table' },
      ],
    }));
  } catch (err) {
    return check('word-colour: the fixture was made', false, err.message);
  }

  let win = null;
  let again = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(async () => (await js(`document.querySelectorAll('.wd-page [data-block]').length`)) > 4, 'the page', 8000);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const runsOf = (i) => (model().blocks[i]?.runs || []).filter((r) => r.text);
    const colourOf = (i) => { const r = runsOf(i); return r.length && r.every((x) => x.fontColour === r[0].fontColour) ? (r[0].fontColour ?? null) : r.map((x) => x.fontColour ?? null); };
    const painted = (i) => js(`(() => {
      const b = document.querySelector('.wd-page [data-block="${i}"]');
      if (!b) return null;
      const spans = [...b.querySelectorAll('span')].filter((s) => s.childNodes.length && [...s.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim()));
      return spans.map((s) => getComputedStyle(s).color);
    })()`);
    const rgb = (hex) => `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`;
    const select = (i, { caret = false } = {}) => js(`(() => {
      const page = document.querySelector('.wd-page');
      const block = page.querySelector('[data-block="${i}"]');
      if (!block) return 'no block';
      page.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      if (${caret}) range.collapse(false);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      page.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      return 'selected';
    })()`);
    const settle = async (i, caret = false) => {
      await select(i, { caret });
      await until(() => { const s = model().selection; return s && s.focus?.block === i && (caret ? s.anchor.offset === s.focus.offset : s.anchor.offset !== s.focus.offset || s.anchor.block !== s.focus.block); }, `the selection in block ${i}`, 3000).catch(() => {});
      await wait(150);
    };
    const press = (start) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith(${JSON.stringify(start)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(start)};
      b.click();
      return 'clicked';
    })()`);
    const palette = () => js(`(() => {
      const box = document.querySelector('.rw-colours');
      if (!box) return null;
      const cells = (sel) => [...box.querySelectorAll(sel)].map((b) => b.dataset.colour);
      return {
        base: cells('.rw-cell.base'), steps: cells('.rw-cell[class*="step"]'), std: cells('.rw-cell.std'), recent: cells('.rw-cell.recent'),
        all: box.querySelectorAll('.rw-cell').length, rows: [...box.querySelectorAll('.rw-colours-row')].map((b) => b.textContent.trim()),
        heads: [...box.querySelectorAll('.rw-menu-head')].map((n) => n.textContent.trim()), highlight: box.classList.contains('rw-colours-hl'),
        values: [...box.querySelectorAll('.rw-cell')].map((b) => b.dataset.value),
      };
    })()`);
    const openPalette = async (name) => {
      const r = await press(`${name}: more colours`);
      await until(async () => Boolean(await palette()), 'the palette', 3000).catch(() => {});
      return r;
    };
    const pickCell = (value) => js(`(() => { const b = document.querySelector('.rw-colours .rw-cell[data-value="${value}"]'); if (!b) return 'no cell ${value}'; b.click(); return 'clicked'; })()`);
    const pickRow = (text) => js(`(() => { const b = [...document.querySelectorAll('.rw-colours .rw-colours-row')].find((n) => n.textContent.trim() === ${JSON.stringify(text)}); if (!b) return 'no row'; b.click(); return 'clicked'; })()`);
    const escape = async () => { wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' }); await until(async () => !(await palette()), 'the palette to close', 2000).catch(() => {}); };
    const grid = themeGrid(model().design?.colors);

    // 1. The palette, laid out as Word's.
    const opened = await openPalette('Font colour');
    const p = await palette();
    await capture(win, 'word-colour-palette.png');
    check('word-colour: the Font colour arrow opens a palette: Automatic, the theme\'s ten colours with five steps under each, ten standard colours and More colours',
      opened === 'clicked' && p && p.base.length === 10 && p.steps.length === 50 && p.std.length === 10 && p.rows[0] === 'Automatic' && p.rows.includes('More colours…')
        && p.base.join() === grid.map((c) => c.hex).join() && p.std.join() === 'C00000,FF0000,FFC000,FFFF00,92D050,00B050,00B0F0,0070C0,002060,7030A0',
      JSON.stringify({ opened, p, theme: grid.map((c) => c.hex) }));
    await escape();

    // 2. Chosen words take a theme step.
    const step = grid[4].steps[2].hex; // Accent 1, lighter 40%
    await settle(1);
    await openPalette('Font colour');
    const picked = await pickCell(step);
    await until(() => colourOf(1) === step, 'the words in the theme step', 4000).catch(() => {});
    // The engine answers in the main process before the page has drawn the answer: wait for the page too.
    await until(async () => (await painted(1))?.every((c) => c === rgb(step)), 'the page to draw the step', 4000).catch(() => {});
    const paint1 = await painted(1);
    check('word-colour: a theme colour\'s step colours the chosen words, and the page paints them in it',
      picked === 'clicked' && colourOf(1) === step && paint1?.length > 0 && paint1.every((c) => c === rgb(step)) && !(await palette()),
      JSON.stringify({ picked, engine: colourOf(1), paint1, want: step }));

    // 3. The face puts the same colour on other words.
    await settle(2);
    const face = await press('Font colour: apply');
    await until(() => colourOf(2) === step, 'the face to colour the words', 4000).catch(() => {});
    const faceTip = await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Font colour: apply'))?.dataset.tip`);
    const bar = await js(`getComputedStyle(document.querySelector('.rw-split[data-kind="word.font"] .rw-split-bar')).backgroundColor`);
    check('word-colour: the button\'s face puts the colour last chosen on the next words, shows it in its bar and names it in its tip',
      face === 'clicked' && colourOf(2) === step && bar === rgb(step) && /lighter 40%/.test(faceTip || ''),
      JSON.stringify({ face, engine: colourOf(2), bar, faceTip }));

    // 4. At the caret in an empty paragraph, the words typed next. With
    // nothing selected a colour goes on the caret's paragraph
    // (D-2026-11-11-11); an empty one has no words yet, so the words typed
    // there take it.
    await settle(3, true);
    await openPalette('Font colour');
    await pickCell('0070C0');
    await wait(200);
    await js(`document.querySelector('.wd-page')?.focus()`);
    wc.insertText('Typed in blue');
    await until(() => runsOf(3).map((r) => r.text).join('') === 'Typed in blue', 'the words typed', 4000).catch(() => {});
    await until(async () => (await painted(3))?.every((c) => c === rgb('0070C0')), 'the page to draw the words typed', 4000).catch(() => {});
    const paint3 = await painted(3);
    check('word-colour: a colour chosen with nothing selected in an empty paragraph colours the words typed there',
      colourOf(3) === '0070C0' && paint3?.length > 0 && paint3.every((c) => c === rgb('0070C0')),
      JSON.stringify({ runs: runsOf(3), paint3 }));

    // 5. A table cell's words.
    const cell = model().blocks.findIndex((b) => /:r0:c1$/.test(b.container || ''));
    await settle(cell);
    await openPalette('Font colour');
    await pickCell('00B050');
    await until(() => colourOf(cell) === '00B050', 'the cell\'s words in green', 4000).catch(() => {});
    await until(async () => (await painted(cell))?.every((c) => c === rgb('00B050')), 'the page to draw the cell green', 4000).catch(() => {});
    const paintCell = await painted(cell);
    check('word-colour: words in a table cell take the colour too',
      cell > 0 && colourOf(cell) === '00B050' && paintCell?.every((c) => c === rgb('00B050')), JSON.stringify({ cell, engine: colourOf(cell), paintCell }));

    // 6. More colours: any colour, by its hex; it joins the recent colours.
    await settle(0);
    await openPalette('Font colour');
    await pickRow('More colours…');
    await until(() => js(`Boolean(document.querySelector('.rw-colours-dialog'))`), 'More colours', 3000).catch(() => {});
    await capture(win, 'word-colour-more.png');
    const typed = await js(`(() => {
      const input = document.querySelector('.rw-colours-dialog .rw-colours-hex');
      if (!input) return 'no hex field';
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      set.call(input, '#12abef');
      input.dispatchEvent(new Event('input', { bubbles: true }));
      return 'typed';
    })()`);
    await wait(150);
    const preview = await js(`document.querySelector('.rw-colours-compare span')?.style.background || ''`);
    await js(`document.querySelector('.rw-colours-dialog .rw-colours-ok')?.click()`);
    await until(() => colourOf(0) === '12ABEF', 'the custom colour', 4000).catch(() => {});
    await openPalette('Font colour');
    const withRecent = await palette();
    await escape();
    check('word-colour: More colours takes a hex, shows it beside the colour it replaces, colours the words, and the colour joins Recent colours',
      typed === 'typed' && /18, 171, 239|12abef/i.test(preview) && colourOf(0) === '12ABEF' && withRecent?.recent?.[0] === '12ABEF' && withRecent.heads.includes('Recent colours'),
      JSON.stringify({ typed, preview, engine: colourOf(0), recent: withRecent?.recent }));

    // 7. Automatic takes the colour off.
    await settle(0);
    await openPalette('Font colour');
    await pickRow('Automatic');
    await until(() => colourOf(0) === null, 'the colour off', 4000).catch(() => {});
    check('word-colour: Automatic takes the colour off again', colourOf(0) === null, JSON.stringify({ engine: colourOf(0) }));

    // 8. Text highlight colour: Word's fifteen, and No highlight.
    await settle(1);
    await openPalette('Text highlight colour');
    const hl = await palette();
    await pickCell('green');
    await until(() => runsOf(1).every((r) => r.highlight === 'green'), 'the highlight', 4000).catch(() => {});
    await until(() => js(`[...document.querySelectorAll('.wd-page [data-block="1"] span')].some((s) => getComputedStyle(s).backgroundColor !== 'rgba(0, 0, 0, 0)')`), 'the page to draw the highlight', 4000).catch(() => {});
    const back = await js(`[...document.querySelectorAll('.wd-page [data-block="1"] span')].map((s) => getComputedStyle(s).backgroundColor).filter((c) => c !== 'rgba(0, 0, 0, 0)')`);
    check('word-colour: Text highlight colour offers Word\'s fifteen highlight colours and No highlight, and Bright green highlights the words',
      hl?.highlight === true && hl.all === 15 && hl.values.join() === 'yellow,green,cyan,magenta,blue,red,darkBlue,darkCyan,darkGreen,darkMagenta,darkRed,darkYellow,darkGray,lightGray,black'
        && hl.rows[0] === 'No highlight' && runsOf(1).every((r) => r.highlight === 'green') && back.length > 0,
      JSON.stringify({ hl, runs: runsOf(1).map((r) => r.highlight), back }));

    // 9. Saved, the file carries them; opened again, so does the page.
    await press('Save');
    await until(() => doc.meta({ id: session.id }).dirty === false, 'the save', 15000).catch(() => false);
    const body = OoxmlPackage.read(fs.readFileSync(file)).text('word/document.xml');
    const inFile = {
      step: body.includes(`<w:color w:val="${step}"/>`), blue: body.includes('<w:color w:val="0070C0"/>'), green: body.includes('<w:color w:val="00B050"/>'),
      highlight: body.includes('<w:highlight w:val="green"/>'), custom: !body.includes('12ABEF'),
    };
    win.close();
    win = null;
    await wait(500);
    again = await open('word', file);
    await until(async () => (await again.webContents.executeJavaScript(`document.querySelectorAll('.wd-page [data-block]').length`)) > 4, 'the page again', 8000).catch(() => {});
    const reopened = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const m2 = doc.model({ id: reopened.id });
    const c2 = (i) => (m2.blocks[i]?.runs || []).filter((r) => r.text).map((r) => r.fontColour ?? null);
    check('word-colour: the saved file carries each colour and the highlight as Word writes them, and opens with them',
      Object.values(inFile).every(Boolean) && c2(1).every((c) => c === step) && c2(3).every((c) => c === '0070C0') && c2(cell).every((c) => c === '00B050'),
      JSON.stringify({ inFile, b1: c2(1), b3: c2(3), cell: c2(cell) }));
  } catch (err) {
    check('word-colour: the checks ran', false, err.stack || err.message);
  } finally {
    try { win?.close(); } catch { /* gone */ }
    try { again?.close(); } catch { /* gone */ }
  }
}
