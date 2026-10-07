// Word: right to left.
//
// A paragraph with w:bidi runs from the right margin, its words drawn
// right to left; Home → Paragraph's direction buttons turn a paragraph
// either way and show which way the caret's runs; aligning a right-to-left
// paragraph puts it where the button says, written mirrored as Word writes
// it; and the file keeps it all. In a second document: the ruler counts
// from the right margin in a right-to-left paragraph and its tab stop lands
// measured from there, a right-to-left table puts its first column at the
// right, and Layout → Columns → Right to left turns the section. Run alone
// with RUTBA_VERIFY_ONLY=rtl.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const ARABIC = 'مرحبا بالعالم، هذه فقرة تبدأ من اليمين.';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordRtl(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const capture = async (win, name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
  };
  const file = path.join(dir, 'rtl.docx');
  // The Arabic paragraph made right to left the way the editor writes it.
  const made = openDocx(buildDocx({ styles: true, paragraphs: [
    { text: 'Direction', style: 'Heading1' },
    { text: 'This paragraph runs left to right, as English does.' },
    { text: ARABIC },
  ] }));
  made.setSelection({ block: 2, offset: 0 });
  made.setParagraphFormat({ rtl: true });
  fs.writeFileSync(file, made.save());
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    // Where a paragraph's words sit against its own box, and which way it runs.
    const placed = (block) => js(`(() => {
      const el = document.querySelector('.wd-page > [data-block="${block}"]');
      if (!el) return null;
      const box = el.getBoundingClientRect();
      const range = document.createRange();
      range.selectNodeContents(el);
      const words = range.getBoundingClientRect();
      return { dir: getComputedStyle(el).direction, left: Math.round(words.left - box.left), right: Math.round(box.right - words.right) };
    })()`);
    const button = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}));
      return b ? b.getAttribute('aria-pressed') : 'missing';
    })()`);
    const pressButton = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button';
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickInto = async (block) => {
      const at = await js(`(() => { const r = document.querySelector('.wd-page > [data-block="${block}"]').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + 8) }; })()`);
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      return until(() => model().selection?.focus?.block === block, `the caret in paragraph ${block}`, 4000).catch(() => false);
    };
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the paragraphs', 8000);

    // The file's right-to-left paragraph, drawn from the right margin.
    const arabic = await placed(2);
    const english = await placed(1);
    check('word: a paragraph the file marks right to left runs from the right margin',
      arabic?.dir === 'rtl' && arabic.right <= 2 && arabic.left > 40 && english?.dir === 'ltr' && english.left <= 2,
      `Arabic ${JSON.stringify(arabic)}; English ${JSON.stringify(english)}`);

    // The buttons show which way the caret's paragraph runs.
    const inArabic = await clickInto(2);
    await wait(250);
    const shown = { rtl: await button('Right-to-left text direction'), ltr: await button('Left-to-right text direction') };
    check('word: Home → Paragraph shows Right-to-left pressed in a right-to-left paragraph',
      inArabic === true && shown.rtl === 'true' && shown.ltr === 'false', JSON.stringify(shown));

    // Aligning it: the button's side of the page, written mirrored.
    const aligned = await pressButton('Align left');
    const moved = await until(async () => { const p = await placed(2); return p && p.left <= 2 && p.right > 40; }, 'the Arabic at the left', 4000).catch(() => false);
    check('word: Align left puts a right-to-left paragraph at the left margin',
      aligned === 'clicked' && moved === true && model().format?.paragraphAlign === 'left', `${aligned}; ${JSON.stringify(await placed(2))}; ${model().format?.paragraphAlign}`);

    // Word's keys: Ctrl+R back to the right margin, Ctrl+E to the middle.
    await press(wc, 'r', { modifiers: ['control'] });
    const keyRight = await until(async () => { const p = await placed(2); return p && p.right <= 2 && p.left > 40; }, 'the Arabic back at the right', 4000).catch(() => false);
    await press(wc, 'e', { modifiers: ['control'] });
    const keyCentre = await until(async () => { const p = await placed(2); return p && Math.abs(p.left - p.right) <= 3; }, 'the Arabic centred', 4000).catch(() => false);
    check('word: Ctrl+R and Ctrl+E align the paragraph, as the buttons\' tooltips say',
      keyRight === true && keyCentre === true && model().format?.paragraphAlign === 'center', `right ${keyRight}, centre ${keyCentre}; ${JSON.stringify(await placed(2))}`);

    // The English paragraph turned right to left, and back.
    await clickInto(1);
    await wait(200);
    const turned = await pressButton('Right-to-left text direction');
    const nowRtl = await until(async () => model().blocks[1].rtl === true && (await placed(1))?.dir === 'rtl', 'paragraph 1 right to left', 4000).catch(() => false);
    const across = await placed(1);
    check('word: Right-to-left text direction turns the caret\'s paragraph to run from the right',
      turned === 'clicked' && nowRtl === true && across.right <= 2, `${turned}; ${JSON.stringify(across)}`);
    await wait(300);
    wc.invalidate();
    await wait(500);
    await capture(win, 'word-rtl.png');
    const back = await pressButton('Left-to-right text direction');
    const nowLtr = await until(async () => (await placed(1))?.dir === 'ltr', 'paragraph 1 left to right', 4000).catch(() => false);
    check('word: Left-to-right text direction turns it back', back === 'clicked' && nowLtr === true && model().blocks[1].rtl === undefined,
      `${back}; ${JSON.stringify(await placed(1))}; rtl ${model().blocks[1].rtl}`);

    // Saved as Word writes it.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    await wait(1400);
    const saved = openDocx(fs.readFileSync(file));
    const xml = saved.doc.doc.xml;
    const blocks = saved.render({ pages: false }).blocks;
    check('word: the file keeps the right-to-left paragraph as w:bidi and its alignment',
      blocks[2].rtl === true && blocks[1].rtl === undefined && /<w:bidi\/>[\s\S]*?<w:jc w:val="center"\/>/.test(xml),
      `rtl ${blocks.map((b) => b.rtl).join(',')}; ${(/<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<w:bidi[\s\S]*?<\/w:pPr>/.exec(xml) || ['no w:bidi'])[0].slice(0, 160)}`);

    const complaints = await errorsIn(win);
    check('word: right to left reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the right-to-left checks ran', false, err.message);
  }

  // The ruler, a tab stop, a table and the section, right to left.
  const second = path.join(dir, 'rtl-ruler.docx');
  const STOP_TWIPS = 2268; // 4 cm
  const parts = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'عنوان\tنص بعد الجدولة' }, { text: 'Before the table.' }] }));
  parts.setSelection({ block: 0, offset: 0 });
  parts.setParagraphFormat({ rtl: true, tabs: [{ align: 'left', posTwips: STOP_TWIPS }] });
  parts.setSelection({ block: 1, offset: 0 });
  parts.insertTable({ rows: 2, cols: 3 });
  const firstCell = parts.render({ pages: false }).blocks.findIndex((b) => /^t\d+:r0:c0$/.test(b.container || ''));
  parts.setSelection({ block: firstCell, offset: 0 });
  parts.insertText('أول');
  parts.tableOp('direction', { rtl: true });
  fs.writeFileSync(second, parts.save());
  try {
    const win = await open('word', second);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="0"] .wd-tab'))`), 'the tab', 8000);
    // The caret into the Arabic paragraph, so the ruler shows it.
    const at = await js(`(() => { const r = document.querySelector('.wd-page > [data-block="0"]').getBoundingClientRect(); return { x: Math.round(r.right - 20), y: Math.round(r.top + 8) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    await until(() => model().selection?.focus?.block === 0, 'the caret in the Arabic', 4000).catch(() => {});
    if (!(await js(`Boolean(document.querySelector('.wd-ruler'))`))) {
      await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'View')?.click(); return 1; })()`);
    }
    await until(() => js(`Boolean(document.querySelector('.wd-ruler[data-dir="rtl"]'))`), 'the ruler right to left', 4000).catch(() => {});
    await wait(300);
    const cm = 96 / 2.54;
    const seen = await js(`(() => {
      const page = document.querySelector('.wd-page').getBoundingClientRect();
      const pad = parseFloat(getComputedStyle(document.querySelector('.wd-page')).paddingRight) || 0;
      const ruler = document.querySelector('.wd-ruler');
      const tabMark = ruler?.querySelector('.wd-ruler-tab');
      const first = ruler?.querySelector('.wd-ruler-first');
      const r = ruler?.getBoundingClientRect();
      const tab = document.querySelector('.wd-page > [data-block="0"] .wd-tab').getBoundingClientRect();
      const cells = [...document.querySelectorAll('.wd-table td')].slice(0, 2).map((td) => Math.round(td.getBoundingClientRect().left));
      return {
        dir: ruler?.dataset.dir || null,
        rightEdge: page.right - pad,
        tabMark: tabMark ? Math.round(tabMark.getBoundingClientRect().left + 4 - r.left) : null,
        first: first ? Math.round(first.getBoundingClientRect().left + 5 - r.left) : null,
        rulerWidth: r ? Math.round(r.width) : null,
        tabEnd: Math.round(tab.left),
        cells,
        tableDir: getComputedStyle(document.querySelector('.wd-table')).direction,
      };
    })()`);
    const stopPx = STOP_TWIPS / 15;
    const fromRight = Math.round(seen.rightEdge - seen.tabEnd);
    check('word: in a right-to-left paragraph the ruler counts from the right margin, its tab stop measured from there',
      seen.dir === 'rtl' && seen.tabMark != null && seen.first != null && seen.first > seen.tabMark && Math.abs((seen.first - seen.tabMark) - stopPx) <= 3,
      `ruler ${seen.dir}; first-line marker at ${seen.first}, tab mark at ${seen.tabMark} of ${seen.rulerWidth} px (${(stopPx / cm).toFixed(2)} cm apart wanted)`);
    check('word: a right-to-left paragraph\'s tab runs to its stop measured from the right margin',
      Math.abs(fromRight - stopPx) <= 3, `the tab ends ${fromRight} px from the right margin, the stop at ${Math.round(stopPx)}`);
    check('word: a right-to-left table puts its first column at the right', seen.tableDir === 'rtl' && seen.cells.length === 2 && seen.cells[0] > seen.cells[1],
      `${seen.tableDir}; first two cells at ${seen.cells.join(', ')}`);
    await wait(200);
    wc.invalidate();
    await wait(400);
    await capture(win, 'word-rtl-ruler.png');

    // Layout → Columns → Right to left: the section turned.
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Layout')?.click(); return 1; })()`);
    await wait(250);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.textContent || '').trim().startsWith('Columns')); b?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); b?.click(); return 1; })()`);
    await wait(250);
    const picked = await js(`(() => { const item = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim() === 'Right to left'); if (!item) return false; item.click(); return true; })()`);
    const turned = await until(() => model().section?.rtl === true, 'the section right to left', 4000).catch(() => false);
    check('word: Layout → Columns → Right to left turns the section', picked === true && turned === true, `menu item ${picked}; section rtl ${model().section?.rtl}`);

    const complaints = await errorsIn(win);
    check('word: the right-to-left ruler, table and section report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the right-to-left ruler checks ran', false, err.message);
  }
}
