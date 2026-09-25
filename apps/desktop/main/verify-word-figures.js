// Word: References → Insert Table of Figures and Update Table, and Insert →
// Quick Parts → Field for PAGE, NUMPAGES and DATE in the body.
//
// Page and NUMPAGES go into a line on page 2 from the Field dialog and read
// the page the window laid them on; DATE reads today in the picture chosen;
// the table of figures lists both captions with their pages; a caption added
// from Insert Caption joins it at Update Table; text added above pushes the
// line to page 3 and F9 brings PAGE and NUMPAGES with it; the saved file
// keeps the TOC \c field and the complex fields. Run alone with
// RUTBA_VERIFY_ONLY=figures.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const FILL = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn despite two price changes and a supply interruption that took most of July to clear. ';

/** The report this check writes, its own: a caption on page 1, another on page 2, a page line at the end. */
export function makeFiguresFixture(dir) {
  const file = path.join(dir, 'figures.docx');
  const view = openDocx(buildDocx({
    styles: true,
    paragraphs: [
      { text: 'Quarterly figures', style: 'Heading1' },
      { text: '' },
      { text: 'Sales rose in every region.' },
      { text: FILL.repeat(20) },
      { text: 'The map shows the regions.' },
      { text: 'Page  of ' },
      { text: 'Printed on ' },
    ],
  }));
  view.setSelection({ block: 2, offset: 0 });
  view.insertCaption({ label: 'Figure', text: 'Sales by region' });
  view.setSelection({ block: 5, offset: 0 });
  view.insertCaption({ label: 'Figure', text: 'The regions' });
  fs.writeFileSync(file, view.save());
  return file;
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordFigures(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  try {
    const file = makeFiguresFixture(dir);
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });

    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickMenu = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(label)}) && !n.disabled);
      if (!b) return 'no menu item ' + ${JSON.stringify(label)};
      b.click();
      return 'clicked';
    })()`);
    const click = (selector) => js(`(() => {
      const b = document.querySelector(${JSON.stringify(selector)});
      if (!b || b.disabled) return 'no ' + ${JSON.stringify(selector)};
      b.click();
      return 'clicked';
    })()`);
    const fill = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no ' + ${JSON.stringify(selector)};
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return 'filled';
    })()`);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const blockText = (i) => js(`(() => { const b = document.querySelector('.wd-page [data-block="${i}"]'); return b ? b.innerText.replace(/\\u200b/g, '') : null; })()`);
    const indexOf = (text) => (model().blocks || []).findIndex((b) => (b.text || '').startsWith(text));
    // A real click at a character offset in a paragraph.
    const caretAt = async (block, offset) => {
      const at = await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="${block}"]');
        if (!b) return null;
        b.scrollIntoView({ block: 'center' });
        const walker = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
        let seen = 0;
        for (let n = walker.nextNode(); n; n = walker.nextNode()) {
          if (seen + n.nodeValue.length >= ${offset}) {
            const r = document.createRange();
            const k = ${offset} - seen;
            if (k > 0) { r.setStart(n, k - 1); r.setEnd(n, k); const box = r.getBoundingClientRect(); return { x: Math.round(box.right), y: Math.round(box.top + box.height / 2) }; }
            r.setStart(n, 0); r.setEnd(n, 1); const box = r.getBoundingClientRect(); return { x: Math.round(box.left), y: Math.round(box.top + box.height / 2) };
          }
          seen += n.nodeValue.length;
        }
        const box = b.getBoundingClientRect();
        return { x: Math.round(box.left + 3), y: Math.round(box.top + 10) };
      })()`);
      if (!at) return false;
      win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      return until(() => model().selection?.focus?.block === block && model().selection.focus.offset === offset, `the caret at ${block}:${offset}`, 5000).catch(() => false);
    };
    const snap = async (name) => {
      win.webContents.invalidate();
      await wait(700);
      await win.webContents.capturePage().catch(() => null);
      win.webContents.invalidate();
      await wait(600);
      await capture(win, name);
    };
    const pageCount = () => js(`(() => { const c = [...document.querySelectorAll('.rw-chip, .chip, span')].map((n) => n.textContent).find((t) => /^Page \\d+ of \\d+$/.test(t)); return c ? Number(/of (\\d+)/.exec(c)[1]) : null; })()`);

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="6"]'))`), 'the report', 8000);
    await wait(600);

    // Insert → Quick Parts → Field… → PAGE, at "Page | of".
    const pageLine = indexOf('Page  of');
    await caretAt(pageLine, 5);
    await clickTab('Insert');
    await wait(200);
    const qp = await clickRibbon('Quick Parts');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the Quick Parts menu', 3000).catch(() => {});
    const fieldItem = await clickMenu('Field');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Field"]'))`), 'the Field dialog', 4000).catch(() => {});
    await snap('word-field-dialog.png');
    await click('.wd-refs-field-ok');
    await wait(300);
    // NUMPAGES at the end of the line.
    const lineLen = (model().blocks[pageLine].text || '').length;
    await caretAt(pageLine, lineLen);
    await clickRibbon('Quick Parts');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the Quick Parts menu', 3000).catch(() => {});
    await clickMenu('Field');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Field"] [data-field="NUMPAGES"]'))`), 'the Field dialog', 4000).catch(() => {});
    await js(`document.querySelector('.rw-dialog[aria-label="Field"] [data-field="NUMPAGES"]')?.click(), 1`);
    await wait(150);
    await click('.wd-refs-field-ok');
    const count = await pageCount();
    const pageText = await until(async () => /^Page \d+ of \d+$/.test((await blockText(pageLine)) || ''), 'Page N of M on the page', 6000).catch(() => false);
    const line = await blockText(pageLine);
    const [, n, of] = /^Page (\d+) of (\d+)$/.exec(line || '') || [];
    check(
      'word: Insert → Quick Parts → Field puts PAGE and NUMPAGES in the body, reading the page the line is on and the page count',
      qp === 'clicked' && fieldItem === 'clicked' && pageText === true && Number(n) >= 2 && Number(of) === count && Number(n) === count,
      JSON.stringify({ line, count, structural: model().blocks[pageLine].structural })
    );

    // DATE in a picture of its own.
    const dateLine = indexOf('Printed on');
    await caretAt(dateLine, (model().blocks[dateLine].text || '').length);
    await clickRibbon('Quick Parts');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the Quick Parts menu', 3000).catch(() => {});
    await clickMenu('Field');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Field"] [data-field="DATE"]'))`), 'the Field dialog', 4000).catch(() => {});
    await js(`document.querySelector('.rw-dialog[aria-label="Field"] [data-field="DATE"]')?.click(), 1`);
    await wait(150);
    await js(`[...document.querySelectorAll('.wd-refs-pictures .wd-refs-row')].find((b) => /^\\d{4}-\\d{2}-\\d{2}$/.test(b.textContent.trim()))?.click(), 1`);
    await wait(150);
    await click('.wd-refs-field-ok');
    const now = new Date();
    const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    const dated = await until(async () => (await blockText(dateLine)) === 'Printed on ' + today, 'the date', 5000).catch(() => false);
    check('word: a DATE field in the picture chosen reads today', dated === true, await blockText(dateLine));

    // References → Insert Table of Figures in the empty paragraph under the title.
    await caretAt(1, 0);
    await clickTab('References');
    await wait(200);
    const tofPressed = await clickRibbon('Insert Table of Figures');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Table of Figures"]'))`), 'the Table of Figures dialog', 4000).catch(() => {});
    await snap('word-figures-dialog.png');
    await click('.wd-refs-tof-ok');
    const entries = () => js(`[...document.querySelectorAll('.wd-page [data-style="TableofFigures"]')].map((b) => b.innerText.replace(/\\u200b/g, '')).filter((t) => t.trim())`);
    const listed = await until(async () => (await entries()).length === 2, 'two entries', 6000).catch(() => false);
    const got = await entries();
    check(
      'word: Insert Table of Figures lists both captions, the second on a later page than the first',
      tofPressed === 'clicked' && listed === true && /^Figure 1: Sales by region\t1$/.test(got[0] || '') && /^Figure 2: The regions\t[2-9]$/.test(got[1] || ''),
      JSON.stringify(got)
    );
    await snap('word-figures.png');

    // A caption from Insert Caption above the second: it takes number 2, the
    // other becomes 3, and Update Table follows both.
    const mapLine = indexOf('The map shows');
    await caretAt(mapLine, 3);
    await clickRibbon('Insert Caption');
    await until(() => js(`Boolean(document.querySelector('.wd-caption-text'))`), 'the Caption dialog', 4000).catch(() => {});
    await fill('.wd-caption-text', 'The map');
    await click('.wd-caption-insert');
    await until(() => indexOf('Figure 3') > 0, 'the third caption', 5000).catch(() => {});
    const upd = await clickRibbon('Update Table — the table of figures');
    const three = await until(async () => (await entries()).length === 3, 'three entries', 6000).catch(() => false);
    check('word: Update Table adds the caption inserted since, and the one after it renumbered', upd === 'clicked' && three === true && (await entries()).some((e) => /^Figure 2: The map\t\d+$/.test(e)) && /^Figure 3: The regions\t\d+$/.test((await entries())[2] || ''), JSON.stringify(await entries()));

    // Text above pushes the page line on; F9 brings PAGE and NUMPAGES with it.
    const fillAt = indexOf('The northern region');
    const before = await pageCount();
    await js(`(async () => {
      const all = await window.rutbaOffice.doc.sessions({});
      const mine = all.filter((s) => s.kind === 'doc').pop();
      await window.rutbaOffice.doc.apply({ id: mine.id, ops: [
        { op: 'setSelection', anchor: { block: ${fillAt}, offset: 0 }, focus: { block: ${fillAt}, offset: 0 } },
        { op: 'insertText', text: ${JSON.stringify(FILL.repeat(14))} },
      ] });
      return 1;
    })()`);
    await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
    await press(win.webContents, 'z', { modifiers: ['control'] });
    await wait(300);
    await press(win.webContents, 'y', { modifiers: ['control'] });
    await until(async () => (await pageCount()) > before, 'the page count to grow', 8000).catch(() => {});
    await wait(500);
    const after = await pageCount();
    await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
    await press(win.webContents, 'F9');
    const f9 = await until(async () => (await blockText(indexOf('Page '))) === `Page ${after} of ${after}`, 'F9 to renumber the page line', 6000).catch(() => false);
    check('word: F9 works PAGE and NUMPAGES out again after the text above moved the line on a page', f9 === true && after > before, JSON.stringify({ before, after, line: await blockText(indexOf('Page ')) }));

    await clickRibbon('Save');
    const saved = () => {
      try {
        const d = openDocx(fs.readFileSync(file)).doc.doc;
        return { tof: d.tablesOfFigures(), body: d._body().body };
      } catch {
        return null;
      }
    };
    await until(() => saved()?.tof?.[0]?.entries?.length === 3, 'the table of figures in the saved file', 8000).catch(() => {});
    const out = saved();
    check(
      'word: the saved file keeps the TOC \\c "Figure" field with three linked entries, and PAGE, NUMPAGES and DATE as complex fields',
      Boolean(out) && out.tof[0]?.label === 'Figure' && out.tof[0].entries.length === 3 && /TOC \\h \\z \\c "Figure"/.test(out.body)
        && /<w:instrText xml:space="preserve"> PAGE \\\* MERGEFORMAT <\/w:instrText>/.test(out.body) && out.body.includes(' NUMPAGES ') && out.body.includes(' DATE \\@ "yyyy-MM-dd"'),
      JSON.stringify(out?.tof)
    );
    const complaints = await errorsIn(win);
    check('word: the table of figures and field checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the table of figures and field checks ran', false, err.message);
  }
}
