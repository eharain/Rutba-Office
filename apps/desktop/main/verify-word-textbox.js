// Word: Insert → Text Box — a box of words floating on the page.
//
// A Simple Text Box goes in from the ribbon; its words are selected, so the
// keys typed next replace them; Bold works in it as in the body; the words
// of its paragraph run beside it. Shape Format appears with it: Shape Fill
// colours it, Align Text moves its words to the middle. Its edge drags it
// somewhere else and a handle sizes it — both written to the file — and
// Draw Text Box makes a second one by dragging on the page. The saved file
// carries both, as Word writes them. Run alone with RUTBA_VERIFY_ONLY=textbox.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn. ';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordTextBox(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'textbox.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [
    { text: 'Text boxes', style: 'Heading1' },
    { text: lorem.repeat(5) },
    { text: lorem.repeat(4) },
    { text: 'The end.' },
  ] }));
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    const boxes = () => (model().drawings || []).filter((d) => d.kind === 'textbox');
    const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickMenuItem = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `the "${label}" menu item`, 3000).catch(() => false);
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'missing'; b.click(); return 'clicked'; })()`);
    };
    const rect = (selector) => js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r ? { left: r.left, top: r.top, right: r.right, bottom: r.bottom, x: r.left + r.width / 2, y: r.top + r.height / 2, width: r.width, height: r.height } : null; })()`);
    const drag = async (x, y, dx, dy) => {
      x = Math.round(x);
      y = Math.round(y);
      wc.sendInputEvent({ type: 'mouseMove', x, y });
      wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: 1 });
      await wait(60);
      for (const f of [0.25, 0.5, 0.75, 1]) {
        wc.sendInputEvent({ type: 'mouseMove', x: Math.round(x + dx * f), y: Math.round(y + dy * f), button: 'left' });
        await wait(40);
      }
      wc.sendInputEvent({ type: 'mouseUp', x: x + dx, y: y + dy, button: 'left', clickCount: 1 });
      await wait(300);
    };
    const typeText = async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(60);
      }
    };
    const shot = async (name) => {
      wc.invalidate();
      await wait(700);
      await capture(win, name);
    };

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the paragraphs', 8000);
    // The caret in the second paragraph, for real.
    const para = await rect('.wd-page [data-block="1"]');
    wc.sendInputEvent({ type: 'mouseDown', x: Math.round(para.left + 14), y: Math.round(para.top + 8), button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: Math.round(para.left + 14), y: Math.round(para.top + 8), button: 'left', clickCount: 1 });
    await wait(400);

    // Insert → Text Box → Simple Text Box.
    await clickTab('Insert');
    await wait(200);
    const opened = await clickRibbon('Text Box');
    const chose = await clickMenuItem('Simple Text Box');
    const made = await until(() => boxes().length === 1, 'the text box in the file', 6000).catch(() => false);
    const drawn = await until(() => js(`Boolean(document.querySelector('.wd-textbox.editable .wd-block'))`), 'the text box on the page', 6000).catch(() => false);
    await wait(400);
    const layout = await js(`(() => {
      const box = document.querySelector('.wd-textbox.editable').getBoundingClientRect();
      const words = document.querySelector('.wd-page > [data-block="1"]');
      const walker = document.createTreeWalker(words, NodeFilter.SHOW_TEXT);
      let node; let first = null;
      while ((node = walker.nextNode())) { if (!node.parentElement.closest('.wd-textbox') && node.textContent.trim()) { first = node; break; } }
      const r = document.createRange(); r.setStart(first, 0); r.setEnd(first, 1);
      const t = r.getBoundingClientRect();
      return { box: { left: Math.round(box.left), top: Math.round(box.top), right: Math.round(box.right), bottom: Math.round(box.bottom) }, text: { left: Math.round(t.left), top: Math.round(t.top) } };
    })()`);
    check('word: Insert → Text Box → Simple Text Box puts a box on the page, the words of its paragraph running beside it',
      opened === 'clicked' && chose === 'clicked' && made === true && drawn === true && layout.text.left >= layout.box.right - 2 && layout.text.top < layout.box.bottom,
      `${opened} ${chose}; ${JSON.stringify(layout)}`);

    // The box's words are selected: typing replaces them.
    await typeText('Hello box');
    const typed = await until(() => (model().blocks || []).some((b) => b.box && b.text === 'Hello box'), 'the typed words in the box', 6000).catch(() => false);
    const shown = await until(() => js(`document.querySelector('.wd-textbox.editable')?.textContent.includes('Hello box')`), 'the typed words on the page', 4000).catch(() => false);
    check('word: typing into a new text box replaces its words, in the box on the page and in the file',
      typed === true && shown === true, JSON.stringify((model().blocks || []).filter((b) => b.box).map((b) => b.text)));

    // Bold, the Home tab's own button, on the box's words.
    await press(wc, 'Home', { modifiers: ['shift'] });
    await wait(400);
    await clickTab('Home');
    await wait(200);
    const bolded = await clickRibbon('Bold');
    const bold = await until(() => (model().blocks || []).some((b) => b.box && (b.runs || []).some((r) => r.bold && /Hello/.test(r.text))), 'bold in the box', 5000).catch(() => false);
    check('word: Home → Bold makes a text box\'s words bold, the usual way', bolded === 'clicked' && bold === true,
      JSON.stringify((model().blocks || []).filter((b) => b.box).map((b) => (b.runs || []).map((r) => [r.text, Boolean(r.bold)]))));

    // Shape Format is there while the box is: Shape Fill, Align Text.
    const hasTab = await until(() => js(`[...document.querySelectorAll('.rw-tab')].some((t) => t.textContent.trim() === 'Shape Format')`), 'the Shape Format tab', 4000).catch(() => false);
    await clickTab('Shape Format');
    await wait(200);
    const filled = await clickRibbon('Shape Fill');
    await clickMenuItem('Light blue');
    const coloured = await until(() => js(`getComputedStyle(document.querySelector('.wd-textbox.editable')).backgroundColor === 'rgb(222, 235, 247)'`), 'the box to turn light blue', 5000).catch(() => false);
    await wait(300);
    const aligned = await clickRibbon('Align Text');
    await clickMenuItem('Middle');
    const middle = await until(() => js(`getComputedStyle(document.querySelector('.wd-textbox.editable')).justifyContent === 'center'`), 'the words in the middle', 5000).catch(() => false);
    const lookNow = boxes()[0];
    check('word: Shape Format appears with the box; Shape Fill colours it and Align Text puts its words in the middle',
      hasTab === true && filled === 'clicked' && coloured === true && aligned === 'clicked' && middle === true,
      `${hasTab} ${filled} ${aligned}; ${JSON.stringify(lookNow && { id: lookNow.id })}`);
    await shot('word-textbox.png');

    // Its edge drags it; a handle sizes it.
    const before = boxes()[0];
    await until(() => js(`Boolean(document.querySelector('.wd-frame-sel .wd-frame-edge.t'))`), 'the box\'s frame', 4000).catch(() => false);
    const edge = await rect('.wd-frame-sel .wd-frame-edge.t');
    // Taken at a quarter of its width: the middle of the top edge is the top handle.
    if (edge) await drag(edge.left + edge.width / 4, edge.y, 150, 50);
    const moved = await until(() => { const d = boxes()[0]; return d && Math.abs((d.hOffsetPx ?? 0) - (before.hOffsetPx ?? 0) - 150) <= 4; }, 'the box to move', 5000).catch(() => false);
    const afterMove = boxes()[0];
    const handle = await rect('.wd-frame-sel .wd-handle[data-handle="se"]');
    if (handle) await drag(handle.x, handle.y, 60, 40);
    const sized = await until(() => { const d = boxes()[0]; return d && Math.abs(d.widthPx - before.widthPx - 60) <= 3; }, 'the box to grow', 5000).catch(() => false);
    const afterSize = boxes()[0];
    const onPage = await rect('.wd-textbox.editable');
    check('word: a text box moves by its edge and sizes by its handle, and the file says where and how big',
      Boolean(edge) && moved === true && sized === true && Math.abs(onPage.width - afterSize.widthPx) <= 3,
      `offset ${before.hOffsetPx}→${afterMove?.hOffsetPx}, ${before.vOffsetPx}→${afterMove?.vOffsetPx}; width ${before.widthPx}→${afterSize?.widthPx}; drawn ${Math.round(onPage?.width)}`);

    // Draw Text Box: a drag on the page is a box of that size.
    await clickTab('Insert');
    await wait(200);
    await clickRibbon('Text Box');
    await clickMenuItem('Draw Text Box');
    await wait(300);
    const para2 = await rect('.wd-page > [data-block="2"]');
    await drag(para2.left + 60, para2.top + 30, 200, 96);
    const two = await until(() => boxes().length === 2, 'the drawn box', 6000).catch(() => false);
    const drawnBox = boxes().find((d) => d.id !== before.id);
    check('word: Insert → Text Box → Draw Text Box draws a box the size of the drag',
      two === true && Boolean(drawnBox) && Math.abs(drawnBox.widthPx - 200) <= 3 && Math.abs(drawnBox.heightPx - 96) <= 3 && drawnBox.block === 2,
      JSON.stringify(drawnBox && { w: drawnBox.widthPx, h: drawnBox.heightPx, block: drawnBox.block, x: drawnBox.hOffsetPx, y: drawnBox.vOffsetPx }));
    await shot('word-textbox-drawn.png');

    // Saved as Word writes them.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    await wait(1400);
    const bytes = fs.readFileSync(file);
    const reopened = openDocx(bytes);
    const xml = reopened.doc.doc.xml;
    const saved = reopened.render({ pages: false });
    const savedBoxes = saved.blocks.flatMap((b) => b.textBoxes || []);
    check('word: the text boxes are saved as Word writes them — wps shapes in markup compatibility content, with their VML twins',
      savedBoxes.length === 2 && (xml.match(/<mc:Choice Requires="wps">/g) || []).length === 2 && (xml.match(/<v:textbox\b/g) || []).length === 2
        && saved.blocks.some((b) => b.box && b.text === 'Hello box') && savedBoxes.some((b) => b.fill === '#DEEBF7' && b.vAnchor === 'middle'),
      JSON.stringify(savedBoxes.map((b) => ({ w: b.widthPx, fill: b.fill, v: b.vAnchor }))));

    const complaints = await errorsIn(win);
    check('word: text boxes report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the text box checks ran', false, err.message);
  }
}
