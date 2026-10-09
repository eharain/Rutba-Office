// Word: Shape Format → Text Effects → Transform — Insert → WordArt, then
// Arch Up: with the caret out of the box its words are drawn along an arc
// (a <textPath>) and the straight ones are unseen; with the caret back in
// they show straight to be edited. The saved file carries a:prstTxWarp
// first in the box's body properties, as Word writes it. Run alone with
// RUTBA_VERIFY_ONLY=wordwordart.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordWordArtTransform({ open, check, until, wait, press, errorsIn }, { dir }) {
  const file = path.join(dir, 'word-wordart.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [
    { text: 'WordArt', style: 'Heading1' },
    { text: 'The words above are WordArt, laid along an arch.' },
    { text: 'The end.' },
  ] }));
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const clickAt = async (x, y) => {
      wc.sendInputEvent({ type: 'mouseDown', x: Math.round(x), y: Math.round(y), button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: Math.round(x), y: Math.round(y), button: 'left', clickCount: 1 });
      await wait(400);
    };
    const rect = (selector) => js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r ? { left: r.left, top: r.top, width: r.width, height: r.height } : null; })()`);

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the paragraphs', 8000);
    const para = await rect('.wd-page [data-block="1"]');
    await clickAt(para.left + 14, para.top + 8);

    // Insert → WordArt → the first style.
    await clickTab('Insert');
    await wait(200);
    await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('WordArt')); b?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b?.click(); return 1; })()`);
    await until(() => js(`[...document.querySelectorAll('.rw-menu button')].some((b) => b.textContent.trim() === 'Fill: blue, shadow')`), 'the WordArt styles', 3000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Fill: blue, shadow')?.click(), 1`);
    const made = await until(() => js(`Boolean(document.querySelector('.wd-textbox.editable .wd-block'))`), 'the WordArt on the page', 6000).then(() => true).catch(() => false);

    // Shape Format → Transform → Arch Up, the caret still in the box.
    await until(() => js(`[...document.querySelectorAll('.rw-tab')].some((t) => t.textContent.trim() === 'Shape Format')`), 'the Shape Format tab', 4000).catch(() => {});
    await clickTab('Shape Format');
    await until(() => js(`Boolean(document.querySelector('.wd-warp[data-preset="textArchUp"]:not(:disabled)'))`), 'the Transform buttons', 3000).catch(() => {});
    await js(`document.querySelector('.wd-warp[data-preset="textArchUp"]')?.click(), 1`);
    const warped = await until(() => js(`Boolean(document.querySelector('.wd-textbox.wd-warped'))`), 'the box transformed', 5000).then(() => true).catch(() => false);

    // The caret out in the body: the words along the arc, the straight ones unseen.
    const body = await rect('.wd-page [data-block="2"]');
    await clickAt(body.left + 14, body.top + 8);
    const arched = await until(() => js(`(() => {
      const box = document.querySelector('.wd-textbox.wd-warped');
      const art = box?.querySelector('.wd-warp-art');
      const t = art?.querySelector('textPath');
      const words = box?.querySelector('.wd-block');
      return Boolean(t && t.textContent.includes('Your words here') && getComputedStyle(art).display !== 'none' && words && getComputedStyle(words).opacity === '0');
    })()`), 'the words along the arc', 5000).then(() => true).catch(() => false);
    if (process.env.RUTBA_VERIFY_CAPTURE) { wc.invalidate(); await wait(500); fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'word-wordart-arch.png'), (await wc.capturePage()).toPNG()); }
    check('word: Shape Format → Transform → Arch Up lays a WordArt\'s words along an arc on the page',
      made && warped && arched, JSON.stringify({ made, warped, arched }));

    // The caret back in its words: straight, to be edited.
    const words = await rect('.wd-textbox.wd-warped .wd-block');
    await clickAt(words.left + words.width / 2, words.top + words.height / 2);
    const straight = await until(() => js(`(() => {
      const box = document.querySelector('.wd-textbox.wd-warped');
      return Boolean(box?.classList.contains('wd-warp-editing') && getComputedStyle(box.querySelector('.wd-warp-art')).display === 'none' && getComputedStyle(box.querySelector('.wd-block')).opacity === '1');
    })()`), 'the words straight to edit', 4000).then(() => true).catch(() => false);
    check('word: with the caret in a transformed WordArt its words show straight, to be edited', straight, String(straight));

    // Saved as Word writes it.
    await clickAt(body.left + 14, body.top + 8);
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    const saved = await until(() => {
      try {
        const reopened = openDocx(fs.readFileSync(file));
        const box = reopened.render({ pages: false }).blocks.flatMap((b) => b.textBoxes || [])[0];
        return box?.warp === 'textArchUp' && /<wps:bodyPr\b[^>]*><a:prstTxWarp prst="textArchUp"><a:avLst\/><\/a:prstTxWarp>/.test(reopened.doc.doc.xml);
      } catch { return false; }
    }, 'the save', 6000).then(() => true).catch(() => false);
    check('word: the saved document carries a:prstTxWarp first in the box\'s body properties, as Word writes it', saved, saved ? 'textArchUp' : 'not in the file');

    const complaints = await errorsIn(win);
    check('word: WordArt\'s Transform reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the WordArt Transform checks ran', false, err?.message || String(err));
  }
}
