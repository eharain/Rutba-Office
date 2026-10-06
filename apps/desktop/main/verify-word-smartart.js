// Documents: Insert → SmartArt, pressed on the ribbon — Basic Chevron
// Process picked, three stages typed in its text pane: one group of
// chevrons holding the words, in a paragraph of its own after the caret's,
// drawn on the page; a click in a chevron's words puts the caret there and
// typing changes them; and the saved document keeps the diagram as a Word
// group of shapes holding their words.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';

/**
 * @param {object} h the harness: open, check, until, wait, doc
 */
export async function verifyWordSmartArt(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'word-smartart.docx');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'The stages follow.' }, { text: 'The end.' }] }));
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG()); };
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const group = () => (model().blocks || []).flatMap((b) => b.groups || []).find((g) => /^Basic Chevron Process \d+$/.test(g.name || ''));
    const textOf = (i) => (model().blocks || []).find((b) => b.index === i)?.text ?? null;

    await until(() => js(`document.querySelectorAll('.wd-page [data-block]').length >= 2`), 'the page', 8000);
    const pressed = await js(`(async () => {
      [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click();
      await new Promise((r) => setTimeout(r, 200));
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'SmartArt');
      if (!b || b.disabled) return 'no button'; b.click(); return 'clicked';
    })()`);
    await until(() => js(`Boolean(document.querySelector('.sa-ok'))`), 'the SmartArt box', 4000).catch(() => {});
    await js(`document.querySelector('.sa-item[data-layout="chevron"]')?.click(), 1`);
    await js(`(() => { const ta = document.querySelector('.sa-text'); const set = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set; set.call(ta, 'Plan\\nBuild\\nShip'); ta.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
    await wait(150);
    await js(`document.querySelector('.sa-ok')?.click(), 1`);
    const made = await until(() => {
      const g = group();
      return g && g.members.filter((m) => m.kind === 'textbox' && m.geom === 'chevron').length === 3;
    }, 'the diagram', 8000).then(() => true).catch(() => false);
    const blocks = (group()?.members || []).map((m) => m.blocks?.[0]);
    const words = blocks.map(textOf);
    const drawn = await until(() => js(`(() => {
      const g = document.querySelector('.wd-page .wd-group');
      return Boolean(g) && g.querySelectorAll('img.wd-member').length === 3 && ['Plan', 'Build', 'Ship'].every((w) => g.textContent.includes(w));
    })()`), 'the diagram on the page', 5000).then(() => true).catch(() => false);
    await capture('word-smartart.png');
    check('documents: Insert → SmartArt → Basic Chevron Process puts the stages typed in as one group of chevrons holding the words, after the paragraph, drawn on the page',
      pressed === 'clicked' && made && words.join(',') === 'Plan,Build,Ship' && drawn, JSON.stringify({ pressed, made, words, drawn, group: group() && { name: group().name, w: group().widthPx, h: group().heightPx } }));

    // The words are the document's: a click in one, and typing changes it.
    // On the word itself: a chevron's box reaches into the next one's notch, which is drawn over it.
    const at = await js(`(() => { const b = document.querySelector('.wd-page .wd-group [data-block="${blocks[1]}"]'); if (!b) return null; const range = document.createRange(); range.selectNodeContents(b); const r = range.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    if (at) {
      wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    }
    await until(() => model().selection?.focus?.block === blocks[1], 'the caret in the chevron', 4000).catch(() => {});
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'End' }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'End' });
    await wait(120);
    wc.insertText('ing');
    const typed = await until(() => textOf(blocks[1]) === 'Building', 'the typed words', 5000).then(() => true).catch(() => false);
    const hit = await js(`(() => { const el = document.elementFromPoint(${at?.x ?? 0}, ${at?.y ?? 0}); const s = window.getSelection(); return { el: el?.tagName + '.' + el?.className, block: el?.closest?.('[data-block]')?.dataset.block, sel: s.anchorNode?.parentElement?.closest?.('[data-block]')?.dataset.block, blocks: ${JSON.stringify(blocks)} }; })()`);
    check('documents: a click in a diagram shape\'s words puts the caret there, and typing changes them', Boolean(at) && typed, JSON.stringify({ at, typed, now: textOf(blocks[1]), focus: model().selection?.focus, hit }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const xml = OoxmlPackage.read(fs.readFileSync(file)).text('word/document.xml');
        return /<wp:docPr id="\d+" name="Basic Chevron Process \d+"\/>/.test(xml) && /<wpg:wgp>/.test(xml)
          && (xml.match(/<a:prstGeom prst="chevron">/g) || []).length === 3 && /<w:t xml:space="preserve">Building<\/w:t>/.test(xml);
      } catch { return false; }
    }, 'the saved diagram', 8000).then(() => true).catch(() => false);
    check('documents: the saved document keeps the diagram as a Word group of shapes holding their words', saved, saved ? 'wpg:wgp of chevrons in the file' : 'not in the file');
    await wait(50);
  } catch (err) {
    check('documents: the SmartArt checks ran', false, err.message);
  }
}
