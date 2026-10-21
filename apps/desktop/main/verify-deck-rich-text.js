// Presentations: the words chosen in a box formatted on their own (owner,
// 2026-10-21: "the entire text box has single formatting and does not allow
// formatting selected text separately; any change applies to entire text
// box"). The check opens a title, chooses one word and presses Bold on the
// ribbon and Ctrl+I, types more at the end, leaves the box, and reads the
// runs in the engine and in the saved file; then the ribbon's Bold on the
// whole box, whose words are partly bold, makes every word bold.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx, Deck } from '@rutba/presentation';

/** In-page: choose the characters from `a` to `b` in the open editor's words. */
const choose = (a, b) => `(() => {
  const el = document.querySelector('.sl-editor');
  el.focus();
  const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  const at = (n) => { let left = n; let t; while ((t = w.nextNode())) { if (left <= t.data.length) return [t, left]; left -= t.data.length; } return null; };
  const s0 = at(${a}); w.currentNode = el; const s1 = at(${b});
  const r = document.createRange(); r.setStart(...s0); r.setEnd(...s1);
  const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  return 1;
})()`;

export async function verifyDeckRichText(h, { dir }) {
  const { open, check, until, wait, doc } = h;
  const file = path.join(dir, 'rich-words.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: 'Words', slides: [{ layout: 'title', title: 'Price list today', body: 'Subtitle' }] }));
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.sl-hit[data-shape]').length > 0`), 'the slide\'s shapes', 8000);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const titleRuns = () => {
      const shape = (doc.model({ id: session.id }).slide?.shapes || []).find((s) => (s.text?.paragraphs || []).some((p) => /^Price/.test(p.plain)));
      return shape ? shape.text.paragraphs.flatMap((p) => p.runs.map((r) => ({ text: r.text, bold: Boolean(r.bold), italic: Boolean(r.italic) }))) : null;
    };
    const bold = () => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => /^Bold/.test(n.dataset.tip || '')); if (!b) return 'no button'; b.click(); return 'pressed'; })()`);
    const boldShown = () => js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => /^Bold/.test(n.dataset.tip || ''))?.getAttribute('aria-pressed')`);

    // A double-click as a hand makes one: a press that chooses the box, then the double-click that opens it.
    await js(`(() => { const hit = document.querySelector('.sl-hit[data-shape]'); hit?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); hit?.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, button: 0 })); return 1; })()`);
    await wait(150);
    await js(`(() => { document.querySelector('.sl-hit[data-shape]')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
    const opened = await until(() => js(`Boolean(document.querySelector('.sl-editor'))`), 'the editor', 4000).then(() => true, () => false);

    // "list" chosen, Bold from the ribbon, then Ctrl+I.
    await js(choose(6, 10));
    await wait(150);
    const pressed = await bold();
    await wait(150);
    await js(`(() => { document.querySelector('.sl-editor').dispatchEvent(new KeyboardEvent('keydown', { key: 'i', ctrlKey: true, bubbles: true, cancelable: true })); return 1; })()`);
    await wait(150);
    const shown = await js(`[...document.querySelectorAll('.sl-editor span')].map((s) => ({ text: s.textContent, bold: /700/.test(s.style.fontWeight), italic: s.style.fontStyle === 'italic' }))`);
    const ribbonOnWord = await boldShown();
    await js(choose(0, 5));
    await wait(200);
    const ribbonOffWord = await boldShown();
    check('presentations: a word chosen in a box is made bold by the ribbon and italic by Ctrl+I, the words round it left as they were, and the ribbon shows the look of the words chosen',
      opened && pressed === 'pressed'
        && JSON.stringify(shown) === JSON.stringify([{ text: 'Price ', bold: false, italic: false }, { text: 'list', bold: true, italic: true }, { text: ' today', bold: false, italic: false }])
        && ribbonOnWord === 'true' && ribbonOffWord === 'false',
      JSON.stringify({ opened, pressed, shown, ribbonOnWord, ribbonOffWord }));

    // More typed at the end, then Esc: the runs reach the engine as they were made.
    await js(`(() => { const el = document.querySelector('.sl-editor'); el.focus(); const s = getSelection(); s.selectAllChildren(el); s.collapseToEnd(); document.execCommand('insertText', false, ' now'); return 1; })()`);
    await wait(150);
    await js(`(() => { document.querySelector('.sl-editor')?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })); return 1; })()`);
    const kept = await until(() => JSON.stringify(titleRuns()) === JSON.stringify([{ text: 'Price ', bold: false, italic: false }, { text: 'list', bold: true, italic: true }, { text: ' today now', bold: false, italic: false }]), 'the runs in the engine', 5000).then(() => true, () => false);
    await doc.save({ id: session.id, path: file });
    const xml = (() => { try { const d = Deck.open(fs.readFileSync(file)); return d.pkg.text(d.slideParts[0].part); } catch { return ''; } })();
    const inFile = /<a:rPr[^>]*\bb="1"[^>]*>(?:(?!<\/a:r>).)*<a:t[^>]*>list<\/a:t>/s.test(xml) && /<a:rPr[^>]*\bi="1"[^>]*>(?:(?!<\/a:r>).)*<a:t[^>]*>list<\/a:t>/s.test(xml)
      && !/<a:rPr[^>]*\bb="1"[^>]*>(?:(?!<\/a:r>).)*<a:t[^>]*>Price <\/a:t>/s.test(xml);
    check('presentations: leaving the box keeps each word\'s own look, words typed after it take the look before them, and the saved file has the bold italic word alone',
      kept && inFile, JSON.stringify({ runs: titleRuns(), inFile, xml: inFile ? undefined : (xml.match(/<a:r>.*?<\/a:r>/gs) || []).slice(0, 4) }));

    // The box chosen, not its words: Bold makes all of it bold, the partly bold words included.
    await wait(300);
    await js(`(() => { document.querySelector('.sl-hit[data-shape]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); document.querySelector('.sl-hit[data-shape]')?.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, button: 0 })); return 1; })()`);
    await wait(200);
    await bold();
    const allBold = await until(() => (titleRuns() || []).every((r) => r.bold), 'the whole title bold', 4000).then(() => true, () => false);
    check('presentations: Bold on a box whose words are partly bold makes every word bold, not each the other way',
      allBold, JSON.stringify(titleRuns()));
  } catch (err) {
    check('presentations: the formatting of chosen words ran', false, err.message);
  }
}
