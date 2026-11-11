// Documents: Font colour as a person uses it, with the real mouse and keys.
// Words are chosen by a double-click and by a drag, Font colour is pressed,
// Red is pressed, and then the page's drawn colour and the engine are read.
// Then a click puts the caret at a line's end, a colour is chosen and words
// are typed. Nothing here is done by script clicks: a press that moves the
// keyboard or drops the selection shows here and nowhere else.
//
// It reads both the old Font colour menu (a list of names) and the palette,
// so the same file runs on an older build to show what a person saw there.
// Run alone with RUTBA_VERIFY_ONLY=word-colour-hand.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';

/**
 * @param {object} h the harness: open, check, until, wait, capture, doc
 */
export async function verifyWordColourByHand(h, { dir }) {
  const { open, check, until, wait, capture, doc } = h;
  const own = path.join(dir, 'word-colour-hand');
  const file = path.join(own, 'by-hand.docx');
  try {
    fs.mkdirSync(own, { recursive: true });
    fs.writeFileSync(file, buildDocx({
      styles: true,
      paragraphs: [{ text: 'A first line' }, { text: 'The quick brown fox jumps' }, { text: 'Second line here' }, { text: 'Last line' }],
    }));
  } catch (err) {
    return check('word-colour-hand: the fixture was made', false, err.message);
  }

  let win = null;
  try {
    win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    await until(async () => (await js(`document.querySelectorAll('.wd-page [data-block]').length`)) > 3, 'the page', 8000);
    await wait(400);
    const session = doc.sessions().filter((s) => s.kind === 'doc' && s.path === file).pop();
    const model = () => doc.model({ id: session.id });
    const runsOf = (i) => (model().blocks[i]?.runs || []).filter((r) => r.text);
    const rgb = (hex) => `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`;

    // Where a word is drawn: its first or last letter's box, or its middle.
    const wordAt = (block, word) => js(`(() => {
      const b = document.querySelector('.wd-page [data-block="${block}"]');
      if (!b) return null;
      const walk = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
      for (let n = walk.nextNode(); n; n = walk.nextNode()) {
        const at = n.textContent.indexOf(${JSON.stringify(word)});
        if (at < 0) continue;
        const r = document.createRange();
        r.setStart(n, at);
        r.setEnd(n, at + ${word.length});
        const box = r.getBoundingClientRect();
        return { left: Math.round(box.left + 1), right: Math.round(box.right - 1), x: Math.round(box.left + box.width / 2), y: Math.round(box.top + box.height / 2), after: Math.round(box.right + 2) };
      }
      return null;
    })()`);
    // What the page and the engine hold: the browser's selection, where the keyboard is, the engine's selection.
    const where = async () => ({
      dom: await js(`(() => { const s = getSelection(); return { text: String(s), collapsed: s.isCollapsed, inPage: Boolean(s.anchorNode && document.querySelector('.wd-page')?.contains(s.anchorNode)) }; })()`),
      focus: await js(`(() => { const a = document.activeElement; return a ? a.tagName.toLowerCase() + (a.className ? '.' + String(a.className).split(' ').join('.') : '') : null; })()`),
      engine: model().selection,
    });
    const mouse = async (type, point, clickCount = 1) => {
      wc.sendInputEvent({ type, x: point.x, y: point.y, button: 'left', clickCount, ...(type === 'mouseMove' ? { modifiers: ['leftButtonDown'] } : {}) });
    };
    const click = async (point, clickCount = 1) => {
      if (!point) return 'nowhere';
      for (let n = 1; n <= clickCount; n++) {
        await mouse('mouseDown', point, n);
        await mouse('mouseUp', point, n);
      }
      await wait(250);
      return 'pressed';
    };
    const drag = async (from, to) => {
      await mouse('mouseDown', from);
      for (let k = 1; k <= 6; k++) {
        await mouse('mouseMove', { x: Math.round(from.x + ((to.x - from.x) * k) / 6), y: Math.round(from.y + ((to.y - from.y) * k) / 6) });
        await wait(30);
      }
      await mouse('mouseUp', to);
      await wait(300);
    };
    const keys = async (text) => {
      for (const ch of text) {
        wc.sendInputEvent({ type: 'keyDown', keyCode: ch });
        wc.sendInputEvent({ type: 'char', keyCode: ch });
        wc.sendInputEvent({ type: 'keyUp', keyCode: ch });
        await wait(60);
      }
    };
    const centre = (selector) => js(`(() => { const n = ${selector}; if (!n) return null; const r = n.getBoundingClientRect(); return r.width ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : { hidden: true }; })()`);

    // Font colour, then Red: the palette's arrow and its Red square, or the
    // older build's Font colour button and the Red in its list.
    const pickRed = async () => {
      const palette = await centre(`document.querySelector('.rw-ribbon .rw-split[data-kind="word.font"] .rw-split-arrow')`);
      const old = palette ? null : await centre(`[...document.querySelectorAll('.rw-ribbon button')].find((b) => (b.dataset.tip || b.title || '') === 'Font colour')`);
      const button = palette || old;
      if (!button || button.hidden) return { pressed: 'no Font colour button on the ribbon', button };
      await click(button);
      await until(() => js(`Boolean(document.querySelector('.rw-colours, .rw-menu'))`), 'the colours', 3000).catch(() => {});
      const opened = await where();
      const red = palette
        ? await centre(`document.querySelector('.rw-colours .rw-cell.std[data-value="FF0000"]')`)
        : await centre(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.querySelector('.rw-menu-label')?.textContent.trim() === 'Red')`);
      if (!red || red.hidden) return { pressed: 'no Red', opened };
      await click(red);
      return { pressed: 'pressed', want: palette ? 'FF0000' : 'E03131', build: palette ? 'palette' : 'menu', opened };
    };
    // The drawn colour of every span holding `word` in a block.
    const paintOf = (block, word) => js(`(() => {
      const b = document.querySelector('.wd-page [data-block="${block}"]');
      if (!b) return null;
      return [...b.querySelectorAll('span')].filter((s) => [...s.childNodes].some((n) => n.nodeType === 3 && n.textContent.includes(${JSON.stringify(word)}))).map((s) => getComputedStyle(s).color);
    })()`);
    const colourOfWord = (block, word) => {
      let at = 0;
      const out = [];
      const text = runsOf(block).map((r) => r.text).join('');
      const from = text.indexOf(word);
      if (from < 0) return null;
      for (const r of runsOf(block)) {
        const end = at + r.text.length;
        if (end > from && at < from + word.length) out.push(r.fontColour ?? null);
        at = end;
      }
      return [...new Set(out)];
    };

    // 1. A word chosen by a double-click.
    const quick = await wordAt(1, 'quick');
    await click(quick ? { x: quick.x, y: quick.y } : null, 2);
    const chosen = await where();
    const r1 = await pickRed();
    await until(() => colourOfWord(1, 'quick')?.join() === r1.want, 'the word in red', 4000).catch(() => {});
    await wait(200);
    const after1 = await where();
    const paint1 = await paintOf(1, 'quick');
    await capture(win, 'word-colour-hand-word.png');
    check('word-colour-hand: a word chosen by a double-click turns red when Font colour and Red are pressed with the mouse, in the engine and as drawn',
      r1.pressed === 'pressed' && colourOfWord(1, 'quick')?.join() === r1.want && paint1?.length > 0 && paint1.every((c) => c === rgb(r1.want)),
      JSON.stringify({ build: r1.build, want: r1.want, chosen, opened: r1.opened, after: after1, engine: colourOfWord(1, 'quick'), paint1, pressed: r1.pressed }));

    // 2. Words chosen by a drag.
    const brown = await wordAt(1, 'brown');
    const fox = await wordAt(1, 'fox');
    if (brown && fox) await drag({ x: brown.left, y: brown.y }, { x: fox.right, y: fox.y });
    const dragged = await where();
    const r2 = await pickRed();
    await until(() => colourOfWord(1, 'brown fox')?.join() === r2.want, 'the dragged words in red', 4000).catch(() => {});
    await wait(200);
    const after2 = await where();
    const paint2 = await paintOf(1, 'fox');
    check('word-colour-hand: words chosen by a drag turn red the same way',
      r2.pressed === 'pressed' && colourOfWord(1, 'brown fox')?.join() === r2.want && paint2?.length > 0 && paint2.every((c) => c === rgb(r2.want)),
      JSON.stringify({ build: r2.build, dragged, opened: r2.opened, after: after2, engine: colourOfWord(1, 'brown fox'), paint2, pressed: r2.pressed }));

    // 3. The caret at a line's end, red chosen, then words typed.
    const here = await wordAt(2, 'here');
    await click(here ? { x: here.after, y: here.y } : null);
    const placed = await where();
    const r3 = await pickRed();
    await wait(300);
    const before = await where();
    await keys(' typed');
    await until(() => runsOf(2).map((r) => r.text).join('').endsWith(' typed'), 'the words typed', 4000).catch(() => {});
    await wait(200);
    const paint3 = await paintOf(2, 'typed');
    check('word-colour-hand: with nothing selected, red chosen with the mouse colours the words typed next',
      r3.pressed === 'pressed' && colourOfWord(2, 'typed')?.join() === r3.want && paint3?.length > 0 && paint3.every((c) => c === rgb(r3.want)),
      JSON.stringify({ build: r3.build, placed, opened: r3.opened, beforeTyping: before, text: runsOf(2).map((r) => [r.text, r.fontColour ?? null]), paint3, pressed: r3.pressed }));
    await capture(win, 'word-colour-hand-typed.png');
  } catch (err) {
    check('word-colour-hand: the checks ran', false, err.stack || err.message);
  } finally {
    try { win?.close(); } catch { /* gone */ }
  }
}
