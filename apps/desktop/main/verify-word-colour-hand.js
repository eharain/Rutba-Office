// Documents: Font colour as a person uses it, with the real mouse and keys.
// Words are chosen by a double-click and by a drag, Font colour is pressed,
// Red is pressed, and then the page's drawn colour and the engine are read.
// Then a click puts the caret at a line's end, a colour is chosen and words
// are typed; the face, a click elsewhere before typing, Text Highlight
// Colour, Shading and the right-click menu follow. Nothing here is done by
// script clicks: a press that moves the keyboard or drops the selection
// shows here and nowhere else.
//
// It reads both the old Font colour menu (a list of names) and the palette,
// so the same file runs on an older build to show what a person saw there.
// Run alone with RUTBA_VERIFY_ONLY=word-colour-hand.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { labelKeyed } from './verify-keys.js';

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
      paragraphs: [{ text: 'A first line' }, { text: 'The quick brown fox jumps' }, { text: 'Second line here' }, { text: 'Last line' },
        { text: 'Fourth paragraph words' }, { text: 'Fifth for undo' }, { text: 'Sixth to highlight' }, { text: 'Bullet zulu item' },
        { table: { rows: [['cellword here', 'other cell']] } }],
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
        : await centre(`((list) => list.find((b) => ${labelKeyed("Red")}(b)) || list.find((b) => b.querySelector('.rw-menu-label')?.textContent.trim() === 'Red'))([...document.querySelectorAll('.rw-menu button')])`);
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

    // 3. The caret at a line's end, red chosen, then words typed: the
    // paragraph is red (D-2026-11-11-11) and so are the words typed.
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
    check('word-colour-hand: with nothing selected, red chosen with the mouse colours the paragraph the caret is in and the words typed next',
      r3.pressed === 'pressed' && colourOfWord(2, 'Second line here')?.join() === r3.want && colourOfWord(2, 'typed')?.join() === r3.want && paint3?.length > 0 && paint3.every((c) => c === rgb(r3.want)),
      JSON.stringify({ build: r3.build, placed, opened: r3.opened, beforeTyping: before, text: runsOf(2).map((r) => [r.text, r.fontColour ?? null]), paint3, pressed: r3.pressed }));
    await capture(win, 'word-colour-hand-typed.png');

    // 4. A whole line chosen by a triple-click: every word of it, drawn red once the page has drawn the answer.
    const last = await wordAt(3, 'Last');
    await click(last ? { x: last.x, y: last.y } : null, 3);
    const line = await where();
    const r4 = await pickRed();
    await until(() => colourOfWord(3, 'Last line')?.join() === r4.want, 'the line in red', 4000).catch(() => {});
    await until(async () => (await paintOf(3, 'line'))?.every((c) => c === rgb(r4.want)), 'the page to draw the line red', 4000).catch(() => {});
    const paint4 = await js(`[...document.querySelectorAll('.wd-page [data-block="3"] span')].filter((s) => [...s.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim())).map((s) => getComputedStyle(s).color)`);
    check('word-colour-hand: a whole line chosen by a triple-click turns red, every word of it as drawn',
      r4.pressed === 'pressed' && colourOfWord(3, 'Last line')?.join() === r4.want && paint4?.length > 0 && paint4.every((c) => c === rgb(r4.want)),
      JSON.stringify({ build: r4.build, line, opened: r4.opened, engine: colourOfWord(3, 'Last line'), paint4, pressed: r4.pressed }));

    // 5. The Font colour button's own face, not its arrow. The palette's face
    // puts on the colour last chosen; the older build has one button, and
    // pressing it opens the list, so there the list opening is what it does.
    const firstWord = await wordAt(0, 'first');
    await click(firstWord ? { x: firstWord.x, y: firstWord.y } : null, 2);
    const face = await centre(`document.querySelector('.rw-ribbon .rw-split[data-kind="word.font"] .rw-split-face')`);
    if (face && !face.hidden) {
      await click(face);
      await until(() => colourOfWord(0, 'first')?.join() === 'FF0000', 'the face to colour the word', 4000).catch(() => {});
      await until(async () => (await paintOf(0, 'first'))?.every((c) => c === rgb('FF0000')), 'the page to draw it', 4000).catch(() => {});
      const paint5 = await paintOf(0, 'first');
      check('word-colour-hand: the Font colour face (not its arrow) puts the last colour, red, on the chosen word',
        colourOfWord(0, 'first')?.join() === 'FF0000' && paint5?.length > 0 && paint5.every((c) => c === rgb('FF0000')),
        JSON.stringify({ build: 'palette', engine: colourOfWord(0, 'first'), paint5, after: await where() }));
    } else {
      const one = await centre(`[...document.querySelectorAll('.rw-ribbon button')].find((b) => (b.dataset.tip || b.title || '') === 'Font colour')`);
      if (one && !one.hidden) await click(one);
      const listed = await js(`[...document.querySelectorAll('.rw-menu .rw-menu-label')].map((n) => n.textContent.trim())`);
      wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
      wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
      await wait(200);
      check('word-colour-hand: the older build\'s single Font colour button opens its list of colours when pressed (it has no face of its own)',
        listed.includes('Red'), JSON.stringify({ build: 'menu', listed, engine: colourOfWord(0, 'first') }));
    }

    // 6. A colour chosen with nothing selected colours the paragraph the
    // caret is in (D-2026-11-11-11), the caret staying put. Then a click
    // elsewhere in that paragraph before typing: the words take its red.
    const z = async () => { wc.sendInputEvent({ type: 'keyDown', keyCode: 'Z', modifiers: ['control'] }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'Z', modifiers: ['control'] }); await wait(400); };
    const end4 = await wordAt(4, 'words');
    await click(end4 ? { x: end4.after, y: end4.y } : null);
    const r6 = await pickRed();
    await until(() => colourOfWord(4, 'Fourth paragraph words')?.join() === r6.want, 'the paragraph in red', 4000).catch(() => {});
    await until(async () => (await paintOf(4, 'Fourth'))?.every((c) => c === rgb(r6.want)), 'the page to draw it', 4000).catch(() => {});
    const coloured = { engine: colourOfWord(4, 'Fourth paragraph words'), paint: await paintOf(4, 'Fourth'), at: await where() };
    check('word-colour-hand: with nothing selected, a colour chosen with the mouse colours the whole paragraph the caret is in, the caret staying put, in the engine and as drawn',
      r6.pressed === 'pressed' && coloured.engine?.join() === r6.want && coloured.paint?.length > 0 && coloured.paint.every((c) => c === rgb(r6.want)) && coloured.at.engine?.collapsed !== false,
      JSON.stringify({ build: r6.build, ...coloured, neighbours: [colourOfWord(3, 'Last'), colourOfWord(5, 'Fifth')] }));
    const fourth = await wordAt(4, 'Fourth');
    await click(fourth ? { x: fourth.after, y: fourth.y } : null);
    await keys('!');
    await until(() => runsOf(4).map((r) => r.text).join('').includes('!'), 'the words typed', 4000).catch(() => {});
    await wait(200);
    check('word-colour-hand: a click elsewhere in the coloured paragraph, then typing: the words typed are red too',
      runsOf(4).map((r) => r.text).join('').includes('!') && colourOfWord(4, '!')?.join() === r6.want, JSON.stringify({ text: runsOf(4).map((r) => [r.text, r.fontColour ?? null]) }));

    // 6b. One Ctrl+Z takes the paragraph's colour back.
    const fifth = await wordAt(5, 'undo');
    await click(fifth ? { x: fifth.x, y: fifth.y } : null);
    const r6b = await pickRed();
    await until(() => colourOfWord(5, 'Fifth for undo')?.join() === r6b.want, 'the paragraph in red', 4000).catch(() => {});
    const red6b = colourOfWord(5, 'Fifth for undo');
    await z();
    await until(() => colourOfWord(5, 'Fifth for undo')?.join() === '', 'the colour undone', 4000).catch(() => {});
    await until(async () => (await paintOf(5, 'Fifth'))?.every((c) => c !== rgb(r6b.want)), 'the page to draw it', 4000).catch(() => {});
    const undone = { engine: colourOfWord(5, 'Fifth for undo'), paint: await paintOf(5, 'Fifth') };
    check('word-colour-hand: one Ctrl+Z takes back the colour a pick with nothing selected put on the paragraph',
      red6b?.join() === r6b.want && undone.engine?.every((c) => c === null) && undone.paint?.every((c) => c !== rgb(r6b.want)),
      JSON.stringify({ red: red6b, ...undone }));

    // 6c. Text Highlight Colour with nothing selected highlights the whole paragraph.
    const sixth = await wordAt(6, 'to');
    await click(sixth ? { x: sixth.x, y: sixth.y } : null);
    const hl6Arrow = await centre(`document.querySelector('.rw-ribbon .rw-split[data-kind="word.highlight"] .rw-split-arrow')`);
    let hl6 = 'no highlight arrow';
    if (hl6Arrow && !hl6Arrow.hidden) {
      await click(hl6Arrow);
      await until(() => js(`Boolean(document.querySelector('.rw-colours'))`), 'the highlights', 3000).catch(() => {});
      const yellow6 = await centre(`document.querySelector('.rw-colours .rw-cell[data-value="yellow"]')`);
      hl6 = yellow6 && !yellow6.hidden ? await click(yellow6) : 'no Yellow';
    }
    const lit = () => [...new Set(runsOf(6).map((r) => r.highlight ?? null))];
    await until(() => lit().join() === 'yellow', 'the paragraph highlighted', 4000).catch(() => {});
    check('word-colour-hand: with nothing selected, Text Highlight Colour highlights the whole paragraph the caret is in',
      hl6 === 'pressed' && lit().join() === 'yellow', JSON.stringify({ pressed: hl6, engine: runsOf(6).map((r) => [r.text, r.highlight ?? null]) }));

    // 7. Text Highlight Colour on a word chosen by a double-click.
    const second = await wordAt(2, 'Second');
    await click(second ? { x: second.x, y: second.y } : null, 2);
    const hlArrow = await centre(`document.querySelector('.rw-ribbon .rw-split[data-kind="word.highlight"] .rw-split-arrow')`);
    const hlOld = hlArrow ? null : await centre(`[...document.querySelectorAll('.rw-ribbon button')].find((b) => (b.dataset.tip || b.title || '') === 'Text highlight colour')`);
    const hlButton = hlArrow || hlOld;
    let hlPressed = 'no Text highlight colour button';
    if (hlButton && !hlButton.hidden) {
      await click(hlButton);
      await until(() => js(`Boolean(document.querySelector('.rw-colours, .rw-menu'))`), 'the highlights', 3000).catch(() => {});
      const yellow = hlArrow
        ? await centre(`document.querySelector('.rw-colours .rw-cell[data-value="yellow"]')`)
        : await centre(`((list) => list.find((b) => ${labelKeyed("Yellow")}(b)) || list.find((b) => b.querySelector('.rw-menu-label')?.textContent.trim() === 'Yellow'))([...document.querySelectorAll('.rw-menu button')])`);
      hlPressed = yellow && !yellow.hidden ? await click(yellow) : 'no Yellow';
    }
    const highlightOf = () => {
      const text = runsOf(2).map((r) => r.text).join('');
      let at = 0;
      const out = [];
      for (const r of runsOf(2)) {
        const end = at + r.text.length;
        if (end > text.indexOf('Second') && at < text.indexOf('Second') + 6) out.push(r.highlight ?? null);
        at = end;
      }
      return [...new Set(out)];
    };
    await until(() => highlightOf().join() === 'yellow', 'the word highlighted', 4000).catch(() => {});
    await wait(300);
    const back7 = await js(`(() => {
      const b = document.querySelector('.wd-page [data-block="2"]');
      const s = b && [...b.querySelectorAll('span')].find((n) => [...n.childNodes].some((c) => c.nodeType === 3 && c.textContent.includes('Second')));
      if (!s) return null;
      for (let n = s; n && n !== b; n = n.parentElement) { const c = getComputedStyle(n).backgroundColor; if (c && c !== 'rgba(0, 0, 0, 0)' && c !== 'transparent') return c; }
      return getComputedStyle(s).backgroundColor;
    })()`);
    check('word-colour-hand: Text Highlight Colour, Yellow, highlights a word chosen by a double-click, in the engine and as drawn',
      hlPressed === 'pressed' && highlightOf().join() === 'yellow' && /^rgba?\(255, 255, 0/.test(back7 || ''),
      JSON.stringify({ build: hlArrow ? 'palette' : 'menu', pressed: hlPressed, engine: highlightOf(), back7 }));

    // 8. Shading behind the paragraph the caret is in.
    const lastWord = await wordAt(3, 'Last');
    await click(lastWord ? { x: lastWord.x, y: lastWord.y } : null);
    const shadeButton = await centre(`[...document.querySelectorAll('.rw-ribbon button')].find((b) => (b.dataset.tip || b.title || '') === 'Shading: a colour behind the paragraph')`);
    let shadePressed = 'no Shading button';
    if (shadeButton && !shadeButton.hidden) {
      await click(shadeButton);
      await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the shades', 3000).catch(() => {});
      const blue = await centre(`((list) => list.find((b) => ${labelKeyed("Light blue")}(b)) || list.find((b) => b.querySelector('.rw-menu-label')?.textContent.trim() === 'Light blue'))([...document.querySelectorAll('.rw-menu button')])`);
      shadePressed = blue && !blue.hidden ? await click(blue) : 'no Light blue';
    }
    const shadingOf = () => String(model().blocks[3]?.shading || '').replace(/^#/, '').toUpperCase();
    await until(() => shadingOf() === 'DEEBF7', 'the paragraph shaded', 4000).catch(() => {});
    await wait(300);
    const back8 = await js(`(() => { const b = document.querySelector('.wd-page [data-block="3"]'); return b ? getComputedStyle(b).backgroundColor : null; })()`);
    await capture(win, 'word-colour-hand-shading.png');
    check('word-colour-hand: Shading, Light blue, shades the paragraph the caret is in, in the engine and as drawn',
      shadePressed === 'pressed' && shadingOf() === 'DEEBF7' && back8 === rgb('DEEBF7'),
      JSON.stringify({ pressed: shadePressed, engine: model().blocks[3]?.shading ?? null, back8 }));

    // 9. The right-click menu on a chosen word: what it offers, and if it
    // offers a colour, that the colour goes on. No build has a mini toolbar.
    const fox2 = await wordAt(1, 'fox');
    await click(fox2 ? { x: fox2.x, y: fox2.y } : null, 2);
    if (fox2) {
      wc.sendInputEvent({ type: 'mouseDown', x: fox2.x, y: fox2.y, button: 'right', clickCount: 1 });
      wc.sendInputEvent({ type: 'mouseUp', x: fox2.x, y: fox2.y, button: 'right', clickCount: 1 });
    }
    await wait(300);
    if (fox2 && !(await js(`Boolean(document.querySelector('.rw-menu'))`))) wc.sendInputEvent({ type: 'contextMenu', x: fox2.x, y: fox2.y });
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the right-click menu', 3000).catch(() => {});
    const offered = await js(`[...document.querySelectorAll('.rw-menu .rw-menu-label')].map((n) => n.textContent.trim())`);
    const colourItem = offered.find((l) => /colou?r/i.test(l));
    let viaMenu = null;
    if (colourItem) {
      await click(await centre(`((list) => list.find((b) => ${labelKeyed(colourItem)}(b)) || list.find((b) => b.querySelector('.rw-menu-label')?.textContent.trim() === ${JSON.stringify(colourItem)}))([...document.querySelectorAll('.rw-menu button')])`));
      viaMenu = await where();
    }
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await wait(200);
    const mini = await js(`Boolean(document.querySelector('.wd-mini, .rw-minibar, .rw-mini-toolbar, [data-role="mini-toolbar"]'))`);
    check('word-colour-hand: the right-click menu opens on a chosen word; it lists what it offers (a colour, if any, is reported), and there is no mini toolbar',
      offered.length > 0 && !mini, JSON.stringify({ offered, colourItem: colourItem || null, viaMenu, mini }));

    // 9b. A word in a bulleted list item: the colour goes on that word, not
    // one letter to the right. 2.24.0 counted the bullet as a character, so
    // "india" chosen in "• List india item" came out as "ndia i" in red.
    const zuluBlock = () => (model().blocks || []).findIndex((b) => (b.text || '').includes('zulu'));
    await click(await wordAt(zuluBlock(), 'zulu'));
    await click(await centre(`[...document.querySelectorAll('.rw-ribbon button')].find((b) => (b.dataset.tip || b.title || '') === 'Bullets')`));
    const hasBullet = () => js(`Boolean(document.querySelector('.wd-page [data-block="${zuluBlock()}"] .wd-marker'))`);
    await until(hasBullet, 'the bullet', 3000).catch(() => {});
    const bulleted = await hasBullet();
    await click(await wordAt(zuluBlock(), 'zulu'), 2);
    const listChosen = await where();
    const r9b = await pickRed();
    await until(() => colourOfWord(zuluBlock(), 'zulu')?.join() === r9b.want, 'the word in red', 4000).catch(() => {});
    await wait(300);
    const listRuns = runsOf(zuluBlock()).map((r) => [r.text, r.fontColour ?? null]);
    const redText = listRuns.filter(([, c]) => c === r9b.want).map(([t]) => t).join('');
    check('word-colour-hand: a word chosen in a bulleted list item turns red, that word and not one letter to the right',
      bulleted && r9b.pressed === 'pressed' && redText.trim() === 'zulu' && (await paintOf(zuluBlock(), 'zulu')).every((c) => c === rgb(r9b.want)),
      JSON.stringify({ bulleted, chosen: listChosen, runs: listRuns }));

    // 10. Ctrl+A, pressed in a table cell, then Bold, Font colour, a
    // highlight, Grow font and Centre: each goes on the whole document, and
    // the selection stays on the page and in the engine after every one.
    // Before, Ctrl+A was the browser's alone: the page showed everything
    // selected, the engine kept the old caret, and a pick went to the caret
    // and dropped the selection.
    const allRuns = () => {
      const out = [];
      const visit = (o) => {
        if (!o || typeof o !== 'object') return;
        if (Array.isArray(o)) return o.forEach(visit);
        if (typeof o.text === 'string' && o.text.trim() && !Array.isArray(o.runs) && !Array.isArray(o.blocks)) out.push(o);
        for (const v of Object.values(o)) visit(v);
      };
      visit(model().blocks);
      return out;
    };
    const allParas = () => (model().blocks || []).filter((b) => (b.runs || []).some((r) => r.text?.trim()));
    const held = async () => {
      const w = await where();
      const s = w.engine;
      const last = await js(`Math.max(...[...document.querySelectorAll('.wd-page [data-block]')].map((n) => Number(n.dataset.block)).filter(Number.isFinite))`);
      return { dom: w.dom.text.length, last, engine: s ? [s.from, s.to, s.collapsed] : null, whole: Boolean(s && !s.collapsed && s.from.block === 0 && s.from.offset === 0 && s.to.block >= last) && w.dom.text.length > 100 };
    };
    const ctrlA = async () => { wc.sendInputEvent({ type: 'keyDown', keyCode: 'A', modifiers: ['control'] }); wc.sendInputEvent({ type: 'keyUp', keyCode: 'A', modifiers: ['control'] }); await wait(500); };
    const ribbonPress = async (title) => click(await centre(`[...document.querySelectorAll('.rw-ribbon button')].find((b) => (b.dataset.tip || b.title || '') === ${JSON.stringify(title)})`));
    const cellWord = await js(`(() => { const walk = document.createTreeWalker(document.querySelector('.wd-page'), NodeFilter.SHOW_TEXT); for (let n = walk.nextNode(); n; n = walk.nextNode()) { const at = n.textContent.indexOf('cellword'); if (at < 0) continue; const r = document.createRange(); r.setStart(n, at); r.setEnd(n, at + 8); const b = r.getBoundingClientRect(); return { x: Math.round(b.left + b.width / 2), y: Math.round(b.top + b.height / 2), inCell: Boolean(n.parentElement.closest('td, .wd-cell')) }; } return null; })()`);
    await click(cellWord ? { x: cellWord.x, y: cellWord.y } : null);
    await ctrlA();
    const selected = await held();
    check('word-colour-hand: Ctrl+A in a table cell selects the whole document in the engine, as the page shows it',
      selected.whole && cellWord?.inCell === true, JSON.stringify({ ...selected, inCell: cellWord?.inCell ?? null }));

    const steps = [
      ['Bold', async () => ribbonPress('Bold (Ctrl+B)'), () => allRuns().every((r) => r.bold === true)],
      ['Font colour', async () => pickRed(), () => allRuns().every((r) => r.fontColour === 'FF0000')],
      ['Text highlight colour', async () => {
        await click(await centre(`document.querySelector('.rw-ribbon .rw-split[data-kind="word.highlight"] .rw-split-arrow')`));
        await until(() => js(`Boolean(document.querySelector('.rw-colours'))`), 'the highlights', 3000).catch(() => {});
        await click(await centre(`document.querySelector('.rw-colours .rw-cell[data-value="yellow"]')`));
      }, () => allRuns().every((r) => r.highlight === 'yellow')],
      ['Grow font', async () => ribbonPress('Grow font'), () => { const s = allRuns().map((r) => r.fontSize); return s.every((x) => Number(x) > 11); }],
      ['Centre', async () => ribbonPress('Centre (Ctrl+E)'), () => allParas().every((b) => b.align === 'center')],
    ];
    for (const [name, press, done] of steps) {
      await press();
      await until(() => done(), `${name} on everything`, 4000).catch(() => {});
      await wait(250);
      const after = await held();
      check(`word-colour-hand: after Ctrl+A, ${name} goes on the whole document and the selection stays`,
        done() && after.whole,
        JSON.stringify({ after, sample: allRuns().slice(0, 3).map((r) => [r.text.slice(0, 12), r.bold, r.fontColour, r.highlight, r.fontSize]), aligns: [...new Set(allParas().map((b) => b.align ?? null))] }));
    }

    // 10b. The ribbon's Select All is the same selection: Bold again takes it off everywhere.
    await click(await wordAt(1, 'quick'));
    await ribbonPress('Select all (Ctrl+A)');
    await wait(300);
    const viaRibbon = await held();
    await ribbonPress('Bold (Ctrl+B)');
    await until(() => allRuns().every((r) => !r.bold), 'bold off everywhere', 4000).catch(() => {});
    check('word-colour-hand: the ribbon\'s Select All selects the same whole document, and Bold then takes bold off everywhere, the selection staying',
      viaRibbon.whole && allRuns().every((r) => !r.bold) && (await held()).whole, JSON.stringify({ viaRibbon, after: await held() }));
    await capture(win, 'word-colour-hand-select-all.png');
  } catch (err) {
    check('word-colour-hand: the checks ran', false, err.stack || err.message);
  } finally {
    try { win?.close(); } catch { /* gone */ }
  }
}
