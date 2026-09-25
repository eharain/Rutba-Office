// Presentation: Review → Comments.
//
// On a deck of its own (three slides, written beside the shared fixtures):
// New Comment with nothing selected opens the Comments pane with a draft,
// and Post writes a thread on the slide — shown as a card, a marker on the
// stage and a count in the strip; a comment on the selected title is
// anchored to it and its marker sits at the title's corner; a reply, then
// Resolve; Next and Previous walk the threads across the slides; Delete
// takes one away and Ctrl+Z brings it back. The saved file carries the
// modern comments part and the authors part, signed with the window's own
// user name. Run alone with RUTBA_VERIFY_ONLY=deckcomments.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { consoleMessage } from './console-message.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where to write the block's own deck
 */
export async function verifyDeckComments(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  try {
    const file = path.join(dir, 'comments.pptx');
    fs.writeFileSync(file, buildPptx({
      title: 'Comment checks',
      slides: [
        { layout: 'title', title: 'Budget review', body: 'Draft for comment' },
        { layout: 'obj', title: 'Spending', body: ['Travel down', 'Tools up'] },
        { layout: 'obj', title: 'Next steps', body: ['Sign off'] },
      ],
    }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const consoleErrors = [];
    wc.on('console-message', (...args) => { const m = consoleMessage(args); if (m.level >= 3) consoleErrors.push(m.text.split('\n')[0].slice(0, 160)); });
    const id = () => sessionFor('deck').id;
    const model = (slide = 0) => doc.model({ id: id(), slide });
    const comments = () => model(0).comments || [];
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const tab = (label) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(label)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.click();
      return 'clicked';
    })()`);
    const clickMenuItem = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `the "${label}" menu item`, 3000).catch(() => {});
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'missing'; b.click(); return 'clicked'; })()`);
    };
    const clickIn = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'missing ' + ${JSON.stringify(selector)}; b.click(); return 'clicked'; })()`);
    const type = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no field ' + ${JSON.stringify(selector)};
      el.focus();
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'typed';
    })()`);
    const goSlide = async (i) => {
      await js(`(() => { document.querySelectorAll('.sl-thumb')[${i}]?.click(); return 1; })()`);
      await until(() => js(`document.querySelectorAll('.sl-thumb')[${i}]?.classList.contains('active')`), `slide ${i + 1}`, 4000).catch(() => {});
      await wait(250);
    };
    const activeSlide = () => js(`[...document.querySelectorAll('.sl-thumb')].findIndex((t) => t.classList.contains('active'))`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 3`), 'the slide sorter', 8000);
    const me = model(0).me;

    // 1. New Comment on the slide: the pane, a draft, Post.
    await tab('Review');
    const newed = await clickRibbon('New Comment');
    const drafting = await until(() => js(`document.activeElement?.classList.contains('sl-cm-input') && Boolean(document.querySelector('.sl-cm-draft'))`), 'the draft, focused', 5000).catch(() => false);
    const typed = await type('.sl-cm-draft .sl-cm-input', 'Is the draft date right?');
    const posted = await clickIn('.sl-cm-draft .sl-cm-post');
    const one = await until(() => comments().length === 1 && comments()[0].slide === 0 && comments()[0].shapeId == null, 'the comment', 5000).catch(() => false);
    const shown = await until(() => js(`document.querySelectorAll('.sl-cm-card[data-comment]').length === 1 && document.querySelectorAll('.sl-cm-marker').length === 1 && document.querySelector('.sl-thumb .sl-thumb-cm')?.dataset.comments === '1'`), 'the card, the marker and the count', 5000).catch(() => false);
    check('slides: New Comment opens the Comments pane with a draft; Post writes a thread on the slide, shown as a card, a marker and a count in the strip, signed with the user\'s name',
      newed === 'clicked' && drafting === true && typed === 'typed' && posted === 'clicked' && one === true && shown === true && comments()[0]?.author === me,
      `${newed}/${typed}/${posted}; drafting ${drafting}; ${JSON.stringify(comments().map((c) => [c.slide, c.author, c.text]))}; shown ${shown}; me ${me}`);
    await capture('slides-comments.png');

    // 2. On the title of slide 2: anchored to it, the marker at its corner.
    await goSlide(1);
    const titleId = model(1).slide.shapes.find((s) => s.placeholder?.type === 'title')?.id;
    await js(`(() => { document.querySelector('.sl-hit[data-shape="${titleId}"]')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-hit.selected[data-shape="${titleId}"]'))`), 'the title selected', 3000).catch(() => {});
    await clickRibbon('New Comment');
    await until(() => js(`Boolean(document.querySelector('.sl-cm-draft .sl-cm-input'))`), 'the draft', 4000).catch(() => {});
    await type('.sl-cm-draft .sl-cm-input', 'Shorter title?');
    await clickIn('.sl-cm-draft .sl-cm-post');
    const anchored = await until(() => comments().some((c) => c.slide === 1 && String(c.shapeId) === String(titleId)), 'the anchored comment', 5000).catch(() => false);
    const corner = await until(() => js(`(() => {
      const m = document.querySelector('.sl-cm-marker')?.getBoundingClientRect();
      const t = document.querySelector('.sl-hit[data-shape="${titleId}"]')?.getBoundingClientRect();
      return Boolean(m && t && Math.abs(m.right - t.right) < 40 && m.bottom <= t.top + 20);
    })()`), 'the marker at the title\'s corner', 5000).catch(() => false);
    check('slides: a comment made with the title selected is anchored to it, and its marker sits at the title\'s top right corner', anchored === true && corner === true, `anchored ${anchored}; corner ${corner}`);

    // 3. A reply, then Resolve.
    const thread = comments().find((c) => c.slide === 1);
    await js(`(() => { document.querySelector('.sl-cm-card[data-comment="${thread.id}"]')?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-cm-card.selected .sl-cm-replybox .sl-cm-input'))`), 'the reply box', 4000).catch(() => {});
    await type('.sl-cm-card.selected .sl-cm-replybox .sl-cm-input', 'Agreed — two words.');
    await clickIn('.sl-cm-card.selected .sl-cm-replybox .sl-cm-post');
    const replied = await until(() => comments().find((c) => c.id === thread.id)?.replies?.length === 1, 'the reply', 5000).catch(() => false);
    await until(() => js(`document.querySelectorAll('.sl-cm-card.selected .sl-cm-reply').length === 1`), 'the reply shown', 4000).catch(() => {});
    await capture('slides-comment-thread.png');
    await clickIn(`.sl-cm-card[data-comment="${thread.id}"] .sl-cm-resolve`);
    const resolved = await until(() => comments().find((c) => c.id === thread.id)?.status === 'resolved', 'resolved', 5000).catch(() => false);
    const faded = await until(() => js(`Boolean(document.querySelector('.sl-cm-card.resolved[data-comment="${thread.id}"]') && document.querySelector('.sl-cm-marker.resolved'))`), 'the card and marker resolved', 4000).catch(() => false);
    check('slides: a reply goes under the thread, and Resolve marks it resolved — the card and the marker dimmed', replied === true && resolved === true && faded === true, `reply ${replied}; resolved ${resolved}; drawn ${faded}`);

    // 4. Next and Previous walk the threads across the slides.
    const next = await clickRibbon('Next');
    const wrapped = await until(async () => (await activeSlide()) === 0, 'Next to slide 1', 4000).catch(() => false);
    const prev = await clickRibbon('Previous');
    const back = await until(async () => (await activeSlide()) === 1, 'Previous to slide 2', 4000).catch(() => false);
    check('slides: Next and Previous move between the comments across the slides, showing each one\'s slide', next === 'clicked' && wrapped === true && prev === 'clicked' && back === true, `${next}/${prev}; wrapped ${wrapped}; back ${back}`);

    // 5. Delete, then Ctrl+Z.
    const del = await clickRibbon('Delete —');
    const item = await clickMenuItem('Delete Comment');
    const deleted = await until(() => comments().length === 1 && comments()[0].slide === 0, 'the thread deleted', 5000).catch(() => false);
    await js(`(() => { document.querySelector('.sl-stage')?.focus(); return 1; })()`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Z', modifiers: ['control'] });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Z', modifiers: ['control'] });
    const restored = await until(() => comments().length === 2, 'the thread back', 5000).catch(() => false);
    check('slides: Review → Delete → Delete Comment takes the selected thread away, and Ctrl+Z brings it back with its reply',
      del === 'clicked' && item === 'clicked' && deleted === true && restored === true && comments().find((c) => c.slide === 1)?.replies.length === 1,
      `${del}/${item}; deleted ${deleted}; restored ${restored}`);

    // 6. Saved: the modern comments part and the authors part.
    await js(`(() => { document.querySelector('.sl-stage')?.focus(); return 1; })()`);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'S', modifiers: ['control'] });
    const inFile = () => { try { return Deck.open(fs.readFileSync(file)); } catch { return null; } };
    const saved = await until(() => (inFile()?.comments() || []).length === 2, 'the saved comments', 8000).catch(() => false);
    const d = inFile();
    const parts = d ? d.pkg.partNames().filter((p) => /modernComment_|authors\.xml/.test(p)) : [];
    check('slides: saved, the file carries PowerPoint 365\'s comments parts — one per slide commented on, and the authors — and opens with both threads',
      saved === true && parts.filter((p) => /modernComment_/.test(p)).length === 2 && parts.includes('ppt/authors.xml') && d.comments().every((c) => c.author === me),
      `saved ${saved}; parts ${parts.join(', ')}`);

    // 7. Show Comments closes the pane.
    const toggled = await clickRibbon('Show Comments');
    const closed = await until(() => js(`!document.querySelector('.sl-cm-pane')`), 'the pane closed', 3000).catch(() => false);
    check('slides: Show Comments closes the pane again, the markers staying on the slide', toggled === 'clicked' && closed === true && (await js(`document.querySelectorAll('.sl-cm-marker').length`)) >= 1, `${toggled}; closed ${closed}`);
    check('slides: the comment checks raised no errors in the window', consoleErrors.length === 0, consoleErrors.slice(0, 3).join(' | '));
  } catch (err) {
    check('slides: the comment checks ran', false, err.stack || err.message);
  }
}
