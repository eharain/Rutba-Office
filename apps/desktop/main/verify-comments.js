// Worksheets: Review → comments, Excel 365's threaded kind.
//
// New Comment on B2 opens a card beside the cell; a comment posted there
// puts Excel's purple corner on the cell and signs the comment with the
// account's name and the time. A reply joins the thread, Resolve greys the
// corner and Reopen brings it back, a reply is edited from its menu. Show
// Comments lists every thread, filtered to the open or the resolved; a
// second thread on D4 and Previous walk between them; Delete takes a
// thread away. Saved, the file carries the threaded part, the person and
// the classic shadow Excel writes. Run alone with RUTBA_VERIFY_ONLY=comments.
import fs from 'node:fs';
import path from 'node:path';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const ROWS = [
  ['Item', 'Cost', 'Qty', 'Total'],
  ['Paint', 40, 2, '=B2*C2'],
  ['Brushes', 12, 3, '=B3*C3'],
  ['Tape', 4, 5, '=B4*C4'],
];

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} where the fixture is written
 */
export async function verifyComments(h, { dir }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'comments.xlsx');
  fs.writeFileSync(file, buildXlsx({ sheets: [{ name: 'Plan', rows: ROWS }] }));
  try {
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('sheet');
    const model = () => doc.model({ id: session.id });
    await until(() => js(`Boolean(document.querySelector('.sh-cell[data-ref="D4"]'))`), 'the grid', 8000);

    const clickIn = async (title) => {
      const find = `[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(title)}) && !n.disabled)`;
      await until(() => js(`Boolean(${find})`), `the ${title} button`, 3000).catch(() => {});
      return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(title)}; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked'; })()`);
    };
    const tab = (name) => js(`(() => { [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const select = async (ref) => {
      await js(`(() => { document.querySelector('.sh-cell[data-ref="${ref}"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 })); return 1; })()`);
      return until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === '${ref}'`), ref + ' active', 4000).catch(() => false);
    };
    // Words into the card's box, and its button pressed — the way a person
    // types and posts, the textarea's own value set.
    const write = (selector, text) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return false;
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(text)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`);
    const pressCard = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b) return 'no ' + ${JSON.stringify(selector)}; b.click(); return 'clicked'; })()`);
    const cardState = () => js(`(() => {
      const card = document.querySelector('.sh-card');
      if (!card) return null;
      return {
        ref: card.dataset.ref,
        comments: [...card.querySelectorAll('.sh-comment')].map((c) => ({ who: c.querySelector('.sh-comment-name')?.textContent || '', when: c.querySelector('.sh-comment-when')?.textContent || '', text: c.querySelector('.sh-comment-text')?.textContent ?? null })),
        resolved: Boolean(card.querySelector('.sh-card-resolved')),
        box: Boolean(card.querySelector('.sh-card-input')),
      };
    })()`);
    const corner = (ref) => js(`(() => {
      const c = document.querySelector('.sh-cell[data-ref="${ref}"]');
      if (!c) return null;
      return { threaded: c.classList.contains('threaded'), resolved: c.classList.contains('resolved'), colour: getComputedStyle(c, '::before').borderTopColor, tip: c.dataset.tip || '' };
    })()`);
    const refresh = async () => { win.webContents.invalidate(); await wait(700); };
    const author = model().commentAuthor;

    // New Comment on B2.
    const onB2 = await select('B2');
    await tab('Review');
    const pressed = await clickIn('New Comment');
    const boxShown = await until(() => js(`Boolean(document.querySelector('.sh-card .sh-card-input'))`), 'the card', 5000).catch(() => false);
    await write('.sh-card .sh-card-input', 'Is 40 the trade price?');
    await pressCard('.sh-card .sh-card-post');
    const posted = await until(async () => (await cardState())?.comments?.length === 1, 'the comment in the card', 5000).catch(() => false);
    const first = await cardState();
    const mark = await until(async () => (await corner('B2'))?.threaded, 'the purple corner', 4000).catch(() => false);
    const c1 = await corner('B2');
    check('sheets: Review → New Comment opens a card beside B2; the comment posted there is signed with the account\'s name and the time, and the cell gets Excel\'s purple corner',
      onB2 === true && pressed === 'clicked' && boxShown === true && posted === true && first.ref === 'B2' && first.comments[0].who === author
        && first.comments[0].text === 'Is 40 the trade price?' && /\d/.test(first.comments[0].when) && mark === true && c1.colour === 'rgb(122, 61, 184)',
      `${pressed}; card ${JSON.stringify(first)}; corner ${JSON.stringify(c1)}; author ${author}`);

    // A reply.
    await write('.sh-card .sh-card-input', 'Yes, before VAT.');
    await pressCard('.sh-card .sh-card-post');
    const replied = await until(async () => (await cardState())?.comments?.length === 2, 'the reply', 5000).catch(() => false);
    const tip = (await corner('B2'))?.tip || '';
    check('sheets: a reply joins the thread under the first comment, and the cell\'s tip counts it',
      replied === true && (await cardState()).comments[1].text === 'Yes, before VAT.' && /1 reply/.test(tip), `${JSON.stringify(await cardState())}; tip ${tip}`);
    await refresh();
    await capture(win, 'sheets-comment-card.png');

    // Resolve, then Reopen.
    await pressCard('.sh-card .sh-card-resolve');
    const resolved = await until(async () => (await cardState())?.resolved && (await corner('B2'))?.resolved, 'resolved', 5000).catch(() => false);
    const grey = (await corner('B2'))?.colour;
    const noBox = (await cardState())?.box === false;
    await pressCard('.sh-card .sh-card-reopen');
    const reopened = await until(async () => !(await cardState())?.resolved && !(await corner('B2'))?.resolved, 'reopened', 5000).catch(() => false);
    check('sheets: Resolve thread marks it Resolved, greys the corner and closes the reply box; Reopen brings it back',
      resolved === true && grey === 'rgb(154, 147, 165)' && noBox && reopened === true, `resolved ${resolved}, colour ${grey}, box gone ${noBox}, reopened ${reopened}`);

    // Edit the reply from its menu.
    await js(`(() => { document.querySelectorAll('.sh-card .sh-comment-more')[1]?.click(); return 1; })()`);
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => /Edit comment/.test(b.textContent)))`), 'the menu', 4000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => /Edit comment/.test(b.textContent))?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sh-card .sh-comment-edit'))`), 'the edit box', 4000).catch(() => {});
    await write('.sh-card .sh-comment-edit', 'Yes — before VAT, per unit.');
    await pressCard('.sh-card .sh-comment-save');
    const edited = await until(async () => (await cardState())?.comments?.[1]?.text === 'Yes — before VAT, per unit.', 'the edit', 5000).catch(() => false);
    check('sheets: a reply\'s menu → Edit comment changes its words in place', edited === true, JSON.stringify(await cardState()));

    // A second thread on D4, then Show Comments and its filter.
    await select('D4');
    await clickIn('New Comment');
    await until(() => js(`document.querySelector('.sh-card')?.dataset.ref === 'D4' && Boolean(document.querySelector('.sh-card .sh-card-input'))`), 'the card on D4', 5000).catch(() => {});
    await write('.sh-card .sh-card-input', 'Check the tape count.');
    await pressCard('.sh-card .sh-card-post');
    await until(async () => (await corner('D4'))?.threaded, 'D4 threaded', 5000).catch(() => {});
    await clickIn('Show Comments');
    const listed = await until(() => js(`[...document.querySelectorAll('.sh-comments-item')].map((n) => n.dataset.ref).join(',') === 'B2,D4'`), 'the pane', 5000).catch(() => false);
    await js(`(() => { document.querySelector('.sh-comments-filter [data-filter="resolved"]').click(); return 1; })()`);
    const none = await until(() => js(`document.querySelectorAll('.sh-comments-item').length === 0 && Boolean(document.querySelector('.sh-comments-empty'))`), 'no resolved', 4000).catch(() => false);
    await js(`(() => { document.querySelector('.sh-comments-filter [data-filter="open"]').click(); return 1; })()`);
    const openOnes = await until(() => js(`document.querySelectorAll('.sh-comments-item').length === 2`), 'two open', 4000).catch(() => false);
    check('sheets: Review → Show Comments lists both threads in a pane, and its Resolved and Open filters sort them',
      listed === true && none === true && openOnes === true, await js(`[...document.querySelectorAll('.sh-comments-item')].map((n) => n.textContent).join(' | ')`));
    await refresh();
    await capture(win, 'sheets-comments-pane.png');

    // Previous from D4 goes to B2 and opens its card; a pane item does the same.
    await clickIn('Previous');
    const walked = await until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === 'B2' && document.querySelector('.sh-card')?.dataset.ref === 'B2'`), 'B2 and its card', 5000).catch(() => false);
    await js(`(() => { document.querySelector('.sh-comments-item[data-ref="D4"]').click(); return 1; })()`);
    const fromPane = await until(() => js(`document.querySelector('.sh-cell.active')?.dataset.ref === 'D4' && document.querySelector('.sh-card')?.dataset.ref === 'D4'`), 'D4 from the pane', 5000).catch(() => false);
    check('sheets: Previous selects the thread before and opens its card; a thread pressed in the pane is gone to', walked === true && fromPane === true, `previous ${walked}, pane ${fromPane}`);

    // Delete takes D4's thread away.
    await clickIn('Delete — the comment');
    const deleted = await until(async () => !(await corner('D4'))?.threaded && !(await js(`Boolean(document.querySelector('.sh-card'))`)), 'D4 cleared', 5000).catch(() => false);
    const count = await until(() => js(`document.querySelectorAll('.sh-comments-item').length === 1`), 'one left in the pane', 4000).catch(() => false);
    check('sheets: Review → Delete takes the thread off D4, and the pane follows', deleted === true && count === true, `deleted ${deleted}, pane ${count}`);
    await clickIn('Show Comments');

    // Saved: the parts Excel writes.
    doc.save({ id: session.id });
    const saved = new SheetView(fs.readFileSync(file));
    const threads = saved.workbook.threads('Plan');
    const pkgText = (name) => (saved.pkg.has(name) ? saved.pkg.text(name) : '');
    const shadow = saved.workbook.comments('Plan');
    check('sheets: the saved file carries the threaded part, the person and the classic "[Threaded comment]" shadow Excel writes',
      threads.length === 1 && threads[0].ref === 'B2' && threads[0].comments.length === 2 && threads[0].comments[0].author === author
        && /<ThreadedComments\b/.test(pkgText('xl/threadedComments/threadedComment1.xml')) && /<person displayName=/.test(pkgText('xl/persons/person.xml'))
        && shadow.length === 1 && /^tc=\{/.test(shadow[0].author) && /\[Threaded comment\][\s\S]*Reply:/.test(shadow[0].text),
      `${JSON.stringify(threads.map((t) => [t.ref, t.comments.length]))}; shadow ${JSON.stringify(shadow.map((s) => s.author + ' ' + s.text.slice(0, 20)))}`);

    const complaints = await errorsIn(win);
    check('sheets: the comment checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('sheets: the comment checks ran', false, err.message);
  }
}
