// Word: pages side by side as the zoom drops, and View, Page Movement.
//
// The owner's report of 2026-10-10, with Word's window beside this one: at
// about half size Word lays two pages side by side, three and more as the
// zoom falls, and View, Multiple Pages shows two whole pages; Documents
// stacked every page at any zoom. This opens a long report in a wide
// window and reads the drawn sheets back: one column at 100%, two or more
// sharing a top edge at 50%, three or more at 30%, every block drawn inside
// a page, a press on the right-hand page's words and on its margin putting
// the caret there (and the ruler over that page), typing at the foot of the
// left page carrying on at the head of the right one, Multiple Pages two
// whole pages across, Side to Side one row moved through sideways by the
// wheel, and Vertical and 100% one column again.
// Run alone with RUTBA_VERIFY_ONLY=sidebyside.

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordPagesAcross(h, { file }) {
  const { open, check, until, wait, errorsIn, capture, doc, sessionFor } = h;
  if (!file) return check('word: the pages side by side fixture was made', false, 'no fixture');
  try {
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const wc = win.webContents;
    win.setSize(1600, 1000);
    await until(() => js(`window.innerWidth >= 1500`), 'the window to widen', 4000).catch(() => {});
    await until(() => js(`document.querySelectorAll('.wd-page > .wd-sheet').length >= 6 && !document.querySelector('.wd-mounting')`), 'the pages to be laid out', 15000);
    await wait(600);
    const model = () => doc.model({ id: sessionFor('doc').id });
    const shot = async (name) => {
      win.webContents.invalidate();
      await wait(700);
      await capture(win, name);
    };

    // The sheets as drawn, in screen pixels, and how many share the first one's top edge.
    const sheets = () => js(`(() => [...document.querySelectorAll('.wd-page > .wd-sheet')].map((s) => {
      const r = s.getBoundingClientRect();
      return { left: Math.round(r.left), top: Math.round(r.top), right: Math.round(r.right), bottom: Math.round(r.bottom) };
    }))()`);
    const firstRow = (list) => list.filter((s) => Math.abs(s.top - list[0].top) <= 1).length;
    const setZoom = async (v) => {
      await js(`(() => {
        const input = document.querySelector('.rw-zoom input[type="range"]');
        if (!input) return false;
        const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        set.call(input, ${JSON.stringify(String(v))});
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
      })()`);
      await until(() => js(`(() => { const p = document.querySelector('.wd-page'); return Math.abs(p.getBoundingClientRect().width / p.offsetWidth - ${v}) < 0.02; })()`), `the page at ${Math.round(v * 100)}%`, 5000).catch(() => {});
      await wait(400);
    };
    const clickRibbon = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(label)};
      b.click();
      return 'clicked';
    })()`);
    const clickTab = (name) => js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(); return 1; })()`);
    const click = async (x, y) => {
      wc.sendInputEvent({ type: 'mouseMove', x, y });
      await wait(40);
      wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: 1 });
      await wait(40);
      wc.sendInputEvent({ type: 'mouseUp', x, y, button: 'left', clickCount: 1 });
    };

    // 100%: one column, as before.
    const at100 = await sheets();
    const oneColumn = (list) => list.length >= 2 && list.every((s) => Math.abs(s.left - list[0].left) <= 1) && list[1].top > list[0].bottom;
    check('word: at 100% in a 1600 px window the pages stand one under another, as before', oneColumn(at100), `${at100.length} pages; first two ${JSON.stringify(at100.slice(0, 2))}`);

    // 50%: two or more across, sharing a top edge, the second to the right of the first.
    await setZoom(0.5);
    await until(async () => firstRow(await sheets()) >= 2, 'the pages to go side by side', 5000).catch(() => {});
    const at50 = await sheets();
    const across50 = firstRow(at50);
    check('word: at 50% two pages share a top edge, the second to the right of the first',
      across50 >= 2 && Math.abs(at50[1].top - at50[0].top) <= 1 && at50[1].left > at50[0].right,
      `${across50} across; first two ${JSON.stringify(at50.slice(0, 2))}`);
    // The rows are centred in the window, with even gaps between the pages.
    const centred = await js(`(() => {
      const s = document.querySelector('.wd-scroll').getBoundingClientRect();
      const list = [...document.querySelectorAll('.wd-page > .wd-sheet')].map((n) => n.getBoundingClientRect());
      const row = list.filter((r) => Math.abs(r.top - list[0].top) <= 1);
      const left = row[0].left - s.left;
      const right = s.right - row[row.length - 1].right;
      const gaps = row.slice(1).map((r, i) => Math.round(r.left - row[i].right));
      return { left: Math.round(left), right: Math.round(right), gaps };
    })()`);
    check('word: the row of pages is centred, the gaps between the pages even',
      Math.abs(centred.left - centred.right) <= 24 && centred.gaps.length >= 1 && centred.gaps.every((g) => g >= 12 && Math.abs(g - centred.gaps[0]) <= 1),
      JSON.stringify(centred));
    // Every block drawn inside the text area of a page.
    const insideNow = () => js(`(() => {
      const page = document.querySelector('.wd-page');
      const z = page.getBoundingClientRect().width / page.offsetWidth;
      const cs = getComputedStyle(page);
      const mt = parseFloat(cs.paddingTop) * z;
      const mb = parseFloat(cs.paddingBottom) * z;
      const list = [...page.querySelectorAll(':scope > .wd-sheet')].map((n) => n.getBoundingClientRect());
      const flow = [...page.children].filter((el) => /wd-block|wd-table/.test(el.className) && !el.classList.contains('wd-frame'));
      const out = [];
      let tall = 0;
      for (const el of flow) {
        const r = el.getBoundingClientRect();
        const s = list.find((q) => r.left >= q.left - 1 && r.right <= q.right + 1 && r.top >= q.top - 1 && r.top <= q.bottom);
        if (!s) { out.push((el.dataset.block || el.dataset.table) + ' on no page'); continue; }
        if (r.height > s.height - mt - mb) { tall += 1; continue; }
        if (r.top < s.top + mt - 2 || r.bottom > s.bottom - mb + 2) out.push((el.dataset.block || el.dataset.table) + ' ' + Math.round(r.top - s.top) + '-' + Math.round(r.bottom - s.top));
      }
      return { blocks: flow.length, tall, out };
    })()`);
    await until(async () => (await insideNow()).out.length === 0, 'the pages laid at 50%', 5000).catch(() => {});
    const inside = await insideNow();
    check('word: at 50% every block of the long report is drawn inside the text area of a page', inside.out.length === 0 && inside.blocks > 20, `${inside.blocks} blocks, ${inside.tall} taller than a page${inside.out.length ? '; outside: ' + inside.out.slice(0, 4).join(' | ') : ''}`);
    await shot('word-pages-50.png');

    // A press on the right-hand page's words puts the caret there, and the ruler follows it.
    const target = await js(`(() => {
      const sheet = document.querySelectorAll('.wd-page > .wd-sheet')[1].getBoundingClientRect();
      const el = [...document.querySelectorAll('.wd-page > .wd-block')].find((b) => {
        const r = b.getBoundingClientRect();
        return r.left >= sheet.left - 1 && r.top >= sheet.top && r.bottom <= sheet.bottom && (b.textContent || '').length > 40 && !b.dataset.from;
      });
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return { block: Number(el.dataset.block), x: Math.round(r.left + r.width / 2), y: Math.round(r.top + 6), sheet: { left: sheet.left, right: sheet.right, top: sheet.top, bottom: sheet.bottom } };
    })()`);
    if (target) {
      await click(target.x, target.y);
      const landed = await until(() => model().selection?.focus?.block === target.block, 'the caret on the right-hand page', 5000).catch(() => false);
      check('word: a press on the right-hand page\'s words puts the caret there', landed === true, `wanted block ${target.block}, caret in ${model().selection?.focus?.block}`);
      const ruler = await until(() => js(`(() => { const r = document.querySelector('.wd-ruler')?.getBoundingClientRect(); return r && Math.abs(r.left - ${target.sheet.left}) <= 2 ? Math.round(r.left) : false; })()`), 'the ruler over the right-hand page', 4000).catch(() => false);
      check('word: the ruler moves over the page the caret is on', ruler !== false, `ruler at ${await js(`Math.round(document.querySelector('.wd-ruler')?.getBoundingClientRect().left)`)}, page at ${Math.round(target.sheet.left)}`);
      // Back to page one first, so the margin press has somewhere to move the caret from.
      const home = await js(`(() => { const b = document.querySelector('.wd-page > .wd-block[data-block="1"]'); const r = b.getBoundingClientRect(); return { x: Math.round(r.left + 20), y: Math.round(r.top + 6) }; })()`);
      await click(home.x, home.y);
      await until(() => model().selection?.focus?.block === 1, 'the caret back on page one', 4000).catch(() => {});
      // A press in that page's left margin, beside a line, off the page element's own box.
      await click(Math.round(target.sheet.left + 12), target.y);
      const margin = await until(() => model().selection?.focus?.block === target.block, 'the caret from the margin', 4000).catch(() => false);
      check('word: a press in the right-hand page\'s margin puts the caret on that page\'s line', margin === true, `wanted block ${target.block}, caret in ${model().selection?.focus?.block}`);
    } else check('word: a press on the right-hand page\'s words puts the caret there', false, 'no paragraph on the second page');

    // Typing at the foot of the left page carries on at the head of the right one.
    const foot = await js(`(() => {
      const list = [...document.querySelectorAll('.wd-page > .wd-sheet')].map((n) => n.getBoundingClientRect());
      const first = list[0];
      const onFirst = [...document.querySelectorAll('.wd-page > .wd-block')].filter((b) => {
        const r = b.getBoundingClientRect();
        return r.left < first.right && r.top >= first.top && r.bottom <= first.bottom;
      });
      const el = onFirst[onFirst.length - 1];
      if (!el) return null;
      const range = document.createRange();
      range.selectNodeContents(el);
      const rects = [...range.getClientRects()].filter((r) => r.width > 0);
      const last = rects[rects.length - 1];
      return { block: Number(el.dataset.block), x: Math.round(last.right - 2), y: Math.round(last.top + last.height / 2), text: (el.textContent || '').length };
    })()`);
    if (foot) {
      await click(foot.x, foot.y);
      await until(() => model().selection?.focus?.block === foot.block, 'the caret at the foot of page one', 5000).catch(() => {});
      await js(`document.querySelector('.wd-page').focus(), 1`);
      // Then to the paragraph's end with the End key, as a person would.
      wc.sendInputEvent({ type: 'keyDown', keyCode: 'End' });
      wc.sendInputEvent({ type: 'keyUp', keyCode: 'End' });
      await wait(300);
      const words = 'The figures for the second half were kept apart from the first, so that the change in the method could be seen on its own. ';
      for (let i = 0; i < 6; i++) {
        wc.insertText(words);
        await wait(350);
      }
      const caretPage = () => js(`(() => {
        const s = getSelection();
        if (!s.rangeCount) return false;
        const c = s.getRangeAt(0).getBoundingClientRect();
        const list = [...document.querySelectorAll('.wd-page > .wd-sheet')].map((n) => n.getBoundingClientRect());
        const k = list.findIndex((q) => c.left >= q.left && c.left <= q.right && c.top >= q.top && c.top <= q.bottom);
        return k === 1 ? { page: k + 1, down: Math.round((c.top - list[1].top) / list[1].height * 100) } : false;
      })()`);
      await until(async () => Boolean(await caretPage()), 'the caret on the right-hand page', 8000).catch(() => false);
      const typed = await caretPage();
      const inModel = (model().blocks[foot.block]?.text || '').includes('kept apart from the first');
      check('word: typing at the foot of the left page carries on at the head of the right one',
        Boolean(typed) && typed.down < 50 && inModel,
        `${typed ? `caret on page ${typed.page}, ${typed.down}% down it` : 'caret not on the right-hand page'}; text in paragraph ${foot.block}: ${inModel}`);
      await shot('word-pages-typed.png');
    } else check('word: typing at the foot of the left page carries on at the head of the right one', false, 'no paragraph at the foot of page one');

    // 30%: three or more across.
    await setZoom(0.3);
    await until(async () => firstRow(await sheets()) >= 3, 'three or more across', 5000).catch(() => {});
    const at30 = await sheets();
    await until(async () => (await insideNow()).out.length === 0, 'the pages laid again at 30%', 6000).catch(() => {});
    const inside30 = await insideNow();
    check('word: at 30% the pages are laid again for the new line breaks, every block inside the text area of a page', inside30.out.length === 0 && inside30.blocks > 20, `${inside30.blocks} blocks, ${inside30.tall} taller than a page${inside30.out.length ? '; outside: ' + inside30.out.slice(0, 4).join(' | ') : ''}`);
    const pct30 = await js(`document.querySelector('.rw-zoom-pct')?.textContent`);
    check('word: at 30% three or more pages share a top edge', pct30 === '30%' && firstRow(at30) >= 3 && at30[2].left > at30[1].right && at30[1].left > at30[0].right, `${firstRow(at30)} across of ${at30.length} at ${pct30}`);
    await shot('word-pages-30.png');

    // View, Multiple Pages: two whole pages across.
    await clickTab('View');
    await wait(200);
    const multiple = await clickRibbon('Multiple Pages');
    await until(async () => firstRow(await sheets()) === 2, 'two across', 5000).catch(() => {});
    await js(`(() => { document.querySelector('.wd-scroll').scrollTop = 0; return 1; })()`);
    await wait(400);
    const two = await js(`(() => {
      const s = document.querySelector('.wd-scroll').getBoundingClientRect();
      const list = [...document.querySelectorAll('.wd-page > .wd-sheet')].map((n) => n.getBoundingClientRect());
      const row = list.filter((r) => Math.abs(r.top - list[0].top) <= 1);
      const whole = row.every((r) => r.left >= s.left && r.right <= s.right && r.top >= s.top && r.bottom <= s.bottom);
      return { across: row.length, whole, zoom: document.querySelector('.rw-zoom-pct')?.textContent };
    })()`);
    check('word: View, Multiple Pages shows two whole pages side by side', multiple === 'clicked' && two.across === 2 && two.whole, JSON.stringify({ multiple, ...two }));
    await shot('word-pages-multiple.png');

    // Side to Side: one row, the wheel moving through it sideways.
    const sideBtn = await clickRibbon('Side to Side');
    await until(() => js(`Boolean(document.querySelector('.wd-scroll.wd-side'))`), 'Side to Side', 4000).catch(() => {});
    await wait(500);
    const row = await sheets();
    const pressed = await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || n.title || '').startsWith('Side to Side'))?.getAttribute('aria-pressed')`);
    check('word: View, Page Movement, Side to Side lays every page in one row', sideBtn === 'clicked' && row.length >= 6 && firstRow(row) === row.length && String(pressed) === 'true', `${firstRow(row)} of ${row.length} in the first row; pressed ${pressed}`);
    const box = await js(`(() => { const r = document.querySelector('.wd-scroll').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), left: document.querySelector('.wd-scroll').scrollLeft }; })()`);
    const scrolled = async () => js(`document.querySelector('.wd-scroll').scrollLeft`);
    let delta = -120;
    wc.sendInputEvent({ type: 'mouseMove', x: box.x, y: box.y });
    wc.sendInputEvent({ type: 'mouseWheel', x: box.x, y: box.y, deltaX: 0, deltaY: delta, wheelTicksX: 0, wheelTicksY: -1, canScroll: true });
    let moved = await until(async () => (await scrolled()) > box.left + 50, 'the row to move sideways', 3000).catch(() => false);
    if (!moved) {
      delta = 120;
      wc.sendInputEvent({ type: 'mouseWheel', x: box.x, y: box.y, deltaX: 0, deltaY: delta, wheelTicksX: 0, wheelTicksY: 1, canScroll: true });
      moved = await until(async () => (await scrolled()) > box.left + 50, 'the row to move sideways', 3000).catch(() => false);
    }
    const vertical = await js(`document.querySelector('.wd-scroll').scrollTop`);
    check('word: in Side to Side the mouse wheel moves through the pages sideways', moved === true, `scrollLeft ${box.left} -> ${await scrolled()}, scrollTop ${vertical}, ${await js(`(() => { const s = document.querySelector('.wd-scroll'); return s.scrollWidth + '/' + s.clientWidth; })()`)}`);
    await shot('word-pages-side.png');

    // Vertical again: the zoom it had before, the pages back in rows.
    await clickRibbon('Vertical');
    await until(() => js(`!document.querySelector('.wd-scroll.wd-side')`), 'Vertical', 4000).catch(() => {});
    await setZoom(1);
    await until(async () => oneColumn(await sheets()), 'one column again', 5000).catch(() => {});
    const back = await sheets();
    check('word: Vertical and 100% put the pages back one under another', oneColumn(back), `${back.length} pages; first two ${JSON.stringify(back.slice(0, 2))}`);

    const complaints = await errorsIn(win);
    check('word: the pages side by side checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    // pages-grid.docx is this check's own fixture and is never saved.
  } catch (err) {
    check('word: the pages side by side checks ran', false, err.message);
  }
}
