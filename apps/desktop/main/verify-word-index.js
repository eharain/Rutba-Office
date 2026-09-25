// Word: References → Mark Entry, Insert Index and Update Index.
//
// A word double-clicked on page 1 is marked with Alt+Shift+X (the Mark Index
// Entry box stays open beside the page, and ¶ comes on to show the hidden XE
// field); a word on page 2 is marked as a subentry and another as a
// cross-reference; Insert Index writes the index with right-aligned page
// numbers behind dots, each entry on the page it was marked on; a later mark
// and Update Index bring it up to date; the saved file keeps the XE fields
// and the INDEX field. Run alone with RUTBA_VERIFY_ONLY=wordindex.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const FILL = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn despite two price changes and a supply interruption that took most of July to clear. ';

/** The paper this check writes, its own, in the run's folder: apples on page 1, bananas and cherries on page 2. */
export function makeIndexFixture(dir) {
  const file = path.join(dir, 'index.docx');
  fs.writeFileSync(file, buildDocx({
    styles: true,
    paragraphs: [
      { text: 'Orchard notes', style: 'Heading1' },
      { text: 'Apples grow on trees in the orchard.' },
      { text: FILL.repeat(24) },
      { text: 'Bananas grow in bunches in the glasshouse.' },
      { text: 'Cherries ripen in June.' },
      { text: '' },
    ],
  }));
  return file;
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordIndex(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  try {
    const file = makeIndexFixture(dir);
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
    // A real double-click on a word: the page selects it, as a person would.
    const selectWord = async (block, word) => {
      const at = await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="${block}"]');
        if (!b) return null;
        b.scrollIntoView({ block: 'center' });
        const walker = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
        for (let n = walker.nextNode(); n; n = walker.nextNode()) {
          const i = n.nodeValue.indexOf(${JSON.stringify(word)});
          if (i < 0) continue;
          const r = document.createRange();
          r.setStart(n, i + 1); r.setEnd(n, i + 2);
          const box = r.getBoundingClientRect();
          return { x: Math.round(box.left + box.width / 2), y: Math.round(box.top + box.height / 2) };
        }
        return null;
      })()`);
      if (!at) return false;
      win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 2 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 2 });
      return until(() => {
        const s = model().selection;
        return s?.anchor?.block === block && s.focus.block === block && Math.abs(s.focus.offset - s.anchor.offset) >= word.length;
      }, `"${word}" selected`, 5000).catch(() => false);
    };
    const caretEnd = async (block) => {
      const at = await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="${block}"]');
        if (!b) return null;
        b.scrollIntoView({ block: 'center' });
        const box = b.getBoundingClientRect();
        return { x: Math.round(box.left + 4), y: Math.round(box.top + Math.min(box.height, 20) / 2) };
      })()`);
      if (!at) return false;
      win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      return until(() => model().selection?.focus?.block === block, `the caret in paragraph ${block}`, 5000).catch(() => false);
    };
    const snap = async (name) => {
      win.webContents.invalidate();
      await wait(700);
      await win.webContents.capturePage().catch(() => null);
      win.webContents.invalidate();
      await wait(600);
      await capture(win, name);
    };
    const xes = () => (model().blocks || []).flatMap((b) => (b.runs || []).filter((r) => r.field?.kind === 'xe').map((r) => ({ block: b.index, instr: r.field.instr.trim() })));

    await until(() => js(`document.querySelectorAll('.wd-sheet, .wd-page [data-block]').length > 3`), 'the paper', 8000);
    await clickTab('References');
    await wait(250);

    // Apples, on page 1: double-click, Alt+Shift+X, Mark.
    const sel1 = await selectWord(1, 'Apples');
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(win.webContents, 'X', { modifiers: ['alt', 'shift'] });
    const panel = await until(() => js(`document.querySelector('.wd-refs-main')?.value === 'Apples'`), 'Mark Index Entry with the selected word', 5000).catch(() => false);
    await click('.wd-refs-mark');
    const marked1 = await until(() => xes().some((x) => x.instr === 'XE "Apples"' && x.block === 1), 'the XE field after Apples', 5000).catch(() => false);
    const shown = await until(() => js(`(() => { const x = document.querySelector('.wd-page.marks .wd-xe'); return x ? getComputedStyle(x, '::after').content : ''; })()`).then((c) => String(c).includes('XE') && String(c).includes('Apples')), 'the field shown with ¶', 5000).catch(() => false);
    const text1 = model().blocks[1].text;
    check(
      'word: Alt+Shift+X opens Mark Index Entry with the selected word, and Mark writes a hidden XE field that ¶ shows',
      sel1 !== false && panel === true && marked1 === true && shown === true && text1 === 'Apples grow on trees in the orchard.',
      JSON.stringify({ sel1, panel, marked1, shown, text1, xes: xes() })
    );

    // Bananas on page 2, as a subentry of Fruit; the box stays open.
    await selectWord(3, 'Bananas');
    await until(() => js(`document.querySelector('.wd-refs-main')?.value === 'Bananas'`), 'the box to follow the selection', 5000).catch(() => {});
    await fill('.wd-refs-main', 'Fruit');
    await fill('.wd-refs-sub', 'bananas');
    await click('.wd-refs-mark');
    const marked2 = await until(() => xes().some((x) => x.instr === 'XE "Fruit:bananas"' && x.block === 3), 'the subentry', 5000).catch(() => false);
    // Orchard: a cross-reference.
    await selectWord(1, 'orchard');
    await until(() => js(`document.querySelector('.wd-refs-main')?.value === 'orchard'`), 'the box to follow the selection', 5000).catch(() => {});
    await js(`document.querySelector('.wd-refs-opt-see')?.click(), 1`);
    await fill('.wd-refs-see', 'See Fruit');
    await snap('word-index-mark.png');
    await click('.wd-refs-mark');
    const marked3 = await until(() => xes().some((x) => x.instr === 'XE "orchard" \\t "See Fruit"'), 'the cross-reference', 5000).catch(() => false);
    check('word: a subentry and a cross-reference are marked from the same open box', marked2 === true && marked3 === true, JSON.stringify(xes()));
    await js(`[...document.querySelectorAll('.wd-refs-float-foot .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);
    await until(() => js(`!document.querySelector('.wd-refs-float')`), 'the box to close', 3000).catch(() => {});

    // Insert Index at the last paragraph: right-aligned pages behind dots.
    await caretEnd(5);
    const pressed = await clickRibbon('Insert Index');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Index"]'))`), 'the Index dialog', 4000).catch(() => {});
    await js(`(() => { const c = document.querySelector('.wd-refs-right'); if (c && !c.checked) c.click(); return 1; })()`);
    await wait(150);
    await snap('word-index-dialog.png');
    await click('.wd-refs-index-ok');
    const lines = () => js(`[...document.querySelectorAll('.wd-page [data-style^="Index"]')].map((b) => b.dataset.style + '|' + b.innerText.replace(/\\u200b/g, ''))`);
    const inserted = await until(async () => (await lines()).length >= 6, 'the index on the page', 6000).catch(() => false);
    const got = await lines();
    const expect = ['IndexHeading|F', 'Index1|Fruit', 'IndexHeading|O', 'Index1|orchard. See Fruit'];
    check(
      'word: Insert Index lists the entries under their letters, bananas on page 2 — the page the window laid it on — right-aligned',
      pressed === 'clicked' && inserted === true && expect.every((l) => got.includes(l)) && got.some((l) => /^Index2\|bananas\t[2-9]$/.test(l)) && got.includes('IndexHeading|A') && got.includes('Index1|Apples\t1'),
      JSON.stringify(got)
    );
    await snap('word-index.png');

    // A later mark, then Update Index.
    await selectWord(4, 'Cherries');
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(win.webContents, 'X', { modifiers: ['alt', 'shift'] });
    await until(() => js(`document.querySelector('.wd-refs-main')?.value === 'Cherries'`), 'the box', 5000).catch(() => {});
    await js(`document.querySelector('.wd-refs-italic')?.click(), 1`);
    await click('.wd-refs-mark');
    await until(() => xes().some((x) => x.instr === 'XE "Cherries" \\i'), 'the italic mark', 5000).catch(() => {});
    await js(`[...document.querySelectorAll('.wd-refs-float-foot .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);
    const updated = await clickRibbon('Update Index');
    const cherries = await until(async () => (await lines()).some((l) => /^Index1\|Cherries\t[2-9]$/.test(l)), 'Update Index to add Cherries', 6000).catch(() => false);
    const italic = await js(`(() => { const b = [...document.querySelectorAll('.wd-page [data-style="Index1"]')].find((x) => x.innerText.startsWith('Cherries')); return b ? [...b.querySelectorAll('span')].some((s) => s.style.fontStyle === 'italic' && /^[2-9]$/.test(s.textContent)) : false; })()`);
    check('word: Update Index adds the entry marked since, its page number in italics as marked', updated === 'clicked' && cherries === true && italic === true, JSON.stringify(await lines()));

    await clickRibbon('Save');
    const saved = () => {
      try {
        const d = openDocx(fs.readFileSync(file)).doc.doc;
        return { entries: d.indexEntries().map((e) => e.instr.trim()), index: d.indexResult() };
      } catch {
        return null;
      }
    };
    await until(() => saved()?.entries?.length === 4, 'the marks in the saved file', 8000).catch(() => {});
    const out = saved();
    check(
      'word: the saved file keeps four XE fields and the INDEX field, right-aligned with a dot leader',
      Boolean(out) && out.entries.length === 4 && out.index?.rightAlign === true && out.index?.leader === 'dot' && out.index.lines.some((l) => /^Cherries\t[2-9]$/.test(l.text)),
      JSON.stringify(out)
    );
    // ¶ back off, for the checks after this one.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    const complaints = await errorsIn(win);
    check('word: the index checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the index checks ran', false, err.message);
  }
}
