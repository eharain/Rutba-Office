// Word: References → Mark Citation, Insert Table of Authorities and Update
// Table.
//
// A case cited on page 1 is selected and marked with Alt+Shift+I (the Mark
// Citation box stays open beside the page), its short form given, and Mark
// All marks the short form where page 2 cites it again; Next Citation finds
// the statute's "U.S.C." and the statute is marked as one; Insert Table of
// Authorities writes a table for each category with the pages the window
// laid the citations on; the saved file keeps the TA fields and the TOA
// fields. Run alone with RUTBA_VERIFY_ONLY=wordtoa.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

const FILL = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn despite two price changes and a supply interruption that took most of July to clear. ';
const CASE = 'Brown v. Board, 347 U.S. 483 (1954)';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordToa(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  try {
    const file = path.join(dir, 'authorities.docx');
    fs.writeFileSync(file, buildDocx({
      styles: true,
      paragraphs: [
        { text: 'Argument', style: 'Heading1' },
        { text: `The court in ${CASE} held otherwise.` },
        { text: FILL.repeat(24) },
        { text: 'Later, Brown was read narrowly.' },
        { text: 'The Clean Air Act, 42 U.S.C. 7401, applies here.' },
        { text: '' },
      ],
    }));
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
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return 'filled';
    })()`);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    /** Where a character of a paragraph's words is on screen. */
    const charAt = (block, text, after) => js(`(() => {
      const b = document.querySelector('.wd-page [data-block="${block}"]');
      if (!b) return null;
      b.scrollIntoView({ block: 'center' });
      const walker = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const i = n.nodeValue.indexOf(${JSON.stringify(text)});
        if (i < 0) continue;
        const k = ${after ? `i + ${text.length} - 1` : 'i'};
        const r = document.createRange();
        r.setStart(n, k); r.setEnd(n, k + 1);
        const box = r.getBoundingClientRect();
        return { x: Math.round(${after ? 'box.right - 1' : 'box.left + 1'}), y: Math.round(box.top + box.height / 2) };
      }
      return null;
    })()`);
    // A click before the words and a Shift+click after them, as a person selects a citation.
    const selectText = async (block, text) => {
      const a = await charAt(block, text, false);
      const b = await charAt(block, text, true);
      if (!a || !b) return false;
      win.webContents.sendInputEvent({ type: 'mouseDown', x: a.x, y: a.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: a.x, y: a.y, button: 'left', clickCount: 1 });
      await wait(120);
      win.webContents.sendInputEvent({ type: 'mouseDown', x: b.x, y: b.y, button: 'left', clickCount: 1, modifiers: ['shift'] });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: b.x, y: b.y, button: 'left', clickCount: 1, modifiers: ['shift'] });
      return until(() => {
        const s = model().selection;
        if (s?.anchor?.block !== block || s.focus.block !== block) return false;
        const words = (model().blocks[block].text || '').slice(Math.min(s.anchor.offset, s.focus.offset), Math.max(s.anchor.offset, s.focus.offset));
        return words === text;
      }, `"${text}" selected`, 5000).catch(() => false);
    };
    const caretAt = async (block) => {
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
    const tas = () => (model().blocks || []).flatMap((b) => (b.runs || []).filter((r) => r.field?.kind === 'ta').map((r) => ({ block: b.index, instr: r.field.instr.trim() })));

    await until(() => js(`document.querySelectorAll('.wd-sheet, .wd-page [data-block]').length > 3`), 'the paper', 8000);
    await clickTab('References');
    await wait(250);

    // The case on page 1: selected, Alt+Shift+I, its short form, Mark All.
    const sel = await selectText(1, CASE);
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(win.webContents, 'I', { modifiers: ['alt', 'shift'] });
    const panel = await until(() => js(`document.querySelector('.wd-toa-long')?.value === ${JSON.stringify(CASE)}`), 'Mark Citation with the selected citation', 5000).catch(() => false);
    await fill('.wd-toa-short', 'Brown');
    await snap('word-toa-mark.png');
    await click('.wd-toa-markall');
    const marked = await until(() => tas().some((t) => t.block === 1 && t.instr === `TA \\l "${CASE}" \\s "Brown" \\c 1`) && tas().some((t) => t.block === 3 && t.instr === 'TA \\s "Brown"'), 'the TA fields', 5000).catch(() => false);
    const shown = await until(() => js(`(() => { const x = document.querySelector('.wd-page.marks .wd-xe'); return x ? getComputedStyle(x, '::after').content : ''; })()`).then((c) => String(c).includes('TA')), 'the field shown with ¶', 5000).catch(() => false);
    check(
      'word: Alt+Shift+I opens Mark Citation with the selected citation, and Mark All marks it there and its short form on page 2, as hidden TA fields ¶ shows',
      sel !== false && panel === true && marked === true && shown === true && model().blocks[1].text === `The court in ${CASE} held otherwise.`,
      JSON.stringify({ sel, panel, marked, shown, tas: tas() })
    );

    // Next Citation stops at the statute's "U.S.C."; the statute is marked under Statutes.
    await click('.wd-toa-next');
    const found = await until(() => {
      const s = model().selection;
      return s?.focus?.block === 4 && (model().blocks[4].text || '').slice(s.anchor.offset, s.focus.offset) === 'U.S.C.';
    }, 'Next Citation on U.S.C.', 5000).catch(() => false);
    await selectText(4, 'Clean Air Act');
    await until(() => js(`document.querySelector('.wd-toa-long')?.value === 'Clean Air Act'`), 'the box to follow the selection', 5000).catch(() => {});
    await fill('.wd-toa-category', '2');
    await click('.wd-toa-mark');
    const statute = await until(() => tas().some((t) => t.block === 4 && t.instr === 'TA \\l "Clean Air Act" \\s "Clean Air Act" \\c 2'), 'the statute marked', 5000).catch(() => false);
    check('word: Next Citation finds the next words that look like a citation, and a statute is marked from the same open box', found === true && statute === true, JSON.stringify({ found, selection: model().selection, tas: tas() }));
    await js(`[...document.querySelectorAll('.wd-refs-float-foot .rw-btn')].find((b) => b.textContent.trim() === 'Close')?.click(), 1`);
    await until(() => js(`!document.querySelector('.wd-refs-float')`), 'the box to close', 3000).catch(() => {});

    // Insert Table of Authorities at the last paragraph: a table for each category.
    await caretAt(5);
    const pressed = await clickRibbon('Insert Table of Authorities');
    await until(() => js(`Boolean(document.querySelector('.wd-toa-ok'))`), 'the Table of Authorities dialog', 4000).catch(() => {});
    await wait(150);
    await snap('word-toa-dialog.png');
    await click('.wd-toa-ok');
    const lines = () => js(`[...document.querySelectorAll('.wd-page [data-style="TOAHeading"], .wd-page [data-style="TableofAuthorities"]')].map((b) => b.dataset.style + '|' + b.innerText.replace(/\\u200b/g, '').trim())`);
    const inserted = await until(async () => (await lines()).length >= 4, 'the table on the page', 6000).catch(() => false);
    const got = await lines();
    check(
      'word: Insert Table of Authorities writes Cases and Statutes, each authority with the pages the window laid its citations on',
      pressed === 'clicked' && inserted === true && got[0] === 'TOAHeading|Cases' && /^TableofAuthorities\|Brown v\. Board, 347 U\.S\. 483 \(1954\)\t1, [2-9]$/.test(got[1]) && got[2] === 'TOAHeading|Statutes' && /^TableofAuthorities\|Clean Air Act\t[2-9]$/.test(got[3]),
      JSON.stringify(got)
    );
    await snap('word-toa.png');

    // Update Table keeps it; Save writes both kinds of field.
    const updated = await clickRibbon('Update Table — the citations');
    await wait(400);
    await clickRibbon('Save');
    const saved = () => {
      try {
        const d = openDocx(fs.readFileSync(file)).doc.doc;
        return { entries: d.authorityEntries().map((e) => e.instr.trim()), tables: d.tablesOfAuthorities().map((t) => t.category) };
      } catch {
        return null;
      }
    };
    await until(() => saved()?.entries?.length === 3, 'the citations in the saved file', 8000).catch(() => {});
    const out = saved();
    check('word: the saved file keeps the TA fields and a TOA field for each category', updated === 'clicked' && Boolean(out) && out.entries.length === 3 && JSON.stringify(out.tables) === '[1,2]', JSON.stringify({ updated, out }));
    const complaints = await errorsIn(win);
    check('word: the table of authorities checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the table of authorities checks ran', false, err.message);
  }
}
