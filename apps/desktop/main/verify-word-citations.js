// Word: References → Citations & Bibliography, driven through the ribbon.
//
// Manage Sources → New… makes a book source (it lands in the profile's master
// list and in the document's list); Insert Citation cites it at a caret put by
// a real click; Style → MLA re-draws the citation; a placeholder is cited and
// then given a source; Bibliography → Works Cited puts the list under the
// text; the saved file carries the sources part, the CITATION field in its
// control and the BIBLIOGRAPHY field. Run alone with
// RUTBA_VERIFY_ONLY=citations.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';

/** The paper this check writes, its own, in the run's folder. */
export function makeCitationsFixture(dir) {
  const file = path.join(dir, 'citations.docx');
  fs.writeFileSync(file, buildDocx({
    styles: true,
    paragraphs: [
      { text: 'Reading and memory', style: 'Heading1' },
      { text: 'Readers remember more of what they read on paper than on a screen' },
      { text: 'Later studies found the gap closing' },
      { text: '' },
    ],
  }));
  return file;
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordCitations(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  let win = null;
  let masterBefore;
  try {
    const file = makeCitationsFixture(dir);
    win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    masterBefore = await js(`window.rutbaOffice.store.get({ key: 'word.bibliography.master', fallback: null })`);
    await js(`window.rutbaOffice.store.set({ key: 'word.bibliography.master', value: '' })`);

    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickMenu = (label) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-menu button')].find((n) => n.textContent.trim().startsWith(${JSON.stringify(label)}) && !n.disabled);
      if (!b) return 'no menu item ' + ${JSON.stringify(label)};
      b.click();
      return 'clicked';
    })()`);
    const click = (selector) => js(`(() => {
      const b = document.querySelector(${JSON.stringify(selector)});
      if (!b || b.disabled) return 'no ' + ${JSON.stringify(selector)};
      b.click();
      return 'clicked';
    })()`);
    // A value typed into a React box: the native setter, then the event React listens for.
    const fill = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return 'no ' + ${JSON.stringify(selector)};
      const proto = el.tagName === 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event(el.tagName === 'SELECT' ? 'change' : 'input', { bubbles: true }));
      return 'filled';
    })()`);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const blockText = (i) => js(`(() => { const b = document.querySelector('.wd-page [data-block="${i}"]'); return b ? b.innerText.replace(/\\u200b/g, '') : null; })()`);
    const caretAt = async (block, where) => {
      const at = await js(`(() => {
        const b = document.querySelector('.wd-page [data-block="${block}"]');
        if (!b) return null;
        b.scrollIntoView({ block: 'center' });
        const r = document.createRange();
        const walker = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
        let last = null;
        for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.nodeValue.length) last = n;
        if (${where === 'end' ? 'true' : 'false'} && last) { r.setStart(last, last.nodeValue.length - 1); r.setEnd(last, last.nodeValue.length); }
        else { r.selectNodeContents(b); r.collapse(true); }
        const rects = [...r.getClientRects()];
        const box = b.getBoundingClientRect();
        const rect = rects.length ? rects[rects.length - 1] : null;
        const x = rect && rect.width + rect.height > 0 ? rect.right + ${where === 'end' ? 2 : 1} : box.left + 3;
        const y = rect && rect.height > 0 ? rect.top + rect.height / 2 : box.top + Math.min(box.height, 20) / 2;
        return { x: Math.round(x), y: Math.round(y) };
      })()`);
      if (!at) return false;
      win.webContents.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      win.webContents.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y, button: 'left', clickCount: 1 });
      const len = (model().blocks?.[block]?.text || '').length;
      return until(() => model().selection?.focus?.block === block && (where !== 'end' || model().selection.focus.offset === len), `the caret in paragraph ${block}`, 5000).catch(() => false);
    };
    const snap = async (name) => {
      win.webContents.invalidate();
      await wait(700);
      await win.webContents.capturePage().catch(() => null);
      win.webContents.invalidate();
      await wait(600);
      await capture(win, name);
    };

    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="2"]'))`), 'the paper', 8000);
    await clickTab('References');
    await wait(250);

    // Manage Sources → New…: a book, in both lists.
    const manage = await clickRibbon('Manage Sources');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Source Manager"]'))`), 'the Source Manager', 4000).catch(() => {});
    await click('.wd-refs-new');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Create Source"] .wd-refs-f-Author'))`), 'Create Source', 4000).catch(() => {});
    const filled = [
      await fill('.wd-refs-f-Author', 'Mangen, Anne'),
      await fill('.wd-refs-f-Title', 'Reading on Paper and Screen'),
      await fill('.wd-refs-f-Year', '2013'),
      await fill('.wd-refs-f-City', 'Oslo'),
      await fill('.wd-refs-f-Publisher', 'Scandinavian University Press'),
    ];
    const tag = await until(async () => (await js(`document.querySelector('.wd-refs-tagname')?.value`)) === 'Man13', 'the tag made from the author and year', 4000).catch(() => false);
    const preview = await js(`document.querySelector('.rw-dialog[aria-label="Create Source"] .wd-refs-preview-body')?.innerText || ''`);
    await snap('word-citations-source.png');
    check(
      'word: Create Source asks for a book\'s fields, makes the tag Man13 from the author and year, and previews the APA entry',
      manage === 'clicked' && filled.every((f) => f === 'filled') && tag === true && preview.includes('Mangen, A. (2013). Reading on Paper and Screen. Scandinavian University Press.'),
      JSON.stringify({ manage, filled, tag, preview })
    );
    await click('.wd-refs-source-ok');
    const listed = await until(() => js(`document.querySelectorAll('.wd-refs-master .wd-refs-row').length === 1 && document.querySelectorAll('.wd-refs-current .wd-refs-row').length === 1`), 'the source in both lists', 4000).catch(() => false);
    await snap('word-citations-manage.png');
    await click('.wd-refs-manage-close');
    const inDoc = await until(() => (model().references?.sources || []).map((s) => s.tag).join() === 'Man13', 'the source in the document', 5000).catch(() => false);
    const master = await js(`window.rutbaOffice.store.get({ key: 'word.bibliography.master', fallback: '' })`);
    check(
      'word: the new source is in the master list (the profile) and in the document\'s list',
      listed === true && inDoc === true && /<b:Tag>Man13<\/b:Tag>/.test(master || ''),
      JSON.stringify({ listed, inDoc, master: String(master || '').slice(0, 120) })
    );

    // Insert Citation at the end of paragraph 1.
    const caret1 = await caretAt(1, 'end');
    await clickRibbon('Insert Citation');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the citation menu', 3000).catch(() => {});
    await snap('word-citations-menu.png');
    const cited = await clickMenu('Mangen, Anne');
    const citedOk = await until(async () => (await blockText(1)) === 'Readers remember more of what they read on paper than on a screen (Mangen, 2013)', 'the citation on the page', 5000).catch(() => false);
    check(
      'word: Insert Citation puts (Mangen, 2013) at the caret, as a field of its own',
      caret1 !== false && cited === 'clicked' && citedOk === true && (model().blocks[1].runs || []).some((r) => r.field?.kind === 'citation'),
      JSON.stringify({ caret1, cited, text: await blockText(1) })
    );

    // Typing after it: the paragraph stays editable.
    await press(win.webContents, '.', { char: true });
    const typed = await until(() => (model().blocks[1].text || '').endsWith('(Mangen, 2013).'), 'a full stop typed after the citation', 5000).catch(() => false);
    check('word: a paragraph with a citation in it takes typing', typed === true, model().blocks[1].text);

    // A placeholder in paragraph 2.
    const caret2 = await caretAt(2, 'end');
    await clickRibbon('Insert Citation');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the citation menu', 3000).catch(() => {});
    await clickMenu('Add New Placeholder');
    await until(() => js(`Boolean(document.querySelector('.wd-refs-placeholder-name'))`), 'the placeholder dialog', 4000).catch(() => {});
    await fill('.wd-refs-placeholder-name', 'Delgado18');
    await click('.wd-refs-placeholder-ok');
    const placeholder = await until(async () => ((await blockText(2)) || '').endsWith('(Delgado18)'), 'the placeholder citation', 5000).catch(() => false);
    check('word: Add New Placeholder cites a name to be filled in later', caret2 !== false && placeholder === true, await blockText(2));

    // Style → MLA: every citation re-drawn.
    const styled = await fill('.wd-refs-style-select', 'mla7');
    const mla = await until(async () => ((await blockText(1)) || '').includes('(Mangen).'), 'the citation in MLA', 6000).catch(() => false);
    check('word: Style → MLA re-draws the citation as (Mangen)', styled === 'filled' && mla === true && model().references?.style === 'mla7', await blockText(1));

    // Bibliography → Works Cited, after the last paragraph.
    await caretAt(3, 'start');
    await clickRibbon('Bibliography');
    await until(() => js(`Boolean(document.querySelector('.rw-menu'))`), 'the bibliography menu', 3000).catch(() => {});
    const worksCited = await clickMenu('Works Cited');
    const listedBib = await until(() => js(`[...document.querySelectorAll('.wd-page [data-block]')].some((b) => b.innerText.includes('Mangen, Anne. Reading on Paper and Screen. Oslo: Scandinavian University Press, 2013. Print.'))`), 'the works cited list on the page', 6000).catch(() => false);
    const heading = await js(`[...document.querySelectorAll('.wd-page [data-block]')].some((b) => b.innerText.trim() === 'Works Cited')`);
    const italic = await js(`(() => { const b = [...document.querySelectorAll('.wd-page [data-block]')].find((x) => x.innerText.includes('Mangen, Anne.')); return b ? [...b.querySelectorAll('span')].some((s) => s.style.fontStyle === 'italic' && s.textContent.includes('Reading on Paper')) : false; })()`);
    check(
      'word: Bibliography → Works Cited lists the source in MLA under its heading, the title in italics',
      worksCited === 'clicked' && listedBib === true && heading === true && italic === true,
      JSON.stringify({ worksCited, listedBib, heading, italic })
    );
    await snap('word-citations.png');

    // The placeholder given a source in Manage Sources: its citation follows.
    await clickRibbon('Manage Sources');
    await until(() => js(`[...document.querySelectorAll('.wd-refs-current .wd-refs-row')].some((r) => r.textContent.includes('Delgado18'))`), 'the placeholder in the current list', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.wd-refs-current .wd-refs-row')].find((r) => r.textContent.includes('Delgado18'))?.click(), 1`);
    await wait(150);
    await click('.wd-refs-edit');
    await until(() => js(`Boolean(document.querySelector('.rw-dialog[aria-label="Edit Source"] .wd-refs-f-Author'))`), 'Edit Source', 4000).catch(() => {});
    await fill('.wd-refs-type-select', 'JournalArticle');
    await wait(150);
    await fill('.wd-refs-f-Author', 'Delgado, Pablo; Vargas, Cristina');
    await fill('.wd-refs-f-Title', 'Don\'t throw away your printed books');
    await fill('.wd-refs-f-JournalName', 'Educational Research Review');
    await fill('.wd-refs-f-Year', '2018');
    await click('.wd-refs-source-ok');
    await wait(200);
    await click('.wd-refs-manage-close');
    const filledIn = await until(async () => ((await blockText(2)) || '').endsWith('(Delgado and Vargas)'), 'the placeholder citation re-drawn', 6000).catch(() => false);
    const bibNow = await until(() => js(`[...document.querySelectorAll('.wd-page [data-block]')].some((b) => b.innerText.startsWith('Delgado, Pablo, and Cristina Vargas.'))`), 'the new entry in the list', 6000).catch(() => false);
    check('word: a placeholder given a source in Manage Sources — its citation and the works cited follow', filledIn === true && bibNow === true, await blockText(2));

    await clickRibbon('Save');
    const saved = () => {
      try {
        const d = openDocx(fs.readFileSync(file)).doc.doc;
        return { sources: d.bibliographySources(), citations: d.citations(), bib: d.bibliographyEntries(), body: d._body().body };
      } catch {
        return null;
      }
    };
    await until(() => saved()?.sources?.sources?.length === 2, 'the sources in the saved file', 8000).catch(() => {});
    const out = saved();
    check(
      'word: the saved file keeps the sources in customXml, each citation a CITATION field in its control, and the BIBLIOGRAPHY field',
      Boolean(out) && out.sources.style === 'mla7' && out.citations.map((c) => c.tags.join()).join('|') === 'Man13|Delgado18'
        && out.bib?.heading === 'Works Cited' && out.bib.entries.length === 2 && /<w:citation\/>/.test(out.body) && /<w:bibliography\/>/.test(out.body),
      JSON.stringify(out && { style: out.sources.style, cites: out.citations.map((c) => c.tags), bib: out.bib })
    );

    const complaints = await errorsIn(win);
    check('word: the citations checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the citations checks ran', false, err.message);
  } finally {
    // The master list is the profile's: put back what was there.
    if (win) await win.webContents.executeJavaScript(`window.rutbaOffice.store.set({ key: 'word.bibliography.master', value: ${JSON.stringify(masterBefore ?? '')} })`).catch(() => {});
  }
}
