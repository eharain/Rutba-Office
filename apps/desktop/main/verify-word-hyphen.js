// Word: Layout → Hyphenation.
//
// Automatic breaks the long words at the ends of lines on the page — soft
// hyphens of the page's own between their syllables — and the PDF breaks
// them at the same places, with a hyphen where it does. A paragraph can be
// left whole; Hyphenation Options writes the zone, the limit and the
// capitals rule; Manual walks the words that would pull a line back and
// puts an optional hyphen where Yes is pressed. Settings and optional
// hyphens are saved as Word writes them. Run alone with RUTBA_VERIFY_ONLY=hyphen.
import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { hyphenPoints } from '@rutba/doc-view/hyphenate';

const long = 'Internationalization characteristically notwithstanding, responsibility demonstrations overwhelmingly institutionalized interdisciplinary considerations throughout. ';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 * @param {{ dir: string }} args where the fixture is written
 */
export async function verifyWordHyphenation(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  const file = path.join(dir, 'hyphen.docx');
  fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [
    { text: 'Hyphenation', style: 'Heading1' },
    { text: long.repeat(3) },
    { text: long.repeat(2) },
    { text: 'Short words stay whole.' },
  ] }));
  try {
    const win = await open('word', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const menuPick = async (button, label) => {
      const pressed = await js(`(() => {
        const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(button)}) && !n.disabled);
        if (!b) return 'no button';
        b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        b.click();
        return 'clicked';
      })()`);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), `"${label}"`, 3000).catch(() => false);
      const picked = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return 'missing'; b.click(); return 'clicked'; })()`);
      return pressed === 'clicked' && picked === 'clicked' ? 'clicked' : `${pressed}/${picked}`;
    };
    // Soft hyphens the page drew, and the ones a line actually broke at (drawn with a width).
    const shy = (block) => js(`(() => {
      const el = document.querySelector('.wd-page > [data-block="${block}"]');
      const all = el ? [...el.querySelectorAll('.wd-shy')] : [];
      return { all: all.length, broken: all.filter((s) => s.getBoundingClientRect().width > 0.5).length };
    })()`);
    const shot = async (name) => {
      wc.invalidate();
      await wait(700);
      await capture(win, name);
    };
    await until(() => js(`Boolean(document.querySelector('.wd-page [data-block="1"]'))`), 'the paragraphs', 8000);

    // Automatic.
    await clickTab('Layout');
    await wait(200);
    const auto = await menuPick('Hyphenation', 'Automatic');
    const on = await until(() => model().hyphenation?.auto === true, 'hyphenation on', 5000).catch(() => false);
    const drawn = await until(async () => (await shy(1)).broken > 0, 'a line broken mid-word', 6000).catch(() => false);
    const first = await shy(1);
    check('word: Layout → Hyphenation → Automatic breaks long words at the ends of lines, with a hyphen',
      auto === 'clicked' && on === true && drawn === true && first.all > first.broken, `${auto}; ${JSON.stringify(first)}`);
    await shot('word-hyphenation.png');

    // The caret counts the document's characters, not the page's soft hyphens.
    const caret = await js(`(() => {
      const el = document.querySelector('.wd-page > [data-block="1"]');
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      let seen = '';
      for (let n = walker.nextNode(); n; n = walker.nextNode()) if (!n.parentElement.closest('.wd-shy')) seen += n.nodeValue;
      return seen.length;
    })()`);
    check('word: the page\'s soft hyphens are not the document\'s — the words on the page are the paragraph\'s, letter for letter',
      caret === model().blocks[1].text.length, `${caret} vs ${model().blocks[1].text.length}`);

    // The printout breaks at the same places.
    const pdfPath = path.join(dir, 'hyphen.pdf');
    doc.export({ id: session.id, format: 'pdf', path: pdfPath });
    const pdf = fs.readFileSync(pdfPath).toString('latin1');
    const ends = [...pdf.matchAll(/([A-Za-z]+)-\) Tj/g)].map((m) => m[1]);
    const words = long.match(/[A-Za-z]+/g);
    // Each line the PDF ends in a hyphen ends with a word's first part — cut
    // at a place the hyphenator gives that word, the one the page asks too.
    const agree = ends.length > 0 && ends.every((frag) => words.some((w) => hyphenPoints(w).some((p) => w.slice(0, p).toLowerCase() === frag.toLowerCase())));
    check('word: the PDF breaks words at the same places the page may — the hyphenator the page uses',
      agree, `${ends.length} hyphenated line end(s): ${ends.slice(0, 6).join(', ')}`);

    // Left whole: this paragraph only.
    const para = await js(`(() => { const r = document.querySelector('.wd-page > [data-block="2"]').getBoundingClientRect(); return { x: Math.round(r.left + 14), y: Math.round(r.top + 8) }; })()`);
    wc.sendInputEvent({ type: 'mouseDown', x: para.x, y: para.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseUp', x: para.x, y: para.y, button: 'left', clickCount: 1 });
    await until(() => model().selection?.focus?.block === 2, 'the caret in paragraph 2', 4000).catch(() => false);
    await wait(300);
    const whole = await menuPick('Hyphenation', "Don't hyphenate this paragraph");
    const left = await until(async () => (await shy(2)).all === 0 && model().blocks[2].noHyphens === true, 'paragraph 2 whole', 5000).catch(() => false);
    check('word: Don\'t hyphenate this paragraph leaves that paragraph\'s words whole and the others broken',
      whole === 'clicked' && left === true && (await shy(1)).all > 0, `${whole}; ${JSON.stringify(await shy(2))} / ${JSON.stringify(await shy(1))}`);

    // Hyphenation Options.
    await menuPick('Hyphenation', 'Hyphenation Options…');
    await until(() => js(`Boolean(document.querySelector('.wd-hyph-zone'))`), 'the options dialog', 4000).catch(() => false);
    await js(`(() => {
      const set = (sel, value) => { const el = document.querySelector(sel); const proto = Object.getPrototypeOf(el); Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, value); el.dispatchEvent(new Event('input', { bubbles: true })); };
      set('.wd-hyph-zone', '1');
      set('.wd-hyph-limit', '2');
      const caps = document.querySelector('.wd-hyph-caps'); if (caps.checked) caps.click();
      return 1;
    })()`);
    await wait(200);
    await shot('word-hyphenation-options.png');
    await js(`document.querySelector('.wd-hyph-ok').click(), 1`);
    const opts = await until(() => { const hy = model().hyphenation; return hy?.zoneTwips === 567 && hy?.limit === 2 && hy?.caps === false; }, 'the options written', 5000).catch(() => false);
    check('word: Hyphenation Options sets the zone, the limit on hyphens in a row and whether capitals are broken',
      opts === true, JSON.stringify(model().hyphenation));

    // Manual: None first, then walk the words and say Yes to the first.
    await menuPick('Hyphenation', 'None');
    await until(() => model().hyphenation?.auto === false, 'hyphenation off', 4000).catch(() => false);
    await wait(400);
    await menuPick('Hyphenation', 'Manual');
    const asked = await until(() => js(`Boolean(document.querySelector('.wd-hyph-word .wd-hyph-point.on'))`), 'a word to hyphenate', 6000).catch(() => false);
    const word = await js(`document.querySelector('.wd-hyph-word')?.dataset.word || null`);
    const before = model().blocks.map((b) => (b.text.match(/­/g) || []).length).reduce((a, b) => a + b, 0);
    await shot('word-hyphenation-manual.png');
    await js(`document.querySelector('.wd-hyph-yes')?.click(), 1`);
    const put = await until(() => model().blocks.map((b) => (b.text.match(/­/g) || []).length).reduce((a, b) => a + b, 0) === before + 1, 'an optional hyphen', 5000).catch(() => false);
    await wait(500);
    await js(`[...document.querySelectorAll('.rw-dialog button, .rw-modal button')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`);
    const inWord = model().blocks.some((b) => b.text.includes('­') && word && b.text.replace(/­/g, '').includes(word) && b.text.split(/\s+/).some((w) => w.replace(/[^A-Za-z­]/g, '').includes('­')));
    check('word: Layout → Hyphenation → Manual offers a word at a line\'s start with its break points, and Yes puts an optional hyphen there',
      asked === true && Boolean(word) && put === true && inWord, `word ${word}; hyphens ${before} → ${before + (put === true ? 1 : 0)}`);

    // Saved as Word writes it.
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    await wait(1400);
    const saved = openDocx(fs.readFileSync(file));
    const settings = saved.doc.doc.pkg.text('word/settings.xml');
    check('word: the file keeps the hyphenation settings, the paragraph left whole and the optional hyphen, as Word writes them',
      /<w:consecutiveHyphenLimit w:val="2"\/><w:hyphenationZone w:val="567"\/><w:doNotHyphenateCaps\/>/.test(settings) && !/<w:autoHyphenation/.test(settings)
        && /<w:suppressAutoHyphens\/>/.test(saved.doc.doc.xml) && /<w:softHyphen\/>/.test(saved.doc.doc.xml),
      settings.replace(/[\s\S]*?(<w:consecutiveHyphenLimit[\s\S]{0,120}).*/, '$1').slice(0, 160));

    const complaints = await errorsIn(win);
    check('word: hyphenation reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the hyphenation checks ran', false, err.message);
  }
}
