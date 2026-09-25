// Word: equations — Office Math read, drawn, edited round and saved.
//
// A document of equations as Word writes them (the quadratic formula, an
// inline πr², a sum, a function, a matrix) opens with each one drawn by the
// page's MathML: a real fraction bar, a radical, stacked limits, italic
// variables, upright function names. The caret steps over an equation as
// one character, a click selects it, Backspace takes it and Undo puts it
// back. Then Insert → Built-in → Quadratic Formula puts one in, a
// double-click opens it in the editor, an edit there updates it, Alt+=
// opens a new one, and Save writes the OMML; printed, every equation is
// Chromium's own picture of its MathML. Run alone with
// RUTBA_VERIFY_ONLY=equations.
import fs from 'node:fs';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { linearToOmml } from '@rutba/ooxml/math-linear';

const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const RPR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
const CTRL = '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>';
const r = (t, sty = '') => `<m:r>${sty ? `<m:rPr><m:sty m:val="${sty}"/></m:rPr>` : ''}${RPR}<m:t>${t}</m:t></m:r>`;
const e = (body) => `<m:e>${body}</m:e>`;

/** The equations the fixture holds, as Word writes them. */
export const FIXTURE_EQUATIONS = {
  // Word's own Quadratic Formula.
  Q: `<m:oMathPara><m:oMath>${r('x=')}<m:f><m:fPr>${CTRL}</m:fPr><m:num>${r('-b±')}<m:rad><m:radPr><m:degHide m:val="1"/>${CTRL}</m:radPr><m:deg/>${e(`<m:sSup><m:sSupPr>${CTRL}</m:sSupPr>${e(r('b'))}<m:sup>${r('2')}</m:sup></m:sSup>${r('-4ac')}`)}</m:rad></m:num><m:den>${r('2a')}</m:den></m:f></m:oMath></m:oMathPara>`,
  // πr², inline among the words.
  A: `<m:oMath>${r('π')}<m:sSup><m:sSupPr>${CTRL}</m:sSupPr>${e(r('r'))}<m:sup>${r('2')}</m:sup></m:sSup></m:oMath>`,
  // A sum with its limits stacked, equal to a fraction.
  S: `<m:oMathPara><m:oMath><m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/>${CTRL}</m:naryPr><m:sub>${r('i=1')}</m:sub><m:sup>${r('n')}</m:sup>${e(r('i'))}</m:nary>${r('=')}<m:f><m:fPr>${CTRL}</m:fPr><m:num>${r('n')}<m:d><m:dPr>${CTRL}</m:dPr>${e(r('n+1'))}</m:d></m:num><m:den>${r('2')}</m:den></m:f></m:oMath></m:oMathPara>`,
  // A function, upright, and a matrix in brackets.
  F: `<m:oMathPara><m:oMath><m:func><m:funcPr>${CTRL}</m:funcPr><m:fName>${r('sin', 'p')}</m:fName>${e(r('θ'))}</m:func>${r('=')}<m:d><m:dPr>${CTRL}</m:dPr>${e(`<m:m><m:mPr>${CTRL}</m:mPr><m:mr>${e(r('a'))}${e(r('b'))}</m:mr><m:mr>${e(r('c'))}${e(r('d'))}</m:mr></m:m>`)}</m:d></m:oMath></m:oMathPara>`,
};

/** equations.docx: words, then the equations among and between them. */
export function makeEquationFixture(file) {
  const bytes = buildDocx({
    styles: true,
    paragraphs: [
      { text: 'Equations', style: 'Title' },
      { text: 'The area of a circle is @A@ where r is its radius.' },
      { text: 'The roots of a quadratic:' },
      { text: '@Q@' },
      { text: 'A sum:' },
      { text: '@S@' },
      { text: 'A function and a matrix:' },
      { text: '@F@' },
      { text: 'The end.' },
    ],
  });
  const pkg = OoxmlPackage.read(bytes);
  let xml = pkg.text('word/document.xml');
  xml = xml.replace(/<w:r>(<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t xml:space="preserve">([^<]*@[A-Z]@[^<]*)<\/w:t><\/w:r>/g, (whole, rPr, text) =>
    text.split(/(@[A-Z]@)/).map((piece) => {
      const name = /^@([A-Z])@$/.exec(piece)?.[1];
      if (name) return FIXTURE_EQUATIONS[name];
      return piece ? `<w:r>${rPr || ''}<w:t xml:space="preserve">${piece}</w:t></w:r>` : '';
    }).join(''));
  xml = xml.replace('<w:document ', `<w:document xmlns:m="${M_NS}" `);
  pkg.write_('word/document.xml', Buffer.from(xml, 'utf8'));
  fs.writeFileSync(file, pkg.write());
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, capture, doc, sessionFor
 */
export async function verifyWordEquations(h, { file }) {
  const { open, check, until, wait, press, errorsIn, capture, doc, sessionFor } = h;
  if (!file) return check('word: the equations fixture was made', false, 'no fixture');
  const original = fs.readFileSync(file);
  try {
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = sessionFor('doc');
    const model = () => doc.model({ id: session.id });
    const MARK = '\uFFFC';

    // Every equation drawn: a host per equation, its MathML in the shadow.
    await until(() => js(`document.querySelectorAll('.wd-page .wd-math').length === 4 && [...document.querySelectorAll('.wd-page .wd-math')].every((h) => h.shadowRoot && h.shadowRoot.querySelector('math'))`), 'four equations drawn', 10000).catch(() => false);
    const drawn = await js(`(() => {
      const hosts = [...document.querySelectorAll('.wd-page .wd-math')];
      return hosts.map((h) => {
        const m = h.shadowRoot && h.shadowRoot.querySelector('math');
        const r = m ? m.getBoundingClientRect() : { width: 0, height: 0 };
        return { block: Number(h.closest('[data-block]')?.dataset.block), display: m?.getAttribute('display'), w: Math.round(r.width), h: Math.round(r.height), text: h.textContent, font: m ? getComputedStyle(m).fontFamily : '' };
      });
    })()`);
    check(
      'word: an equations document opens with each equation drawn as MathML — one inline among the words, three on lines of their own',
      drawn.length === 4 && drawn[0].display === 'inline' && drawn.slice(1).every((d) => d.display === 'block') && drawn.every((d) => d.w > 8 && d.h > 8),
      JSON.stringify(drawn)
    );
    check(
      'word: each equation is one character of its paragraph to the caret (the page holds U+FFFC for it, the letters live in the MathML)',
      drawn.every((d) => d.text === MARK) && model().blocks[3].text === MARK && model().blocks[1].text.includes(MARK),
      JSON.stringify(model().blocks[1].text)
    );
    check('word: equations are set in Cambria Math', drawn.every((d) => /Cambria Math/.test(d.font)), drawn[0]?.font || '');

    // The quadratic formula laid out as Word lays it: the numerator above
    // the bar, the denominator below it, the radical's overbar spanning
    // b²−4ac, a superscript raised, x italic.
    const shape = await js(`(() => {
      const host = document.querySelector('.wd-page [data-block="3"] .wd-math');
      const root = host.shadowRoot;
      const frac = root.querySelector('mfrac');
      const [num, den] = frac ? [...frac.children] : [];
      const sqrt = root.querySelector('msqrt');
      const sup = root.querySelector('msup');
      const [base, exp] = sup ? [...sup.children] : [];
      const rect = (n) => n ? n.getBoundingClientRect() : null;
      const x = root.querySelector('mi');
      const line = document.querySelector('.wd-page [data-block="2"]').getBoundingClientRect().height;
      return {
        numAbove: rect(num).bottom <= rect(den).top + 1,
        sqrtWider: rect(sqrt).width > rect(sqrt.firstElementChild).width + 4,
        sqrtTaller: rect(sqrt).height > rect(sqrt.firstElementChild).height,
        supRaised: rect(exp).top < rect(base).top,
        italic: getComputedStyle(x).textTransform,
        height: Math.round(rect(host).height),
        line: Math.round(line),
      };
    })()`);
    check(
      'word: the quadratic formula is laid out as Word sets it — numerator over the bar, the radical over b²−4ac, the square raised, x italic',
      shape.numAbove && shape.sqrtWider && shape.sqrtTaller && shape.supRaised && shape.italic === 'math-auto',
      JSON.stringify(shape)
    );
    check(
      'word: the on-screen page makes room for it — the equation\'s paragraph is taller than a line of words',
      shape.height > shape.line * 1.6,
      `${shape.height} px against a ${shape.line} px line`
    );
    const sumAndFunc = await js(`(() => {
      const sum = document.querySelector('.wd-page [data-block="5"] .wd-math').shadowRoot;
      const uo = sum.querySelector('munderover');
      const [op, under, over] = uo ? [...uo.children].map((n) => n.getBoundingClientRect()) : [];
      const func = document.querySelector('.wd-page [data-block="7"] .wd-math').shadowRoot;
      const sin = [...func.querySelectorAll('mi')].find((m) => m.textContent === 'sin');
      const table = func.querySelector('mtable');
      return {
        stacked: Boolean(op && under.top >= op.bottom - 2 && over.bottom <= op.top + 2),
        // MathML Core sets a one-letter identifier italic and a longer one
        // upright: "sin" is ONE identifier of three letters, not three.
        sinUpright: Boolean(sin) && sin.textContent.length === 3 && !sin.hasAttribute('mathvariant'),
        rows: table ? table.querySelectorAll('mtr').length : 0,
      };
    })()`);
    check(
      'word: the sum\'s limits are stacked under and over it, "sin" is upright and the matrix has its two rows',
      sumAndFunc.stacked && sumAndFunc.sinUpright && sumAndFunc.rows === 2,
      JSON.stringify(sumAndFunc)
    );

    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'word-equations.png');

    // The caret steps over the inline equation as one character: with the
    // caret just before it, one ArrowRight lands just after it.
    const at = model().blocks[1].text.indexOf(MARK);
    const inlineRect = await js(`(() => { const r = document.querySelector('.wd-page [data-block="1"] .wd-math').getBoundingClientRect(); return { x: Math.round(r.left - 12), y: Math.round(r.top + r.height / 2) }; })()`);
    win.webContents.sendInputEvent({ type: 'mouseDown', x: inlineRect.x, y: inlineRect.y, button: 'left', clickCount: 1 });
    win.webContents.sendInputEvent({ type: 'mouseUp', x: inlineRect.x, y: inlineRect.y, button: 'left', clickCount: 1 });
    await until(() => model().selection?.focus?.block === 1, 'the caret in the paragraph', 5000).catch(() => false);
    // A click lands on a letter's near side; walk up to the equation.
    for (let i = 0; i < 4 && (model().selection?.focus?.offset ?? at) < at; i++) {
      const was = model().selection.focus.offset;
      await press(win.webContents, 'Right');
      await until(() => model().selection?.focus?.offset !== was, 'the caret to move', 3000).catch(() => false);
    }
    const before = model().selection?.focus;
    await press(win.webContents, 'Right');
    await until(() => model().selection?.focus?.offset === at + 1, 'the caret past the equation', 5000).catch(() => false);
    const after = model().selection?.focus;
    check(
      'word: one press of the right arrow steps over an inline equation — it is one character to the caret',
      before?.offset === at && after?.block === 1 && after?.offset === at + 1,
      `before ${JSON.stringify(before)} after ${JSON.stringify(after)}`
    );

    // A click on the display equation selects it whole; Backspace takes it;
    // Ctrl+Z puts it back, the same XML.
    const quadratic = model().blocks[3].runs.find((x) => x.math)?.math?.xml;
    const q = await js(`(() => { const r = document.querySelector('.wd-page [data-block="3"] .wd-math').shadowRoot.querySelector('math').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    win.webContents.sendInputEvent({ type: 'mouseDown', x: q.x, y: q.y, button: 'left', clickCount: 1 });
    win.webContents.sendInputEvent({ type: 'mouseUp', x: q.x, y: q.y, button: 'left', clickCount: 1 });
    const picked = await until(async () => {
      const s = model().selection;
      return s?.from?.block === 3 && s.from.offset === 0 && s.to.offset === 1 && (await js(`document.querySelector('.wd-page [data-block="3"] .wd-math').classList.contains('sel')`));
    }, 'the equation selected', 5000).catch(() => false);
    check('word: a click on an equation selects it whole, and the page shows it selected', picked === true, JSON.stringify(model().selection));
    await press(win.webContents, 'Backspace');
    const gone = await until(async () => model().blocks[3].text === '' && (await js(`!document.querySelector('.wd-page [data-block="3"] .wd-math')`)), 'the equation removed', 5000).catch(() => false);
    check('word: Backspace on the selected equation removes it from the page and the model', gone === true, JSON.stringify(model().blocks[3].text));
    await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
    await press(win.webContents, 'z', { modifiers: ['control'] });
    const back = await until(async () => model().blocks[3].runs.some((x) => x.math?.xml === quadratic) && (await js(`Boolean(document.querySelector('.wd-page [data-block="3"] .wd-math')?.shadowRoot?.querySelector('mfrac'))`)), 'the equation back', 5000).catch(() => false);
    check('word: Undo puts the equation back, the same OMML, drawn again', back === true);

    /* ── Item B: inserting and editing ─────────────────────────────────── */

    const clickRibbon = (tip) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(tip)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(tip)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickTab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 'tab'`);
    const setEditor = (value) => js(`(() => {
      const el = document.querySelector('.wd-eq-input');
      if (!el) return false;
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    })()`);
    const QUAD = 'x=(-b±√(b^2-4ac))/2a';
    const QUAD_EDITED = 'x=(-b±√(b^2-4ac))/2 a';

    // A fresh empty paragraph after "The end." — Enter at its end — and
    // Insert → Built-in → Quadratic Formula: a display equation, drawn.
    const endRect = await js(`(() => { const r = document.querySelector('.wd-page [data-block="8"]').getBoundingClientRect(); return { x: Math.round(r.right - 6), y: Math.round(r.top + r.height / 2) }; })()`);
    win.webContents.sendInputEvent({ type: 'mouseDown', x: endRect.x, y: endRect.y, button: 'left', clickCount: 1 });
    win.webContents.sendInputEvent({ type: 'mouseUp', x: endRect.x, y: endRect.y, button: 'left', clickCount: 1 });
    await until(() => model().selection?.focus?.block === 8, 'the caret at the end', 5000).catch(() => false);
    await press(win.webContents, 'End');
    await wait(150);
    // Enter types a paragraph break through its character, as a key does.
    await press(win.webContents, 'Return', { char: true });
    await until(() => model().blocks.length === 10 && model().selection?.focus?.block === 9, 'a new paragraph', 5000).catch(() => false);
    await clickTab('Insert');
    await wait(250);
    const built = await clickRibbon('Built-in equations');
    await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Quadratic Formula'))`), 'the gallery menu', 4000).catch(() => false);
    await js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Quadratic Formula')?.click(); return 1; })()`);
    const galleried = await until(async () => {
      const run = model().blocks[9]?.runs?.find((x) => x.math);
      return run?.math?.linear === QUAD && run.math.display === true
        && (await js(`Boolean(document.querySelector('.wd-page [data-block="9"] .wd-math')?.shadowRoot?.querySelector('math[display="block"] mfrac msqrt'))`));
    }, 'the quadratic formula from the gallery, drawn', 6000).catch(() => false);
    check(
      'word: Insert → Built-in → Quadratic Formula puts Word\'s quadratic formula on the empty line as a display equation, drawn with its fraction and radical',
      built === 'clicked' && galleried === true,
      `${built}; ${model().blocks.length} blocks, caret ${JSON.stringify(model().selection?.focus)}; ${JSON.stringify(model().blocks[9]?.runs?.map((x) => x.math?.linear ?? x.text))}`
    );

    // Double-click it: the editor opens on its linear form.
    const qRect = await js(`(() => { const r = document.querySelector('.wd-page [data-block="9"] .wd-math').shadowRoot.querySelector('math').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    win.webContents.sendInputEvent({ type: 'mouseDown', x: qRect.x, y: qRect.y, button: 'left', clickCount: 1 });
    win.webContents.sendInputEvent({ type: 'mouseUp', x: qRect.x, y: qRect.y, button: 'left', clickCount: 1 });
    win.webContents.sendInputEvent({ type: 'mouseDown', x: qRect.x, y: qRect.y, button: 'left', clickCount: 2 });
    win.webContents.sendInputEvent({ type: 'mouseUp', x: qRect.x, y: qRect.y, button: 'left', clickCount: 2 });
    const reopened = await until(async () => (await js(`document.querySelector('.wd-eq-input')?.value || null`)) === QUAD, 'the editor on the equation', 6000).catch(() => false);
    check('word: a double-click on an equation opens the editor with its linear form, as Word\'s own editor shows it', reopened === true, String(await js(`document.querySelector('.wd-eq-input')?.value ?? 'no editor'`)));

    // 2a → 2 a: a space ends an operand, so the denominator is 2 and a
    // follows the fraction. The preview follows the typing.
    await setEditor(QUAD_EDITED);
    const previewed = await until(() => js(`(() => { const f = document.querySelector('.wd-eq-built mfrac'); return Boolean(f) && f.children[1]?.textContent === '2'; })()`), 'the preview of the edit', 5000).catch(() => false);
    await js(`document.querySelector('.wd-eq-struct[data-struct="Fraction"]')?.click(), 1`);
    await wait(300);
    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'word-equation-editor.png');
    await js(`document.querySelector('.wd-eq-struct[data-struct="Fraction"]')?.click(), 1`);
    await js(`document.querySelector('.wd-eq-ok')?.click(), 1`);
    const updated = await until(async () => model().blocks[9]?.runs?.find((x) => x.math)?.math?.linear === QUAD_EDITED
      && (await js(`(() => { const f = document.querySelector('.wd-page [data-block="9"] .wd-math')?.shadowRoot?.querySelector('mfrac'); return Boolean(f) && f.children[1]?.textContent === '2' && !document.querySelector('.wd-eq-input'); })()`)), 'the equation updated on the page', 6000).catch(() => false);
    check(
      'word: the editor\'s live preview follows the typing, and Update replaces the equation — its denominator now 2, drawn on the page',
      previewed === true && updated === true,
      JSON.stringify(model().blocks[9]?.runs?.find((x) => x.math)?.math?.linear)
    );

    win.webContents.invalidate();
    await wait(700);
    await capture(win, 'word-equations-inserted.png');
    // The Symbols group sits at the Insert tab's far end: scrolled into view for a look.
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      await js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Equation')); let el = b?.parentElement; while (el && el.scrollWidth <= el.clientWidth + 1) el = el.parentElement; if (el) el.scrollLeft = el.scrollWidth; return 1; })()`);
      win.webContents.invalidate();
      await wait(700);
      await capture(win, 'word-equations-ribbon.png');
    }

    // Alt+= opens a new editor; a sum typed in it previews with its limits
    // stacked; Cancel leaves the document as it was.
    await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
    const blocksBefore = model().blocks.length;
    await press(win.webContents, '=', { modifiers: ['alt'] });
    const altOpen = await until(() => js(`Boolean(document.querySelector('.wd-eq-input'))`), 'Alt+= to open the editor', 5000).catch(() => false);
    await setEditor('\\sum_(i=1)^n i');
    const sumPreview = await until(() => js(`Boolean(document.querySelector('.wd-eq-built munderover'))`), 'the sum previewed', 5000).catch(() => false);
    await setEditor('(a+b');
    const told = await until(() => js(`(document.querySelector('.wd-eq-error')?.textContent || '').includes('not closed') && document.querySelector('.wd-eq-ok')?.disabled === true`), 'the mistake reported', 5000).catch(() => false);
    await js(`[...document.querySelectorAll('.rw-dialog .rw-btn')].find((b) => b.textContent.trim() === 'Cancel')?.click(), 1`);
    await until(() => js(`!document.querySelector('.wd-eq-input')`), 'the editor closed', 4000).catch(() => false);
    check(
      'word: Alt+= opens the equation editor, \\sum_(i=1)^n i previews as a sum with stacked limits, an unclosed bracket is reported (Insert disabled), and Cancel changes nothing',
      altOpen === true && sumPreview === true && told === true && model().blocks.length === blocksBefore,
      `${altOpen} ${sumPreview} ${told}`
    );

    // The edited equation selected with a click, taken with Backspace,
    // brought back with Ctrl+Z.
    const q2 = await js(`(() => { const r = document.querySelector('.wd-page [data-block="9"] .wd-math').shadowRoot.querySelector('math').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    win.webContents.sendInputEvent({ type: 'mouseDown', x: q2.x, y: q2.y, button: 'left', clickCount: 1 });
    win.webContents.sendInputEvent({ type: 'mouseUp', x: q2.x, y: q2.y, button: 'left', clickCount: 1 });
    await until(() => model().selection?.from?.block === 9 && model().selection.to.offset === 1, 'the new equation selected', 5000).catch(() => false);
    await press(win.webContents, 'Backspace');
    const removed2 = await until(() => model().blocks[9]?.text === '', 'the new equation removed', 5000).catch(() => false);
    await js(`document.querySelector('.wd-page')?.focus(), 'focused'`);
    await press(win.webContents, 'z', { modifiers: ['control'] });
    const restored2 = await until(async () => model().blocks[9]?.runs?.find((x) => x.math)?.math?.linear === QUAD_EDITED
      && (await js(`Boolean(document.querySelector('.wd-page [data-block="9"] .wd-math')?.shadowRoot?.querySelector('mfrac'))`)), 'the equation back', 5000).catch(() => false);
    check('word: the inserted equation goes with Backspace and comes back with Ctrl+Z', removed2 === true && restored2 === true);

    // Save: the file holds the OMML Word writes for what was typed.
    await wait(300);
    await clickTab('Home');
    await clickRibbon('Save');
    const inFile = () => {
      try {
        const x = OoxmlPackage.read(fs.readFileSync(file)).text('word/document.xml');
        return x;
      } catch {
        return '';
      }
    };
    const expected = linearToOmml(QUAD_EDITED, { display: true }).xml;
    const saved = await until(() => inFile().includes(expected), 'the equation in the saved file', 8000).catch(() => false);
    const xml = inFile();
    check(
      'word: Save writes the equation into document.xml as Word writes OMML — m:oMathPara, Cambria Math runs, the fraction and radical — beside the original equations untouched',
      saved === true && xml.includes(FIXTURE_EQUATIONS.Q) && xml.includes(FIXTURE_EQUATIONS.S) && /<w:document xmlns:m=/.test(xml),
      `${xml.length} characters; the new one ${xml.includes(expected) ? 'found' : 'missing'}`
    );

    // Paper: the PDF writer places Chromium's own picture of each equation
    // (five pictures, the new one too), not the linear-form fallback.
    const pdfPath = file.replace(/\.docx$/, '.pdf');
    try { fs.rmSync(pdfPath, { force: true }); } catch { /* none yet */ }
    await js(`window.rutbaOffice.print.pdf({ id: ${JSON.stringify(session.id)}, path: ${JSON.stringify(pdfPath)}, options: {} }).then(() => 'written', (e) => 'failed ' + e.message)`);
    await until(() => fs.existsSync(pdfPath) && fs.statSync(pdfPath).size > 1000, 'the PDF', 15000).catch(() => false);
    const pdf = fs.existsSync(pdfPath) ? fs.readFileSync(pdfPath).toString('latin1') : '';
    const pictures = (pdf.match(/\/Subtype \/Image/g) || []).length;
    check(
      'word: printing to PDF draws each equation from the MathML Chromium laid out — five pictures, no linear-form stand-in',
      pictures === 5 && !pdf.includes('sqrt'),
      `${pictures} pictures, ${pdf.length} bytes`
    );
    if (process.env.RUTBA_VERIFY_CAPTURE && pdf) fs.copyFileSync(pdfPath, `${process.env.RUTBA_VERIFY_CAPTURE}/word-equations.pdf`);

    const complaints = await errorsIn(win);
    check('word: the equation checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the equation checks ran', false, err.message);
  } finally {
    // equations.docx is this block's own fixture; put it back as it was.
    try { fs.writeFileSync(file, original); } catch { /* the temp folder goes with the run */ }
  }
}
