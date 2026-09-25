// Word: equations — Office Math read, drawn, edited round and saved.
//
// A document of equations as Word writes them (the quadratic formula, an
// inline πr², a sum, a function, a matrix) opens with each one drawn by the
// page's MathML: a real fraction bar, a radical, stacked limits, italic
// variables, upright function names. The caret steps over an equation as
// one character, a click selects it, Backspace takes it and Undo puts it
// back. Run alone with RUTBA_VERIFY_ONLY=equations.
import fs from 'node:fs';
import { buildDocx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml/package';

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

    // Paper: the PDF writer places Chromium's own picture of each equation
    // (four pictures), not the linear-form fallback.
    const pdfPath = file.replace(/\.docx$/, '.pdf');
    try { fs.rmSync(pdfPath, { force: true }); } catch { /* none yet */ }
    await js(`window.rutbaOffice.print.pdf({ id: ${JSON.stringify(session.id)}, path: ${JSON.stringify(pdfPath)}, options: {} }).then(() => 'written', (e) => 'failed ' + e.message)`);
    await until(() => fs.existsSync(pdfPath) && fs.statSync(pdfPath).size > 1000, 'the PDF', 15000).catch(() => false);
    const pdf = fs.existsSync(pdfPath) ? fs.readFileSync(pdfPath).toString('latin1') : '';
    const pictures = (pdf.match(/\/Subtype \/Image/g) || []).length;
    check(
      'word: printing to PDF draws each equation from the MathML Chromium laid out — four pictures, no linear-form stand-in',
      pictures === 4 && !pdf.includes('sqrt'),
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
