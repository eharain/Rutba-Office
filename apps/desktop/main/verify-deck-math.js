// Presentation: equations on a slide.
//
// On a deck of its own holding an equation as PowerPoint 365 writes one (an
// mc:AlternateContent: Office Math in a14:m, and its picture fallback) and
// words with an inline equation among them: the stage and the strip draw
// both as MathML laid out by the page, the printed page carries the same
// MathML; Insert → Equation opens Word's equation editor and writes a new
// equation with a picture Chromium drew as its fallback; a double-click
// opens it again to change it; the arrow keys move the pair as one; and
// the saved file keeps PowerPoint's equation byte for byte. Run alone with
// RUTBA_VERIFY_ONLY=deckmath.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';
import { consoleMessage } from './console-message.js';

const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const RPR = '<a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>';
const r = (t) => `<m:r>${RPR}<m:t>${t}</m:t></m:r>`;
/** The sum of the first n integers, as PowerPoint 365 writes it, with its picture fallback. */
const POWERPOINT_EQUATION = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">' +
  '<p:sp><p:nvSpPr><p:cNvPr id="7" name="TextBox 6"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="3500000" y="2300000"/><a:ext cx="5200000" cy="1500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="none" anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a14:m><m:oMathPara ${M}><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath>` +
  `<m:nary><m:naryPr><m:chr m:val="∑"/><m:ctrlPr>${RPR}</m:ctrlPr></m:naryPr><m:sub>${r('i=1')}</m:sub><m:sup>${r('n')}</m:sup><m:e>${r('i')}</m:e></m:nary>${r('=')}<m:f><m:fPr><m:ctrlPr>${RPR}</m:ctrlPr></m:fPr><m:num>${r('n')}<m:d><m:dPr><m:ctrlPr>${RPR}</m:ctrlPr></m:dPr><m:e>${r('n+1')}</m:e></m:d></m:num><m:den>${r('2')}</m:den></m:f>` +
  '</m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US" sz="3200" dirty="0"/></a:p></p:txBody></p:sp>' +
  '</mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="7" name="TextBox 6"/><p:cNvSpPr txBox="1"><a:spLocks noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="3500000" y="2300000"/><a:ext cx="5200000" cy="1500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>sum of i from 1 to n = n(n+1)/2</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>';
const INLINE = `<p:sp><p:nvSpPr><p:cNvPr id="8" name="TextBox 7"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1200000" y="4600000"/><a:ext cx="9000000" cy="700000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>A circle's area is </a:t></a:r><a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><m:oMath ${M}>${r('π')}<m:sSup><m:e>${r('r')}</m:e><m:sup>${r('2')}</m:sup></m:sSup></m:oMath></a14:m><a:r><a:rPr lang="en-US" sz="2400"/><a:t>, which grows fast.</a:t></a:r></a:p></p:txBody></p:sp>`;

export function makeMathFixture(file) {
  const pkg = OoxmlPackage.read(buildPptx({
    title: 'Equation checks',
    slides: [
      { layout: 'title', title: 'Counting', body: 'A short proof' },
      { layout: 'obj', title: 'The sum', body: [] },
    ],
  }));
  const part = 'ppt/slides/slide2.xml';
  pkg.write_(part, pkg.text(part).replace('</p:spTree>', `${POWERPOINT_EQUATION}${INLINE}</p:spTree>`));
  fs.writeFileSync(file, pkg.write());
  return file;
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args where to write the block's own deck
 */
export async function verifyDeckMath(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  try {
    const file = makeMathFixture(path.join(dir, 'maths.pptx'));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const consoleErrors = [];
    wc.on('console-message', (...args) => { const m = consoleMessage(args); if (m.level >= 3) consoleErrors.push(m.text.split('\n')[0].slice(0, 160)); });
    const id = () => sessionFor('deck').id;
    const model = (slide = 1) => doc.model({ id: id(), slide });
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
    const typeEq = (value) => js(`(() => {
      const el = document.querySelector('.wd-eq-input');
      if (!el) return 'no editor';
      el.focus();
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      return 'typed';
    })()`);
    const mathBox = (scope) => js(`(() => {
      const list = [...document.querySelectorAll(${JSON.stringify(scope)})];
      return list.map((m) => { const r = m.getBoundingClientRect(); return { w: Math.round(r.width), h: Math.round(r.height), frac: Boolean(m.querySelector('mfrac')), sum: /∑/.test(m.textContent) }; });
    })()`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 2`), 'the slide sorter', 8000);
    await js(`(() => { document.querySelectorAll('.sl-thumb')[1]?.click(); return 1; })()`);
    await until(() => js(`document.querySelectorAll('.sl-thumb')[1]?.classList.contains('active')`), 'slide 2', 4000).catch(() => {});

    // 1. Drawn: the stage and the strip, as MathML laid out by the page.
    const drawn = await until(async () => {
      const boxes = await mathBox('.sl-svg math');
      return boxes.length === 2 && boxes.some((b) => b.frac && b.sum && b.w > 150 && b.h > 40) && boxes.some((b) => !b.frac && b.w > 20);
    }, 'the equations on the stage', 8000).catch(() => false);
    const stageBoxes = await mathBox('.sl-svg math');
    const inThumb = await until(async () => (await mathBox('.sl-thumb .sl-thumb-pic math')).length === 2, 'the equations in the strip', 8000).catch(() => false);
    const words = await js(`/A circle's area is/.test(document.querySelector('.sl-svg')?.textContent || '') && /which grows fast/.test(document.querySelector('.sl-svg')?.textContent || '')`);
    check('slides: an equation PowerPoint wrote, and one inline among words, are drawn on the stage and in the strip as MathML — a sum with its limits, a fraction, πr²',
      drawn === true && inThumb === true && words, `stage ${JSON.stringify(stageBoxes)}; strip ${inThumb}; words ${words}`);
    const printed = (() => { try { return doc.printSource({ id: id(), options: {} }).html || ''; } catch (err) { return `error ${err.message}`; } })();
    check('slides: the printed deck carries the same MathML', /<math style="font-family:'Cambria Math'[^"]*" xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML" display="block">/.test(printed) && /<mfrac>/.test(printed), `${(printed.match(/<math /g) || []).length} equations in the page printed`);
    await capture('slides-equation-read.png');

    // 2. Insert → Equation: Word's editor, a new equation with a picture fallback.
    const before = model(1).slide.shapes.length;
    await tab('Insert');
    const opened = await clickRibbon('Equation');
    await until(() => js(`Boolean(document.querySelector('.wd-eq-input'))`), 'the equation editor', 5000).catch(() => {});
    const typed = await typeEq('a^2+b^2=c^2');
    const preview = await until(() => js(`Boolean(document.querySelector('.wd-eq-preview math msup'))`), 'the preview', 4000).catch(() => false);
    await capture('slides-equation-editor.png');
    const inserted = await js(`(() => { const b = document.querySelector('.wd-eq-ok'); if (!b || b.disabled) return 'no insert'; b.click(); return 'clicked'; })()`);
    const added = await until(() => model(1).slide.shapes.length === before + 1, 'the new equation', 12000).catch(() => false);
    const eq = model(1).slide.shapes[model(1).slide.shapes.length - 1];
    const linear = eq?.text?.paragraphs?.[0]?.runs?.[0]?.math?.linear;
    const onStage = await until(async () => (await mathBox('.sl-svg math')).length === 3, 'the new equation drawn', 6000).catch(() => false);
    check('slides: Insert → Equation opens Word\'s equation editor with a live preview, and Insert puts the equation on the slide, drawn at once',
      opened === 'clicked' && typed === 'typed' && preview === true && inserted === 'clicked' && added === true && linear === 'a^2+b^2=c^2' && onStage === true,
      `${opened}/${typed}/${inserted}; preview ${preview}; added ${added}; linear ${linear}; drawn ${onStage}`);
    await capture('slides-equation.png');

    // 3. A double-click opens it again; Update changes it.
    const eqId = eq?.id;
    await js(`(() => { const el = document.querySelector('.sl-hit[data-shape="${eqId}"]'); if (!el) return 'no hit'; el.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 'dbl'; })()`);
    const reopened = await until(() => js(`document.querySelector('.wd-eq-input')?.value === 'a^2+b^2=c^2'`), 'the editor with the equation', 5000).catch(() => false);
    await typeEq('E=mc^2');
    await until(() => js(`Boolean(document.querySelector('.wd-eq-preview math'))`), 'the preview', 4000).catch(() => {});
    await js(`(() => { const b = document.querySelector('.wd-eq-ok'); if (!b || b.disabled) return 'no update'; b.click(); return 'clicked'; })()`);
    const changed = await until(() => model(1).slide.shapes.find((s) => String(s.id) === String(eqId))?.text?.paragraphs?.[0]?.runs?.[0]?.math?.linear === 'E=mc^2', 'the equation changed', 12000).catch(() => false);
    check('slides: double-clicking an equation opens it in the editor in its linear form, and Update changes it in place', reopened === true && changed === true, `reopened ${reopened}; changed ${changed}`);

    // 4. The arrow keys move the pair as one; saved, both halves moved and PowerPoint's own equation is untouched.
    await js(`(() => { document.querySelector('.sl-hit[data-shape="${eqId}"]')?.click(); document.querySelector('.sl-stage')?.focus(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-hit.selected[data-shape="${eqId}"]'))`), 'the equation selected', 3000).catch(() => {});
    const x0 = model(1).slide.shapes.find((s) => String(s.id) === String(eqId))?.geometry?.x;
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Right', modifiers: ['shift'] });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Right', modifiers: ['shift'] });
    const moved = await until(() => Math.round(model(1).slide.shapes.find((s) => String(s.id) === String(eqId))?.geometry?.x - x0) === 10, 'the equation nudged', 5000).catch(() => false);
    await wait(300);
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control'] });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'S', modifiers: ['control'] });
    const read = () => { try { return OoxmlPackage.read(fs.readFileSync(file)); } catch { return null; } };
    const saved = await until(() => (read()?.text('ppt/slides/slide2.xml').match(/<mc:AlternateContent/g) || []).length === 2, 'the saved file', 8000).catch(() => false);
    const pkg = read();
    const xml = pkg?.text('ppt/slides/slide2.xml') || '';
    const ours = /<mc:AlternateContent(?:(?!<\/mc:AlternateContent>)[\s\S])*?Equation[\s\S]*?<\/mc:AlternateContent>/.exec(xml)?.[0] || '';
    const offs = [...ours.matchAll(/<a:off x="(\d+)" y="(\d+)"\/>/g)].map((m) => m[0]);
    const fallbackPart = /<mc:Fallback>[\s\S]*?r:embed="(rId\d+)"/.exec(ours)?.[1];
    const rels = pkg?.text('ppt/slides/_rels/slide2.xml.rels') || '';
    const target = fallbackPart ? new RegExp(`Id="${fallbackPart}"[^>]*Target="\\.\\./media/([^"]+)"`).exec(rels)?.[1] : null;
    const png = target && pkg.has(`ppt/media/${target}`) ? pkg.read(`ppt/media/${target}`) : null;
    check('slides: the arrow keys move the equation and its picture fallback as one, and the saved file keeps PowerPoint\'s equation byte for byte beside ours, with a PNG Chromium drew of it',
      moved === true && saved === true && xml.includes(POWERPOINT_EQUATION) && offs.length === 2 && offs[0] === offs[1] && Boolean(png) && png.length > 200 && png.slice(1, 4).toString() === 'PNG' && /<a14:m><m:oMathPara xmlns:m=/.test(ours) && !/<w:/.test(ours),
      `moved ${moved}; saved ${saved}; pair offs ${offs.join(' ')}; fallback ${target} ${png ? png.length : 0} bytes; ours ${ours.length}`);
    const deck = pkg ? Deck.open(fs.readFileSync(file)) : null;
    check('slides: the saved deck opens again with all three equations', deck?.slide(1).shapes.filter((s) => (s.text?.paragraphs || []).some((p) => p.runs.some((x) => x.math))).length === 3, String(deck?.slide(1).shapes.length));
    check('slides: the equation checks raised no errors in the window', consoleErrors.length === 0, consoleErrors.slice(0, 3).join(' | '));
  } catch (err) {
    check('slides: the equation checks ran', false, err.stack || err.message);
  }
}
