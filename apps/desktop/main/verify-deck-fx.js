// Presentation: the Format pane's Fill (Gradient, transparency) and Home →
// Shape Effects (Glow, Soft Edges, Reflection, beside Shadow).
//
// One rectangle, added through the ribbon, carries every effect at once by
// the end — each pane press or menu click is read straight off the engine's
// model, and each one is checked to have kept the others rather than wiping
// them, the way asking for a shadow alone must not turn a glow off. The
// stage's own SVG is polled (not just read once) for the drawing itself, and
// the saved file's raw XML is read back for the schema order. Everything
// this adds comes off again before it returns. Run alone with
// RUTBA_VERIFY_ONLY=deckfx.

import fs from 'node:fs';
import path from 'node:path';
import { Deck } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ file: string }} args the deck fixture (files.pptx)
 */
export async function verifyDeckFx(h, { file }) {
  const { open, check, until, wait, errorsIn, doc, sessionFor } = h;
  if (!file) return check('slides: the fill and effects checks have a fixture', false, 'no files.pptx');
  try {
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
    const shapeOf = (id) => model().slide.shapes.find((s) => String(s.id) === String(id));

    // Matched by its own tip when the tip starts with the button's label —
    // true for most buttons — or, when the tip reads as a sentence instead
    // (Shape Fill's own does), by the label span's exact text.
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => !n.disabled && ((n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) || (n.querySelector('span')?.textContent || '').trim() === ${JSON.stringify(title)}));
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickMenuItem = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `the "${label}" menu item`, 3000);
      return js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'missing'; b.click(); return 'clicked'; })()`);
    };
    const setRange = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return false;
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      setter.call(el, ${JSON.stringify(String(value))});
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    })()`);
    const stageHas = (re) => js(`${re.toString()}.test(document.querySelector('.sl-svg')?.innerHTML || '')`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 1`), 'the slide sorter', 8000);
    const before = model().slide.shapes.length;

    await clickRibbon('Shapes');
    await clickMenuItem('Rectangle');
    await until(() => model().slide.shapes.length === before + 1, 'the rectangle', 5000);
    const id = model().slide.shapes[model().slide.shapes.length - 1].id;
    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${id}"]'))`), 'its hit area', 4000);

    // The Format pane, opened at Fill from Shape Fill.
    const opened = await clickRibbon('Shape Fill');
    await until(() => js(`Boolean(document.querySelector('.sl-format-fill'))`), 'the Format pane', 4000);

    // Gradient: the Gradient chip, then the pane's own angle slider.
    await until(() => js(`Boolean([...document.querySelectorAll('.sl-format-fill .sl-chip')].find((b) => b.textContent.trim() === 'Gradient'))`), 'the Gradient chip', 3000);
    await js(`(() => { [...document.querySelectorAll('.sl-format-fill .sl-chip')].find((b) => b.textContent.trim() === 'Gradient')?.click(); return 1; })()`);
    const gradient = await until(() => shapeOf(id).fill?.type === 'gradient' && shapeOf(id).fill.stops?.length === 2, 'the gradient fill', 4000).catch(() => false);
    check('slides: the Fill section\'s Gradient chip writes a light-variation gradient, read back with its stops', opened === 'clicked' && gradient === true, `${opened}; fill ${JSON.stringify(shapeOf(id).fill)}`);

    await until(() => js(`Boolean(document.querySelector('.sl-format-fill input[type="range"]'))`), 'the angle slider', 3000);
    await setRange('.sl-format-fill input[type="range"]', 135);
    const angled = await until(() => Math.round(shapeOf(id).fill?.angle ?? -1) === 135, 'the angle', 4000).catch(() => false);
    check('slides: the angle slider turns the gradient', angled === true, `angle ${shapeOf(id).fill?.angle}`);

    // The stage SVG redraws a render after the engine has it.
    const drawnGradient = await until(() => stageHas(/<linearGradient/), 'the gradient drawn on the stage', 10000).catch(() => false);
    check('slides: the stage draws the gradient', drawnGradient === true, drawnGradient ? 'a linearGradient is in the SVG' : 'not found');
    await wait(300);
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-gradient.png'), (await win.webContents.capturePage()).toPNG());

    // Solid fill with transparency, from the same section.
    await js(`(() => { [...document.querySelectorAll('.sl-format-fill .sl-chip')].find((b) => b.textContent.trim() === 'Solid')?.click(); return 1; })()`);
    await until(() => shapeOf(id).fill?.type === 'solid', 'solid fill again', 4000).catch(() => {});
    await until(() => js(`Boolean(document.querySelector('.sl-format-fill .sl-format-slider input[type="range"]'))`), 'the transparency slider', 3000);
    await setRange('.sl-format-fill .sl-format-slider input[type="range"]', 60);
    const transparent = await until(() => Math.round((1 - (shapeOf(id).fill?.alpha ?? 1)) * 100) === 60, 'the transparency', 4000).catch(() => false);
    check('slides: the transparency slider writes an alpha on the solid fill', transparent === true, `fill ${JSON.stringify(shapeOf(id).fill)}`);

    // Shape Effects: Glow, then Soft Edges, then Reflection — each merged
    // in beside what the shape already carries, never wiping the others.
    await clickRibbon('Shape Effects');
    await clickMenuItem('Glow: 8 pt');
    const glowed = await until(() => Math.round(shapeOf(id).effects?.glow?.radiusPt ?? -1) === 8, 'the glow', 4000).catch(() => false);
    check('slides: Shape Effects → Glow: 8 pt writes a glow', glowed === true, JSON.stringify(shapeOf(id).effects));
    const drawnGlow = await until(() => stageHas(/feGaussianBlur/), 'the glow drawn on the stage', 10000).catch(() => false);
    check('slides: the stage draws the glow', drawnGlow === true, drawnGlow ? 'a blur filter is in the SVG' : 'not found');
    // The engine has the glow the instant the model above reads it, but the
    // *window's* own copy of the shape — what the next press merges into —
    // catches up on its own round trip, over IPC. The Format pane is open
    // and redraws from that same copy, so its own chip marking "current" is
    // waited for, not a blind pause that might land before it.
    await until(() => js(`document.querySelector('.sl-format-glow .sl-chip.current')?.textContent.trim() === '8 pt'`), "the window's own copy to catch up on the glow", 4000).catch(() => {});

    await clickRibbon('Shape Effects');
    await clickMenuItem('Soft Edges: 5 pt');
    const softened = await until(() => Math.round(shapeOf(id).effects?.softEdge?.radiusPt ?? -1) === 5, 'soft edges', 4000).catch(() => false);
    const glowKept1 = Math.round(shapeOf(id).effects?.glow?.radiusPt ?? -1) === 8;
    check('slides: Shape Effects → Soft Edges: 5 pt writes soft edges and keeps the glow', softened === true && glowKept1 === true, JSON.stringify(shapeOf(id).effects));
    await until(() => js(`document.querySelector('.sl-format-softedge .sl-chip.current')?.textContent.trim() === '5 pt'`), "the window's own copy to catch up on the soft edge", 4000).catch(() => {});

    await clickRibbon('Shape Effects');
    await clickMenuItem('Reflection: half');
    const reflected = await until(() => shapeOf(id).effects?.reflection === 'half', 'the reflection', 4000).catch(() => false);
    const keptBoth = Math.round(shapeOf(id).effects?.glow?.radiusPt ?? -1) === 8 && Math.round(shapeOf(id).effects?.softEdge?.radiusPt ?? -1) === 5;
    check('slides: Shape Effects → Reflection: half writes a reflection and keeps the glow and the soft edges', reflected === true && keptBoth === true, JSON.stringify(shapeOf(id).effects));
    await wait(300);
    if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, 'slides-effects.png'), (await win.webContents.capturePage()).toPNG());

    // Save, and read the schema order back off the file's own XML.
    await clickRibbon('Save');
    const spXmlInFile = () => {
      try {
        const d = Deck.open(fs.readFileSync(file));
        const xml = d.pkg.text(d.slideParts[0].part);
        const m = new RegExp(`<p:cNvPr id="${id}"[\\s\\S]*?</p:sp>`).exec(xml);
        return m ? m[0] : '';
      } catch { return ''; }
    };
    // By now the fill is solid again (the Solid chip replaced the gradient
    // at line 82, on the way to testing transparency) with an alpha on its
    // colour, and the effects are glow, soft edge and reflection together.
    const saved = await until(() => /<a:solidFill><a:srgbClr[^>]*><a:alpha/.test(spXmlInFile()) && /<a:glow/.test(spXmlInFile()), 'the fill and effects in the file', 8000).catch(() => false);
    const spXml = spXmlInFile();
    const order = ['a:glow', 'a:reflection', 'a:softEdge'].map((tag) => spXml.indexOf(`<${tag}`));
    check('slides: the saved file carries the transparent solid fill and the effects, the effects in schema order (glow, reflection, soft edge)',
      saved === true && order.every((n) => n >= 0) && order[0] < order[1] && order[1] < order[2],
      `order ${JSON.stringify(order)}; ${spXml.slice(spXml.indexOf('<a:effectLst'), spXml.indexOf('<a:effectLst') + 260)}`);

    // Clean up: the rectangle comes off, and the deck is saved as it was found.
    await doc.apply({ id: sessionFor('deck').id, ops: [{ op: 'removeShape', slide: 0, shape: id }] });
    await until(() => model().slide.shapes.length === before, 'the extra shape to go', 4000).catch(() => {});
    await clickRibbon('Save');
    await until(() => { try { return Deck.open(fs.readFileSync(file)).slide(0).shapes.length === before; } catch { return false; } }, 'the deck back to how it was found', 8000).catch(() => false);
    const restored = (() => { try { return Deck.open(fs.readFileSync(file)).slide(0).shapes.length; } catch { return -1; } })();
    check('slides: the fill and effects checks leave the deck as they found it', restored === before, `now ${restored} shapes, started with ${before}`);

    const complaints = await errorsIn(win);
    check('slides: the fill and effects checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('slides: the fill and effects checks ran', false, err.message);
  }
}
