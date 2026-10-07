// Presentations: Format Shape → Fill → Pattern, 3-D Format and 3-D Rotation,
// pressed in the pane — a rectangle given the Dark upward diagonal pattern
// in its own colours, drawn as its tile; a Circle bevel and 18 pt of depth;
// Isometric: Right Up, drawn turned with the depth behind; and the saved
// deck keeps them as PowerPoint writes them.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

/**
 * @param {object} h the harness: open, check, until, wait, doc, sessionFor
 */
export async function verifyDeck3d(h, { dir }) {
  const { open, check, until, wait, doc, sessionFor } = h;
  const file = path.join(dir, 'shape-3d.pptx');
  try {
    fs.writeFileSync(file, buildPptx({ title: '3-D', slides: [{ layout: 'blank' }] }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => { if (process.env.RUTBA_VERIFY_CAPTURE) fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG()); };
    const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
    const ribbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => !n.disabled && ((n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) || (n.querySelector('span')?.textContent || '').trim() === ${JSON.stringify(title)}));
      if (!b) return 'no button'; b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); b.click(); return 'clicked';
    })()`);
    const menu = async (label) => {
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}))`), label, 3000);
      return js(`(() => { [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}).click(); return 'clicked'; })()`);
    };
    const choose = (selector, value) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(selector)});
      if (!el) return false;
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    })()`);
    const stage = () => js(`document.querySelector('.sl-svg')?.innerHTML || ''`);

    await until(() => js(`document.querySelectorAll('.sl-thumb').length >= 1`), 'the slides', 8000);
    await ribbon('Shapes');
    await menu('Rectangle');
    await until(() => model().slide.shapes.length === 1, 'the rectangle', 5000);
    const id = model().slide.shapes[0].id;
    const shape = () => model().slide.shapes.find((s) => String(s.id) === String(id));
    await ribbon('Shape Fill');
    await until(() => js(`Boolean(document.querySelector('.sl-format-fill'))`), 'the Format pane', 4000);
    await js(`(() => { [...document.querySelectorAll('.sl-format-fill .sl-chip')].find((b) => b.textContent.trim() === 'Pattern')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-pattern[data-pattern="dkUpDiag"]'))`), 'the pattern gallery', 4000).catch(() => {});
    await js(`document.querySelector('.sl-pattern[data-pattern="dkUpDiag"]')?.click(), 1`);
    const patterned = await until(() => shape()?.fill?.type === 'pattern' && shape().fill.preset === 'dkUpDiag', 'the pattern', 5000).then(() => true).catch(() => false);
    const tiled = await until(async () => /<pattern id="[^"]*pat\d+"/.test(await stage()), 'the tile on the stage', 4000).then(() => true).catch(() => false);
    check('presentations: Format Shape → Fill → Pattern gives the shape a pattern from the gallery, drawn as its tile',
      patterned && tiled, JSON.stringify({ patterned, tiled, fill: shape()?.fill }));

    // 3-D Format: a Circle bevel and 18 pt of depth; 3-D Rotation: Isometric: Right Up.
    await js(`document.querySelector('.sl-format-3d [data-bevel="circle"]')?.click(), 1`);
    await until(() => shape()?.shape3d?.bevel?.prst === 'circle', 'the bevel', 5000).catch(() => {});
    await until(() => js(`document.querySelector('.sl-format-3d [data-bevel="circle"]')?.classList.contains('current')`), 'the pane to catch up', 3000).catch(() => {});
    await choose('.sl-depth', '18');
    await until(() => shape()?.shape3d?.depth === 18, 'the depth', 5000).catch(() => {});
    await choose('.sl-camera', 'isometricRightUp');
    const three = await until(() => shape()?.shape3d?.camera === 'isometricRightUp' && shape().shape3d.depth === 18 && shape().shape3d.bevel?.prst === 'circle', 'the rotation', 5000).then(() => true).catch(() => false);
    const drawn = await until(async () => { const svg = await stage(); return /feSpecularLighting/.test(svg) && /matrix\(0\.866 -0\.5 0 1 /.test(svg); }, 'the 3-D on the stage', 4000).then(() => true).catch(() => false);
    await capture('slides-3d.png');
    check('presentations: 3-D Format gives a bevel and depth, 3-D Rotation turns the shape, drawn lit and turned with its depth behind',
      three && drawn, JSON.stringify({ three, drawn, shape3d: shape()?.shape3d }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => {
      try {
        const xml = Deck.open(fs.readFileSync(file)).pkg.text('ppt/slides/slide1.xml');
        return /<a:pattFill prst="dkUpDiag">/.test(xml) && /<a:camera prst="isometricRightUp"\/>/.test(xml) && /<a:sp3d extrusionH="228600"><a:bevelT w="76200" h="76200"\/>/.test(xml);
      } catch { return false; }
    }, 'the saved 3-D', 8000).then(() => true).catch(() => false);
    check('presentations: the saved deck keeps the pattern and the 3-D as PowerPoint writes them', saved, saved ? 'pattFill, scene3d and sp3d in the file' : 'not in the file');
    await wait(50);
  } catch (err) {
    check('presentations: the pattern and 3-D checks ran', false, err.message);
  }
}
