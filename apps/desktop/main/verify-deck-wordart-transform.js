// Presentations: Shape Format → Text Effects → Transform — a text box
// selected, Arch Up pressed: its words are drawn along an arc on the stage
// (a <textPath>), the button reads pressed, and the saved deck carries
// a:prstTxWarp as PowerPoint writes it; No Transform puts them straight
// again. Run alone with RUTBA_VERIFY_ONLY=wordarttransform.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

export async function verifyDeckWordArtTransform({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'wordart-transform.pptx');
  try {
    const deck = Deck.open(buildPptx({ title: 'Art', slides: [{ layout: 'blank' }] }));
    const placed = deck.addTextBox(0, { x: 200, y: 150, w: 600, h: 300, paragraphs: [{ align: 'center', runs: [{ text: 'Rutba Office', size: 54 }] }], name: 'WordArt' });
    const id = String(placed?.id ?? placed);
    fs.writeFileSync(file, deck.save());
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const model = () => doc.model({ id: sessionFor('deck').id, slide: 0 });
    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${id}"]'))`), 'the text box', 8000);
    await js(`(() => { document.querySelector('.sl-hit[data-shape="${id}"]')?.click(); return 1; })()`);
    const tabShown = await until(() => js(`[...document.querySelectorAll('.rw-tab')].some((t) => t.textContent.trim() === 'Shape Format')`), 'the Shape Format tab', 4000).then(() => true).catch(() => false);
    await js(`(() => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Shape Format')?.click(); return 1; })()`);
    await until(() => js(`Boolean(document.querySelector('.sl-warp[data-preset="textArchUp"]'))`), 'the Transform buttons', 3000).catch(() => {});
    await js(`document.querySelector('.sl-warp[data-preset="textArchUp"]')?.click(), 1`);
    const drawn = await until(() => js(`(() => { const t = document.querySelector('.sl-stage textPath, .sl-canvas textPath, svg textPath'); return Boolean(t) && t.textContent.includes('Rutba Office'); })()`), 'the words along the arc', 5000).then(() => true).catch(() => false);
    const warp = model().slide.shapes.find((s) => String(s.id) === id)?.text?.warp?.preset ?? null;
    if (process.env.RUTBA_VERIFY_CAPTURE) {
      // A picture of each preset, for a person to look at.
      for (const preset of ['textArchUp', 'textArchDown', 'textCircle', 'textButton']) {
        await js(`document.querySelector('.sl-warp[data-preset="${preset}"]')?.click(), 1`);
        await wait(600);
        fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, `deck-wordart-${preset}.png`), (await win.webContents.capturePage()).toPNG());
      }
      await js(`document.querySelector('.sl-warp[data-preset="textArchUp"]')?.click(), 1`);
      await wait(600);
    }
    const pressed = await js(`document.querySelector('.sl-warp[data-preset="textArchUp"]')?.getAttribute('aria-pressed') === 'true' || document.querySelector('.sl-warp[data-preset="textArchUp"]')?.classList.contains('pressed') || false`);
    check('presentations: Shape Format → Transform → Arch Up lays the words along an arc on the slide, the button pressed',
      tabShown && drawn && warp === 'textArchUp' && pressed, JSON.stringify({ tabShown, drawn, warp, pressed }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => { try { return /<a:prstTxWarp prst="textArchUp">/.test(Deck.open(fs.readFileSync(file)).pkg.text('ppt/slides/slide1.xml')); } catch { return false; } }, 'the save', 6000).then(() => true).catch(() => false);
    await js(`document.querySelector('.sl-warp[data-preset="textNoShape"]')?.click(), 1`);
    await wait(400);
    const straight = model().slide.shapes.find((s) => String(s.id) === id)?.text?.warp == null && !(await js(`Boolean(document.querySelector('svg textPath'))`));
    check('presentations: the saved deck carries a:prstTxWarp as PowerPoint writes it, and No Transform puts the words straight again', saved && straight, JSON.stringify({ saved, straight }));

    const complaints = await errorsIn(win);
    check('presentations: WordArt\'s Transform reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the WordArt Transform checks ran', false, err?.message || String(err));
  }
}
