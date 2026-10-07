// A PowerPoint 97-2003 presentation's model (from msppt.js), as the .pptx
// Presentations edits: blank slides at the presentation's own size, each
// with its background, and on it every shape where it stood — text boxes
// with their paragraphs and runs, shapes in their fill and outline with
// their words (a freeform with its own outline), lines, pictures — turned
// and flipped as they were, and the speaker notes.
import zlib from 'node:zlib';
import { Deck, buildPptx } from '@rutba/presentation';

const PICTURES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/x-emf', 'image/x-wmf']);

export function pptModelToDeck(model) {
  const slides = model.slides?.length ? model.slides : [{ shapes: [], notes: '' }];
  const size = model.size || { width: 960, height: 720 };
  const deck = Deck.open(buildPptx({
    title: 'Presentation',
    size: { cx: Math.round(size.width * 9525), cy: Math.round(size.height * 9525) },
    slides: slides.map(() => ({ layout: 'blank' })),
  }));
  slides.forEach((slide, i) => {
    if (slide.background && slide.background !== 'none') {
      try { deck.setBackground(i, { colour: slide.background.replace('#', '') }); } catch { /* a background this cannot write leaves the default */ }
    }
    for (const sh of slide.shapes || []) {
      try {
        const box = { x: Math.round(sh.x), y: Math.round(sh.y), w: Math.max(1, Math.round(sh.w)), h: Math.max(1, Math.round(sh.h)) };
        let id = null;
        if (sh.type === 'text') {
          id = deck.addTextBox(i, { ...box, name: sh.name || 'TextBox', paragraphs: sh.paragraphs });
          if ((sh.fill && sh.fill !== 'none') || (sh.line && sh.line !== 'none')) deck.setShapeStyle(i, id, { fill: sh.fill || 'none', line: sh.line || 'none' });
        } else if (sh.type === 'shape') {
          id = deck.addShape(i, { ...box, preset: sh.preset, name: sh.name || null, fill: sh.fill || 'none', line: sh.line || 'none', text: sh.paragraphs?.length ? sh.paragraphs : null });
          if (sh.path) deck.setShapePath(i, id, { commands: sh.path.commands, w: box.w, h: box.h, filled: sh.path.filled });
        } else if (sh.type === 'line') {
          id = deck.addShape(i, { preset: 'line', x: box.x, y: box.y, w: Math.max(0, Math.round(sh.w)), h: Math.max(0, Math.round(sh.h)) || (Math.round(sh.w) ? 0 : 1), name: sh.name || null, fill: 'none', line: sh.line === 'none' ? { color: '#000000', width: 0.75 } : sh.line });
        } else if (sh.type === 'picture') {
          const img = model.images?.[sh.image];
          if (!img || !PICTURES.has(img.contentType)) continue;
          let data = img.bytes;
          if (img.deflated) { try { data = zlib.inflateSync(Buffer.from(data)); } catch { data = Buffer.from(data); } }
          const placed = deck.addPicture(i, { ...box, name: sh.name || 'Picture', data: Buffer.from(data), contentType: img.contentType });
          id = placed?.id ?? placed;
        }
        if (id != null && (sh.rotation || sh.flipH || sh.flipV)) deck.setGeometry(i, id, { ...box, rot: sh.rotation || 0, flipH: Boolean(sh.flipH), flipV: Boolean(sh.flipV) });
      } catch {
        // A drawing this cannot express is left out; the slide keeps the rest.
      }
    }
    if (slide.notes) deck.setNotes(i, slide.notes);
  });
  return deck.save();
}
