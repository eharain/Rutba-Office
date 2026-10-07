// A PowerPoint 97-2003 presentation's model (from msppt.js), as the .pptx
// Presentations edits: blank slides at the presentation's own size, each
// with its background, and on it every shape where it stood — text boxes
// with their paragraphs and runs, shapes in their fill and outline with
// their words (a freeform with its own outline), lines, pictures and
// tables — turned and flipped as they were, each built in the show as it
// was, and the speaker notes; a slide hidden from the show stays hidden,
// and comes on with its transition.
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
      const g = slide.background.gradient;
      const hex = (c) => String(typeof c === 'object' ? c.color : c).replace('#', '');
      const spec = g ? { gradient: { from: { colour: hex(g.stops[0].color) }, to: { colour: hex(g.stops[1].color) }, angle: g.angle } } : { colour: hex(slide.background) };
      try { deck.setBackground(i, spec); } catch { /* a background this cannot write leaves the default */ }
    }
    const builds = [];
    for (const sh of slide.shapes || []) {
      try {
        const box = { x: Math.round(sh.x), y: Math.round(sh.y), w: Math.max(1, Math.round(sh.w)), h: Math.max(1, Math.round(sh.h)) };
        let id = null;
        if (sh.type === 'text') {
          id = deck.addTextBox(i, { ...box, name: sh.name || 'TextBox', paragraphs: sh.paragraphs });
          if ((sh.fill && sh.fill !== 'none') || (sh.line && sh.line !== 'none')) deck.setShapeStyle(i, id, { fill: sh.fill || 'none', line: sh.line || 'none' });
        } else if (sh.type === 'shape') {
          id = deck.addShape(i, { ...box, preset: sh.preset, name: sh.name || null, fill: sh.fill?.gradient ? 'none' : sh.fill || 'none', line: sh.line || 'none', text: sh.paragraphs?.length ? sh.paragraphs : null });
          if (sh.fill?.gradient) deck.setShapeStyle(i, id, { fill: sh.fill });
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
        } else if (sh.type === 'table') {
          // A table as a table: its own columns and rows, each cell its words, fill, borders and merging; no table style over them.
          deck.addTable(i, {
            ...box, rows: sh.rows.length, cols: sh.columns.length, columnWidths: sh.columns, rowHeights: sh.rows, styled: false,
            cells: sh.cells.map((row) => row.map((c) => ({ ...c, fill: c.fill === 'none' ? null : c.fill }))),
          });
        }
        if (id != null && sh.anchor && (sh.type === 'text' || (sh.type === 'shape' && sh.paragraphs?.length))) deck.setBodyProps(i, id, { anchor: sh.anchor });
        if (id != null && sh.shadow) deck.setShapeStyle(i, id, { effects: { shadow: sh.shadow } });
        if (id != null && (sh.rotation || sh.flipH || sh.flipV)) deck.setGeometry(i, id, { ...box, rot: sh.rotation || 0, flipH: Boolean(sh.flipH), flipV: Boolean(sh.flipV) });
        if (id != null) builds.push({ id, spid: sh.spid, animation: sh.animation });
      } catch {
        // A drawing this cannot express is left out; the slide keeps the rest.
      }
    }
    // Its effects: a later PowerPoint's, each on the shape it names, or else
    // the PowerPoint 97 builds, in the slide's own order.
    const byId = new Map(builds.filter((b) => b.spid != null).map((b) => [b.spid, b.id]));
    const effects = slide.timing?.length
      ? slide.timing.filter((e) => byId.has(e.spid)).map((e) => ({ ...e, id: byId.get(e.spid) }))
      : builds.filter((b) => b.animation).sort((x, y) => x.animation.order - y.animation.order).map((b) => ({ id: b.id, kind: 'entr', ...b.animation }));
    for (const e of effects) {
      try {
        deck.addAnimation(i, e.id, { kind: e.kind, effect: e.effect, ...(e.direction ? { direction: e.direction } : {}), ...(e.duration ? { duration: e.duration } : {}), trigger: e.trigger, delay: e.delay || 0 });
      } catch { /* an effect this cannot write is left out */ }
    }
    if (slide.notes) deck.setNotes(i, slide.notes);
    // How it comes on in the show: hidden from it, and its transition.
    if (slide.hidden) deck.setSlideHidden(i, true);
    if (slide.transition) { try { deck.setTransition(i, slide.transition); } catch { /* a transition this cannot write leaves none */ } }
  });
  return deck.save();
}
