/**
 * Design → Design Ideas, worked out on this computer: for the slide in hand,
 * a few layouts its own content suits, as PowerPoint's Designer offers them
 * — the picture beside the words on one side or the other, the picture
 * across the slide with the title on a band over it, the pictures in a row
 * under the title, the title on a band of the theme's accent, an accent bar
 * down the side, the title alone in the middle.
 *
 * An idea is a list of changes in the deck's own terms — shapes moved
 * (`moves`), shapes added behind the rest in the theme's colours (`adds`),
 * shapes' words given one colour (`colours`), shapes sent to the back
 * first (`back`) — so the same idea draws as a
 * preview (`previewIdea`, the scene changed in memory) and is applied by the
 * deck's own methods (Deck#applyDesignIdea), one undo step.
 */

const hex = (c, fallback) => '#' + String(c || fallback).replace('#', '').toUpperCase();

/** A picture's box fitted inside an area, its own proportions kept, centred. */
function fitInside(pic, area) {
  const ratio = pic.geometry.w > 0 && pic.geometry.h > 0 ? pic.geometry.w / pic.geometry.h : area.w / area.h;
  let w = area.w;
  let h = w / ratio;
  if (h > area.h) { h = area.h; w = h * ratio; }
  return { x: area.x + (area.w - w) / 2, y: area.y + (area.h - h) / 2, w, h };
}

/**
 * A picture filling an area edge to edge, its own proportions kept: the area
 * is its box, and the crop (a:srcRect, in thousandths of a percent of each
 * side) takes off what overflows, top and bottom or left and right equally.
 */
function cover(pic, area) {
  const ratio = pic.geometry.w > 0 && pic.geometry.h > 0 ? pic.geometry.w / pic.geometry.h : area.w / area.h;
  const want = area.w / area.h;
  const crop = { l: 0, t: 0, r: 0, b: 0 };
  if (want > ratio) crop.t = crop.b = Math.round(((1 - ratio / want) / 2) * 100000);
  else if (want < ratio) crop.l = crop.r = Math.round(((1 - want / ratio) / 2) * 100000);
  return { geometry: { x: Math.round(area.x), y: Math.round(area.y), w: Math.round(area.w), h: Math.round(area.h) }, crop };
}

const hasWords = (s) => (s.text?.paragraphs || []).some((p) => (p.runs || []).some((r) => String(r.text || '').trim()));
const wordCount = (s) => (s.text?.paragraphs || []).reduce((n, p) => n + (p.runs || []).reduce((m, r) => m + String(r.text || '').split(/\s+/).filter(Boolean).length, 0), 0);

/**
 * The ideas for a slide scene (Deck#slide). `colors` are the theme's, by
 * name (accent1, dk2, lt1 …), from Deck#designInfo.
 */
export function designIdeas(scene, { colors = {} } = {}) {
  const W = scene?.size?.width || 960;
  const H = scene?.size?.height || 540;
  const shapes = (scene?.shapes || []).filter((s) => !s.hidden && s.geometry && !s.groupId);
  const title = shapes.find((s) => ['title', 'ctrTitle'].includes(s.placeholder?.type)) || null;
  const pictures = shapes.filter((s) => s.kind === 'picture');
  const bodies = shapes.filter((s) => s !== title && s.kind === 'shape' && hasWords(s));
  const accent = hex(colors.accent1, '2B5FD9');
  const deep = hex(colors.dk2, '1F3864');
  const light = hex(colors.lt1, 'FFFFFF');
  const m = Math.round(W * 0.06);
  const ideas = [];
  const words = (area) => {
    // The title at the top of the area, the words under it, shared down the area's height.
    const out = [];
    if (title) out.push({ shape: title.id, geometry: { x: area.x, y: area.y, w: area.w, h: Math.round(area.h * 0.24) } });
    const top = area.y + (title ? Math.round(area.h * 0.28) : 0);
    const room = area.y + area.h - top;
    bodies.forEach((b, i) => {
      const each = room / bodies.length;
      out.push({ shape: b.id, geometry: { x: area.x, y: Math.round(top + i * each), w: area.w, h: Math.round(each - (i < bodies.length - 1 ? m * 0.3 : 0)) } });
    });
    return out;
  };

  if (pictures.length === 1 && (title || bodies.length)) {
    const pic = pictures[0];
    const half = Math.round(W * 0.46);
    ideas.push({
      id: 'picture-right',
      name: 'The picture on the right, the words beside it',
      moves: [{ shape: pic.id, ...cover(pic, { x: W - half, y: 0, w: half, h: H }) }, ...words({ x: m, y: Math.round(H * 0.1), w: W - half - 2 * m, h: Math.round(H * 0.8) })],
      adds: [],
      colours: [],
    });
    ideas.push({
      id: 'picture-left',
      name: 'The picture on the left, the words beside it',
      moves: [{ shape: pic.id, ...cover(pic, { x: 0, y: 0, w: half, h: H }) }, ...words({ x: half + m, y: Math.round(H * 0.1), w: W - half - 2 * m, h: Math.round(H * 0.8) })],
      adds: [],
      colours: [],
    });
    if (title && bodies.length <= 1 && bodies.every((b) => wordCount(b) <= 25)) {
      const band = { x: 0, y: Math.round(H * 0.66), w: W, h: H - Math.round(H * 0.66) };
      ideas.push({
        id: 'picture-band',
        name: 'The picture across the slide, the title on a band over it',
        moves: [
          { shape: pic.id, ...cover(pic, { x: 0, y: 0, w: W, h: H }) },
          { shape: title.id, geometry: { x: m, y: band.y + Math.round(band.h * 0.12), w: W - 2 * m, h: Math.round(band.h * (bodies.length ? 0.45 : 0.76)) } },
          ...bodies.map((b) => ({ shape: b.id, geometry: { x: m, y: band.y + Math.round(band.h * 0.58), w: W - 2 * m, h: Math.round(band.h * 0.34) } })),
        ],
        // The picture to the back, the band in front of it, the words in front of both.
        back: [pic.id],
        adds: [{ ...band, fill: deep, scheme: 'tx2', front: pic.id }],
        colours: [title, ...bodies].map((s) => ({ shape: s.id, colour: light })),
      });
    }
  }

  if (pictures.length >= 2 && pictures.length <= 4) {
    const top = title ? Math.round(H * 0.3) : m;
    const gap = Math.round(m / 2);
    const cell = (W - 2 * m - gap * (pictures.length - 1)) / pictures.length;
    ideas.push({
      id: 'pictures-row',
      name: 'The pictures in a row under the title',
      moves: [
        ...(title ? [{ shape: title.id, geometry: { x: m, y: Math.round(H * 0.08), w: W - 2 * m, h: Math.round(H * 0.18) } }] : []),
        ...pictures.map((p, i) => ({ shape: p.id, geometry: fitInside(p, { x: m + i * (cell + gap), y: top, w: cell, h: H - top - m }) })),
      ],
      adds: [],
      colours: [],
    });
  }

  if (title && pictures.length === 0) {
    const bandH = Math.round(H * 0.28);
    ideas.push({
      id: 'title-band',
      name: 'The title on a band of the theme\'s accent',
      moves: [
        { shape: title.id, geometry: { x: m, y: Math.round(bandH * 0.18), w: W - 2 * m, h: Math.round(bandH * 0.64) } },
        ...bodies.map((b, i) => {
          const room = H - bandH - 2 * m;
          const each = room / bodies.length;
          return { shape: b.id, geometry: { x: m, y: Math.round(bandH + m + i * each), w: W - 2 * m, h: Math.round(each - m * 0.3) } };
        }),
      ],
      adds: [{ x: 0, y: 0, w: W, h: bandH, fill: accent, scheme: 'accent1' }],
      colours: [{ shape: title.id, colour: light }],
    });
    const bar = Math.round(W * 0.03);
    ideas.push({
      id: 'side-bar',
      name: 'An accent bar down the left, the words beside it',
      moves: words({ x: bar + m, y: Math.round(H * 0.1), w: W - bar - 2 * m, h: Math.round(H * 0.8) }),
      adds: [{ x: 0, y: 0, w: bar, h: H, fill: accent, scheme: 'accent1' }],
      colours: [],
    });
    if (bodies.every((b) => wordCount(b) <= 30)) {
      ideas.push({
        id: 'centred',
        name: 'The title in the middle, the words under it',
        moves: [
          { shape: title.id, geometry: { x: m * 2, y: Math.round(H * (bodies.length ? 0.28 : 0.36)), w: W - 4 * m, h: Math.round(H * 0.22) } },
          ...bodies.map((b, i) => ({ shape: b.id, geometry: { x: m * 2, y: Math.round(H * 0.54 + i * H * 0.14), w: W - 4 * m, h: Math.round(H * 0.13) } })),
        ],
        adds: [{ x: Math.round(W / 2 - W * 0.06), y: Math.round(H * (bodies.length ? 0.51 : 0.6)), w: Math.round(W * 0.12), h: Math.max(3, Math.round(H * 0.008)), fill: accent, scheme: 'accent1' }],
        colours: [],
        centre: [title.id, ...bodies.map((b) => b.id)],
      });
    }
    ideas.push({
      id: 'deep-ground',
      name: 'Light words on the theme\'s dark colour',
      moves: words({ x: m, y: Math.round(H * 0.1), w: W - 2 * m, h: Math.round(H * 0.8) }),
      adds: [{ x: 0, y: 0, w: W, h: H, fill: deep, scheme: 'tx2' }, { x: m, y: Math.round(H * 0.1 + H * 0.8 * 0.25), w: Math.round(W * 0.08), h: Math.max(3, Math.round(H * 0.008)), fill: accent, scheme: 'accent1' }],
      colours: [title, ...bodies].map((s) => ({ shape: s.id, colour: light })),
    });
  }
  return ideas;
}

/** The scene as it would be with the idea applied: a copy, drawn by the deck's renderer as a preview. */
export function previewIdea(scene, idea) {
  const copy = JSON.parse(JSON.stringify(scene));
  for (const mv of idea.moves) {
    const s = copy.shapes.find((x) => String(x.id) === String(mv.shape));
    if (!s) continue;
    s.geometry = { ...s.geometry, ...mv.geometry };
    // The renderer fills a box with an uncropped picture as the crop would leave it.
    if (mv.crop) s.crop = { l: mv.crop.l / 100000, t: mv.crop.t / 100000, r: mv.crop.r / 100000, b: mv.crop.b / 100000 };
  }
  for (const c of idea.colours) {
    const s = copy.shapes.find((x) => String(x.id) === String(c.shape));
    if (!s?.text?.paragraphs) continue;
    for (const p of s.text.paragraphs) for (const r of p.runs || []) if (!r.break && !r.field) r.color = c.colour;
  }
  for (const id of idea.centre || []) {
    const s = copy.shapes.find((x) => String(x.id) === String(id));
    if (s?.text?.paragraphs) for (const p of s.text.paragraphs) p.align = 'center';
  }
  for (const id of [...(idea.back || [])].reverse()) {
    const at = copy.shapes.findIndex((x) => String(x.id) === String(id));
    if (at > 0) copy.shapes.unshift(...copy.shapes.splice(at, 1));
  }
  // Added shapes go behind everything, or just in front of the shape they cover.
  idea.adds.forEach((a, i) => {
    const shape = { kind: 'shape', id: `idea${i}`, name: 'Design idea', hidden: false, preset: 'rect', geometry: { x: a.x, y: a.y, w: a.w, h: a.h, rot: 0, flipH: false, flipV: false }, fill: { type: 'solid', color: a.fill.toLowerCase(), alpha: a.alpha ?? 1 }, line: null, effects: null, text: null };
    const after = a.front ? copy.shapes.findIndex((x) => String(x.id) === String(a.front)) : -1;
    copy.shapes.splice(after >= 0 ? after + 1 : i, 0, shape);
  });
  return copy;
}
