// Review → Show Changes: what is different about a deck since it was last
// open on this computer.
//
// PowerPoint's Show Changes names what co-authors changed while you were
// away; this suite has no server to ask, so it compares instead. When a deck
// is opened or saved, a fingerprint of it is kept on this computer: each
// slide by its own id (`p:sldId`, which survives moving a slide), its title,
// its words, and each shape's name, kind, box and words. Opening the deck
// again compares the file with that fingerprint: slides added, removed and
// moved, and on a slide the shapes added and removed, the words changed, the
// shapes moved or resized, and anything else about the slide (a colour, a
// transition) as "formatting changed". A file changed elsewhere — on a shared
// drive, or sent back by e-mail — shows what is new in it.

const fnv = (text) => {
  let h = 0x811c9dc5;
  const s = String(text);
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(36);
};

const wordsOf = (shape) => (shape.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text || '').join('')).join('\n').trim();
const box = (g) => (g ? [g.x, g.y, g.w, g.h, g.rot || 0].map((v) => Math.round(Number(v) || 0)).join(',') : '');

/**
 * A deck's fingerprint: `{ savedBy, saved, slides: [{ id, title, hash,
 * shapes: [{ id, name, kind, box, words }] }] }` — small enough to keep for
 * every deck opened, the words of each shape cut to a hash and a short start.
 */
export function deckFingerprint(deck) {
  const core = deck.pkg.has('docProps/core.xml') ? deck.pkg.text('docProps/core.xml') : '';
  const tag = (name) => {
    const m = new RegExp(`<${name}\\b[^>]*>([^<]*)</${name}>`).exec(core);
    return m ? m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&') : null;
  };
  const slides = deck.slideParts.map((entry, index) => {
    const scene = deck.slide(index);
    const shapes = (scene.shapes || []).map((s) => {
      const words = wordsOf(s);
      return { id: String(s.id), name: s.name || '', kind: s.kind, box: box(s.geometry), words: words ? fnv(words) : '', start: words.replace(/\s*\n\s*/g, ' / ').slice(0, 40) };
    });
    const titleShape = (scene.shapes || []).find((s) => s.placeholder?.type === 'title' || s.placeholder?.type === 'ctrTitle');
    return {
      id: String(entry.id),
      title: titleShape ? wordsOf(titleShape).split('\n')[0].slice(0, 80) : '',
      hash: fnv(deck.pkg.text(entry.part)),
      shapes,
    };
  });
  return { savedBy: tag('cp:lastModifiedBy'), saved: tag('dcterms:modified'), slides };
}

/**
 * What changed from `before` to `now` (two fingerprints): `[{ kind, id,
 * index, was, title, details }]` in the order of the deck as it is now, the
 * removed slides last. `kind` is added, removed, moved or changed; a slide
 * both moved and changed is changed, with "moved" among its details.
 */
export function compareFingerprints(before, now) {
  if (!before?.slides || !now?.slides) return [];
  const old = new Map(before.slides.map((s, i) => [s.id, { ...s, index: i }]));
  const kept = new Set();
  const out = [];
  // Moved: the fewest slides that account for the new order — those outside
  // the longest run of slides still in their old order — so one slide
  // dragged to the front is the one moved, not every slide it passed.
  const shared = now.slides.filter((s) => old.has(s.id));
  const seq = shared.map((s) => old.get(s.id).index);
  const tails = [];
  const prev = new Array(seq.length).fill(-1);
  seq.forEach((v, i) => {
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (seq[tails[mid]] < v) lo = mid + 1; else hi = mid; }
    prev[i] = lo > 0 ? tails[lo - 1] : -1;
    tails[lo] = i;
  });
  const stay = new Set();
  for (let i = tails.length ? tails[tails.length - 1] : -1; i >= 0; i = prev[i]) stay.add(shared[i].id);
  now.slides.forEach((s, index) => {
    const was = old.get(s.id);
    const title = s.title || was?.title || '';
    if (!was) {
      out.push({ kind: 'added', id: s.id, index, was: null, title, details: [] });
      return;
    }
    kept.add(s.id);
    const details = [];
    if (!stay.has(s.id)) details.push(`moved from slide ${was.index + 1}`);
    if (s.hash !== was.hash) {
      const then = new Map(was.shapes.map((x) => [x.id, x]));
      const ids = new Set(s.shapes.map((x) => x.id));
      for (const x of s.shapes) {
        const y = then.get(x.id);
        const label = x.name || `a ${x.kind}`;
        if (!y) details.push(`${label} added`);
        else {
          if (x.words !== y.words) details.push(`words of ${label} changed${x.start ? `: “${x.start}${x.start.length >= 40 ? '…' : ''}”` : ''}`);
          if (x.box !== y.box) details.push(`${label} moved or resized`);
        }
      }
      for (const y of was.shapes) if (!ids.has(y.id)) details.push(`${y.name || `a ${y.kind}`} removed`);
      if (!details.some((d) => !d.startsWith('moved from'))) details.push('formatting changed');
    }
    if (!details.length) return;
    const onlyMoved = details.length === 1 && details[0].startsWith('moved from');
    out.push({ kind: onlyMoved ? 'moved' : 'changed', id: s.id, index, was: was.index, title, details });
  });
  for (const s of before.slides) {
    if (!kept.has(s.id) && !now.slides.some((n) => n.id === s.id)) out.push({ kind: 'removed', id: s.id, index: null, was: old.get(s.id).index, title: s.title, details: [] });
  }
  return out;
}
