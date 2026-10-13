// The gallery's decks, written through the suite's own presentation engine.
//
// Each starts as File > New's deck (buildPptx), takes one of the suite's
// themes from Design, gains the layouts a real deck has in Slide Master view
// (Section Header, Title Only, Picture with Caption beside the built-in Title
// Slide and Title and Content), and is then built slide by slide with the
// same operations the Presentations window sends: new slides on a layout,
// words in their placeholders, shapes, tables, charts, pictures, slide
// numbers and speaker notes.

import zlib from 'node:zlib';
import { Deck, buildPptx } from '@rutba/presentation';
import { themeById } from '@rutba/office-formats/themes';

/* ---- pictures for the album: drawn here, not taken from anywhere -------- */

const crcTable = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();
const crc32 = (buf) => {
  let c = -1;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
};
const chunk = (type, data) => {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(data.length, 0);
  head.write(type, 4, 'latin1');
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([head.subarray(4), data])), 0);
  return Buffer.concat([head, data, crc]);
};

/** A PNG of `w` x `h` RGB pixels from `paint(x, y)` -> [r, g, b]. */
export function pngOf(w, h, paint) {
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) {
    const row = y * (w * 3 + 1);
    raw[row] = 0;
    for (let x = 0; x < w; x++) {
      const [r, g, b] = paint(x, y);
      raw[row + 1 + x * 3] = r;
      raw[row + 2 + x * 3] = g;
      raw[row + 3 + x * 3] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
}

const hexRgb = (hex) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
const mix = (a, b, t) => a.map((v, i) => Math.round(v + (b[i] - v) * t));

/**
 * A landscape: a sky from `top` to `horizon`, a sun, and two rows of hills,
 * all soft; `seed` moves the hills so no two pictures are the same.
 */
function landscape({ w = 800, h = 534, top, horizon, sun, near, far, seed = 1 }) {
  const [T, H, S, N, F] = [top, horizon, sun, near, far].map(hexRgb);
  const sunX = w * (0.25 + (seed % 5) * 0.12);
  const sunY = h * 0.38;
  const ridge = (x, k, amp, base) => base + Math.sin(x / (w / (2.2 + k)) + seed) * amp + Math.sin(x / (w / (5.1 + k)) + seed * 2.3) * amp * 0.4;
  return pngOf(w, h, (x, y) => {
    let c = mix(T, H, Math.min(1, y / (h * 0.62)));
    const d = Math.hypot(x - sunX, y - sunY);
    if (d < h * 0.09) c = S;
    else if (d < h * 0.2) c = mix(S, c, (d - h * 0.09) / (h * 0.11));
    const farY = ridge(x, 0, h * 0.05, h * 0.6);
    if (y > farY) c = mix(F, H, Math.max(0, 0.35 - (y - farY) / h));
    const nearY = ridge(x, 1.7, h * 0.07, h * 0.76);
    if (y > nearY) c = mix(N, F, Math.max(0, 0.25 - (y - nearY) / (h * 0.8)));
    return c;
  });
}

const SCENES = [
  { top: '1B2A4E', horizon: 'F6A15A', sun: 'FFE9B0', far: '6A4C7A', near: '2E2A44', seed: 1 },
  { top: '5AA9E6', horizon: 'CDEBFA', sun: 'FFFFFF', far: '6FAE8C', near: '2F6B4F', seed: 2 },
  { top: '2B1E3F', horizon: 'E06C75', sun: 'FFD38A', far: '8E4B6E', near: '3A2440', seed: 3 },
  { top: '7FB7D9', horizon: 'F2E6C9', sun: 'FFF6D6', far: 'B79B6C', near: '6B5434', seed: 4 },
  { top: '0F2A3A', horizon: '3E8A9A', sun: 'D9F3F0', far: '1E5560', near: '0E2F36', seed: 5 },
  { top: 'F7C59F', horizon: 'FCE8D5', sun: 'FFFFFF', far: 'D08C60', near: '7A4A2A', seed: 6 },
];

/* ---- building a deck ------------------------------------------------------ */

/**
 * A deck on a theme, with the layouts a real deck carries: the two
 * buildPptx makes (Title Slide, Title and Content) and Section Header,
 * Title Only and Picture with Caption added through Slide Master view's
 * own operations.
 */
function start(theme, first) {
  const deck = Deck.open(Buffer.from(buildPptx({ title: first.title, slides: [{ layout: 'title', title: first.title, body: first.subtitle }] })));
  deck.applyTheme(theme);
  const master = deck.masterParts()[0];
  // Section Header: the title large across the middle, a line of words under it.
  const section = deck.insertLayout(master, { name: 'Section Header' });
  deck.setGeometry(section, 2, { x: 96, y: 236, w: 1088, h: 170 });
  deck.setTextStyle(section, 2, { size: 48 });
  const sectionText = deck.insertPlaceholder(section, 'text', { x: 96, y: 416, w: 1088, h: 90 });
  deck.setTextStyle(section, sectionText, { size: 24 });
  const titleOnly = deck.insertLayout(master, { name: 'Title Only' });
  // Picture with Caption: the picture on the right, its title and words on the left.
  const picture = deck.insertLayout(master, { name: 'Picture with Caption' });
  deck.setGeometry(picture, 2, { x: 64, y: 96, w: 420, h: 190 });
  deck.setTextStyle(picture, 2, { size: 32 });
  deck.insertPlaceholder(picture, 'picture', { x: 520, y: 70, w: 700, h: 580 });
  const pictureText = deck.insertPlaceholder(picture, 'text', { x: 64, y: 300, w: 420, h: 330 });
  deck.setTextStyle(picture, pictureText, { size: 18 });
  const layouts = Object.fromEntries(deck.layoutList().map((l) => [l.name, l.part]));
  return { deck, layouts: { ...layouts, section, titleOnly, picture } };
}

/** Bulleted paragraphs for a placeholder. */
const bullets = (lines) => lines.map((l) => (typeof l === 'string' ? { runs: [{ text: l }] } : l));

/**
 * The builder's verbs, in the order a person works: a slide on a layout
 * with its title (and its body's lines), then shapes on it.
 */
function verbs(deck, layouts) {
  const last = () => deck.slideCount - 1;
  const titleShape = (i) => deck.slide(i).shapes.find((s) => s.placeholder && /title/i.test(s.placeholder.type || ''));
  const bodyShape = (i) => deck.slide(i).shapes.find((s) => s.placeholder && !/title/i.test(s.placeholder.type || ''));
  return {
    /** Title and Content: a title and bulleted lines. */
    content(title, lines) {
      const i = deck.insertSlide(last(), { layoutPart: layouts['Title and Content'], title, body: '' });
      deck.setText(i, bodyShape(i).id, bullets(lines));
      return i;
    },
    /** Title Only: a title, and the slide's own shapes after. */
    titled(title) {
      const i = deck.insertSlide(last(), { layoutPart: layouts.titleOnly, title, body: '' });
      const body = bodyShape(i);
      if (body) deck.removeShape(i, body.id);
      return i;
    },
    /** Section Header: a big title and a line under it. */
    section(title, line) {
      const i = deck.insertSlide(last(), { layoutPart: layouts.section, title, body: line });
      return i;
    },
    /** Title Slide again, as a closing slide. */
    closing(title, line) {
      return deck.insertSlide(last(), { layoutPart: layouts['Title Slide'], layout: 'title', title, body: line });
    },
    titleShape,
    bodyShape,
  };
}

const text = (t, o = {}) => ({ align: o.align, spaceAfter: o.after, bullet: o.bullet ?? false, runs: [{ text: t, size: o.size, bold: o.bold, color: o.color }] });

/** Slide numbers on every slide but the first, as a deck that is presented carries them. */
function numberSlides(deck) {
  for (let i = 1; i < deck.slideCount; i++) deck.setFooter(i, { slideNumber: true });
}

/* ---- the templates ------------------------------------------------------- */

function pitch() {
  const theme = 'harbour';
  const p = themeById(theme).palette;
  const { deck, layouts } = start(theme, { title: 'Company Name', subtitle: 'One line about what you do, and for whom' });
  const v = verbs(deck, layouts);
  // The title slide's own picture: three discs in the theme's accents.
  deck.addShape(0, { preset: 'ellipse', x: 870, y: 80, w: 300, h: 300, fill: { scheme: 'accent1', alpha: 0.9 }, line: 'none' });
  deck.addShape(0, { preset: 'ellipse', x: 1010, y: 300, w: 200, h: 200, fill: { scheme: 'accent2', alpha: 0.85 }, line: 'none' });
  deck.addShape(0, { preset: 'ellipse', x: 850, y: 400, w: 130, h: 130, fill: { scheme: 'accent3', alpha: 0.9 }, line: 'none' });
  v.content('The problem', [
    'Who has the problem, in a sentence',
    'How often it happens, and what it costs them',
    'Why the ways people deal with it today fall short',
  ]);
  const s = v.titled('Our solution');
  [['Fast', 'Set up in minutes, not weeks'], ['Simple', 'One screen, nothing to learn'], ['Fair', 'One price, paid monthly']].forEach(([head, line], k) => {
    deck.addShape(s, { preset: 'roundRect', x: 96 + k * 372, y: 220, w: 340, h: 300, fill: { scheme: k === 1 ? 'accent2' : 'accent1' }, line: 'none',
      text: [{ align: 'center', runs: [{ text: head, size: 32, bold: true, color: '#FFFFFF' }] }, { align: 'center', runs: [{ text: line, size: 18, color: '#FFFFFF' }] }] });
  });
  const m = v.titled('The market');
  deck.addChart(m, { type: 'column', title: 'Market size, £ million', categories: ['2024', '2025', '2026', '2027', '2028'], series: [{ name: 'Market size', values: [120, 148, 181, 220, 268] }], x: 96, y: 190, w: 700, h: 440 });
  deck.addTextBox(m, { x: 840, y: 230, w: 360, h: 360, paragraphs: [
    text('£268m', { size: 54, bold: true, color: '#' + p.accent1 }),
    text('spent each year by 2028 on the problem we solve, growing at about 22% a year.', { size: 20, after: 18 }),
    text('Source: name the study the figures come from.', { size: 14, color: '#' + p.dk2 }),
  ] });
  const t = v.titled('Traction');
  [['1,200', 'paying customers'], ['£48k', 'monthly revenue'], ['92%', 'still with us after a year']].forEach(([n, line], k) => {
    deck.addTextBox(t, { x: 96 + k * 372, y: 250, w: 340, h: 240, paragraphs: [text(n, { size: 66, bold: true, align: 'center', color: '#' + p.accent1 }), text(line, { size: 22, align: 'center' })] });
  });
  const team = v.titled('The team');
  [['AB', 'Alex Brown', 'Chief executive'], ['SK', 'Sam Khan', 'Product and design'], ['JL', 'Jo Lee', 'Sales and partners']].forEach(([initials, name, role], k) => {
    const x = 170 + k * 340;
    deck.addShape(team, { preset: 'ellipse', x, y: 200, w: 180, h: 180, fill: { scheme: ['accent1', 'accent2', 'accent3'][k] }, line: 'none', text: [{ align: 'center', runs: [{ text: initials, size: 44, bold: true, color: '#FFFFFF' }] }] });
    deck.addTextBox(team, { x: x - 60, y: 400, w: 300, h: 110, paragraphs: [text(name, { size: 22, bold: true, align: 'center' }), text(role, { size: 18, align: 'center', color: '#' + p.dk2 })] });
  });
  const ask = v.titled('The ask: £500,000');
  deck.addChart(ask, { type: 'doughnut', title: 'How the money is spent', categories: ['Product', 'Sales', 'Marketing', 'Operations'], series: [{ name: 'Share', values: [40, 30, 20, 10] }], x: 96, y: 180, w: 520, h: 460 });
  deck.addTextBox(ask, { x: 680, y: 220, w: 520, h: 400, paragraphs: [
    text('Eighteen months of runway to reach 5,000 customers.', { size: 24, bold: true, after: 18 }),
    text('Hire four engineers and two in sales', { size: 20, bullet: { type: 'char', char: '•' }, after: 8 }),
    text('Launch in two more markets', { size: 20, bullet: { type: 'char', char: '•' }, after: 8 }),
    text('Break even by the end of 2028', { size: 20, bullet: { type: 'char', char: '•' } }),
  ] });
  v.closing('Thank you', 'Your name  |  you@example.com  |  example.com');
  numberSlides(deck);
  return Buffer.from(deck.save());
}

function status() {
  const theme = 'basalt';
  const p = themeById(theme).palette;
  const { deck, layouts } = start(theme, { title: 'Project Name: status report', subtitle: '[Date]  |  [Project manager]' });
  const v = verbs(deck, layouts);
  // The title slide's own picture: a timeline, three milestones of five reached.
  deck.addShape(0, { preset: 'rect', x: 160, y: 584, w: 960, h: 4, fill: { scheme: 'accent1' }, line: 'none' });
  for (let k = 0; k < 5; k++) deck.addShape(0, { preset: 'ellipse', x: 146 + k * 240, y: 572, w: 28, h: 28, fill: k < 3 ? { scheme: 'accent1' } : { scheme: 'bg1' }, line: { color: { scheme: 'accent1' }, width: 3 } });
  const s = v.titled('Summary');
  [['Schedule', 'On track', '2E7D32'], ['Budget', 'At risk', 'F9A825'], ['Scope', 'On track', '2E7D32'], ['Quality', 'Needs attention', 'C62828']].forEach(([what, state, colour], k) => {
    deck.addShape(s, { preset: 'roundRect', x: 96 + k * 278, y: 200, w: 250, h: 180, fill: '#' + colour, line: 'none',
      text: [{ align: 'center', runs: [{ text: what, size: 20, color: '#FFFFFF' }] }, { align: 'center', runs: [{ text: state, size: 26, bold: true, color: '#FFFFFF' }] }] });
  });
  deck.addTextBox(s, { x: 96, y: 420, w: 1088, h: 200, paragraphs: [
    text('In one sentence: where the project stands, and the one thing the reader must know.', { size: 22, after: 12 }),
    text('Green is on track, amber is at risk, red needs a decision.', { size: 16, color: '#' + p.dk2 }),
  ] });
  v.content('Progress this period', [
    'Finished: the design was signed off by the steering group',
    'Finished: the supplier contract was agreed',
    'Started: building the first release',
    'Next period: test the first release with ten users',
  ]);
  const ms = v.titled('Milestones');
  const head = (t) => ({ paragraphs: [{ runs: [{ text: t, bold: true, color: '#FFFFFF' }] }], fill: '#' + p.accent1 });
  deck.addTable(ms, {
    rows: 6, cols: 3, x: 96, y: 190, w: 1088, columnWidths: [560, 264, 264],
    cells: [
      [head('Milestone'), head('Due'), head('Status')],
      ['Design signed off', '15 March', 'Done'],
      ['Contract agreed', '30 March', 'Done'],
      ['First release built', '30 April', 'In progress'],
      ['User testing', '20 May', 'Not started'],
      ['Go live', '30 June', 'Not started'],
    ],
  });
  const b = v.titled('Budget');
  deck.addChart(b, { type: 'column', title: 'Planned and spent, £ thousand', categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], series: [{ name: 'Planned', values: [40, 45, 50, 55, 50, 40] }, { name: 'Spent', values: [38, 47, 56, 61, null, null] }], x: 96, y: 180, w: 760, h: 460 });
  deck.addTextBox(b, { x: 900, y: 220, w: 300, h: 380, paragraphs: [
    text('£202k', { size: 48, bold: true, color: '#' + p.accent1 }),
    text('spent of £280k, 6% over plan to date.', { size: 20, after: 18 }),
    text('Why, and what is being done about it.', { size: 16, color: '#' + p.dk2 }),
  ] });
  const r = v.titled('Risks and issues');
  deck.addTable(r, {
    rows: 4, cols: 4, x: 96, y: 190, w: 1088, columnWidths: [400, 160, 180, 348],
    cells: [
      [head('Risk or issue'), head('Impact'), head('Owner'), head('What is being done')],
      ['Supplier may deliver late', 'High', '[Name]', 'Weekly check-in; second supplier lined up'],
      ['Two people leave in May', 'Medium', '[Name]', 'Handover notes; cover agreed'],
      ['Testing room not booked', 'Low', '[Name]', 'Booking this week'],
    ],
  });
  const next = v.content('Next steps', ['Agree the budget change at the next steering group', 'Finish the first release', 'Book the user testing sessions']);
  deck.setText(next, v.bodyShape(next).id, ['Agree the budget change at the next steering group', 'Finish the first release', 'Book the user testing sessions'].map((l) => ({ bullet: { type: 'number' }, runs: [{ text: l }] })));
  numberSlides(deck);
  return Buffer.from(deck.save());
}

function album() {
  const theme = 'nightfall';
  const { deck, layouts } = start(theme, { title: 'Photo Album', subtitle: '[The place or the occasion, and the date]' });
  const v = verbs(deck, layouts);
  const pics = SCENES.map((s) => landscape(s));
  const picture = (i, k, x, y, w, h) => deck.addPicture(i, { data: pics[k], contentType: 'image/png', name: `Photo ${k + 1}`, x, y, w, h });
  const caption = (i, x, y, w, words, align = 'center') => deck.addTextBox(i, { x, y, w, h: 50, paragraphs: [text(words, { size: 18, align })] });
  // On the title slide, the first picture behind a band of words.
  picture(0, 0, 0, 0, 1280, 720);
  deck.addShape(0, { preset: 'rect', x: 0, y: 170, w: 1280, h: 400, fill: { scheme: 'bg1', alpha: 0.65 }, line: 'none' });
  // The title slide's words come above the picture and the band.
  for (const shape of deck.slide(0).shapes.filter((x) => x.placeholder)) deck.reorderShape(0, shape.id, 'front');
  const one = v.titled('A day by the sea');
  picture(one, 1, 140, 150, 1000, 500);
  caption(one, 140, 655, 1000, 'A caption: where, when, and who is in the picture.');
  const two = v.titled('Two of a kind');
  picture(two, 2, 96, 170, 528, 352);
  picture(two, 3, 656, 170, 528, 352);
  caption(two, 96, 540, 528, 'The first picture\'s caption');
  caption(two, 656, 540, 528, 'The second picture\'s caption');
  const side = deck.insertSlide(deck.slideCount - 1, { layoutPart: layouts.picture, title: 'The view from the top', body: '' });
  const words = v.bodyShape(side);
  if (words) deck.setText(side, words.id, [text('Write a few lines about this picture: the story behind it, or what you remember of the day.', { size: 20 })]);
  picture(side, 4, 520, 70, 700, 580);
  const three = v.titled('Around town');
  [5, 0, 1].forEach((k, n) => picture(three, k, 96 + n * 368, 210, 352, 234));
  ['Morning', 'Afternoon', 'Evening'].forEach((w, n) => caption(three, 96 + n * 368, 460, 352, w));
  v.closing('Thank you for looking', '[Your name]');
  return Buffer.from(deck.save());
}

function training() {
  const theme = 'meadow';
  const p = themeById(theme).palette;
  const { deck, layouts } = start(theme, { title: 'Course Title', subtitle: 'Session 1: [the topic]  |  [Trainer\'s name]' });
  const v = verbs(deck, layouts);
  // The title slide's own picture: a ring and two shapes in the theme's accents.
  deck.addShape(0, { preset: 'donut', x: 880, y: 110, w: 290, h: 290, fill: { scheme: 'accent1' }, line: 'none' });
  deck.addShape(0, { preset: 'ellipse', x: 1060, y: 380, w: 120, h: 120, fill: { scheme: 'accent2' }, line: 'none' });
  deck.addShape(0, { preset: 'roundRect', x: 880, y: 440, w: 100, h: 100, fill: { scheme: 'accent3' }, line: 'none' });
  deck.setNotes(0, 'Welcome everyone, say who you are, and how long the session runs.');
  const goals = v.content('What you will learn', [
    'The three ideas the session is built on',
    'How to put them to use in your own work',
    'Where to find help afterwards',
  ]);
  deck.setNotes(goals, 'Ask what people hope to get from the session, and note anything to come back to.');
  const agenda = v.titled('Agenda');
  const head = (t) => ({ paragraphs: [{ runs: [{ text: t, bold: true, color: '#FFFFFF' }] }], fill: '#' + p.accent1 });
  deck.addTable(agenda, {
    rows: 6, cols: 2, x: 240, y: 180, w: 800, columnWidths: [200, 600],
    cells: [[head('Time'), head('What')], ['10:00', 'Welcome and introductions'], ['10:15', 'Part 1: the basics'], ['11:00', 'Break'], ['11:15', 'Part 2: in practice'], ['12:00', 'Quick check and summary']],
  });
  v.section('Part 1: The basics', 'The ideas everything else is built on');
  const idea = v.content('The key idea', [
    'State the idea in one sentence',
    'Give an example from the learners\' own work',
    'Say why it matters, and what goes wrong without it',
  ]);
  deck.setNotes(idea, 'Pause here for questions before moving on.');
  const two = v.titled('Two ways to do it');
  const column = (x, colour, head, lines) => {
    deck.addShape(two, { preset: 'rect', x, y: 180, w: 520, h: 60, fill: { scheme: colour }, line: 'none', text: [{ align: 'center', runs: [{ text: head, size: 24, bold: true, color: '#FFFFFF' }] }] });
    deck.addTextBox(two, { x, y: 260, w: 520, h: 360, paragraphs: lines.map((l) => text(l, { size: 20, after: 10, bullet: { type: 'char', char: '•' } })) });
  };
  column(96, 'accent1', 'The quick way', ['When it suits', 'The steps, in order', 'What to watch for']);
  column(664, 'accent2', 'The thorough way', ['When it suits', 'The steps, in order', 'What to watch for']);
  const check = v.content('Quick check', ['Question one about the key idea?', 'Question two: which way suits this case?', 'Question three: what would you do first?']);
  deck.setText(check, v.bodyShape(check).id, ['Question one about the key idea?', 'Question two: which way suits this case?', 'Question three: what would you do first?'].map((l) => ({ bullet: { type: 'number' }, runs: [{ text: l }] })));
  deck.setNotes(check, 'Give people two minutes, then go through the answers together.');
  const end = v.content('Summary', ['The key idea, in one sentence', 'The two ways, and when each suits', 'Where to find help: [link or name]']);
  deck.setNotes(end, 'Thank everyone, hand out the notes, and ask for feedback.');
  numberSlides(deck);
  return Buffer.from(deck.save());
}

/** The decks, by template id. */
export const DECKS = {
  'deck-pitch': pitch,
  'deck-status': status,
  'deck-album': album,
  'deck-training': training,
};
