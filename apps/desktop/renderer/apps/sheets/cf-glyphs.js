// Conditional formatting drawn in a cell: a data bar behind the number, and
// an icon set's icon at the cell's left — the sheet view says which (a bar's
// length and look, an icon's set and place in it); these are the glyphs.
// Excel's sets, each lowest first: arrows, traffic lights, signs, symbols,
// flags, ratings, quarters, red to black, and Excel 2010's stars, triangles
// and boxes. A set not known here draws nothing.

import React from 'react';

const GREEN = '#5da67e';
const YELLOW = '#ebb94f';
const RED = '#d8583e';
const BLACK = '#3b3b3b';
const GREY = '#8e8e8e';
const PINK = '#e89a8e';
const BLUE = '#4472c4';
const PALE = '#d0d0d0';
const GOLD = '#f2b800';

/**
 * A data bar as a cell's background layer: from the cell's left, as long as
 * its share of the range puts it between its shortest and longest; faded to
 * white as Excel 2007 draws one, or solid, with its border.
 */
export function barStyle(bar, width) {
  const min = bar.min ?? 0.1;
  const max = bar.max ?? 0.9;
  const length = Math.max(0, Math.min(1, min + (bar.fraction || 0) * (max - min)));
  const colour = bar.colour || '#638ec6';
  return {
    width: Math.max(0, (width - 3) * length),
    background: bar.gradient === false ? colour : `linear-gradient(to right, ${colour}, color-mix(in srgb, ${colour} 10%, white))`,
    border: bar.border ? `1px solid ${bar.border}` : undefined,
  };
}

const arrow = (turn, colour) => (
  <path d="M8 1.5 L14.5 8 L10.6 8 L10.6 14.5 L5.4 14.5 L5.4 8 L1.5 8 Z" fill={colour} transform={`rotate(${turn} 8 8)`} />
);
const TURNS = { up: 0, upRight: 45, side: 90, downRight: 135, down: 180 };
const ARROWS = { 3: ['down', 'side', 'up'], 4: ['down', 'downRight', 'upRight', 'up'], 5: ['down', 'downRight', 'side', 'upRight', 'up'] };
const ARROW_COLOURS = { 3: [RED, YELLOW, GREEN], 4: [RED, YELLOW, YELLOW, GREEN], 5: [RED, YELLOW, YELLOW, YELLOW, GREEN] };
const disc = (colour, r = 6.5) => <circle cx="8" cy="8" r={r} fill={colour} stroke="rgba(0,0,0,0.25)" strokeWidth="0.8" />;
const MARKS = {
  check: 'M4.4 8.4 L7 11 L11.8 5.4',
  cross: 'M5.2 5.2 L10.8 10.8 M10.8 5.2 L5.2 10.8',
  bang: 'M8 3.6 L8 9.4',
};
const mark = (name, colour, width) => (
  <>
    <path d={MARKS[name]} fill="none" stroke={colour} strokeWidth={width} strokeLinecap="round" strokeLinejoin="round" />
    {name === 'bang' ? <circle cx="8" cy="12" r={width / 1.6} fill={colour} /> : null}
  </>
);
const bars = (filled, of = 4) => [0, 1, 2, 3].slice(0, of).map((i) => (
  <rect key={i} x={1.5 + i * 3.5} y={13.5 - (i + 1) * 3} width="2.6" height={(i + 1) * 3} fill={i < filled ? BLUE : PALE} />
));
const STAR = 'M8 1.2 L10 6 L15 6.3 L11.1 9.5 L12.4 14.5 L8 11.7 L3.6 14.5 L4.9 9.5 L1 6.3 L6 6 Z';
const quarter = (n) => {
  if (n <= 0) return <circle cx="8" cy="8" r="6" fill="#fff" stroke={BLACK} strokeWidth="1.2" />;
  if (n >= 4) return <circle cx="8" cy="8" r="6" fill={BLACK} stroke={BLACK} strokeWidth="1.2" />;
  const a = (n / 4) * 2 * Math.PI;
  const x = 8 + 6 * Math.sin(a);
  const y = 8 - 6 * Math.cos(a);
  return (
    <>
      <circle cx="8" cy="8" r="6" fill="#fff" stroke={BLACK} strokeWidth="1.2" />
      <path d={`M8 8 L8 2 A6 6 0 ${n > 2 ? 1 : 0} 1 ${x.toFixed(2)} ${y.toFixed(2)} Z`} fill={BLACK} />
    </>
  );
};

/** One icon of a set: the glyph for its place, lowest first. */
function glyph(set, index, count) {
  const at = (list) => list[Math.max(0, Math.min(list.length - 1, index))];
  switch (set) {
    case '3Arrows': case '4Arrows': case '5Arrows':
      return arrow(TURNS[at(ARROWS[count] || ARROWS[3])], at(ARROW_COLOURS[count] || ARROW_COLOURS[3]));
    case '3ArrowsGray': case '4ArrowsGray': case '5ArrowsGray':
      return arrow(TURNS[at(ARROWS[count] || ARROWS[3])], GREY);
    case '3TrafficLights1':
      return disc(at([RED, YELLOW, GREEN]));
    case '3TrafficLights2':
      return <><rect x="1" y="1" width="14" height="14" rx="3" fill={BLACK} />{disc(at([RED, YELLOW, GREEN]), 5)}</>;
    case '4TrafficLights':
      return disc(at([BLACK, RED, YELLOW, GREEN]));
    case '4RedToBlack':
      return disc(at([BLACK, GREY, PINK, RED]));
    case '3Signs':
      return at([
        <path key="d" d="M8 1.5 L14.5 8 L8 14.5 L1.5 8 Z" fill={RED} />,
        <path key="t" d="M8 2 L14.5 13.8 L1.5 13.8 Z" fill={YELLOW} />,
        disc(GREEN),
      ]);
    case '3Symbols':
      return <>{disc(at([RED, YELLOW, GREEN]), 7)}{mark(at(['cross', 'bang', 'check']), '#fff', 1.8)}</>;
    case '3Symbols2':
      return mark(at(['cross', 'bang', 'check']), at([RED, YELLOW, GREEN]), 2.4);
    case '3Flags':
      return <><rect x="3" y="1.5" width="1.4" height="13" fill="#555" /><path d="M4.4 2 L13.5 4.8 L4.4 8 Z" fill={at([RED, YELLOW, GREEN])} /></>;
    case '4Rating':
      return <>{bars(index + 1)}</>;
    case '5Rating':
      return <>{bars(index)}</>;
    case '5Quarters':
      return quarter(index);
    case '3Stars':
      return index >= 2
        ? <path d={STAR} fill={GOLD} stroke="#b88a00" strokeWidth="0.6" />
        : index === 1
          ? <><defs><clipPath id="sh-half-star"><rect x="0" y="0" width="8" height="16" /></clipPath></defs><path d={STAR} fill="#fff" stroke="#b88a00" strokeWidth="0.6" /><path d={STAR} fill={GOLD} clipPath="url(#sh-half-star)" /></>
          : <path d={STAR} fill="#fff" stroke="#b88a00" strokeWidth="0.6" />;
    case '3Triangles':
      return at([
        <path key="d" d="M2 4.5 L14 4.5 L8 12.5 Z" fill={RED} />,
        <rect key="m" x="3" y="6.8" width="10" height="2.4" fill={YELLOW} />,
        <path key="u" d="M2 11.5 L14 11.5 L8 3.5 Z" fill={GREEN} />,
      ]);
    case '5Boxes':
      return <>{[0, 1, 2, 3].map((i) => <rect key={i} x={2 + (i % 2) * 6.5} y={2 + Math.floor(i / 2) * 6.5} width="5.5" height="5.5" fill={i < index ? BLUE : PALE} />)}</>;
    default:
      return null;
  }
}

/** An icon set's icon for a cell, at its left and as tall as the cell allows. */
export function CfIcon({ icon, height }) {
  const drawn = glyph(icon.set, icon.index, icon.count);
  if (!drawn) return null;
  const size = Math.max(8, Math.min(16, (height || 20) - 4));
  return (
    <svg className="sh-cf-icon" data-set={icon.set} data-index={icon.index} width={size} height={size} viewBox="0 0 16 16" aria-hidden="true">{drawn}</svg>
  );
}

export const CF_GLYPHS_CSS = `
/* A data bar behind the number: the cell its own layer, the bar above its fill and under its words. */
.sh-cell.barred { isolation: isolate; }
.sh-cf-bar { position: absolute; left: 1px; top: 2px; bottom: 2px; z-index: -1; box-sizing: border-box; pointer-events: none; }
/* An icon set's icon at the cell's left; the words keep clear of it. */
.sh-cf-icon { position: absolute; left: 2px; top: 50%; transform: translateY(-50%); pointer-events: none; }
.sh-cell.iconed { padding-left: 22px; }
`;
